#!/usr/bin/env python3
"""
ABOUTME: Applies audit results to Word documents with track changes and comments
ABOUTME: Reads JSONL export from audit report and modifies the source document
"""

import argparse
import hashlib
import json
import sys
import difflib
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import List, Dict, Tuple, Optional, Generator

from docx import Document
from docx.opc.part import Part
from docx.opc.packuri import PackURI
from lxml import etree

# ============================================================
# Constants
# ============================================================

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"

# ============================================================
# Data Classes
# ============================================================

@dataclass
class EditItem:
    """Single edit item from JSONL export"""
    uuid: str                    # Paragraph ID (w14:paraId)
    violation_text: str          # Text to find
    violation_reason: str        # Reason for violation  
    fix_action: str              # delete | replace | manual
    revised_text: str            # Replacement text or suggestion
    category: str                # Category
    rule_id: str                 # Rule ID
    heading: str = ''            # Violation heading/title
    content: str = ''            # Original paragraph content

@dataclass
class EditResult:
    """Result of processing an edit item"""
    success: bool
    item: EditItem
    error_message: Optional[str] = None

# ============================================================
# Main Class: AuditEditApplier
# ============================================================

class AuditEditApplier:
    """
    Applies audit edit results to Word documents.
    
    Supports three fix_action types:
    - delete: Remove text with track changes
    - replace: Replace text using diff-based minimal changes
    - manual: Add Word comment on the text
    """
    
    def __init__(self, jsonl_path: str, output_path: str = None,
                 skip_hash: bool = False, verbose: bool = False,
                 author: str = 'AI-Assistant', initials: str = 'AI'):
        self.jsonl_path = Path(jsonl_path)
        self.skip_hash = skip_hash
        self.verbose = verbose
        self.author = author
        self.initials = initials if initials else author[:2] if len(author) >= 2 else author
        
        # Load JSONL
        self.meta, self.edit_items = self._load_jsonl()
        
        # Determine paths
        self.source_path = Path(self.meta['source_file'])
        self.output_path = Path(output_path) if output_path else \
            self.source_path.with_stem(self.source_path.stem + '_edited')
        
        # Document objects (lazy loaded)
        self.doc: Document = None
        self.body_elem = None
        
        # Comment management
        self.next_comment_id = 0
        self.comments: List[Dict] = []
        
        # Track change ID management
        self.next_change_id = 0
        
        # Results tracking
        self.results: List[EditResult] = []
    
    # ==================== JSONL & Hash ====================
    
    def _load_jsonl(self) -> Tuple[Dict, List[EditItem]]:
        """Load JSONL export file"""
        meta = {}
        items = []
        
        with open(self.jsonl_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                data = json.loads(line)
                if data.get('type') == 'meta':
                    meta = data
                else:
                    items.append(EditItem(
                        uuid=data.get('uuid', ''),
                        violation_text=data.get('violation_text', ''),
                        violation_reason=data.get('violation_reason', ''),
                        fix_action=data.get('fix_action', 'manual'),
                        revised_text=data.get('revised_text', ''),
                        category=data.get('category', ''),
                        rule_id=data.get('rule_id', ''),
                        heading=data.get('heading', ''),
                        content=data.get('content', '')
                    ))
        
        if not meta:
            raise ValueError("JSONL file missing meta line")
        
        return meta, items
    
    def _verify_hash(self) -> bool:
        """Verify document hash matches expected value"""
        expected_hash = self.meta.get('source_hash', '')
        if not expected_hash:
            raise ValueError(
                "JSONL file missing source_hash field.\n"
                "Cannot verify document integrity. Use --skip-hash to bypass verification."
            )
        
        sha256 = hashlib.sha256()
        with open(self.source_path, 'rb') as f:
            for chunk in iter(lambda: f.read(4096), b""):
                sha256.update(chunk)
        
        actual_hash = f"sha256:{sha256.hexdigest()}"
        return actual_hash == expected_hash
    
    # ==================== Paragraph Location ====================
    
    def _find_para_node_by_id(self, para_id: str):
        """
        Find paragraph by w14:paraId using XPath deep search.
        Handles paragraphs nested in tables.
        """
        xpath_expr = f'.//w:p[@w14:paraId="{para_id}"]'
        nodes = self.body_elem.xpath(xpath_expr)
        return nodes[0] if nodes else None
    
    def _iter_paragraphs_following(self, start_node) -> Generator:
        """
        Generator: iterate paragraphs from start_node in document order.
        Handles transitions from table cells to main body and vice versa.
        """
        all_paras = self.body_elem.xpath('.//w:p')
        
        try:
            start_index = all_paras.index(start_node)
            for p in all_paras[start_index:]:
                yield p
        except ValueError:
            return
    
    # ==================== Run Processing ====================
    
    def _collect_runs_info(self, para_elem) -> List[Dict]:
        """
        Collect information about all runs in a paragraph.
        Returns: [{'elem': run, 'text': str, 'start': int, 'end': int, 'rPr': element}, ...]
        """
        runs_info = []
        pos = 0
        
        for run in para_elem.findall('.//w:r', NS):
            text_elems = run.findall('w:t', NS)
            if not text_elems:
                continue
            
            run_text = ''.join(t.text or '' for t in text_elems)
            if not run_text:
                continue
                
            rPr = run.find('w:rPr', NS)
            
            runs_info.append({
                'elem': run,
                'text': run_text,
                'start': pos,
                'end': pos + len(run_text),
                'rPr': rPr
            })
            pos += len(run_text)
        
        return runs_info
    
    def _get_combined_text(self, runs_info: List[Dict]) -> str:
        """Get combined text from runs"""
        return ''.join(r['text'] for r in runs_info)
    
    def _find_affected_runs(self, runs_info: List[Dict], 
                           match_start: int, match_end: int) -> List[Dict]:
        """Find runs that overlap with the target text range"""
        affected = []
        for info in runs_info:
            if info['end'] > match_start and info['start'] < match_end:
                affected.append(info)
        return affected
    
    def _get_rPr_xml(self, rPr_elem) -> str:
        """Convert rPr element to XML string"""
        if rPr_elem is None:
            return ''
        return etree.tostring(rPr_elem, encoding='unicode')
    
    # ==================== Diff Calculation ====================
    
    def _calculate_diff(self, old_text: str, new_text: str) -> List[Tuple[str, str]]:
        """
        Calculate minimal diff between two texts.
        Returns: [('equal', text), ('delete', text), ('insert', text), ...]
        """
        matcher = difflib.SequenceMatcher(None, old_text, new_text)
        operations = []
        
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                operations.append(('equal', old_text[i1:i2]))
            elif tag == 'delete':
                operations.append(('delete', old_text[i1:i2]))
            elif tag == 'insert':
                operations.append(('insert', new_text[j1:j2]))
            elif tag == 'replace':
                operations.append(('delete', old_text[i1:i2]))
                operations.append(('insert', new_text[j1:j2]))
        
        return operations
    
    # ==================== ID Management ====================
    
    def _init_comment_id(self):
        """Initialize next comment ID by scanning existing comments"""
        max_id = -1
        
        # Check comments.xml via OPC API
        try:
            from docx.opc.constants import RELATIONSHIP_TYPE as RT
            comments_part = self.doc.part.part_related_by(RT.COMMENTS)
            comments_xml = etree.fromstring(comments_part.blob)
            
            for comment in comments_xml.findall(f'{{{NS["w"]}}}comment'):
                cid = comment.get(f'{{{NS["w"]}}}id')
                if cid:
                    try:
                        max_id = max(max_id, int(cid))
                    except ValueError:
                        pass
        except (KeyError, AttributeError):
            pass
        
        # Also check document.xml for comment references
        for tag in ('commentRangeStart', 'commentRangeEnd', 'commentReference'):
            for elem in self.body_elem.iter(f'{{{NS["w"]}}}{tag}'):
                cid = elem.get(f'{{{NS["w"]}}}id')
                if cid:
                    try:
                        max_id = max(max_id, int(cid))
                    except ValueError:
                        pass
        
        self.next_comment_id = max_id + 1
    
    def _init_change_id(self):
        """Initialize next track change ID by scanning existing changes"""
        max_id = -1
        
        for tag in ('ins', 'del'):
            for elem in self.body_elem.iter(f'{{{NS["w"]}}}{tag}'):
                cid = elem.get(f'{{{NS["w"]}}}id')
                if cid:
                    try:
                        max_id = max(max_id, int(cid))
                    except ValueError:
                        pass
        
        self.next_change_id = max_id + 1
    
    def _get_next_change_id(self) -> str:
        """Get next track change ID and increment counter"""
        cid = str(self.next_change_id)
        self.next_change_id += 1
        return cid
    
    # ==================== XML Helpers ====================
    
    def _escape_xml(self, text: str) -> str:
        """Escape XML special characters"""
        return (text
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;'))
    
    def _create_run(self, text: str, rPr_xml: str = '') -> etree.Element:
        """Create a new w:r element with text"""
        run_xml = f'<w:r xmlns:w="{NS["w"]}">{rPr_xml}<w:t>{self._escape_xml(text)}</w:t></w:r>'
        return etree.fromstring(run_xml)
    
    def _replace_runs(self, para_elem, affected_runs: List[Dict], 
                     new_elements: List[etree.Element]):
        """Replace affected runs with new elements in the paragraph"""
        if not affected_runs or not new_elements:
            return
        
        first_run = affected_runs[0]['elem']
        parent = first_run.getparent()
        insert_idx = list(parent).index(first_run)
        
        # Remove old runs
        for info in affected_runs:
            parent.remove(info['elem'])
        
        # Insert new elements
        for i, elem in enumerate(new_elements):
            parent.insert(insert_idx + i, elem)
    
    # ==================== Delete Operation ====================
    
    def _apply_delete(self, para_elem, violation_text: str) -> bool:
        """Apply delete operation with track changes"""
        runs_info = self._collect_runs_info(para_elem)
        combined = self._get_combined_text(runs_info)
        
        match_start = combined.find(violation_text)
        if match_start == -1:
            return False
        
        match_end = match_start + len(violation_text)
        affected = self._find_affected_runs(runs_info, match_start, match_end)
        
        if not affected:
            return False
        
        rPr_xml = self._get_rPr_xml(affected[0]['rPr'])
        timestamp = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
        change_id = self._get_next_change_id()
        
        # Calculate split points
        first_run = affected[0]
        last_run = affected[-1]
        before_text = first_run['text'][:match_start - first_run['start']]
        after_text = last_run['text'][match_end - last_run['start']:]
        
        new_elements = []
        
        # Before text (unchanged)
        if before_text:
            new_elements.append(self._create_run(before_text, rPr_xml))
        
        # Deleted text
        del_xml = f'''<w:del xmlns:w="{NS['w']}" w:id="{change_id}" w:author="{self.author}" w:date="{timestamp}">
            <w:r>{rPr_xml}<w:delText>{self._escape_xml(violation_text)}</w:delText></w:r>
        </w:del>'''
        new_elements.append(etree.fromstring(del_xml))
        
        # After text (unchanged)
        if after_text:
            new_elements.append(self._create_run(after_text, rPr_xml))
        
        self._replace_runs(para_elem, affected, new_elements)
        return True
    
    # ==================== Replace Operation ====================
    
    def _apply_replace(self, para_elem, violation_text: str, 
                      revised_text: str) -> bool:
        """Apply replace operation with diff-based track changes"""
        runs_info = self._collect_runs_info(para_elem)
        combined = self._get_combined_text(runs_info)
        
        match_start = combined.find(violation_text)
        if match_start == -1:
            return False
        
        match_end = match_start + len(violation_text)
        affected = self._find_affected_runs(runs_info, match_start, match_end)
        
        if not affected:
            return False
        
        rPr_xml = self._get_rPr_xml(affected[0]['rPr'])
        timestamp = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
        
        # Calculate split points
        first_run = affected[0]
        last_run = affected[-1]
        before_text = first_run['text'][:match_start - first_run['start']]
        after_text = last_run['text'][match_end - last_run['start']:]
        
        # Calculate diff
        diff_ops = self._calculate_diff(violation_text, revised_text)
        
        new_elements = []
        
        # Before text
        if before_text:
            new_elements.append(self._create_run(before_text, rPr_xml))
        
        # Apply diff operations
        for op, text in diff_ops:
            if op == 'equal':
                new_elements.append(self._create_run(text, rPr_xml))
            elif op == 'delete':
                change_id = self._get_next_change_id()
                del_xml = f'''<w:del xmlns:w="{NS['w']}" w:id="{change_id}" w:author="{self.author}" w:date="{timestamp}">
                    <w:r>{rPr_xml}<w:delText>{self._escape_xml(text)}</w:delText></w:r>
                </w:del>'''
                new_elements.append(etree.fromstring(del_xml))
            elif op == 'insert':
                change_id = self._get_next_change_id()
                ins_xml = f'''<w:ins xmlns:w="{NS['w']}" w:id="{change_id}" w:author="{self.author}" w:date="{timestamp}">
                    <w:r>{rPr_xml}<w:t>{self._escape_xml(text)}</w:t></w:r>
                </w:ins>'''
                new_elements.append(etree.fromstring(ins_xml))
        
        # After text
        if after_text:
            new_elements.append(self._create_run(after_text, rPr_xml))
        
        self._replace_runs(para_elem, affected, new_elements)
        return True
    
    # ==================== Manual (Comment) Operation ====================
    
    def _apply_manual(self, para_elem, violation_text: str,
                     violation_reason: str, revised_text: str) -> bool:
        """Apply manual operation by adding a Word comment"""
        runs_info = self._collect_runs_info(para_elem)
        combined = self._get_combined_text(runs_info)
        
        match_start = combined.find(violation_text)
        if match_start == -1:
            return False
        
        match_end = match_start + len(violation_text)
        affected = self._find_affected_runs(runs_info, match_start, match_end)
        
        if not affected:
            return False
        
        comment_id = self.next_comment_id
        self.next_comment_id += 1
        
        rPr_xml = self._get_rPr_xml(affected[0]['rPr'])
        
        # Calculate split points
        first_run = affected[0]
        last_run = affected[-1]
        before_text = first_run['text'][:match_start - first_run['start']]
        after_text = last_run['text'][match_end - last_run['start']:]
        
        new_elements = []
        
        # Before text
        if before_text:
            new_elements.append(self._create_run(before_text, rPr_xml))
        
        # Comment range start
        range_start = etree.fromstring(
            f'<w:commentRangeStart xmlns:w="{NS["w"]}" w:id="{comment_id}"/>'
        )
        new_elements.append(range_start)
        
        # Commented text
        new_elements.append(self._create_run(violation_text, rPr_xml))
        
        # Comment range end
        range_end = etree.fromstring(
            f'<w:commentRangeEnd xmlns:w="{NS["w"]}" w:id="{comment_id}"/>'
        )
        new_elements.append(range_end)
        
        # Comment reference
        ref_xml = f'''<w:r xmlns:w="{NS['w']}">
            <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
            <w:commentReference w:id="{comment_id}"/>
        </w:r>'''
        new_elements.append(etree.fromstring(ref_xml))
        
        # After text
        if after_text:
            new_elements.append(self._create_run(after_text, rPr_xml))
        
        self._replace_runs(para_elem, affected, new_elements)
        
        # Record comment content
        comment_text = violation_reason
        if revised_text:
            comment_text += f"\nSuggestion: {revised_text}"
        
        self.comments.append({
            'id': comment_id,
            'text': comment_text
        })
        
        return True
    
    # ==================== Comment Saving ====================
    
    def _save_comments(self):
        """Save comments to comments.xml using OPC API"""
        if not self.comments:
            return
        
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        timestamp = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
        
        # Try to get existing comments.xml
        try:
            comments_part = self.doc.part.part_related_by(RT.COMMENTS)
            comments_xml = etree.fromstring(comments_part.blob)
        except (KeyError, AttributeError):
            # Create new comments.xml
            comments_xml = etree.fromstring(
                f'<w:comments xmlns:w="{NS["w"]}" xmlns:w14="{NS["w14"]}"/>'
            )
        
        # Add comments
        for comment in self.comments:
            comment_elem = etree.SubElement(
                comments_xml, f'{{{NS["w"]}}}comment'
            )
            comment_elem.set(f'{{{NS["w"]}}}id', str(comment['id']))
            comment_elem.set(f'{{{NS["w"]}}}author', self.author)
            comment_elem.set(f'{{{NS["w"]}}}date', timestamp)
            comment_elem.set(f'{{{NS["w"]}}}initials', self.initials)

            # Add paragraph with text
            p = etree.SubElement(comment_elem, f'{{{NS["w"]}}}p')
            r = etree.SubElement(p, f'{{{NS["w"]}}}r')
            t = etree.SubElement(r, f'{{{NS["w"]}}}t')
            t.text = comment['text']
        
        # Save via OPC
        blob = etree.tostring(
            comments_xml, xml_declaration=True, encoding='UTF-8'
        )
        
        try:
            comments_part = self.doc.part.part_related_by(RT.COMMENTS)
            comments_part._blob = blob
        except (KeyError, AttributeError):
            # Create new Part
            comments_part = Part(
                PackURI('/word/comments.xml'),
                COMMENTS_CONTENT_TYPE,
                blob,
                self.doc.part.package
            )
            self.doc.part.relate_to(comments_part, RT.COMMENTS)
    
    # ==================== Main Processing ====================
    
    def _process_item(self, item: EditItem) -> EditResult:
        """Process a single edit item"""
        try:
            # 1. Find anchor paragraph by ID
            anchor_para = self._find_para_node_by_id(item.uuid)
            if anchor_para is None:
                return EditResult(False, item,
                    f"Paragraph ID {item.uuid} not found (may be in header/footer or ID changed)")
            
            # 2. Search for text from anchor paragraph
            target_para = None
            for para in self._iter_paragraphs_following(anchor_para):
                runs_info = self._collect_runs_info(para)
                combined = self._get_combined_text(runs_info)
                
                if item.violation_text in combined:
                    target_para = para
                    break
            
            if target_para is None:
                return EditResult(False, item,
                    f"Text not found after anchor: {item.violation_text[:30]}...")
            
            # 3. Apply operation based on fix_action
            if item.fix_action == 'delete':
                success = self._apply_delete(target_para, item.violation_text)
            elif item.fix_action == 'replace':
                success = self._apply_replace(
                    target_para, item.violation_text, item.revised_text
                )
            elif item.fix_action == 'manual':
                success = self._apply_manual(
                    target_para, item.violation_text,
                    item.violation_reason, item.revised_text
                )
            else:
                return EditResult(False, item, f"Unknown action type: {item.fix_action}")
            
            if success:
                return EditResult(True, item)
            else:
                return EditResult(False, item, "Operation failed")
                
        except Exception as e:
            return EditResult(False, item, str(e))
    
    def apply(self) -> List[EditResult]:
        """Execute all edit operations"""
        # 1. Verify hash
        if not self.skip_hash:
            if not self._verify_hash():
                raise ValueError(
                    f"Document hash mismatch\n"
                    f"Expected: {self.meta.get('source_hash', 'N/A')}\n"
                    f"Document may have been modified. Use --skip-hash to bypass."
                )
        
        # 2. Load document
        self.doc = Document(str(self.source_path))
        self.body_elem = self.doc._element.body
        
        # 3. Initialize IDs
        self._init_comment_id()
        self._init_change_id()
        
        # 4. Process each item
        for i, item in enumerate(self.edit_items):
            if self.verbose:
                print(f"[{i+1}/{len(self.edit_items)}] {item.fix_action}: "
                      f"{item.violation_text[:40]}...")
            
            result = self._process_item(item)
            self.results.append(result)
            
            if self.verbose:
                status = "✓" if result.success else "✗"
                print(f"  [{status}]", end="")
                if not result.success:
                    print(f" {result.error_message}")
                else:
                    print()
        
        # 5. Save comments
        self._save_comments()
        
        return self.results
    
    def save(self, dry_run: bool = False):
        """Save modified document"""
        if dry_run:
            print(f"[DRY RUN] Would save to: {self.output_path}")
            return
        
        self.doc.save(str(self.output_path))
        print(f"Saved to: {self.output_path}")
    
    def save_failed_items(self) -> Optional[Path]:
        """
        Save failed edit items to JSONL file for retry.
        
        Returns:
            Path to failed items file if any failures exist, None otherwise
        """
        failed_results = [r for r in self.results if not r.success]
        
        if not failed_results:
            return None  # No failures, no file created
        
        # Generate output path: <input>_fail.jsonl
        fail_path = self.jsonl_path.with_stem(self.jsonl_path.stem + '_fail')
        
        with open(fail_path, 'w', encoding='utf-8') as f:
            # Write enhanced meta line
            meta_line = {
                **self.meta,  # Include all original meta fields
                'type': 'meta',  # Explicitly ensure type field exists (required for retry)
                'original_export': self.jsonl_path.name,
                'generated_at': datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S%z'),
                'failed_count': len(failed_results),
                'total_count': len(self.edit_items)
            }
            json.dump(meta_line, f, ensure_ascii=False)
            f.write('\n')
            
            # Write failed items with error information (matching HTML export field order)
            for result in failed_results:
                item = result.item
                data = {
                    'category': item.category,
                    'fix_action': item.fix_action,
                    'violation_reason': item.violation_reason,
                    'violation_text': item.violation_text,
                    'revised_text': item.revised_text,
                    'rule_id': item.rule_id,
                    'uuid': item.uuid,
                    'heading': item.heading,
                    'content': item.content,
                    '_error': result.error_message  # Add error info for debugging
                }
                json.dump(data, f, ensure_ascii=False)
                f.write('\n')
        
        return fail_path

# ============================================================
# Main Function
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Apply audit results to Word document"
    )
    parser.add_argument('jsonl_file', help='Audit export file (JSONL format)')
    parser.add_argument('-o', '--output', help='Output file path')
    parser.add_argument('--author', default='AI', 
                       help='Author name for track changes/comments (default: AI)')
    parser.add_argument('--initials', 
                       help='Author initials for comments (default: first 2 chars of author)')
    parser.add_argument('--skip-hash', action='store_true', 
                       help='Skip hash verification')
    parser.add_argument('--dry-run', action='store_true', 
                       help='Validate only, do not save')
    parser.add_argument('-v', '--verbose', action='store_true', 
                       help='Verbose output')
    
    args = parser.parse_args()
    
    try:
        applier = AuditEditApplier(
            args.jsonl_file,
            output_path=args.output,
            author=args.author,
            initials=args.initials,
            skip_hash=args.skip_hash,
            verbose=args.verbose
        )
        
        print(f"Source file: {applier.source_path}")
        print(f"Output to: {applier.output_path}")
        print(f"Edit items: {len(applier.edit_items)}")
        print("-" * 50)
        
        results = applier.apply()
        
        # Statistics
        success_count = sum(1 for r in results if r.success)
        fail_count = len(results) - success_count
        
        print("-" * 50)
        print(f"Completed: {success_count} succeeded, {fail_count} failed")
        
        if fail_count > 0:
            print("\nFailed items:")
            for r in results:
                if not r.success:
                    print(f"  - [{r.item.rule_id}] {r.error_message}")
        
        if not args.dry_run:
            applier.save()
        else:
            applier.save(dry_run=True)
        
        # Save failed items for retry
        fail_file = applier.save_failed_items()
        if fail_file:
            print(f"\n{'=' * 50}")
            print(f"Failed items saved to: {fail_file}")
            print(f"  → You can modify and retry with this file")
            print(f"  → Command: python {sys.argv[0]} {fail_file} --skip-hash")
            print(f"{'=' * 50}")
            
        sys.exit(0 if fail_count == 0 else 1)
        
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == '__main__':
    main()
