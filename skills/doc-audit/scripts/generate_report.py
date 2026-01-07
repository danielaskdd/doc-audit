#!/usr/bin/env python3
"""
ABOUTME: Generates HTML audit reports from audit manifest
ABOUTME: Includes statistics, issue details, and source tracing
"""

import argparse
import html
import json
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Optional

try:
    from jinja2 import Environment
except ImportError:
    print("Error: jinja2 not installed. Run: pip install jinja2", file=sys.stderr)
    sys.exit(1)


def load_manifest(file_path: str) -> tuple:
    """
    Load audit results and metadata from manifest JSONL file.

    Args:
        file_path: Path to manifest file

    Returns:
        Tuple of (metadata dict, list of audit result dictionaries)
    """
    metadata = {}
    results = []
    with open(file_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line:
                entry = json.loads(line)
                # Check if this is metadata
                if entry.get('type') == 'meta':
                    metadata = entry
                else:
                    results.append(entry)
    return metadata, results


def load_rules(file_path: str) -> dict:
    """
    Load rules from JSON file.

    Args:
        file_path: Path to rules JSON file

    Returns:
        Dictionary mapping rule_id to rule details
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        rules_data = json.load(f)
    
    rules_dict = {}
    for rule in rules_data.get('rules', []):
        rule_id = rule.get('id')
        if rule_id:
            rules_dict[rule_id] = {
                'id': rule_id,
                'description': rule.get('description', ''),
                'severity': rule.get('severity', 'medium'),
                'category': rule.get('category', 'other')
            }
    return rules_dict


def generate_report_data(manifest: list, rules_file_dict: dict = None) -> dict:
    """
    Generate report data from manifest.

    Args:
        manifest: List of audit result entries
        rules_file_dict: Optional dictionary of rules loaded from rules file

    Returns:
        Dictionary with report data
    """
    violations = []
    category_counts = Counter()
    rules = {}

    if rules_file_dict is None:
        rules_file_dict = {}

    for entry in manifest:
        if not entry.get('is_violation', False):
            continue

        # Handle multiple violations per block
        entry_violations = entry.get('violations', [])
        if entry_violations:
            for v in entry_violations:
                category = v.get('category', 'other')
                rule_id = v.get('rule_id', '')

                violations.append({
                    'uuid': entry.get('uuid', ''),
                    'heading': entry.get('p_heading', ''),
                    'content': entry.get('p_content', ''),
                    'category': category,
                    'rule_id': rule_id,
                    'violation_text': v.get('violation_text', ''),
                    'violation_reason': v.get('violation_reason', ''),
                    'fix_action': v.get('fix_action', 'manual'),
                    'revised_text': v.get('revised_text', '')
                })

                category_counts[category] += 1

                # Collect unique rule information
                if rule_id and rule_id not in rules:
                    # Try to get rule info from rules file first
                    if rule_id in rules_file_dict:
                        rules[rule_id] = rules_file_dict[rule_id].copy()
                    else:
                        # Fallback to data from manifest
                        rules[rule_id] = {
                            'id': rule_id,
                            'category': category,
                            'severity': v.get('severity', 'medium'),
                            'description': v.get('rule_description', '')
                        }
        else:
            # Single violation (backward compatibility)
            category = entry.get('category', entry.get('issue_type', 'other'))
            rule_id = entry.get('rule_id', '')

            violations.append({
                'uuid': entry.get('uuid', ''),
                'heading': entry.get('p_heading', ''),
                'content': entry.get('p_content', ''),
                'category': category,
                'rule_id': rule_id,
                'violation_text': entry.get('violation_text', ''),
                'violation_reason': entry.get('violation_reason', ''),
                'suggestion': entry.get('suggestion', '')
            })

            category_counts[category] += 1

            # Collect unique rule information
            if rule_id and rule_id not in rules:
                # Try to get rule info from rules file first
                if rule_id in rules_file_dict:
                    rules[rule_id] = rules_file_dict[rule_id].copy()
                else:
                    # Fallback to data from manifest
                    rules[rule_id] = {
                        'id': rule_id,
                        'category': category,
                        'severity': entry.get('severity', 'medium'),
                        'description': entry.get('rule_description', '')
                    }

    return {
        'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'total_blocks': len(manifest),
        'violation_count': len(violations),
        'violations': violations,
        'category_counts': dict(category_counts),
        'max_category_count': max(category_counts.values()) if category_counts else 0,
        'rules': rules
    }


def render_report(data: dict, template_path: Optional[str] = None, trusted_html: bool = False) -> str:
    """
    Render HTML report from data.

    Args:
        data: Report data dictionary
        template_path: Path to Jinja2 template file (required)
        trusted_html: If True, disable HTML escaping (use only for trusted inputs)

    Returns:
        Rendered HTML string
    """
    if Environment is None:
        # Fallback without Jinja2
        return render_report_simple(data, trusted_html=trusted_html)

    # Load template (required)
    if not template_path:
        raise ValueError("Template path is required.")
    template_file = Path(template_path)
    if not template_file.exists():
        raise FileNotFoundError(f"Template file not found: {template_path}")
    template_str = template_file.read_text(encoding='utf-8')

    env = Environment(autoescape=not trusted_html)
    template = env.from_string(template_str)
    return template.render(**data)


def render_report_simple(data: dict, trusted_html: bool = False) -> str:
    """
    Render a simple HTML report without Jinja2.

    Args:
        data: Report data dictionary

    Returns:
        Rendered HTML string
    """
    def maybe_escape(value: str) -> str:
        if trusted_html:
            return value
        return html.escape(value, quote=True)

    violations_html = ""
    for v in data['violations']:
        heading = maybe_escape(str(v['heading']))
        category = maybe_escape(str(v['category']))
        violation_reason = maybe_escape(str(v['violation_reason']))
        content = maybe_escape(str(v['content'])[:200])
        suggestion = maybe_escape(str(v['suggestion']))
        violations_html += f"""
        <div style="border: 1px solid #ddd; padding: 15px; margin: 10px 0; border-left: 4px solid #2563eb;">
            <h4>{heading}</h4>
            <p><strong>Category:</strong> {category}</p>
            <p><strong>Reason:</strong> {violation_reason}</p>
            <p><strong>Source:</strong> {content}...</p>
            {f"<p><strong>Suggestion:</strong> {suggestion}</p>" if v['suggestion'] else ""}
        </div>
        """

    return f"""<!DOCTYPE html>
<html>
<head>
    <title>Document Audit Report</title>
    <style>
        body {{ font-family: sans-serif; padding: 20px; }}
        h1 {{ color: #333; }}
        .stat {{ display: inline-block; margin: 10px; padding: 10px 20px; background: #f5f5f5; }}
    </style>
</head>
<body>
    <h1>Document Audit Report</h1>
    <p>Generated: {data['generated_at']}</p>
    <div>
        <div class="stat">Total Blocks: {data['total_blocks']}</div>
        <div class="stat">Issues Found: {data['violation_count']}</div>
    </div>
    <h2>Issues</h2>
    {violations_html if violations_html else "<p>No issues found.</p>"}
</body>
</html>"""


def main():
    parser = argparse.ArgumentParser(
        description="Generate HTML audit report from manifest"
    )
    parser.add_argument(
        "manifest",
        type=str,
        help="Path to audit manifest JSONL file"
    )
    parser.add_argument(
        "--output", "-o",
        type=str,
        default="audit_report.html",
        help="Output HTML file path (default: audit_report.html)"
    )
    parser.add_argument(
        "--template", "-t",
        type=str,
        required=True,
        help="Path to Jinja2 HTML template (required)"
    )
    parser.add_argument(
        "--trusted-html",
        action="store_true",
        help="Render report without HTML escaping (only for trusted inputs)"
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Also output report data as JSON"
    )
    parser.add_argument(
        "--rules", "-r",
        type=str,
        help="Path to rules JSON file (optional, for loading rule descriptions)"
    )

    args = parser.parse_args()

    # Validate input
    manifest_path = Path(args.manifest)
    if not manifest_path.exists():
        print(f"Error: Manifest file not found: {args.manifest}", file=sys.stderr)
        sys.exit(1)

    template_path = Path(args.template)
    if not template_path.exists():
        print(f"Error: Template file not found: {args.template}", file=sys.stderr)
        sys.exit(1)

    # Load rules if provided
    rules_dict = {}
    if args.rules:
        rules_path = Path(args.rules)
        if rules_path.exists():
            print(f"Loading rules: {args.rules}")
            rules_dict = load_rules(args.rules)
            print(f"Loaded {len(rules_dict)} rules")
        else:
            print(f"Warning: Rules file not found: {args.rules}", file=sys.stderr)

    # Load and process manifest
    print(f"Loading manifest: {args.manifest}")
    metadata, manifest = load_manifest(args.manifest)
    print(f"Loaded {len(manifest)} entries")
    
    # Display source file info if available
    if metadata:
        print(f"Source file: {metadata.get('source_file', 'Unknown')}")
        print(f"File hash: {metadata.get('source_hash', 'Unknown')[:20]}...")

    # Generate report data
    data = generate_report_data(manifest, rules_dict)
    
    # Add metadata to report data
    if metadata:
        data['source_file'] = metadata.get('source_file', 'Unknown')
        data['source_hash'] = metadata.get('source_hash', '')
        data['parsed_at'] = metadata.get('parsed_at', '')
        data['audited_at'] = metadata.get('audited_at', '')
    else:
        data['source_file'] = 'Unknown'
        data['source_hash'] = ''
        data['parsed_at'] = ''
        data['audited_at'] = ''
    
    print(f"Found {data['violation_count']} issues")

    # Render HTML
    html = render_report(data, args.template, trusted_html=args.trusted_html)

    # Save HTML
    output_path = Path(args.output)
    output_path.write_text(html, encoding='utf-8')
    print(f"HTML report saved to: {output_path}")

    # Optionally save JSON
    if args.json:
        json_path = output_path.with_suffix('.json')
        json_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding='utf-8')
        print(f"JSON data saved to: {json_path}")

    # Summary
    print("\n--- Summary ---")
    print(f"Total blocks: {data['total_blocks']}")
    print(f"Issues found: {data['violation_count']}")
    print(f"By category: {dict(data['category_counts'])}")


if __name__ == "__main__":
    main()
