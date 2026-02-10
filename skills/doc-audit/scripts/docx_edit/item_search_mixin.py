"""
This mixin class handles item text-location logic before applying edits.
It encapsulates single-paragraph, cross-paragraph, table, and boundary fallbacks.
"""

from typing import Dict, List, Optional, Tuple, Any

from .common import (
    NS,
    EditItem,
    EditResult,
    extract_longest_segment,
    format_text_preview,
    build_numbering_variants,
    strip_numbering_by_mode,
    strip_table_row_number_only,
    strip_table_row_numbering,
    normalize_table_json,
    dedupe_search_attempts,
    extract_matching_table_row,
    DEBUG_MARKER,
)


class ItemSearchMixin:
    def _locate_item_match(
        self,
        item: EditItem,
        anchor_para,
        violation_text: str,
        revised_text: str,
    ) -> Dict[str, Any]:
        """
        Locate violation text within uuid -> uuid_end range.

        Returns:
            Dict with:
            - target_para
            - matched_runs_info
            - matched_start
            - violation_text
            - revised_text
            - numbering_stripped
            - is_cross_paragraph
            - early_result (EditResult | None)
        """
        target_para = None
        matched_runs_info = None
        matched_start = -1
        numbering_stripped = False
        is_cross_paragraph = False
        boundary_error = None

        # Used for "not found" reason generation.
        tables_in_range = []
        body_segments = []

        # Handle mixed body/table content (violation_text contains <table> tags)
        has_table_tag = '<table>' in violation_text or '</table>' in violation_text
        if has_table_tag:
            if item.fix_action in ('delete', 'replace'):
                reason = "Mixed body/table content is invalid"
                self._apply_fallback_comment(anchor_para, item, reason)
                if self.verbose:
                    print(f"  [Mixed content] {reason}")
                return {
                    'target_para': target_para,
                    'matched_runs_info': matched_runs_info,
                    'matched_start': matched_start,
                    'violation_text': violation_text,
                    'revised_text': revised_text,
                    'numbering_stripped': numbering_stripped,
                    'is_cross_paragraph': is_cross_paragraph,
                    'early_result': EditResult(
                        success=True,
                        item=item,
                        error_message=reason,
                        warning=True,
                    ),
                }
            else:  # manual
                longest = extract_longest_segment(violation_text)
                if longest:
                    if self.verbose:
                        print(
                            f"  [Mixed content] Extracted longest segment for comment: "
                            f"'{format_text_preview(longest)}'"
                        )
                    violation_text = longest

        numbering_variants = build_numbering_variants(violation_text)

        # Try original text first (using revision-free view) - single paragraph
        for para in self._iter_paragraphs_in_range(anchor_para, item.uuid_end):
            runs_info_orig, _ = self._collect_runs_info_original(para)

            pos, matched_override = self._find_in_runs_with_normalization(
                runs_info_orig, violation_text
            )
            if pos != -1:
                target_para = para
                matched_runs_info = runs_info_orig
                matched_start = pos
                if matched_override is not None:
                    violation_text = matched_override
                break

        # Fallback 1: Try stripping auto-numbering if original match failed
        if target_para is None:
            for stripped_violation, strip_mode in numbering_variants:
                for para in self._iter_paragraphs_in_range(anchor_para, item.uuid_end):
                    runs_info_orig, _ = self._collect_runs_info_original(para)

                    pos, matched_override = self._find_in_runs_with_normalization(
                        runs_info_orig, stripped_violation
                    )
                    if pos != -1:
                        target_para = para
                        matched_runs_info = runs_info_orig
                        matched_start = pos
                        numbering_stripped = True
                        violation_text = matched_override or stripped_violation

                        # Handle revised_text for replace operation
                        if item.fix_action == 'replace':
                            stripped_revised, revised_has_numbering = strip_numbering_by_mode(
                                revised_text,
                                strip_mode,
                            )
                            if revised_has_numbering:
                                revised_text = stripped_revised
                        break
                if target_para is not None:
                    break

        # Fallback 2: Try cross-paragraph search (within uuid → uuid_end range)
        if target_para is None:
            cross_runs, cross_text, is_multi_para, boundary_error = self._collect_runs_info_across_paragraphs(
                anchor_para, item.uuid_end
            )

            if boundary_error:
                if self.verbose:
                    print(f"  [Boundary] {boundary_error}")
            elif is_multi_para:
                search_attempts: List[Tuple[str, Optional[str]]] = [(violation_text, None)]
                for stripped_violation, strip_mode in numbering_variants:
                    search_attempts.append((stripped_violation, strip_mode))
                if '\\n' in violation_text:
                    search_attempts.append((violation_text.replace('\\n', '\n'), None))
                if self._is_table_mode(cross_runs) and violation_text.startswith('["'):
                    stripped_row_only, was_row_stripped = strip_table_row_number_only(violation_text)
                    if was_row_stripped:
                        search_attempts.append((stripped_row_only, "table_row_number"))
                    stripped_table_text, was_table_stripped = strip_table_row_numbering(violation_text)
                    if was_table_stripped and stripped_table_text != stripped_row_only:
                        search_attempts.append((stripped_table_text, "table_row"))

                for search_text, strip_mode in search_attempts:
                    if self._is_table_mode(cross_runs):
                        pos = cross_text.find(search_text)
                        matched_override = None
                        if pos == -1 and search_text:
                            normalized_table_text, norm_to_orig = self._normalize_table_text_for_search(cross_runs)
                            pos_norm = normalized_table_text.find(search_text)
                            if pos_norm != -1:
                                norm_end = pos_norm + len(search_text) - 1
                                if 0 <= pos_norm < len(norm_to_orig) and 0 <= norm_end < len(norm_to_orig):
                                    orig_start = norm_to_orig[pos_norm]
                                    orig_end = norm_to_orig[norm_end] + 1
                                    pos = orig_start
                                    matched_override = cross_text[orig_start:orig_end]
                    else:
                        pos, matched_override = self._find_in_runs_with_normalization(
                            cross_runs, search_text
                        )

                    if pos != -1:
                        target_para = anchor_para
                        matched_runs_info = cross_runs
                        matched_start = pos
                        is_cross_paragraph = True
                        violation_text = matched_override if matched_override is not None else search_text

                        if strip_mode:
                            numbering_stripped = True
                            if item.fix_action == 'replace':
                                if strip_mode == "table_row_number":
                                    stripped_revised, revised_has_numbering = strip_table_row_number_only(revised_text)
                                elif strip_mode == "table_row":
                                    stripped_revised, revised_has_numbering = strip_table_row_numbering(revised_text)
                                else:
                                    stripped_revised, revised_has_numbering = strip_numbering_by_mode(
                                        revised_text, strip_mode
                                    )
                                if revised_has_numbering:
                                    revised_text = stripped_revised

                        if self.verbose:
                            print("  [Success] Found in cross-paragraph mode")
                        break

        # Fallback 3: Try table search if violation_text looks like JSON array
        if target_para is None and (violation_text.startswith('["') or violation_text.startswith('[["')):
            violation_text = normalize_table_json(violation_text)
            if item.fix_action == 'replace':
                revised_text = normalize_table_json(revised_text)

            tables_in_range = self._find_tables_in_range(anchor_para, item.uuid_end)

            search_attempts = [(violation_text, None)]

            stripped_row_only, was_row_stripped = strip_table_row_number_only(violation_text)
            if was_row_stripped:
                search_attempts.append((stripped_row_only, "table_row_number"))

            stripped_table_text, was_table_stripped = strip_table_row_numbering(violation_text)
            if was_table_stripped and stripped_table_text != stripped_row_only:
                search_attempts.append((stripped_table_text, "table_row"))

            for search_text, strip_mode in search_attempts:
                for table_elem in tables_in_range:
                    table_paras = list(table_elem.iter(f'{{{NS["w"]}}}p'))
                    if not table_paras:
                        continue

                    first_table_para = table_paras[0]
                    last_table_para = table_paras[-1]

                    first_para_id = first_table_para.get(
                        '{http://schemas.microsoft.com/office/word/2010/wordml}paraId'
                    )
                    last_para_id = last_table_para.get(
                        '{http://schemas.microsoft.com/office/word/2010/wordml}paraId'
                    )
                    if not first_para_id or not last_para_id:
                        continue

                    try:
                        table_runs, table_text, _, _ = self._collect_runs_info_in_table(
                            first_table_para, last_para_id, table_elem
                        )

                        pos = table_text.find(search_text)
                        matched_text_override = None
                        if pos == -1 and search_text:
                            normalized_table_text, norm_to_orig = self._normalize_table_text_for_search(table_runs)
                            pos_norm = normalized_table_text.find(search_text)
                            if pos_norm != -1:
                                norm_end = pos_norm + len(search_text) - 1
                                if 0 <= pos_norm < len(norm_to_orig) and 0 <= norm_end < len(norm_to_orig):
                                    orig_start = norm_to_orig[pos_norm]
                                    orig_end = norm_to_orig[norm_end] + 1
                                    pos = orig_start
                                    matched_text_override = table_text[orig_start:orig_end]

                        if pos == -1 and DEBUG_MARKER and (
                            DEBUG_MARKER in table_text or DEBUG_MARKER in search_text
                        ):
                            print(
                                f"\n  [DEBUG] Table matching failed for row containing "
                                f"'{DEBUG_MARKER}':"
                            )

                            table_row = extract_matching_table_row(table_text, DEBUG_MARKER)
                            search_row = extract_matching_table_row(search_text, DEBUG_MARKER)

                            if table_row:
                                print("  [DEBUG] Table row content:")
                                print(f"    {table_row}")
                            else:
                                print(
                                    f"  [DEBUG] Table content (marker '{DEBUG_MARKER}' "
                                    f"not found in individual row):"
                                )
                                print(f"    {table_text[:300]}...")

                            if search_row:
                                print("  [DEBUG] Searching for:")
                                print(f"    {search_row}")
                            else:
                                print(
                                    f"  [DEBUG] Search content (marker '{DEBUG_MARKER}' "
                                    f"not found in individual row):"
                                )
                                print(f"    {search_text[:300]}...")

                            if table_row and search_row:
                                min_len = min(len(table_row), len(search_row))
                                for i in range(min_len):
                                    if table_row[i] != search_row[i]:
                                        print(f"  [DEBUG] First difference at position {i}:")
                                        print(
                                            f"    Table: ..."
                                            f"{repr(table_row[max(0, i - 10):i + 30])}..."
                                        )
                                        print(
                                            f"    Search: ..."
                                            f"{repr(search_row[max(0, i - 10):i + 30])}..."
                                        )
                                        break
                                else:
                                    if len(table_row) != len(search_row):
                                        print(
                                            f"  [DEBUG] Length mismatch: "
                                            f"table={len(table_row)}, search={len(search_row)}"
                                        )

                        if pos != -1:
                            target_para = first_table_para
                            matched_runs_info = table_runs
                            matched_start = pos
                            is_cross_paragraph = True

                            violation_text = matched_text_override or search_text

                            if strip_mode and item.fix_action == 'replace':
                                if strip_mode == "table_row_number":
                                    stripped_revised, revised_was_stripped = strip_table_row_number_only(revised_text)
                                else:
                                    stripped_revised, revised_was_stripped = strip_table_row_numbering(revised_text)
                                if revised_was_stripped:
                                    revised_text = stripped_revised

                            if self.verbose:
                                if strip_mode:
                                    print("  [Success] Found in table after stripping row numbering")
                                else:
                                    print("  [Success] Found in table (JSON format)")
                            break

                    except (ValueError, KeyError, IndexError, AttributeError) as e:
                        if self.verbose:
                            print(f"  [Warning] Skipping table: {e}")
                        continue

                if target_para is not None:
                    break

        # Fallback 2.5: Try non-JSON table search (raw text mode)
        if target_para is None and not violation_text.startswith('["') and not violation_text.startswith('[["'):
            tables_in_range = self._find_tables_in_range(anchor_para, item.uuid_end)

            for table_elem in tables_in_range:
                result = self._search_in_table_cell_raw(
                    table_elem, violation_text, anchor_para, item.uuid_end
                )

                if result:
                    target_para, matched_runs_info, matched_start, matched_text, strip_mode = result
                    violation_text = matched_text
                    is_cross_paragraph = False
                    if strip_mode:
                        numbering_stripped = True
                        if item.fix_action == 'replace':
                            stripped_revised, revised_has_numbering = strip_numbering_by_mode(
                                revised_text, strip_mode
                            )
                            if revised_has_numbering:
                                revised_text = stripped_revised

                    if self.verbose:
                        print("  [Success] Found in table cell (plain text mode)")
                    break

        if target_para is None:
            if boundary_error:
                if boundary_error == 'boundary_crossed':
                    tables_in_range = self._find_tables_in_range(anchor_para, item.uuid_end)

                    for table_elem in tables_in_range:
                        for tc in table_elem.iter(f'{{{NS["w"]}}}tc'):
                            cell_paras = tc.findall(f'{{{NS["w"]}}}p')
                            if not cell_paras:
                                continue

                            cell_text_parts = []
                            cell_para_runs_map = {}

                            for cell_para in cell_paras:
                                para_runs, para_text = self._collect_runs_info_original(cell_para)
                                cell_text_parts.append(para_text)
                                cell_para_runs_map[id(cell_para)] = (para_runs, para_text)

                            cell_combined_text = '\n'.join(cell_text_parts)
                            cell_normalized = self._normalize_text_for_search(cell_combined_text)

                            search_attempts: List[Tuple[str, Optional[str]]] = [(violation_text, None)]
                            search_attempts.extend(build_numbering_variants(violation_text))
                            if '\\n' in violation_text:
                                newline_text = violation_text.replace('\\n', '\n')
                                search_attempts.append((newline_text, None))
                                search_attempts.extend(build_numbering_variants(newline_text))

                            deduped_attempts = dedupe_search_attempts(search_attempts)

                            match_pos = -1
                            matched_search_text = violation_text
                            matched_strip_mode: Optional[str] = None
                            for search_text, strip_mode in deduped_attempts:
                                match_pos = cell_normalized.find(search_text)
                                if match_pos != -1:
                                    matched_search_text = search_text
                                    matched_strip_mode = strip_mode
                                    break

                            if match_pos != -1:
                                if self.verbose:
                                    print("  [Success] Found in table cell (raw text match)")
                                if matched_strip_mode:
                                    numbering_stripped = True
                                    violation_text = matched_search_text
                                    if item.fix_action == 'replace':
                                        stripped_revised, revised_has_numbering = strip_numbering_by_mode(
                                            revised_text, matched_strip_mode
                                        )
                                        if revised_has_numbering:
                                            revised_text = stripped_revised
                                else:
                                    violation_text = matched_search_text

                                current_offset = 0
                                matched_para = None
                                matched_para_runs = None
                                matched_para_start = -1

                                for cell_para in cell_paras:
                                    para_id_obj = id(cell_para)
                                    if para_id_obj not in cell_para_runs_map:
                                        continue

                                    para_runs, para_text = cell_para_runs_map[para_id_obj]
                                    para_len = len(para_text)

                                    if current_offset <= match_pos < current_offset + para_len:
                                        matched_para = cell_para
                                        matched_para_runs = para_runs
                                        matched_para_start = match_pos - current_offset
                                        break

                                    current_offset += para_len + 1

                                if matched_para is not None:
                                    target_para = matched_para
                                    matched_runs_info = matched_para_runs
                                    matched_start = matched_para_start
                                    is_cross_paragraph = False
                                    break
                            else:
                                marker_idx = cell_combined_text.find(DEBUG_MARKER)
                                if DEBUG_MARKER and marker_idx != -1:
                                    snippet_len = len(DEBUG_MARKER) + 60
                                    snippet = cell_combined_text[marker_idx:marker_idx + snippet_len]
                                    print(
                                        "  [DEBUG] Non-JSON cell content from marker: "
                                        f"{repr(snippet)}"
                                    )

                        if target_para is not None:
                            break

                if target_para is None:
                    if self.verbose:
                        print("  [Boundary] Table search failed, trying body text...")

                    body_segments = []
                    current_segment = []

                    for para in self._iter_paragraphs_in_range(anchor_para, item.uuid_end):
                        if self._is_paragraph_in_table(para):
                            if current_segment:
                                body_segments.append(current_segment)
                                current_segment = []
                            continue

                        para_runs, para_text = self._collect_runs_info_original(para)
                        if not para_text.strip():
                            continue

                        current_segment.append((para, para_runs, para_text))

                    if current_segment:
                        body_segments.append(current_segment)

                    for segment_idx, body_paras_data in enumerate(body_segments):
                        all_runs = []
                        pos = 0

                        for i, (para, para_runs, para_text) in enumerate(body_paras_data):
                            for run in para_runs:
                                run_copy = dict(run)
                                run_copy['para_elem'] = para
                                run_copy['start'] = run['start'] + pos
                                run_copy['end'] = run['end'] + pos
                                all_runs.append(run_copy)

                            pos += len(para_text)

                            if i < len(body_paras_data) - 1:
                                all_runs.append({
                                    'text': '\n',
                                    'start': pos,
                                    'end': pos + 1,
                                    'para_elem': para,
                                    'is_para_boundary': True,
                                })
                                pos += 1

                        search_attempts: List[Tuple[str, Optional[str]]] = [(violation_text, None)]
                        numbering_variants = build_numbering_variants(violation_text)
                        search_attempts.extend(numbering_variants)

                        match_pos = -1
                        matched_text = violation_text
                        matched_override = None
                        matched_strip_mode: Optional[str] = None

                        for search_text, strip_mode in search_attempts:
                            match_pos, matched_override = self._find_in_runs_with_normalization(
                                all_runs, search_text
                            )
                            if match_pos != -1:
                                matched_text = matched_override or search_text
                                matched_strip_mode = strip_mode
                                break

                        if match_pos != -1:
                            target_para = body_paras_data[0][0]
                            matched_runs_info = all_runs
                            matched_start = match_pos
                            is_cross_paragraph = len(body_paras_data) > 1
                            violation_text = matched_text
                            numbering_stripped = matched_strip_mode is not None
                            if matched_strip_mode and item.fix_action == 'replace':
                                stripped_revised, _ = strip_numbering_by_mode(
                                    revised_text, matched_strip_mode
                                )
                                revised_text = stripped_revised

                            if self.verbose:
                                if is_cross_paragraph:
                                    print(
                                        f"  [Success] Found in body segment {segment_idx + 1} "
                                        f"(cross-paragraph)"
                                    )
                                else:
                                    print(f"  [Success] Found in body segment {segment_idx + 1}")
                            break

                if target_para is None:
                    if boundary_error == 'boundary_crossed':
                        reason = "Violation text not found(C)"
                    else:
                        reason = f"Boundary error: {boundary_error}"

                    if DEBUG_MARKER:
                        try:
                            for para in self._iter_paragraphs_in_range(anchor_para, item.uuid_end):
                                if self._is_paragraph_in_table(para):
                                    continue
                                _, para_text = self._collect_runs_info_original(para)
                                marker_idx = para_text.find(DEBUG_MARKER)
                                if marker_idx != -1:
                                    print(f"  [DEBUG] Body search target: {repr(violation_text)}")
                                    break
                        except Exception:
                            pass

                    self._apply_fallback_comment(anchor_para, item, reason)
                    if self.verbose:
                        print(f"  [Boundary] {reason}")
                    return {
                        'target_para': target_para,
                        'matched_runs_info': matched_runs_info,
                        'matched_start': matched_start,
                        'violation_text': violation_text,
                        'revised_text': revised_text,
                        'numbering_stripped': numbering_stripped,
                        'is_cross_paragraph': is_cross_paragraph,
                        'early_result': EditResult(
                            success=True,
                            item=item,
                            error_message=reason,
                            warning=True,
                        ),
                    }

            if target_para is None:
                if DEBUG_MARKER:
                    try:
                        for para in self._iter_paragraphs_in_range(anchor_para, item.uuid_end):
                            if self._is_paragraph_in_table(para):
                                continue
                            _, para_text = self._collect_runs_info_original(para)
                            marker_idx = para_text.find(DEBUG_MARKER)
                            if marker_idx != -1:
                                print(f"  [DEBUG] Body search target: {repr(violation_text)}")
                                break
                    except Exception:
                        pass

                tables_count = len(tables_in_range)
                body_count = len(body_segments)
                total_segments = tables_count + body_count

                if total_segments == 1:
                    reason = "Violation text not found(S)"
                else:
                    reason = "Violation text not found(M)"

                if item.fix_action == 'manual':
                    self._apply_error_comment(anchor_para, item)
                    return {
                        'target_para': target_para,
                        'matched_runs_info': matched_runs_info,
                        'matched_start': matched_start,
                        'violation_text': violation_text,
                        'revised_text': revised_text,
                        'numbering_stripped': numbering_stripped,
                        'is_cross_paragraph': is_cross_paragraph,
                        'early_result': EditResult(
                            success=True,
                            item=item,
                            error_message=reason,
                            warning=True,
                        ),
                    }
                self._apply_error_comment(anchor_para, item)
                return {
                    'target_para': target_para,
                    'matched_runs_info': matched_runs_info,
                    'matched_start': matched_start,
                    'violation_text': violation_text,
                    'revised_text': revised_text,
                    'numbering_stripped': numbering_stripped,
                    'is_cross_paragraph': is_cross_paragraph,
                    'early_result': EditResult(False, item, reason),
                }

        return {
            'target_para': target_para,
            'matched_runs_info': matched_runs_info,
            'matched_start': matched_start,
            'violation_text': violation_text,
            'revised_text': revised_text,
            'numbering_stripped': numbering_stripped,
            'is_cross_paragraph': is_cross_paragraph,
            'early_result': None,
        }
