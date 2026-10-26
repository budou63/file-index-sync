"""Same synthetic records, pure-Python reference model, NOT Excel timing.

Run: python -B tests/benchmark_sync_review_second_pass.py
379 input files + 6 internal-only records; 380 normal and 561 internal groups.
The model deliberately DOES NOT simulate MSForms/Excel COM latency. It reports
source-derived COM write/validation call counts separately from CPU timings.
"""
from __future__ import annotations
import json
import statistics
import time
from pathlib import Path
from typing import Any

HERE = Path(__file__).resolve().parents[1]
SOURCE = (HERE / 'vba/frmSyncReview.txt').read_text(encoding='utf-8')


def make_records():
    records = []
    for n in range(385):
        groups = [True] if n < 379 else []
        if n == 0:
            groups.append(True)
        internal = 561 // 385 + (n < 561 % 385)
        groups.extend([False] * internal)
        records.append((f'stable:{n:04d}', groups, (f'file:{n}', f'folder:{n}', f'year:{n}')))
    assert sum(any(grp for grp in r[1]) for r in records) == 379
    assert sum(sum(grp for grp in r[1]) for r in records) == 380
    assert sum(sum(not grp for grp in r[1]) for r in records) == 561
    assert sum(not any(grp for grp in r[1]) for r in records) == 6
    return records


def render(records, full, new):
    rows = []
    input_groups = 0
    internal_groups = 0
    needs_calls = 0
    cell_writes = 0
    bulk_writes = 0

    def needs(groups):
        nonlocal needs_calls
        needs_calls += 1
        return any(groups)

    for key, groups, names in records:
        # Same visible set and summary; old code calls ERecordNeedsConfirmation
        # in the count, filter and internal-only label branches.
        needed = needs(groups)
        if needed:
            input_groups += sum(groups)
        internal_groups += sum(not grp for grp in groups)
        if new:
            visible = full or needed
            only_internal = not needed
        else:
            visible = full or needs(groups)
            only_internal = visible and not needs(groups)
        if visible:
            label = ('内部のみ・要調査: ' if only_internal else '') + names[0]
            if new:
                rows.append((label, names[1], names[2]))
            else:
                rows.append([label, '', ''])
                rows[-1][1] = names[1]
                rows[-1][2] = names[2]
                cell_writes += 3
    if new and rows:
        # A single MSForms.List assignment vs 3 property writes per row.
        bulk_writes = 1
    return {
        'keys': [r[0] for r in records if full or any(r[1])],
        'rows': [tuple(row) for row in rows],
        'count': (len(rows), input_groups, internal_groups),
        'needs_calls': needs_calls, 'cell_writes': cell_writes,
        'bulk_writes': bulk_writes,
    }


def validate_for_selected_field(fields):
    # Synthetic scan cost. The real VBA validation still runs on each gesture;
    # the source of truth and live CSV/workbook checks are not cached.
    checksum = 0
    for decision in ('system', 'excel', 'hold'):
        for key, system, excel in fields:
            checksum += (len(key) + len(system if decision == 'system' else excel))
    return checksum


def timing(fn, repetitions=40):
    samples = []
    for _ in range(repetitions):
        start = time.perf_counter_ns()
        fn()
        samples.append((time.perf_counter_ns() - start) / 1_000_000)
    return round(statistics.median(samples), 3)


def main():
    assert 'Private Sub mLstRows_MouseDown(' in SOURCE
    assert 'Private Sub mLstRows_KeyDown(' in SOURCE
    assert 'If fieldSelectionKey = eRenderedFieldSelection Then Exit Sub' in SOURCE
    assert 'ShowECategory True, False' in SOURCE
    assert 'mLstCases.List = listData' in SOURCE
    records = make_records()
    fields = [(str(n), str(n + 1), str(n + 2)) for n in range(561)]
    result: dict[str, Any] = {'note': 'Python-only synthetic model; no Excel/VBA event or COM timing.'}
    for full, label in [(False, 'input_only'), (True, 'all_items')]:
        before, after = render(records, full, False), render(records, full, True)
        assert before['keys'] == after['keys'] and before['rows'] == after['rows']
        assert before['count'] == after['count']
        result[label] = {
            'same_display_and_counts': True, 'visible_files': before['count'][0],
            'before_model_ms': timing(lambda: render(records, full, False)),
            'after_model_ms': timing(lambda: render(records, full, True)),
            'before_needs_scans': before['needs_calls'],
            'after_needs_scans': after['needs_calls'],
            'before_listbox_cell_writes': before['cell_writes'],
            'after_listbox_bulk_writes': after['bulk_writes'],
        }
    for repeat, label in [(1, 'right_single'), (10, 'right_ten')]:
        old = lambda: [validate_for_selected_field(fields) for _ in range(repeat * 2)]
        new = lambda: [validate_for_selected_field(fields) for _ in range(repeat)]
        assert old()[-1] == new()[-1]
        result[label] = {'before_model_ms': timing(old), 'after_model_ms': timing(new),
                         'before_live_validation_sets': 2 * repeat, 'after_live_validation_sets': repeat,
                         'live_checks_per_set': 3}
    result['left_selection'] = {'before_model_ms': None, 'after_model_ms': None,
                                'note': 'unchanged render path; Excel timing required'}
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
