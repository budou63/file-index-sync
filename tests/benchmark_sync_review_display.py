"""Executable Python cost/equivalence model, NOT an Excel/VBA benchmark.

Mirrors the old column scan and the new per-render header map. Uses the current
form's actual CompareFieldSpec aliases; supplies synthetic headers and values.
No COM, workbook, CSV, adoption, or synchronization operations are performed.
Run: python -B tests/benchmark_sync_review_display.py --records 379 --repeats 3
"""
from collections import Counter
from pathlib import Path
import argparse
import json
import re
import statistics
from time import perf_counter
import unicodedata

ROOT = Path(__file__).resolve().parents[1]
FORM = (ROOT / 'vba/frmSyncReview.txt').read_text(encoding='utf-8')


def key(value):
    # Test approximation of StrConv(vbNarrow), not a claim of locale equivalence.
    return unicodedata.normalize('NFKC', value.strip().replace(' ', '').replace('　', '')).lower()


def aliases():
    match = re.search(r'(?ms)^Private Sub CompareFieldSpec\b.*?^End Sub', FORM)
    assert match is not None
    result = {}
    for index, block in re.findall(r'Case (\d+)\s*(.*?)(?=\n\s*Case |\n\s*End Select)', match.group(), re.S):
        result[int(index)] = {}
        for side in ('csv', 'excel'):
            array = re.search(side + r'Aliases = Array\(((?:"[^"]*"\s*,?\s*)*)\)', block)
            result[int(index)][side] = re.findall(r'"([^"]*)"', array.group(1)) if array else []
    return result


class Sheet:
    def __init__(self, name, headers):
        self.name, self.headers = name, headers
        self.calls = Counter()

    def cell(self, col):
        self.calls['header_cell_value_reads'] += 1
        return self.headers[col]

    def header_range(self):
        self.calls['header_range_reads'] += 1
        return list(self.headers)


class OldLookup:
    def column(self, sheet, names):
        found = []
        for col in range(len(sheet.headers)):
            if sheet.cell(col) is not None:  # None models an error header.
                header = key(sheet.cell(col))
                for name in names:
                    if header == key(name):
                        found.append(col + 1)
                        break
        return found[0] if len(found) == 1 else 0


class FrameLookup:
    def __init__(self):
        self.maps = {}
        self.normalized = {}

    def norm(self, value):
        if value not in self.normalized:
            self.normalized[value] = key(value)
        return self.normalized[value]

    def column(self, sheet, names):
        if sheet.name not in self.maps:
            mapping = {}
            for col, value in enumerate(sheet.header_range(), 1):
                if value is not None:
                    header = self.norm(value)
                    mapping[header] = 0 if header in mapping else col
            self.maps[sheet.name] = mapping
        mapping, matched = self.maps[sheet.name], set()
        for name in names:
            header = self.norm(name)
            if header in mapping:
                col = mapping[header]
                if col == 0:
                    return 0
                matched.add(col)
        return next(iter(matched)) if len(matched) == 1 else 0


class SelectionGate:
    def __init__(self):
        self.rendered = ''
        self.rebuilds = 0

    def reset(self):
        self.rendered = ''

    def event(self, mode, focus, selected):
        stamp = f'{mode}:{focus}:' + ''.join(f'{i};' for i in sorted(selected))
        if stamp == self.rendered:
            return False
        self.rebuilds += 1
        self.rendered = stamp
        return True


def run_operation(operation, optimized, records, columns):
    spec = aliases()
    sheets = {}
    for side in ('csv', 'excel'):
        headers = [spec[index][side][0] for index in range(1, 9)]
        headers += [f'untouched_{n}' for n in range(columns - len(headers))]
        sheets[side] = Sheet(side, headers)
    outputs, rebuilds = [], 0

    def frame(indices, fields):
        reader = FrameLookup() if optimized else OldLookup()
        out = []
        for row in indices:
            for field in fields:
                for side, sheet in sheets.items():
                    col = reader.column(sheet, spec[field][side])
                    # The production code still reads Value2/Text freshly. Here only
                    # a deterministic displayed value/column is compared, not Excel formatting.
                    out.append((row, field, side, col, f'{side}:{row}:{col}'))
        return out

    if operation in ('open', 'input_to_all', 'all_to_input'):
        outputs.append(frame(range(records), (7, 8, 5)))
        if operation != 'open':
            outputs.append(frame((records // 2,), (7, 8, 5, 6, 1, 2, 3)))
        rebuilds = 1
    else:
        positions = {'first': [0], 'middle': [records // 2], 'last': [records - 1],
                     'continuous': list(range(records))}[operation]
        gate = SelectionGate()
        for pos in positions:
            gate.reset()  # a new mouse/keyboard gesture
            last = None
            for event in ('Change', 'Click', 'MouseUp'):
                if not optimized or gate.event(0, pos, {pos}):
                    last = frame((pos,), (7, 8, 5, 6, 1, 2, 3))
                    rebuilds += 1
            outputs.append(last)
    counts = sum((sheet.calls for sheet in sheets.values()), Counter())
    return outputs, dict(counts), rebuilds


def benchmark(records=379, columns=128, repeats=3):
    report = {'kind': 'Python model only; not Excel, COM or VBA timings',
              'data': 'synthetic', 'records': records, 'columns_per_sheet': columns,
              'repeats': repeats, 'operations': {}}
    for operation in ('open', 'first', 'middle', 'last', 'continuous', 'input_to_all', 'all_to_input'):
        results = []
        for optimized in (False, True):
            times = []
            output, counts, rebuilds = [], {}, 0
            for _ in range(repeats):
                started = perf_counter()
                output, counts, rebuilds = run_operation(operation, optimized, records, columns)
                times.append((perf_counter() - started) * 1000)
            results.append((output, {'median_python_model_ms': round(statistics.median(times), 3),
                                     'reads': counts, 'list_rebuilds_model': rebuilds}))
        assert results[0][0] == results[1][0], operation
        report['operations'][operation] = {'before': results[0][1], 'after': results[1][1],
                                           'model_display_values_equal': True}
    return report


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--records', type=int, default=379)
    parser.add_argument('--columns', type=int, default=128)
    parser.add_argument('--repeats', type=int, default=3)
    args = parser.parse_args()
    if args.records < 1 or args.columns < 8 or args.repeats < 1:
        parser.error('records/repeats must be positive and columns >= 8')
    print(json.dumps(benchmark(args.records, args.columns, args.repeats), ensure_ascii=False, indent=2))
