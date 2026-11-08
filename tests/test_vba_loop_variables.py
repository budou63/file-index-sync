"""Static VBA loop-variable checks; these do not replace a VBE compile."""
from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]
START = re.compile(r"^For\s+(?:Each\s+)?([\w\u0080-\uffff]+)\s*(?:In\b|=)", re.I)
END = re.compile(r"^Next(?:\s+([\w\u0080-\uffff]+(?:\s*,\s*[\w\u0080-\uffff]+)*))?\s*$", re.I)
STRINGS = re.compile(r'"(?:[^"]|"")*"')


def loop_errors(source):
    stack, errors = [], []
    for line_no, line in enumerate(source.splitlines(), 1):
        # Mask string literals before removing comments and splitting statements.
        code = STRINGS.sub('""', line).split("'", 1)[0]
        for statement in code.split(':'):
            statement = statement.strip()
            start = START.match(statement)
            if start:
                stack.append((start.group(1), line_no))
                continue
            end = END.match(statement)
            if not end:
                continue
            names = end.group(1).split(',') if end.group(1) else [None]
            for name in names:
                if not stack:
                    errors.append((line_no, 'Next without For'))
                    continue
                expected, opened = stack.pop()
                if name and name.strip().casefold() != expected.casefold():
                    errors.append((line_no, f'Next {name.strip()} does not match For {expected} on line {opened}'))
    errors.extend((line_no, f'For {name} has no Next') for name, line_no in stack)
    return errors


class VbaLoopVariables(unittest.TestCase):
    def test_detects_original_bud151_mismatch(self):
        source = 'For Each headerGroup In groups\n  For Each name In headerGroup\n  Next name\nNext group'
        self.assertEqual(loop_errors(source), [(4, 'Next group does not match For headerGroup on line 1')])

    def test_valid_nested_and_inline_loops_and_quoted_text(self):
        source = 'For Each x In xs\nFor i = 1 To 3: Debug.Print "Next y: For z = 1" : Next i\nNext x'
        self.assertEqual(loop_errors(source), [])

    def test_all_managed_vba_loop_variables_match(self):
        paths = sorted((ROOT / 'vba').glob('*.txt'))
        self.assertTrue(paths, 'No managed VBA sources found')
        for path in paths:
            with self.subTest(component=path.name):
                self.assertEqual(loop_errors(path.read_text(encoding='utf-8-sig')), [])


if __name__ == '__main__':
    unittest.main()
