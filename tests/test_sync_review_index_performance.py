"""D detector indexing contracts; the Python model is not an Excel timing test."""
from pathlib import Path
import re
from collections import defaultdict
from time import perf_counter
import unittest


def pad(value):
    if value and value.isascii() and value.isdecimal():
        return value.lstrip("0") or "0"
    return value


def name(row):
    if not all(row[k] for k in ("created", "g1", "g2", "folder")):
        return ""
    return (row["created"], row["g1"], row["g2"], row["folder"])


def detect(csv, dest, year, indexed):
    """Independent Python decision mirror of the D-only detector, not VBA execution."""
    keys = ("id", "bar", "no")
    def value(row, key):
        return pad(row[key]) if key == "no" else row[key]

    if indexed:
        source_indexes = {k: defaultdict(list) for k in keys}
        dest_indexes = {k: defaultdict(list) for k in keys}
        for i, row in enumerate(csv):
            for k in keys:
                if value(row, k):
                    source_indexes[k][value(row, k)].append(i)
        for i, row in enumerate(dest):
            for k in keys:
                if value(row, k):
                    dest_indexes[k][value(row, k)].append(i)

    def other_match(row, source_index):
        if indexed:
            return any(any(i != source_index for i in source_indexes[k].get(value(row, k), ()))
                       for k in keys if value(row, k))
        return any(i != source_index and any(value(other, k) == value(row, k) and value(row, k)
                                             for k in keys)
                   for i, other in enumerate(csv))

    results = []
    for s, source in enumerate(csv):
        if source["year"] != year or not all(value(source, k) for k in keys) or not name(source):
            continue
        if other_match(source, s):
            continue
        if indexed:
            matches = sorted({i for k in keys for i in dest_indexes[k].get(value(source, k), ())})
        else:
            matches = [i for i, row in enumerate(dest)
                       if any(value(row, k) == value(source, k) for k in keys)]
        keepers = [i for i in matches if all(value(dest[i], k) == value(source, k) for k in keys)
                   and dest[i]["year"] == year and name(dest[i]) == name(source)]
        if len(keepers) != 1 or dest[keepers[0]]["status"] not in ("登録予定", "同期済"):
            continue
        keeper = keepers[0]
        candidates = [i for i in matches if i != keeper]
        if not candidates or any(dest[i]["status"] != "登録予定" or not name(dest[i])
                                 or name(dest[i]) == name(source) or other_match(dest[i], s)
                                 for i in candidates):
            continue
        results.append((s, keeper, tuple(candidates)))
    return results


def row(id, no, bar, folder, year=2026, created=2024, g1="guide", g2="section", status="登録予定"):
    return dict(id=id, no=no, bar=bar, folder=folder, year=year,
                created=created, g1=g1, g2=g2, status=status)


def test_d_oracle_edge_cases():
    csv = [row("A", "0007", "b7", "Alpha"), row("B", "8", "b8", "Beta"),
           row("C", "9", "b9", "Gamma"), row("X", "100", "bx", "Other", year=2025),
           row("X", "101", "by", "Other2"), row("E", "12", "be", "Epsilon"),
           row("G", "13", "bg", "Guide"), row("H", "14", "bh", "Same")]
    dest = [row("A", "7", "b7", "Alpha"), row("A", "7", "b7", "Different"),
            row("B", "8", "b8", "Beta"), row("B", "008", "b8", "Different"),
            row("C", "9", "b9", "Gamma"), row("C", "9", "b9", "Other", status="同期済"),
            row("E", "12", "be", "Epsilon"), row("E", "12", "be", "Epsilon", year=2025),
            row("G", "13", "bg", "Guide"), row("G", "13", "bg", "Different", g1="else"),
            row("H", "14", "bh", "Same"), row("H", "14", "bh", "Same")]
    assert detect(csv, dest, 2026, False) == [(0, 0, (1,)), (1, 2, (3,)), (6, 8, (9,))]
    assert detect(csv, dest, 2026, True) == detect(csv, dest, 2026, False)


def test_d_oracle_cross_year_duplicates_and_candidate_safeguards():
    csv = [row("ID", "00042", "barcode", "Folder"),
           row("OTHER", "43", "other", "Folder", year=2025)]
    dest = [row("ID", "42", "barcode", "Folder"),
            row("ID", "00042", "barcode", "Another")]
    assert detect(csv, dest, 2026, False) == [(0, 0, (1,))]
    for conflicting in (row("ID", "44", "unrelated", "Foreign", year=2025),
                        row("unrelated", "00042", "unrelated", "Foreign", year=2025),
                        row("unrelated", "44", "barcode", "Foreign", year=2025)):
        changed = csv + [conflicting]
        assert detect(changed, dest, 2026, False) == []
        assert detect(changed, dest, 2026, True) == []
    # A different CSV source claiming a candidate identifier vetoes the D case.
    changed = csv + [row("OTHER2", "55", "other2", "Foreign"),
                     row("OTHER3", "56", "other3", "Foreign", year=2025)]
    changed_dest = dest.copy()
    changed_dest[1] = dict(dest[1], bar="other2")
    assert detect(changed, changed_dest, 2026, False) == []
    assert detect(changed, changed_dest, 2026, True) == []


def test_d_oracle_deterministic_random_cases():
    import random
    rng = random.Random(4917)
    for _ in range(25):
        csv = [row(rng.choice(("", "A", "B", "C", "D")), rng.choice(("", "0", "00", "1", "01")),
                   rng.choice(("", "x", "y")), rng.choice(("", "Folder", "Other")),
                   year=rng.choice((2025, 2026)), g1=rng.choice(("G1", "G2"))) for _ in range(15)]
        dest = [row(rng.choice(("", "A", "B", "C", "D")), rng.choice(("", "0", "00", "1", "01")),
                    rng.choice(("", "x", "y")), rng.choice(("", "Folder", "Other")),
                    year=rng.choice((2025, 2026)), g1=rng.choice(("G1", "G2")),
                    status=rng.choice(("登録予定", "同期済", "削除予定"))) for _ in range(18)]
        assert detect(csv, dest, 2026, True) == detect(csv, dest, 2026, False)


def benchmark(size):
    csv = [row(f"I{i}", str(i), f"B{i}", f"Folder{i}", year=2026 if i % 5 == 0 else 2025)
           for i in range(size)]
    dest = [dict(r, status="同期済") for r in csv]
    times = []
    for indexed in (False, True):
        start = perf_counter()
        result = detect(csv, dest, 2026, indexed)
        times.append((perf_counter() - start, result))
    assert times[0][1] == times[1][1] == []
    return tuple(t for t, _ in times)


if __name__ == "__main__":
    for size in (5000, 10000):
        old, new = benchmark(size)
        print(f"Python model {size} CSV + {size} destination: nested={old:.3f}s indexed={new:.3f}s ratio={old/new:.1f}x")

SOURCE = Path(__file__).resolve().parents[1] / "vba/modSyncReview.txt"


def procedure(name):
    text = SOURCE.read_text(encoding="utf-8")
    match = re.search(r"(?im)^Public Function " + name + r"\([^\n]*\) As Collection\s*$", text)
    assert match, name
    return text[match.end():text.index("End Function", match.end())]


def test_d_detector_uses_indexes_not_nested_sheet_scans():
    body = procedure("SyncReviewCopyIdCandidates")
    assert 'Set csvIds = CreateObject("Scripting.Dictionary")' in body
    assert 'Set dstIds = CreateObject("Scripting.Dictionary")' in body
    assert 'ReviewMapAdd csvIds, ReviewField(src, sRow, "ファイルID"), sRow' in body
    assert 'ReviewMapAdd dstIds, ReviewField(dst, dRow, "ファイルID"), dRow' in body
    assert 'ReviewOtherRowsMatch(csvIds, fid, sRow)' in body
    assert 'ReviewMatchingRows(dstIds, dstBars, dstNumbers, fid, barcode, fileNo)' in body
    assert 'ReviewOtherRowsMatch(csvIds, ReviewField(dst, dRow, "ファイルID"), sRow)' in body
    assert "For otherRow = 2 To srcLast" not in body
    assert body.count("For dRow = 2 To dstLast") == 1  # one-time index construction only
    assert body.index("datasetSignature = ReviewCopyDatasetSignature") < body.index('Set csvIds = CreateObject')
    assert body.index('ReviewMapAdd csvIds') < body.index('If ReviewFiscal(src, sRow) <> yearValue')
    assert 'ReviewMapAdd csvNumbers, ReviewPadding(ReviewField(src, sRow, "ファイル番号"))' in body
    assert 'ReviewMapAdd dstNumbers, ReviewPadding(ReviewField(dst, dRow, "ファイル番号"))' in body
    assert 'ReviewOtherRowsMatch(csvBars, ReviewField(dst, dRow, "バーコード番号"), sRow)' in body
    assert 'ReviewOtherRowsMatch(csvNumbers, ReviewPadding(ReviewField(dst, dRow, "ファイル番号")), sRow)' in body
    assert 'If keeperRows.Count <> 1 Then GoTo NextCopySource' in body
    assert 'If Not valid Or candidates.Count = 0 Then GoTo NextCopySource' in body
    assert 'copyCase.Add "copyDatasetSignature", datasetSignature' in body


def test_index_helpers_keep_binary_keys_and_ordered_unique_rows():
    text = SOURCE.read_text(encoding="utf-8")
    helper = text[text.index("Private Function ReviewOtherRowsMatch("):text.index("Public Function SyncReviewCopyIdCandidates(")]
    assert 'If Len(key) = 0 Then Exit Function' in helper
    assert 'rows.Count > 1 Or CLng(rows(1)) <> excludedRow' in helper
    assert 'candidate < nextRow' in helper
    assert 'If CLng(lists(i)(positions(i))) = nextRow Then positions(i) = positions(i) + 1' in helper
    assert 'Set ReviewMatchingRows = result' in helper
    # ReviewMapAdd is the shared binary-key index builder; do not change its CompareMode.
    map_helper = text[text.index("Private Sub ReviewMapAdd("):text.index("Private Function ReviewRows(")]
    assert 'map.Add key, rows' in map_helper
    assert 'CompareMode' not in map_helper


class IndexPerformanceContracts(unittest.TestCase):
    """Keep the decision-equivalence checks in standard unittest discovery."""

    def test_edge_cases(self):
        test_d_oracle_edge_cases()

    def test_cross_year_duplicates(self):
        test_d_oracle_cross_year_duplicates_and_candidate_safeguards()

    def test_deterministic_cases(self):
        test_d_oracle_deterministic_random_cases()

    def test_vba_index_wiring(self):
        test_d_detector_uses_indexes_not_nested_sheet_scans()

    def test_vba_index_order(self):
        test_index_helpers_keep_binary_keys_and_ordered_unique_rows()
