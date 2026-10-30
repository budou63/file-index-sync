"""Independent data-format oracle, NOT a VBA interpreter or Excel performance test."""
from __future__ import annotations

META = ['Excel管理行ID', '__同期基準形式', '__確定済項目', '__項目確定日時']
VERSION = 'HB1'


def text(v):
    return '' if v is None else str(v)


def canonical(v):
    h = ''.join(ch for ch in text(v) if ch not in ' \u3000\r\n\t')
    aliases = {'分類名1': '分類名１', '分類名2': '分類名２', '第1ガイド': '分類名２', '第１ガイド': '分類名２', '分類名3': '分類名３', '第2ガイド': '分類名３', '第２ガイド': '分類名３', '個別フォルダー': 'タイトル', '個別フォルダ': 'タイトル'}
    if h in {'n+1', 'N+1', 'ｎ+1', 'Ｎ+1', 'n＋1', 'N＋1', 'Ｎ＋1', 'ｎ＋1'}:
        return 'ｎ＋1'
    return aliases.get(h, h)


def key(rid, header):
    return text(rid).strip().upper(), canonical(header).casefold()


def pack(values):
    return ''.join(f'{len(text(v).encode("utf-16-le")) // 2}:{text(v)}' for v in values)


def unpack(packed, count=None):
    raw = packed.encode('utf-16-le')
    pos = 0
    result = []
    while pos < len(raw):
        end = raw.find(b':\x00', pos)
        if end < pos or (end - pos) % 2:
            raise ValueError('invalid timestamp pack')
        length = raw[pos:end].decode('utf-16-le')
        if not length.isascii() or not length.isdigit():
            raise ValueError('invalid timestamp length')
        start = end + 2
        stop = start + 2 * int(length)
        if stop > len(raw):
            raise ValueError('truncated timestamp')
        result.append(raw[start:stop].decode('utf-16-le'))
        pos = stop
    if count is not None and len(result) != count:
        raise ValueError('incorrect token count')
    return result


def legacy_maps(records):
    result = {}
    for rid, h, value, stamp, version in records:
        rid, h = text(rid).strip(), canonical(h)
        if not rid and not h:
            if text(value) or text(stamp) or text(version):
                raise ValueError('unkeyed legacy payload')
            continue
        if not rid or not h or text(version) != '1':
            raise ValueError('invalid legacy record')
        k = key(rid, h)
        if k in result:
            raise ValueError('duplicate legacy key')
        result[k] = text(value)
    return result


def pivot(records):
    records = list(records)
    original = legacy_maps(records)
    ids, fields = {}, {}
    for rid, h, value, stamp, version in records:
        if not text(rid).strip() and not canonical(h):
            continue
        ids.setdefault(text(rid).strip().upper(), text(rid).strip())
        fields.setdefault(canonical(h).casefold(), canonical(h))
    data = [META + list(fields.values())]
    timestamps = {(key(rid, h)): text(stamp) for rid,h,v,stamp,ver in records if text(rid).strip() and canonical(h)}
    for normalized_id, rid in ids.items():
        keys = [(normalized_id, h) for h in fields]
        mask = ''.join('1' if k in original else '0' for k in keys)
        data.append([rid, VERSION, mask, pack(timestamps.get(k, '') for k in keys)] + [original.get(k, '') for k in keys])
    if load(data) != original:
        raise ValueError('migration differs')
    return data


def load(data):
    if not data or data[0][:4] != META:
        raise ValueError('not an owned baseline table')
    fields = [canonical(h).casefold() for h in data[0][4:]]
    if any(not h for h in fields) or len(set(fields)) != len(fields):
        raise ValueError('ambiguous fields')
    if any(len(text(value).encode('utf-16-le'))//2 > 32767 for row in data for value in row):
        raise ValueError('cell text limit')
    ids, result = set(), {}
    for row in data[1:]:
        if len(row) != len(data[0]):
            raise ValueError('row width')
        rid = text(row[0]).strip().upper()
        if not rid or rid in ids or row[1] != VERSION:
            raise ValueError('identity or version')
        ids.add(rid)
        mask = row[2]
        if len(mask) != len(fields) or set(mask) - {'0', '1'}:
            raise ValueError('mask')
        stamps = unpack(row[3], len(fields))
        for i, h in enumerate(fields):
            value = text(row[i + 4])
            if mask[i] == '1':
                result[rid, h] = value
            elif value or stamps[i]:
                raise ValueError('unconfirmed value or timestamp')
    return result


def _matches_production_copy(candidate_headers, destination_headers):
    """Recognize a named copy by a unique set of destination headers incl. managed ID."""
    candidate = [canonical(h).casefold() for h in candidate_headers]
    destination = [canonical(h).casefold() for h in destination_headers]
    if len(candidate) < 2 or 'excel管理行id' not in candidate:
        return False
    if any(not h for h in candidate) or len(set(candidate)) != len(candidate):
        return False
    return all(destination.count(h) == 1 for h in candidate)


def ensure_model(existing_headers, destination_headers, legacy_records, destination_has_rows,
                 existing_data=None, mixed_number_format=False):
    """Model migration source choice; mixed cell formats are deliberately non-semantic."""
    if existing_headers is not None and existing_headers[:4] == META:
        if existing_data is None or load(existing_data) is None:
            raise ValueError('invalid canonical table')
        return 'canonical', existing_data, False
    if existing_headers is not None:
        markers = set(existing_headers[:4]) & set(META)
        if markers:
            raise ValueError('partial horizontal schema')
        if not _matches_production_copy(existing_headers, destination_headers):
            raise ValueError('unknown sheet structure')
        if legacy_records is None:
            raise ValueError('copy has no legacy migration source')
        migrated = pivot(legacy_records)
        if destination_has_rows and not load(migrated):
            raise ValueError('legacy baseline is insufficient')
        return 'migrated_copy', migrated, True
    if legacy_records is None:
        if destination_has_rows:
            raise ValueError('existing rows have no baseline')
        return 'empty', [META], False
    migrated = pivot(legacy_records)
    if destination_has_rows and not load(migrated):
        raise ValueError('legacy baseline is insufficient')
    return 'migrated_legacy', migrated, False


def migrate_copy_transactionally(state, migrated, fail_at=None):
    """Fault-injection model for archive/replace rollback, not VBA execution."""
    import copy
    before = copy.deepcopy(state)
    try:
        state['archive'] = copy.deepcopy(state['copy_sheet'])
        if fail_at == 'after_archive':
            raise RuntimeError('migration interrupted after archive')
        state['baseline'] = copy.deepcopy(migrated)
        if fail_at == 'after_write' or load(state['baseline']) != load(migrated):
            raise RuntimeError('migration write/readback failed')
    except RuntimeError:
        state.clear()
        state.update(before)
        assert state == before
        raise
    return state


def apply_confirmed(data, pending, stamp='2026/10/30 00:00:00'):
    """Pure oracle: only explicit confirmed keys update; no current-sheet values are read."""
    values = load(data)
    stamps = {}
    for row in data[1:]:
        dates = unpack(row[3], len(data[0])-4)
        for i, header in enumerate(data[0][4:]):
            if row[2][i] == '1':
                stamps[key(row[0], header)] = dates[i]
    updates = {}
    for (rid, header), value in pending.items():
        k = key(rid, header)
        if not k[0] or not k[1] or k in updates:
            raise ValueError('missing or duplicate pending identity')
        updates[k] = text(value)
    records = [(rid, header, value, stamps[(rid, header)], '1')
               for (rid, header), value in values.items() if (rid, header) not in updates]
    records += [(rid, header, value, stamp, '1') for (rid, header), value in updates.items()]
    return pivot(records)


def transactional_model(state, new, fail_at=None, rollback_fail=False):
    """Fault-injection model only. This does not execute VBA/Excel rollback."""
    import copy
    original = copy.deepcopy(state)
    load(new)
    try:
        state[:] = copy.deepcopy(new)
        if fail_at == 'write':
            state[-1][-1] = 'partial-corruption'
            raise RuntimeError('write failed')
        if fail_at == 'readback' or state != new:
            raise RuntimeError('readback mismatch')
    except RuntimeError:
        if rollback_fail:
            raise RuntimeError('保存せず復旧コピーから再開')
        state[:] = original
        assert state == original
        raise


def history_baseline_matches(values, rid, snapshot):
    """Only the persisted-baseline part of replay. No CSV/Excel identity or live revalidation."""
    tokens = unpack(snapshot)
    flagged = bool(tokens and tokens[0] == 'BSC2')
    tokens = tokens[1:] if flagged else tokens
    width = 6 if flagged else 5
    if not tokens or len(tokens) % width:
        raise ValueError('invalid history field pack')
    matches = True
    for start in range(0, len(tokens), width):
        fields = tokens[start:start+width]
        known = fields[5] if flagged else '1'
        if known not in ('0', '1'):
            raise ValueError('invalid history known flag')
        k = key(rid, fields[0])
        if known == '1':
            matches = matches and k in values and values[k] == fields[2]
        else:
            matches = matches and k not in values and fields[2] == ''
    return matches


def decide(baseline, excel, csv):
    if excel == csv:
        return 'confirmed'
    if baseline is None:
        return 'initial_hold'
    if excel != baseline and csv == baseline:
        return 'excel_only_hold'
    if excel == baseline and csv != baseline:
        return 'csv_only_apply'
    return 'conflict_hold'
