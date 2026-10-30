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


def unpack(packed, count):
    raw = packed.encode('utf-16-le')
    pos = 0
    result = []
    for _ in range(count):
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
    if pos != len(raw):
        raise ValueError('extra timestamps')
    return result


def legacy_maps(records):
    result = {}
    for rid, h, value, stamp, version in records:
        rid, h = text(rid).strip(), canonical(h)
        if not rid and not h:
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
