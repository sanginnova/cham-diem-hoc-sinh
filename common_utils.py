# -*- coding: utf-8 -*-
"""Shared utilities and Grade 12 master keys for Cham_Diem_Hoc_Sinh."""

import unicodedata as _ud

_VN_CHAR_ORDER = {c: i for i, c in enumerate("aăâbcdđeêfghijklmnoôơpqrstuưvwxyz")}
_VN_TONE_ORDER = {"": 0, "\u0300": 1, "\u0301": 2, "\u0309": 3, "\u0303": 4, "\u0323": 5}


def _vn_word_key(w: str):
    nfd = _ud.normalize("NFD", w.lower())
    base, tones = [], []
    for ch in nfd:
        if ch in _VN_TONE_ORDER:
            tones.append(_VN_TONE_ORDER[ch])
        elif ch == "\u0306":
            if base and base[-1] == "a":
                base[-1] = "ă"
        elif ch == "\u0302":
            if base and base[-1] in ("a", "e", "o"):
                base[-1] = {"a": "â", "e": "ê", "o": "ô"}[base[-1]]
        elif ch == "\u031b":
            if base and base[-1] in ("o", "u"):
                base[-1] = {"o": "ơ", "u": "ư"}[base[-1]]
        else:
            base.append(ch)
    return ([_VN_CHAR_ORDER.get(c, 100 + ord(c)) for c in base], tones)


def vn_name_sort_key(name: str):
    """Sort key for Vietnamese full names: Ten first, then Ho & Ten dem."""
    parts = name.strip().split()
    if not parts:
        return ([], [])
    return (_vn_word_key(parts[-1]), [_vn_word_key(p) for p in parts[:-1]])


MASTER_KEYS_12 = {
    "101": {
        "p1": ["A", "C", "D", "C", "D", "B", "D", "A", "B", "D", "B", "C"],
        "p2_13": ["Đ", "Đ", "S", "S"],
        "p2_14": ["S", "Đ", "Đ", "S"],
        "p3": ["2", "1", "7", "18"],
    },
    "102": {
        "p1": ["A", "D", "A", "D", "D", "D", "D", "D", "C", "B", "A", "C"],
        "p2_13": ["S", "Đ", "Đ", "S"],
        "p2_14": ["Đ", "Đ", "S", "S"],
        "p3": ["7", "18", "1", "2"],
    },
    "103": {
        "p1": ["B", "A", "B", "D", "C", "A", "A", "D", "B", "C", "A", "C"],
        "p2_13": ["Đ", "Đ", "S", "S"],
        "p2_14": ["S", "Đ", "Đ", "S"],
        "p3": ["18", "2", "1", "7"],
    },
    "104": {
        "p1": ["D", "A", "B", "A", "D", "C", "A", "D", "C", "C", "D", "D"],
        "p2_13": ["Đ", "Đ", "S", "S"],
        "p2_14": ["S", "Đ", "Đ", "S"],
        "p3": ["2", "1", "7", "18"],
    },
    "201": {
        "p1": ["D", "B", "C", "A", "C", "D", "B", "C", "B", "A", "B", "D"],
        "p2_13": ["Đ", "S", "S", "Đ"],
        "p2_14": ["Đ", "S", "Đ", "Đ"],
        "p3": ["7", "20", "1", "1"],
    },
    "202": {
        "p1": ["D", "C", "A", "B", "B", "D", "A", "A", "A", "B", "A", "D"],
        "p2_13": ["Đ", "S", "Đ", "Đ"],
        "p2_14": ["Đ", "S", "S", "Đ"],
        "p3": ["1", "20", "1", "7"],
    },
    "203": {
        "p1": ["A", "B", "D", "A", "B", "A", "C", "B", "A", "D", "C", "C"],
        "p2_13": ["Đ", "S", "Đ", "Đ"],
        "p2_14": ["S", "Đ", "S", "Đ"],
        "p3": ["1", "1", "20", "7"],
    },
    "204": {
        "p1": ["B", "D", "A", "B", "D", "B", "C", "A", "C", "D", "A", "A"],
        "p2_13": ["Đ", "S", "Đ", "Đ"],
        "p2_14": ["Đ", "S", "Đ", "S"],
        "p3": ["1", "1", "7", "20"],
    },
}


def get_core_keys_12():
    """Return 8-code dict (+ 4-digit '0xxx' aliases) with merged 'p2' list."""
    res = {
        k: {"p1": v["p1"], "p2": v["p2_13"] + v["p2_14"], "p3": v["p3"]}
        for k, v in MASTER_KEYS_12.items()
    }
    for k in list(res.keys()):
        res[f"0{k}"] = res[k]
    return res
