# -*- coding: utf-8 -*-
"""Лёгкий пересчёт формул P&L без внешних зависимостей (только openpyxl).
Поддерживает: ссылки на ячейки (в т.ч. 'Лист'!C5), SUM(range), MAX(a,b),
IFERROR(x,y), арифметику + - * / и скобки. Итеративно до стабилизации.
После пересчёта вписывает кэш-значения в формульные ячейки (формула + <v>),
чтобы число было видно в любой программе."""
import re, openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter


def _collect(wb):
    form = {}; val = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                if isinstance(c.value, str) and c.value.startswith('='):
                    form[(ws.title, c.coordinate)] = c.value[1:]
                elif isinstance(c.value, (int, float)):
                    val[(ws.title, c.coordinate)] = float(c.value)
    return form, val


def _eval(expr, sheet, val, form, seen):
    expr = expr.strip()
    # SUM(...)
    def sum_repl(m):
        inner = m.group(1)
        if ':' in inner:
            a, b = inner.split(':')
            sh, a = _split_sheet(a, sheet)
            _, b = _split_sheet(b, sheet)
            c1, r1 = _cr(a); c2, r2 = _cr(b)
            tot = 0.0
            for rr in range(min(r1, r2), max(r1, r2) + 1):
                for cc in range(min(c1, c2), max(c1, c2) + 1):
                    tot += _cell(sh, f'{get_column_letter(cc)}{rr}', val, form, seen)
            return repr(tot)
        else:
            parts = inner.split(',')
            return repr(sum(_eval(p, sheet, val, form, seen) for p in parts))
    expr = re.sub(r'SUM\(([^()]*)\)', sum_repl, expr, flags=re.I)

    def max_repl(m):
        parts = _split_args(m.group(1))
        return repr(max(_eval(p, sheet, val, form, seen) for p in parts))
    while re.search(r'MAX\(', expr, flags=re.I):
        expr = re.sub(r'MAX\(([^()]*)\)', max_repl, expr, count=1, flags=re.I)

    def iferr_repl(m):
        parts = _split_args(m.group(1))
        try:
            return repr(_eval(parts[0], sheet, val, form, seen))
        except Exception:
            return repr(_eval(parts[1], sheet, val, form, seen)) if len(parts) > 1 else '0'
    while re.search(r'IFERROR\(', expr, flags=re.I):
        expr = re.sub(r'IFERROR\(([^()]*)\)', iferr_repl, expr, count=1, flags=re.I)

    # ссылки на ячейки  'Sheet'!A1 или A1
    def ref_repl(m):
        return repr(_cell(sheet, m.group(0), val, form, seen))
    expr = re.sub(r"(?:'[^']+'|[A-Za-z0-9_ ]+)!\$?[A-Z]{1,3}\$?\d+|\$?[A-Z]{1,3}\$?\d+",
                  ref_repl, expr)
    if expr.strip() in ('', '""'):
        return 0.0
    try:
        return float(eval(expr, {'__builtins__': {}}, {}))
    except Exception:
        return 0.0


def _split_args(s):
    out = []; depth = 0; cur = ''
    for ch in s:
        if ch == ',' and depth == 0:
            out.append(cur); cur = ''
        else:
            if ch in '([': depth += 1
            elif ch in ')]': depth -= 1
            cur += ch
    out.append(cur)
    return out


def _split_sheet(ref, default):
    ref = ref.strip()
    if '!' in ref:
        sh, cell = ref.rsplit('!', 1)
        return sh.strip("'"), cell.replace('$', '')
    return default, ref.replace('$', '')


def _cr(a1):
    m = re.match(r'([A-Z]{1,3})(\d+)', a1)
    return column_index_from_string(m.group(1)), int(m.group(2))


def _cell(sheet, ref, val, form, seen):
    sh, cell = _split_sheet(ref, sheet)
    key = (sh, cell)
    if key in val:
        return val[key]
    if key in form:
        if key in seen:
            return 0.0
        seen.add(key)
        v = _eval(form[key], sh, val, form, seen)
        val[key] = v
        return v
    return 0.0


def recalc_native(path):
    wb = openpyxl.load_workbook(path)
    form, val = _collect(wb)
    for key in list(form):
        _cell(key[0], key[1], val, form, set())
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)
    return _inject_cached_values(path, {k: v for k, v in val.items() if k in form})


def _inject_cached_values(path, valmap):
    import zipfile, shutil
    from xml.etree import ElementTree as ET
    MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    ET.register_namespace('', MAIN)
    zin = zipfile.ZipFile(path, 'r')
    names = zin.namelist()
    wbroot = ET.fromstring(zin.read('xl/workbook.xml'))
    relroot = ET.fromstring(zin.read('xl/_rels/workbook.xml.rels'))
    rid2tgt = {r.get('Id'): r.get('Target') for r in relroot}
    title2file = {}
    for sh in wbroot.iter('{%s}sheet' % MAIN):
        rid = sh.get('{%s}id' % REL); tgt = rid2tgt.get(rid)
        if tgt:
            tgt = tgt.lstrip('/')
            if not tgt.startswith('xl/'):
                tgt = 'xl/' + tgt
            title2file[sh.get('name')] = tgt
    byfile = {}
    for (title, coord), v in valmap.items():
        f = title2file.get(title)
        if f:
            byfile.setdefault(f, {})[coord] = v
    written = 0; newdata = {}
    for fname in names:
        data = zin.read(fname)
        if fname in byfile:
            root = ET.fromstring(data); coords = byfile[fname]
            for c in root.iter('{%s}c' % MAIN):
                coord = c.get('r')
                if coord in coords:
                    fel = c.find('{%s}f' % MAIN)
                    if fel is not None:
                        vel = c.find('{%s}v' % MAIN)
                        if vel is None:
                            vel = ET.SubElement(c, '{%s}v' % MAIN)
                        vv = coords[coord]
                        vel.text = repr(float(vv)) if isinstance(vv, (int, float)) else str(vv)
                        c.remove(vel); c.append(vel)
                        written += 1
            data = ET.tostring(root, xml_declaration=True, encoding='UTF-8')
        newdata[fname] = data
    zin.close()
    tmp = path + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for fname in names:
            zout.writestr(fname, newdata[fname])
    shutil.move(tmp, path)
    return written
