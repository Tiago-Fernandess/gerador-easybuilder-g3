from __future__ import annotations

import argparse
import copy
import re
import shutil
import sqlite3
import tempfile
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", NS_MAIN)


def q(tag: str) -> str:
    return f"{{{NS_MAIN}}}{tag}"


def col_index(ref: str) -> int:
    letters = re.match(r"[A-Z]+", ref or "")
    if not letters:
        return 0
    n = 0
    for ch in letters.group(0):
        n = n * 26 + ord(ch) - 64
    return n


def read_xlsx(path: Path):
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall(q("si")):
                shared.append("".join(t.text or "" for t in si.findall(f".//{q('t')}")))

        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        relmap = {x.attrib["Id"]: x.attrib["Target"] for x in rels}
        sheet = wb.find(f".//{q('sheet')}")
        target = relmap[sheet.attrib[f"{{{NS_REL}}}id"]]
        sheet_path = target if target.startswith("xl/") else "xl/" + target.lstrip("/")
        xml = z.read(sheet_path)
        root = ET.fromstring(xml)
        rows = []
        for row in root.findall(f".//{q('sheetData')}/{q('row')}"):
            values = {}
            for cell in row.findall(q("c")):
                v = cell.find(q("v"))
                value = "" if v is None else (v.text or "")
                if cell.attrib.get("t") == "s" and value:
                    value = shared[int(value)]
                values[col_index(cell.attrib.get("r", ""))] = value
            rows.append((row, values))
        return rows, z.namelist(), sheet_path, xml


def source_items(rows):
    """Select only recipe fields: axis cotas plus IHM_RET tempos/parametros.

    The source workbook also contains comments, diagnostics and other program
    variables. BOOL fields under IHM_RET.Bts are converted to WORD recipe items.
    """
    items = []
    for _, values in rows:
        name = str(values.get(2, "")).strip()
        typ = str(values.get(5, "")).strip().upper()
        owner = str(values.get(5, "")).strip()
        if not name or name in {"Item count", "Individual transfer"}:
            continue
        # This is a Sysmac declaration export: columns B/C are NAME/DATATYPE.
        typ = str(values.get(3, "")).strip().upper()
        if name in {"Receita_G3", "Receita", "ID"}:
            items.append({"source": name, "type": typ or "TEXT", "base": True})
            continue
        if typ.startswith("BOOL") and name.startswith("IHM_RET.Bts."):
            items.append({"source": name, "type": "WORD", "base": False, "kind": "bool"})
            continue
        if not typ or typ.startswith("BOOL") or typ in {"COMMENT", "BYTE", "DWORD", "LREAL"}:
            continue
        axis_member = re.match(r"^Struct_Eixo_\d+$", owner, re.IGNORECASE)
        # Recipe positions are the numbered Cota_1, Cota_2, ... fields. Other
        # axis members such as Cota_Estica_Correia and Limites are not recipe
        # items in the working EasyBuilder export.
        is_axis_cota = bool(axis_member and re.match(r"^[^.]+\.Cota_\d+(?:_|$)", name, re.IGNORECASE))
        is_tempo_param = name.startswith("IHM_RET.Tempos.") or name.startswith("IHM_RET.Parametros.")
        if is_axis_cota or is_tempo_param:
            items.append({"source": name, "type": typ, "base": False})
    # The EasyBuilder recipe always starts with these two definition items.
    seen = {x["source"] for x in items}
    prefix = []
    if "Receita" not in seen:
        prefix.append({"source": "Receita", "type": "TEXT", "base": True})
    if "ID" not in seen:
        prefix.append({"source": "ID", "type": "INT", "base": True})
    bases = [x for x in items if x["source"] in {"Receita", "ID"}]
    bases_by_name = {x["source"]: x for x in bases}
    ordered_bases = []
    for base_name in ("Receita", "ID"):
        if base_name in bases_by_name:
            ordered_bases.append(bases_by_name[base_name])
        elif any(x["source"] == base_name for x in prefix):
            ordered_bases.append(next(x for x in prefix if x["source"] == base_name))
    items = ordered_bases + [x for x in items if x["source"] not in {"Receita", "ID"}]
    # Remove duplicate names while retaining order.
    result, seen = [], set()
    for item in items:
        if item["source"] not in seen:
            result.append(item)
            seen.add(item["source"])
    base_items = [x for x in result if x["source"] in {"Receita", "ID"}]
    axis_items = [x for x in result if re.match(r"^[^.]+\.Cota_\d+", x["source"], re.IGNORECASE)]
    tempo_items = [x for x in result if x["source"].startswith("IHM_RET.Tempos.")]
    param_items = [x for x in result if x["source"].startswith("IHM_RET.Parametros.")]
    bool_items = [x for x in result if x.get("kind") == "bool"]
    for index, item in enumerate(bool_items):
        item["rw_base"] = 5 + (index * 2)
    return base_items + axis_items + tempo_items + param_items + bool_items


def db_name(source: str) -> str:
    return source.replace(".", "_").replace(" ", "_")


def sql_type(source_type: str) -> str:
    t = source_type.upper()
    if t == "REAL":
        return "REAL"
    if t == "WORD":
        return "INT UNSIGNED"
    if t in {"INT", "DINT", "LINT", "UINT", "UDINT", "ULINT"}:
        return "INT"
    if t in {"STRING", "TEXT"}:
        return "TEXT"
    raise ValueError(f"Tipo não suportado: {source_type}")


def template_info(template: Path):
    conn = sqlite3.connect(template)
    tables = [r[0] for r in conn.execute("select name from sqlite_master where type='table'")]
    if "Receita_G3.info" not in tables or "Receita_G3" not in tables:
        conn.close()
        raise ValueError("O banco molde não contém as tabelas Receita_G3 e Receita_G3.info.")
    cols = conn.execute('pragma table_info("Receita_G3")').fetchall()
    info = conn.execute('select * from "Receita_G3.info"').fetchall()
    info_cols = [r[1] for r in conn.execute('pragma table_info("Receita_G3.info")')]
    data = conn.execute('select * from "Receita_G3" limit 1').fetchone()
    conn.close()
    return cols, info_cols, info, data


def generate_db(template: Path, output: Path, items):
    cols, info_cols, info_rows, data_row = template_info(template)
    by_name = {row[0]: row for row in info_rows}
    type_codes = {}
    for row in info_rows:
        name = row[0]
        col = next((c for c in cols if c[1] == name), None)
        if col:
            type_codes.setdefault(col[2].upper(), Counter())[row[1]] += 1
    for sqlt in ("REAL", "INT UNSIGNED", "INT", "TEXT"):
        if sqlt not in type_codes:
            raise ValueError(f"O banco molde não fornece código de ItemType para {sqlt}.")

    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()
    conn = sqlite3.connect(output)
    col_defs = []
    for item in items:
        name = db_name(item["source"])
        if name == "Receita":
            typ = "TEXT"
        else:
            typ = sql_type(item["type"])
        col_defs.append((name, typ))
    qcols = ",".join(f'"{name}" {typ}' for name, typ in col_defs)
    conn.execute(f'CREATE TABLE "Receita_G3" ({qcols})')
    info_defs = '"ItemName" TEXT,"ItemType" INT UNSIGNED,"Size" INT UNSIGNED,"Decimal_PT" INT UNSIGNED,"DisplayWidth" INT UNSIGNED,"Alignment" INT UNSIGNED,"LabelName" TEXT,"LimitEnable" INT UNSIGNED,"LimitMinimum" REAL,"LimitMaximum" REAL'
    conn.execute(f'CREATE TABLE "Receita_G3.info" ({info_defs})')
    info_values = []
    for item, (name, typ) in zip(items, col_defs):
        old = by_name.get(name)
        if old:
            vals = list(old)
            vals[0] = name
        else:
            code = type_codes[typ].most_common(1)[0][0]
            size = 40 if typ == "TEXT" else 4 if typ == "REAL" else 2
            width = 5 if typ == "TEXT" else 1
            vals = [name, code, size, 0, width, 0, "", 0, 0.0, 99.0]
        info_values.append(vals)
    conn.executemany('insert into "Receita_G3.info" values (?,?,?,?,?,?,?,?,?,?)', info_values)
    # Preserve the reference's initial recipe row where matching columns exist.
    if data_row:
        old_names = [c[1] for c in cols]
        old_values = dict(zip(old_names, data_row))
        values = [old_values.get(name, 0 if typ != "TEXT" else "Receita 1") for name, typ in col_defs]
    else:
        values = ["Receita 1" if typ == "TEXT" else 0 for _, typ in col_defs]
    placeholders = ",".join("?" for _ in values)
    conn.execute(f'insert into "Receita_G3" values ({placeholders})', values)
    conn.commit()
    conn.execute("vacuum")
    conn.close()
    return len(col_defs)


def _set_shared(cell, value, shared_map, shared_root):
    value = str(value)
    if value not in shared_map:
        si = ET.SubElement(shared_root, q("si"))
        t = ET.SubElement(si, q("t"))
        t.text = value
        shared_map[value] = len(shared_map)
    cell.attrib["t"] = "s"
    for child in list(cell):
        cell.remove(child)
    v = ET.SubElement(cell, q("v"))
    v.text = str(shared_map[value])


def update_xlsx(template: Path, output: Path, items):
    """Build the workbook from the known-good EasyBuilder Transfer Data export."""
    with zipfile.ZipFile(template) as zin:
        files = {n: zin.read(n) for n in zin.namelist()}
    shared_name = "xl/sharedStrings.xml"
    shared_root = ET.fromstring(files[shared_name])
    shared_values = ["".join(t.text or "" for t in si.findall(f".//{q('t')}")) for si in shared_root.findall(q("si"))]
    shared_map = {value: i for i, value in enumerate(shared_values)}
    worksheet_name = next(n for n in files if n.startswith("xl/worksheets/"))
    root = ET.fromstring(files[worksheet_name])
    sheet_data = root.find(f".//{q('sheetData')}")
    original_rows = list(sheet_data)
    base_rows = {int(r.attrib["r"]): r for r in original_rows if int(r.attrib.get("r", "0")) in (5, 6)}
    ref_rows, _, _, _ = read_xlsx(template)
    ref_row_by_name = {}
    for row, values in ref_rows:
        name = values.get(2, "")
        if name:
            rno = int(row.attrib.get("r", "0"))
            ref_row_by_name[name] = next((x for x in original_rows if int(x.attrib.get("r", "0")) == rno), None)
    exemplar_by_type = {}
    for row, values in ref_rows:
        rno = int(row.attrib.get("r", "0"))
        if rno >= 7 and values.get(5) in {"REAL", "WORD", "INT"}:
            exemplar_by_type.setdefault(values[5], next(x for x in original_rows if int(x.attrib.get("r", "0")) == rno))
    for row in list(sheet_data):
        if int(row.attrib.get("r", "0")) >= 7:
            sheet_data.remove(row)

    def add_row(number, item):
        if item["source"] == "Receita":
            row = copy.deepcopy(base_rows[5])
            values = {2: "Receita", 3: "Disable"}
        elif item["source"] == "ID":
            row = copy.deepcopy(base_rows[6])
            values = {2: "ID", 3: "Disable"}
        else:
            typ = item["type"].upper()
            source = item["source"]
            if source in ref_row_by_name and ref_row_by_name[source] is not None:
                # Preserve every field from the known-good export, including
                # its exact VAR:// link and any EasyBuilder-specific details.
                row = copy.deepcopy(ref_row_by_name[source])
                values = {2: db_name(source)}
            else:
                if typ not in exemplar_by_type:
                    raise ValueError(f"O modelo Transfer Data não possui linha-modelo para o tipo {typ}.")
                row = copy.deepcopy(exemplar_by_type[typ])
                if item.get("kind") == "bool":
                    values = {2: db_name(source), 3: "Enable", 4: "Local HMI", 5: "RW",
                              6: "System Tag: Off", 7: "User-defined Tag: Off",
                              8: str(item["rw_base"]), 9: "IDX: null"}
                else:
                    values = {2: source, 3: "Enable", 4: "OMRON EtherNet/IP (NJ/NX Series)", 5: typ,
                              6: "System Tag: Off", 7: "User-defined Tag: Off", 8: f"VAR://{source}", 9: "IDX: null"}
        row.attrib["r"] = str(number)
        for cell in row.findall(q("c")):
            idx = col_index(cell.attrib.get("r", ""))
            if idx <= 26:
                cell.attrib["r"] = f"{chr(64 + idx)}{number}"
            if idx in values:
                _set_shared(cell, values[idx], shared_map, shared_root)
        sheet_data.append(row)

    for number, item in enumerate(items, 5):
        if number < 7:
            continue
        add_row(number, item)
    for c in root.findall(f".//{q('c')}"):
        if c.attrib.get("r") == "B4":
            _set_shared(c, f"Item count {len(items)}", shared_map, shared_root)
            break
    files[worksheet_name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    shared_root.attrib["count"] = str(len(shared_map))
    shared_root.attrib["uniqueCount"] = str(len(shared_map))
    files[shared_name] = ET.tostring(shared_root, encoding="utf-8", xml_declaration=True)
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)


def write_bool_macros(out_dir: Path, bool_items):
    """Write the tested two-macro pattern for the IHM_RET.Bts WORD mirrors."""
    if not bool_items:
        return []

    def declarations(items):
        return [f"bool {db_name(item['source'])}" for item in items]

    write = [
        'macro_definition_begin',
        '    "id": "0"',
        '    "name": "Executa Receita BOOLs"',
        '    "startup": "false"',
        '    "periodic": [ "false" ]',
        '    "interlock": [ "false" ]',
        'macro_definition_end',
        '',
        'macro_command main()',
        '',
        '// Macro WRITE: Recipe RW-5..RW-19 -> OMRON EtherNet/IP',
        '',
    ]
    write += declarations(bool_items)
    write += [
        '', 'short selection',
        '// Read the current selected record',
        'GetData(selection, "Local HMI", RECIPE, "Receita_G3.Selection")',
        'short offset',
        'offset = selection * 16',
        '',
        '//////////////////////Declara Variavel Gerar Pop_Up Loading//////////////////////////////////////////////',
        'bool AlwaysTrue',
        'bool AlwaysFalse',
        'AlwaysTrue = true',
        'AlwaysFalse = false',
        'SetData(AlwaysTrue, "Local HMI", LB, 120, 1)',
        '',
        '//////////////////////Declara Variavel Receita Salva//////////////////////////////////////////////',
        'bool Receita_Salva',
        '',
        'GetData(Receita_Salva, "Local HMI", RW, 10000 + selection, 1)',
        '',
        'if (Receita_Salva == false) then',
        'SetData(AlwaysFalse, "Local HMI", LB, 120, 1)',
        'SetData(AlwaysTrue, "Local HMI", LB, 104, 1)',
        '',
        'return',
        'end if',
        '',
        'DELAY(200)',
        '',
    ]
    for item in bool_items:
        var = db_name(item['source'])
        write.append(f"GetData({var}, \"Local HMI\", RW, {item['rw_base']} + offset, 1)")
    write += ['', '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////', '']
    for item in bool_items:
        var = db_name(item['source'])
        write.append(f"SetData({var}, \"OMRON EtherNet/IP (NJ/NX Series)\", \"VAR://{item['source']}\", 1)")
    write += [
        '',
        'SetData(AlwaysTrue, "Local HMI", "Bit_Pop_Up_Executa_Receita", 1)',
        'SetData(AlwaysFalse, "Local HMI", LB, 120, 1)',
        '',
        'SetData(selection,"OMRON EtherNet/IP (NJ/NX Series)", "VAR://Index_Receita", 1)',
        '',
        'end macro_command',
    ]

    read = [
        'macro_definition_begin',
        '    "id": "1"',
        '    "name": "Salva Receita BOOLs"',
        '    "startup": "false"',
        '    "periodic": [ "false" ]',
        '    "interlock": [ "false" ]',
        'macro_definition_end',
        '',
        'macro_command main()',
        '',
        '// Macro READ: OMRON EtherNet/IP -> Recipe RW-5..RW-19',
        '',
    ]
    read += declarations(bool_items)
    read += [
        '', 'short selection',
        '// Read the current selected record',
        'GetData(selection, "Local HMI", RECIPE, "Receita_G3.Selection")',
        'short offset',
        'offset = selection * 16',
        '',
        '//////////////////////Declara Variavel Gerar Pop_Up Loading//////////////////////////////////////////////',
        'bool AlwaysTrue',
        'bool AlwaysFalse',
        'AlwaysTrue = true',
        'AlwaysFalse = false',
        'SetData(AlwaysTrue, "Local HMI", LB, 120, 1)',
        '',
        '//////////////////////Declara Variavel Receita Salva E Libera Receita Para Ser Executada///////////////////////',
        'bool Receita_Salva = true',
        'SetData(Receita_Salva, "Local HMI", RW, 10000 + selection, 1)',
        '',
    ]
    for item in bool_items:
        var = db_name(item['source'])
        read.append(f"GetData({var}, \"OMRON EtherNet/IP (NJ/NX Series)\", \"VAR://{item['source']}\", 1)")
    read += ['', '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////', '']
    for item in bool_items:
        var = db_name(item['source'])
        read.append(f"SetData({var}, \"Local HMI\", RW, {item['rw_base']} + offset, 1)")
    read += [
        '',
        'SetData(AlwaysTrue, "Local HMI", "Bit_Pop_Up_Executa_Receita", 1)',
        'SetData(AlwaysFalse, "Local HMI", LB, 120, 1)',
        '',
        'SetData(selection,"OMRON EtherNet/IP (NJ/NX Series)", "VAR://Index_Receita", 1)',
        '',
        'end macro_command',
    ]

    out_dir.mkdir(parents=True, exist_ok=True)
    paths = [out_dir / "0_Executa Receita BOOLs.ebm", out_dir / "1_Salva Receita BOOLs.ebm"]
    paths[0].write_text("\n".join(write) + "\n", encoding="utf-8-sig")
    paths[1].write_text("\n".join(read) + "\n", encoding="utf-8-sig")
    return paths


def generate(input_xlsx: Path, template_db: Path, out_dir: Path, transfer_template: Path | None = None):
    rows, _, _, _ = read_xlsx(input_xlsx)
    items = source_items(rows)
    if not items:
        raise ValueError("Nenhuma variável não-BOOL foi encontrada na planilha.")
    db_out = out_dir / "Definition_Variables.db"
    xlsx_out = out_dir / "Transfer_Variables.xlsx"
    count = generate_db(template_db, db_out, items)
    transfer_template = transfer_template or Path(__file__).resolve().parent.parent / "modelos" / "modelo_transfer_datas.xlsx"
    update_xlsx(transfer_template, xlsx_out, items)
    macros = write_bool_macros(out_dir, [item for item in items if item.get("kind") == "bool"])
    return db_out, xlsx_out, count, macros
