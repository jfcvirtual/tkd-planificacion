"""Generador del proyecto VBA (vbaProject.bin) y conversión de .xlsx a .xlsm, sin dependencias externas.

Implementa la compresión MS-OVBA y el formato de archivo compuesto MS-CFB.

Los módulos se guardan solo como código fuente comprimido (sin p-code);
Excel lo recompila automáticamente al abrir el libro.
"""
import math
import struct
import uuid

# ----------------------------------------------------------------------
# MS-OVBA 2.4.1  Compresión / descompresión
# ----------------------------------------------------------------------

def _copy_token_help(diff):
    bit_count = max(4, math.ceil(math.log2(diff)) if diff > 1 else 0)
    length_mask = 0xFFFF >> bit_count
    offset_mask = (~length_mask) & 0xFFFF
    maximum_length = (0xFFFF >> bit_count) + 3
    return length_mask, offset_mask, bit_count, maximum_length


def compress(data: bytes) -> bytes:
    out = bytearray(b"\x01")
    pos = 0
    n = len(data)
    while pos < n:
        chunk_start = pos
        chunk_end = min(pos + 4096, n)
        chunk = bytearray()
        dpos = pos
        while dpos < chunk_end:
            flag_index = len(chunk)
            chunk.append(0)
            flags = 0
            for bit in range(8):
                if dpos >= chunk_end:
                    break
                # buscar la coincidencia más larga hacia atrás dentro del chunk
                best_len, best_off = 0, 0
                _, _, _, max_len = _copy_token_help(dpos - chunk_start)
                cand = dpos - 1
                while cand >= chunk_start:
                    l = 0
                    while (dpos + l < chunk_end and l < max_len
                           and data[cand + l] == data[dpos + l]):
                        l += 1
                    if l > best_len:
                        best_len, best_off = l, dpos - cand
                    cand -= 1
                if best_len >= 3:
                    length_mask, offset_mask, bit_count, _ = _copy_token_help(dpos - chunk_start)
                    token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                    chunk += struct.pack("<H", token)
                    flags |= 1 << bit
                    dpos += best_len
                else:
                    chunk.append(data[dpos])
                    dpos += 1
            chunk[flag_index] = flags
        if len(chunk) < 4096:
            header = 0xB000 | (len(chunk) + 2 - 3)     # chunk comprimido
            out += struct.pack("<H", header) + chunk
        else:                                          # chunk sin comprimir
            raw = bytes(data[chunk_start:chunk_end]).ljust(4096, b"\x00")
            out += struct.pack("<H", 0x3000 | (4098 - 3)) + raw
        pos = chunk_end
    return bytes(out)


def decompress(buf: bytes) -> bytes:
    assert buf[0] == 1, "firma inválida"
    out = bytearray()
    i = 1
    while i < len(buf):
        header = struct.unpack_from("<H", buf, i)[0]
        size = (header & 0x0FFF) + 3
        flag = header >> 15
        chunk_end = i + size
        i += 2
        dstart = len(out)
        if not flag:
            out += buf[i:i + 4096]
            i += 4096
            continue
        while i < chunk_end:
            flags = buf[i]; i += 1
            for bit in range(8):
                if i >= chunk_end:
                    break
                if flags & (1 << bit):
                    token = struct.unpack_from("<H", buf, i)[0]; i += 2
                    length_mask, offset_mask, bit_count, _ = _copy_token_help(len(out) - dstart)
                    length = (token & length_mask) + 3
                    offset = ((token & offset_mask) >> (16 - bit_count)) + 1
                    src = len(out) - offset
                    for k in range(length):
                        out.append(out[src + k])
                else:
                    out.append(buf[i]); i += 1
    return bytes(out)


# ----------------------------------------------------------------------
# MS-CFB  Escritor de archivo compuesto (versión 3, sectores de 512 bytes)
# ----------------------------------------------------------------------
FREESECT, ENDOFCHAIN, FATSECT, NOSTREAM = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFF


class Entry:
    def __init__(self, name, kind, data=b""):
        self.name, self.kind, self.data = name, kind, data   # kind: 1 storage, 2 stream, 5 root
        self.children = []
        self.left = self.right = self.child = NOSTREAM
        self.color = 1
        self.start, self.size = ENDOFCHAIN, 0


def _key(e):
    return (len(e.name), e.name.upper())


def _build_tree(nodes, depth, h):
    """Árbol binario balanceado; los nodos del último nivel incompleto van en rojo."""
    if not nodes:
        return NOSTREAM
    mid = len(nodes) // 2
    node = nodes[mid]
    node.color = 0 if depth >= h else 1
    node.left = _build_tree(nodes[:mid], depth + 1, h)
    node.right = _build_tree(nodes[mid + 1:], depth + 1, h)
    return node.sid


def write_cfb(root: Entry) -> bytes:
    entries = []

    def walk(e):
        e.sid = len(entries)
        entries.append(e)
        for c in e.children:
            walk(c)
    walk(root)
    for e in entries:
        if e.children:
            kids = sorted(e.children, key=_key)
            h = int(math.floor(math.log2(len(kids) + 1)))
            e.child = _build_tree(kids, 0, h)

    streams = [e for e in entries if e.kind == 2]
    small = [e for e in streams if len(e.data) < 4096]
    large = [e for e in streams if len(e.data) >= 4096]

    # Mini stream
    mini = bytearray()
    minifat = []
    for e in small:
        e.size = len(e.data)
        if e.size == 0:
            e.start = ENDOFCHAIN
            continue
        nsec = math.ceil(e.size / 64)
        e.start = len(mini) // 64
        for k in range(nsec):
            minifat.append(e.start + k + 1 if k < nsec - 1 else ENDOFCHAIN)
        mini += e.data.ljust(nsec * 64, b"\x00")

    n_dir = math.ceil(len(entries) * 128 / 512)
    n_minifat = math.ceil(len(minifat) * 4 / 512)
    n_mini = math.ceil(len(mini) / 512)
    n_large = sum(math.ceil(len(e.data) / 512) for e in large)
    base = n_dir + n_minifat + n_mini + n_large
    n_fat = 1
    while n_fat * 128 < base + n_fat:
        n_fat += 1
    assert n_fat <= 109

    fat = []
    sectors = []

    def alloc(data_bytes, nsec):
        start = len(fat)
        for k in range(nsec):
            fat.append(start + k + 1 if k < nsec - 1 else ENDOFCHAIN)
        sectors.append(bytes(data_bytes).ljust(nsec * 512, b"\x00"))
        return start

    for _ in range(n_fat):
        fat.append(FATSECT)
    sectors.append(None)                       # se llena al final
    dir_start = len(fat)
    for _ in range(n_dir):
        fat.append(0)                          # provisional
    fat[dir_start:dir_start + n_dir] = [dir_start + k + 1 if k < n_dir - 1 else ENDOFCHAIN
                                        for k in range(n_dir)]
    sectors.append(None)
    minifat_start = ENDOFCHAIN
    if n_minifat:
        mf = b"".join(struct.pack("<I", v) for v in minifat)
        mf = mf.ljust(n_minifat * 512, b"\xff")
        minifat_start = alloc(mf, n_minifat)
    if n_mini:
        root.start = alloc(mini, n_mini)
        root.size = len(mini)
    for e in large:
        e.size = len(e.data)
        e.start = alloc(e.data, math.ceil(e.size / 512))

    # Directorio
    d = bytearray()
    for e in entries:
        name16 = e.name.encode("utf-16-le") + b"\x00\x00"
        rec = name16.ljust(64, b"\x00")
        rec += struct.pack("<HBB", len(name16), e.kind, e.color)
        rec += struct.pack("<III", e.left, e.right, e.child)
        rec += b"\x00" * 16 + struct.pack("<I", 0) + b"\x00" * 16
        rec += struct.pack("<IQ", e.start if e.kind != 1 else 0, e.size if e.kind != 1 else 0)
        d += rec
    empty = (b"\x00" * 64 + struct.pack("<HBB", 0, 0, 0) + struct.pack("<III", NOSTREAM, NOSTREAM, NOSTREAM)
             + b"\x00" * 36 + struct.pack("<IQ", 0, 0))
    while len(d) < n_dir * 512:
        d += empty
    sectors[1] = bytes(d)

    fat_bytes = b"".join(struct.pack("<I", v) for v in fat).ljust(n_fat * 512, b"\xff")
    sectors[0] = fat_bytes

    hdr = bytearray()
    hdr += bytes.fromhex("D0CF11E0A1B11AE1") + b"\x00" * 16
    hdr += struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + b"\x00" * 6
    hdr += struct.pack("<IIIIIIIII", 0, n_fat, dir_start, 0, 4096,
                       minifat_start, n_minifat, ENDOFCHAIN, 0)
    difat = [k for k in range(n_fat)] + [FREESECT] * (109 - n_fat)
    hdr += b"".join(struct.pack("<I", v) for v in difat)
    assert len(hdr) == 512
    return bytes(hdr) + b"".join(sectors)


# ----------------------------------------------------------------------
# Proyecto VBA
# ----------------------------------------------------------------------
# Identificador y campos de protección (CMG/DPB/GC) generados por Excel para
# un proyecto sin contraseña; deben ir juntos.
PROJECT_ID = "{9E394C0B-697E-4AEE-9FA6-446F51FB30DC}"
CMG = "41435A5A5E5A5E5A5E5A5E"
DPB = "BCBEA7A2591C5A1C5A1C"
GC = "37352C2BDCDD56DE56DEA9"

WB_GUID = "{00020819-0000-0000-C000-000000000046}"
WS_GUID = "{00020820-0000-0000-C000-000000000046}"


def _rec(rid, payload=b""):
    return struct.pack("<HI", rid, len(payload)) + payload


def build_vba_project(doc_modules, std_modules, dir_header: bytes, cp="cp1252") -> bytes:
    """doc_modules: [(nombre, guid, código)] ; std_modules: [(nombre, código con Attribute VB_Name)]"""
    sources = []
    for name, guid, code in doc_modules:
        hdr = (f'Attribute VB_Name = "{name}"\nAttribute VB_Base = "0{guid}"\n'
               'Attribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\n'
               'Attribute VB_PredeclaredId = True\nAttribute VB_Exposed = True\n'
               'Attribute VB_TemplateDerived = False\nAttribute VB_Customizable = True\n')
        sources.append((name, True, hdr + code))
    for name, code in std_modules:
        sources.append((name, False, code))

    def norm(t):
        t = t.replace("\r\n", "\n").replace("\n", "\r\n")
        if not t.endswith("\r\n"):
            t += "\r\n"
        return t.encode(cp)

    # dir stream
    d = bytearray(dir_header)
    d += _rec(0x000F, struct.pack("<H", len(sources)))
    d += _rec(0x0013, struct.pack("<H", 0xFFFF))
    for name, is_doc, _ in sources:
        nb, nu = name.encode(cp), name.encode("utf-16-le")
        d += _rec(0x0019, nb) + _rec(0x0047, nu)
        d += _rec(0x001A, nb) + _rec(0x0032, nu)
        d += _rec(0x001C) + _rec(0x0048)
        d += _rec(0x0031, struct.pack("<I", 0))
        d += _rec(0x001E, struct.pack("<I", 0))
        d += _rec(0x002C, struct.pack("<H", 0xFFFF))
        d += _rec(0x0022 if is_doc else 0x0021)
        d += _rec(0x002B)
    d += _rec(0x0010)

    # PROJECT stream
    lines = [f'ID="{PROJECT_ID}"']
    for name, is_doc, _ in sources:
        lines.append(f"Document={name}/&H00000000" if is_doc else f"Module={name}")
    lines += ['Name="VBAProject"', 'HelpContextID="0"', 'VersionCompatible32="393222000"',
              f'CMG="{CMG}"', f'DPB="{DPB}"', f'GC="{GC}"', "",
              "[Host Extender Info]",
              "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000", "",
              "[Workspace]"]
    for name, is_doc, _ in sources:
        lines.append(f"{name}=0, 0, 0, 0, C")
    project_txt = ("\r\n".join(lines) + "\r\n").encode(cp)

    # PROJECTwm
    wm = bytearray()
    for name, _, _ in sources:
        wm += name.encode(cp) + b"\x00" + name.encode("utf-16-le") + b"\x00\x00"
    wm += b"\x00\x00"

    root = Entry("Root Entry", 5)
    vba = Entry("VBA", 1)
    vba.children.append(Entry("_VBA_PROJECT", 2, struct.pack("<HHBH", 0x61CC, 0xFFFF, 0, 3)))
    vba.children.append(Entry("dir", 2, compress(bytes(d))))
    for name, _, src in sources:
        vba.children.append(Entry(name, 2, compress(norm(src))))
    root.children += [vba, Entry("PROJECT", 2, project_txt), Entry("PROJECTwm", 2, bytes(wm))]
    return write_cfb(root)


# Cabecera del stream "dir" tomada de un proyecto creado por Excel
# (sistema, idioma, página de códigos 1252 y referencias stdole + Office).
DIR_HEADER = bytes.fromhex(
    "010004000000030000004A000400000003000000020004000000090400001400040000000904"
    "0000030002000000E40404000A00000056424150726F6A656374050000000000400000000000"
    "0600000000003D00000000000700040000000000000008000400000000000000090004000000"
    "5702BE6511000C00000000003C00000000001600060000007374646F6C653E000C0000007300"
    "740064006F006C0065000D00680000005E0000002A5C477B30303032303433302D303030302D"
    "303030302D433030302D3030303030303030303034367D23322E30233023433A5C57696E646F"
    "77735C53797374656D33325C7374646F6C65322E746C62234F4C45204175746F6D6174696F6E"
    "0000000000001600060000004F66666963653E000C0000004F00660066006900630065000D00"
    "9E000000940000002A5C477B32444638443034432D354246412D313031422D424445352D3030"
    "414130303434444535327D23322E30233023433A5C50726F6772616D2046696C65735C436F6D"
    "6D6F6E2046696C65735C4D6963726F736F6674205368617265645C4F464649434531365C4D53"
    "4F2E444C4C234D6963726F736F6674204F66666963652031362E30204F626A656374204C6962"
    "72617279000000000000"
)


# ----------------------------------------------------------------------
# Conversión .xlsx -> .xlsm
# ----------------------------------------------------------------------
import io
import re
import zipfile
from pathlib import Path

CARPETA_VBA = Path(__file__).resolve().parent.parent / "vba"


def _leer(nombre: str) -> str:
    return (CARPETA_VBA / nombre).read_text(encoding="utf-8")


def inyectar_vba(xlsx: bytes) -> bytes:
    """Recibe un .xlsx (con codeName en el libro y en cada hoja) y devuelve el .xlsm con la macro."""
    zin = zipfile.ZipFile(io.BytesIO(xlsx))
    partes = {n: zin.read(n) for n in zin.namelist()}
    wbx = partes["xl/workbook.xml"].decode("utf-8")
    if 'codeName="ThisWorkbook"' not in wbx:
        wbx = re.sub(r"<workbookPr\b", '<workbookPr codeName="ThisWorkbook"', wbx, count=1)
        partes["xl/workbook.xml"] = wbx.encode("utf-8")
    rels = partes["xl/_rels/workbook.xml.rels"].decode("utf-8")

    # codeName de cada hoja, en el orden del libro
    codenames = []
    for rid in re.findall(r'<sheet [^>]*r:id="(rId\d+)"', wbx):
        rel = re.search(rf'<Relationship[^>]*Id="{rid}"[^>]*/>', rels).group(0)
        destino = re.search(r'Target="/?(?:xl/)?([^"]+)"', rel).group(1)
        hoja = partes["xl/" + destino].decode("utf-8")
        codenames.append(re.search(r'<sheetPr[^>]*codeName="([^"]+)"', hoja).group(1))

    docs = [("ThisWorkbook", WB_GUID, _leer("ThisWorkbook_code.txt"))]
    docs += [(c, WS_GUID, _leer("HojaDatos_code.txt") if c == "HojaDatos" else "") for c in codenames]
    partes["xl/vbaProject.bin"] = build_vba_project(docs, [("modMenu", _leer("modMenu.bas"))], DIR_HEADER)

    ids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)]
    rels = rels.replace("</Relationships>",
        f'<Relationship Id="rId{max(ids) + 1}" '
        'Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
        'Target="vbaProject.bin"/></Relationships>')
    partes["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")

    ct = partes["[Content_Types].xml"].decode("utf-8")
    ct = ct.replace("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                    "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
    ct = re.sub(r'<Default Extension="bin"[^>]*/>', "", ct)
    ct = ct.replace("<Default ", '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default ', 1)
    partes["[Content_Types].xml"] = ct.encode("utf-8")

    salida = io.BytesIO()
    with zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", partes.pop("[Content_Types].xml"))
        for nombre, datos in partes.items():
            z.writestr(nombre, datos)
    return salida.getvalue()
