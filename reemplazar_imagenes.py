"""
Aplicación para reemplazar imágenes placeholder en plantillas Word (.docx)
Las capturas deben tener el mismo nombre que la descripción/título del placeholder.

Uso:
  1. Abre la plantilla .docx con tus imágenes placeholder.
  2. La app detecta los placeholders y muestra sus nombres.
  3. Selecciona una carpeta con tus capturas (el nombre del archivo debe
     coincidir con el nombre del placeholder, sin importar la extensión)
  4. Activa "También generar PDF" si lo deseas.
  5. Pulsa "Generar documento" y elige dónde guardarlo.

Requisitos:
  pip install python-docx pillow lxml
  Conversión a PDF: Microsoft Word instalado (Windows) o LibreOffice (Linux/macOS)
"""

import os
import sys
import zipfile
import shutil
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

# ── Dependencias opcionales ──────────────────────────────────────────────────
try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.shared import Inches
    from lxml import etree
    HAVE_DOCX = True
except ImportError:
    HAVE_DOCX = False

try:
    from PIL import Image, ImageTk
    HAVE_PIL = True
except ImportError:
    HAVE_PIL = False

# ── Constantes XML ───────────────────────────────────────────────────────────
NS = {
    'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'mc':  'http://schemas.openxmlformats.org/markup-compatibility/2006',
}
CT_IMAGE = {
    '.png':  'image/png',
    '.jpg':  'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.gif':  'image/gif',
    '.bmp':  'image/bmp',
    '.tiff': 'image/tiff',
    '.tif':  'image/tiff',
    '.webp': 'image/webp',
}

# ── Paleta de colores ────────────────────────────────────────────────────────
CLR_BG        = "#F0F4F8"
CLR_HEADER    = "#1E3A5F"
CLR_ACCENT    = "#2563EB"
CLR_ACCENT2   = "#1D4ED8"
CLR_SUCCESS   = "#16A34A"
CLR_SUCCESS2  = "#15803D"
CLR_BTN_FG    = "#FFFFFF"
CLR_FRAME_BG  = "#FFFFFF"
CLR_BORDER    = "#CBD5E1"
CLR_TEXT      = "#1E293B"
CLR_SUBTEXT   = "#64748B"
FOUND_COLOR   = "#DCFCE7"
MISSING_COLOR = "#FEF3C7"
PLACEHOLDER_COLOR = "#EFF6FF"


# ════════════════════════════════════════════════════════════════════════════
#  Detección y conversión a PDF
# ════════════════════════════════════════════════════════════════════════════

def _find_word_exe() -> str | None:
    """
    Localiza WINWORD.EXE vía registro de Windows o rutas conocidas.
    Devuelve la ruta si la encuentra, None si no.
    """
    if sys.platform != 'win32':
        return None
    # ── Registro de Windows ──
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE'
        )
        path_val, _ = winreg.QueryValueEx(key, '')
        winreg.CloseKey(key)
        if path_val and Path(path_val).exists():
            return path_val
    except Exception:
        pass
    # ── Rutas comunes Office 365 / 2019 / 2016 / 2013 ──
    for p in [
        r'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE',
        r'C:\Program Files\Microsoft Office\root\Office15\WINWORD.EXE',
        r'C:\Program Files\Microsoft Office\Office16\WINWORD.EXE',
        r'C:\Program Files\Microsoft Office\Office15\WINWORD.EXE',
        r'C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE',
        r'C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE',
    ]:
        if Path(p).exists():
            return p
    return None


def convert_to_pdf(docx_path: str, pdf_path: str) -> str:
    """
    Convierte un .docx a .pdf.
    - Intenta primero Microsoft Word via COM (Windows).
    - Si no está disponible, usa LibreOffice.
    - Si ya existe un PDF en pdf_path, lo sobreescribe.
    Devuelve el método usado ('word' o 'libreoffice').
    Lanza RuntimeError si ninguno está disponible.
    """
    docx_path = str(Path(docx_path).resolve())
    pdf_path  = str(Path(pdf_path).resolve())

    # ── FALLO 2 FIX: Eliminar PDF existente antes de convertir ───────────
    if Path(pdf_path).exists():
        try:
            Path(pdf_path).unlink()
        except Exception as e:
            raise RuntimeError(
                f"No se puede sobreescribir el PDF existente:\n{pdf_path}\n\n"
                f"Cierra el archivo si está abierto y vuelve a intentarlo.\n({e})"
            )

    # ── Intento 1: Microsoft Word via COM (solo Windows) ──────────────────
    # FALLO 1 FIX: comprobamos que Word existe ANTES de intentar COM,
    # así el error es explícito y no cae silenciosamente a LibreOffice.
    if sys.platform == 'win32' and _find_word_exe() is not None:
        try:
            import comtypes.client
            word = comtypes.client.CreateObject('Word.Application')
            word.Visible = False
            try:
                doc = word.Documents.Open(docx_path)
                doc.SaveAs(pdf_path, FileFormat=17)   # wdFormatPDF = 17
                doc.Close(False)
            finally:
                word.Quit()
            if Path(pdf_path).exists():
                return 'word'
        except Exception:
            pass  # COM falló a pesar de encontrar Word → intentar LibreOffice

    # ── Intento 2: LibreOffice ─────────────────────────────────────────────
    import subprocess
    out_dir = str(Path(pdf_path).parent)

    soffice = None
    for c in [
        'libreoffice', 'soffice',
        r'C:\Program Files\LibreOffice\program\soffice.exe',
        r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
        '/usr/bin/soffice', '/usr/bin/libreoffice',
        '/Applications/LibreOffice.app/Contents/MacOS/soffice',
    ]:
        if shutil.which(c) or Path(c).exists():
            soffice = c
            break

    if soffice is None:
        raise RuntimeError(
            "No se encontró Microsoft Word ni LibreOffice en este equipo.\n\n"
            "Para generar PDF instala una de estas aplicaciones:\n"
            "  • Microsoft Word (Office 365, 2019, 2016…)\n"
            "  • LibreOffice (gratuito): https://www.libreoffice.org"
        )

    result = subprocess.run(
        [soffice, '--headless', '--convert-to', 'pdf',
         '--outdir', out_dir, docx_path],
        capture_output=True, text=True, timeout=60
    )
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice falló:\n{result.stderr}")

    # LibreOffice guarda con el mismo stem del .docx en out_dir
    lo_output = Path(out_dir) / (Path(docx_path).stem + '.pdf')
    if str(lo_output) != pdf_path and lo_output.exists():
        shutil.move(str(lo_output), pdf_path)

    return 'libreoffice'


# ════════════════════════════════════════════════════════════════════════════
#  Lógica de reemplazo
# ════════════════════════════════════════════════════════════════════════════

def _get_drawing_title(drawing_elem):
    """Extrae el título/descripción alt de un elemento <w:drawing>."""
    for docPr in drawing_elem.iter(
        '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr'
    ):
        title = docPr.get('title') or docPr.get('descr') or docPr.get('name') or ''
        if title:
            return title.strip()
    return ''


def detect_placeholders(docx_path: str) -> list[dict]:
    """
    Recorre el docx y devuelve la lista de placeholders de imagen encontrados.
    Cada elemento: {'name': str, 'rId': str, 'media': str}
    """
    placeholders = []
    seen = set()

    with zipfile.ZipFile(docx_path, 'r') as z:
        rels = {}
        if 'word/_rels/document.xml.rels' in z.namelist():
            rels_tree = etree.fromstring(z.read('word/_rels/document.xml.rels'))
            for rel in rels_tree:
                rid = rel.get('Id', '')
                typ = rel.get('Type', '')
                tgt = rel.get('Target', '')
                if 'image' in typ:
                    rels[rid] = tgt

        doc_xml = z.read('word/document.xml')
        tree    = etree.fromstring(doc_xml)

        for drawing in tree.iter(
            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing'
        ):
            title = _get_drawing_title(drawing)
            rid = None
            for blip in drawing.iter(
                '{http://schemas.openxmlformats.org/drawingml/2006/main}blip'
            ):
                rid = blip.get(
                    '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'
                )
                break

            if rid and title and title not in seen:
                seen.add(title)
                placeholders.append({
                    'name':  title,
                    'rId':   rid,
                    'media': rels.get(rid, ''),
                })

    return placeholders


def replace_images(docx_path: str, mappings: dict, output_path: str) -> tuple[list, list]:
    """
    Reemplaza imágenes placeholder en el docx.
    mappings: {placeholder_name: ruta_captura}
    Devuelve (reemplazados, omitidos)
    """
    replaced = []
    skipped  = []

    tmp_dir = Path(output_path).parent / '_docx_tmp'
    if tmp_dir.exists():
        shutil.rmtree(tmp_dir)
    tmp_dir.mkdir(parents=True)

    with zipfile.ZipFile(docx_path, 'r') as z:
        z.extractall(tmp_dir)

    rels_path = tmp_dir / 'word' / '_rels' / 'document.xml.rels'
    rels_tree = etree.parse(str(rels_path))
    rels_root = rels_tree.getroot()

    rid_to_target = {}
    for rel in rels_root:
        rid = rel.get('Id', '')
        typ = rel.get('Type', '')
        tgt = rel.get('Target', '')
        if 'image' in typ:
            rid_to_target[rid] = tgt

    doc_path = tmp_dir / 'word' / 'document.xml'
    doc_tree = etree.parse(str(doc_path))
    doc_root = doc_tree.getroot()

    ct_path = tmp_dir / '[Content_Types].xml'
    ct_tree = etree.parse(str(ct_path))
    ct_root = ct_tree.getroot()
    CT_NS   = 'http://schemas.openxmlformats.org/package/2006/content-types'

    existing_extensions = {
        e.get('Extension', '').lower()
        for e in ct_root.findall(f'{{{CT_NS}}}Default')
    }

    max_rid = max(
        (int(re.sub(r'\D', '', r.get('Id', 'rId0'))) for r in rels_root),
        default=0
    )

    for drawing in doc_root.iter(
        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing'
    ):
        title = _get_drawing_title(drawing)
        if not title or title not in mappings:
            continue

        capture_path = Path(mappings[title])
        if not capture_path.exists():
            skipped.append(title)
            continue

        ext     = capture_path.suffix.lower()
        ct_mime = CT_IMAGE.get(ext, 'image/png')

        new_media_name = f"capture_{title.replace(' ', '_')}{ext}"
        media_dest     = tmp_dir / 'word' / 'media' / new_media_name
        shutil.copy2(capture_path, media_dest)

        max_rid += 1
        new_rid  = f'rId{max_rid}'
        REL_NS   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        etree.SubElement(rels_root, 'Relationship', {
            'Id':     new_rid,
            'Type':   f'{REL_NS}/image',
            'Target': f'media/{new_media_name}',
        })

        ext_clean = ext.lstrip('.')
        if ext_clean not in existing_extensions:
            etree.SubElement(ct_root, f'{{{CT_NS}}}Default', {
                'Extension':   ext_clean,
                'ContentType': ct_mime,
            })
            existing_extensions.add(ext_clean)

        for blip in drawing.iter(
            '{http://schemas.openxmlformats.org/drawingml/2006/main}blip'
        ):
            blip.set(
                '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed',
                new_rid
            )

        replaced.append(title)

    doc_tree.write(str(doc_path),   xml_declaration=True, encoding='UTF-8', standalone=True)
    rels_tree.write(str(rels_path), xml_declaration=True, encoding='UTF-8', standalone=True)
    ct_tree.write(str(ct_path),     xml_declaration=True, encoding='UTF-8', standalone=True)

    if Path(output_path).exists():
        os.remove(output_path)
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for fpath in tmp_dir.rglob('*'):
            if fpath.is_file():
                zout.write(fpath, fpath.relative_to(tmp_dir))

    shutil.rmtree(tmp_dir)
    return replaced, skipped


# ════════════════════════════════════════════════════════════════════════════
#  Interfaz gráfica — Widgets personalizados
# ════════════════════════════════════════════════════════════════════════════

class RoundedButton(tk.Canvas):
    """Botón con esquinas redondeadas sobre Canvas."""

    def __init__(self, parent, text, command,
                 bg=CLR_ACCENT, fg=CLR_BTN_FG, hover_bg=CLR_ACCENT2,
                 font=("Segoe UI", 9, "bold"),
                 width=130, height=34, radius=8, state='normal', **kwargs):
        super().__init__(parent, width=width, height=height,
                         bg=parent.cget('bg'), highlightthickness=0, **kwargs)
        self._bg       = bg
        self._hover_bg = hover_bg
        self._fg       = fg
        self._text     = text
        self._font     = font
        self._command  = command
        self._radius   = radius
        self._w        = width
        self._h        = height
        self._state    = state
        self._draw()
        self.bind('<Enter>',          self._on_enter)
        self.bind('<Leave>',          self._on_leave)
        self.bind('<ButtonRelease-1>', self._on_click)

    def _rounded_rect(self, color):
        self.delete('all')
        r, w, h = self._radius, self._w, self._h
        for x0, y0, x1, y1, start in [
            (0, 0, 2*r, 2*r, 90),
            (w-2*r, 0, w, 2*r, 0),
            (0, h-2*r, 2*r, h, 180),
            (w-2*r, h-2*r, w, h, 270),
        ]:
            self.create_arc(x0, y0, x1, y1,
                            start=start, extent=90, fill=color, outline=color)
        self.create_rectangle(r, 0,   w-r, h,   fill=color, outline=color)
        self.create_rectangle(0, r,   w,   h-r, fill=color, outline=color)

    def _draw(self, color=None):
        c  = color or (self._bg if self._state == 'normal' else CLR_BORDER)
        fg = self._fg if self._state == 'normal' else CLR_SUBTEXT
        self._rounded_rect(c)
        self.create_text(self._w // 2, self._h // 2,
                         text=self._text, fill=fg,
                         font=self._font, anchor='center')

    def _on_enter(self, _):
        if self._state == 'normal':
            self._draw(self._hover_bg)
            self.config(cursor='hand2')

    def _on_leave(self, _):
        self._draw()

    def _on_click(self, _):
        if self._state == 'normal' and self._command:
            self._command()

    def config_state(self, state):
        self._state = state
        self._draw()


# ════════════════════════════════════════════════════════════════════════════
#  Ventana principal
# ════════════════════════════════════════════════════════════════════════════

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Reemplazador de Imágenes · PKD Generator")
        self.resizable(True, True)
        self.minsize(760, 580)
        self.configure(bg=CLR_BG)

        self.template_path = tk.StringVar()
        self.captures_dir  = tk.StringVar()
        self.placeholders  = []
        self.capture_files = {}
        self.generate_pdf  = tk.BooleanVar(value=False)

        self._build_ui()
        self._center()

    # ── Construcción UI ──────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Cabecera ──────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=CLR_HEADER, height=64)
        hdr.pack(fill='x')
        hdr.pack_propagate(False)

        tk.Label(hdr, text="📋  Reemplazador de Imágenes en Word",
                 bg=CLR_HEADER, fg="white",
                 font=("Segoe UI", 14, "bold")).pack(
            side='left', padx=20, pady=18)

        tk.Label(hdr, text="PKD Generator v1.1",
                 bg=CLR_HEADER, fg="#94A3B8",
                 font=("Segoe UI", 8)).pack(
            side='right', padx=20, pady=24)

        # ── Área de contenido ─────────────────────────────────────────────
        main = tk.Frame(self, bg=CLR_BG)
        main.pack(fill='both', expand=True, padx=16, pady=12)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(2, weight=1)

        # ── Sección 1: Plantilla ──
        self._section_file(
            main, row=0,
            title="1 · Plantilla Word (.docx)",
            var=self.template_path,
            cmd=self._browse_template,
        )

        # ── Sección 2: Carpeta capturas ──
        self._section_file(
            main, row=1,
            title="2 · Carpeta con capturas",
            var=self.captures_dir,
            cmd=self._browse_captures,
        )

        # ── Tabla de placeholders ──
        self._build_table(main, row=2)

        # ── Leyenda ──
        self._build_legend(main, row=3)

        # ── Barra inferior (opciones + botón) ──
        self._build_bottom_bar(main, row=4)

        # ── Barra de estado ──
        status_bar = tk.Frame(self, bg="#1E293B", height=28)
        status_bar.pack(fill='x', side='bottom')
        status_bar.pack_propagate(False)

        self.status = tk.StringVar(
            value="  Selecciona una plantilla Word para comenzar."
        )
        tk.Label(status_bar, textvariable=self.status,
                 bg="#1E293B", fg="#94A3B8",
                 font=("Segoe UI", 8), anchor='w').pack(
            side='left', padx=8, fill='both')

    def _section_file(self, parent, row, title, var, cmd):
        frm = tk.LabelFrame(parent, text=f"  {title}  ",
                            bg=CLR_FRAME_BG,
                            font=("Segoe UI", 9, "bold"),
                            fg=CLR_HEADER, bd=1, relief='solid')
        frm.grid(row=row, column=0, sticky='ew', pady=(0, 6))

        tk.Entry(frm, textvariable=var, state='readonly',
                 width=62, font=("Segoe UI", 9),
                 readonlybackground=CLR_FRAME_BG,
                 fg=CLR_TEXT, relief='flat').pack(
            side='left', padx=(12, 6), pady=8, fill='x', expand=True)

        RoundedButton(frm, text="Examinar…", command=cmd,
                      width=110, height=30,
                      font=("Segoe UI", 9, "bold")).pack(
            side='right', padx=10, pady=8)

    def _build_table(self, parent, row):
        frm = tk.LabelFrame(parent,
                            text="  3 · Placeholders detectados  ",
                            bg=CLR_FRAME_BG,
                            font=("Segoe UI", 9, "bold"),
                            fg=CLR_HEADER, bd=1, relief='solid')
        frm.grid(row=row, column=0, sticky='nsew', pady=(0, 4))

        style = ttk.Style()
        style.theme_use('clam')
        style.configure("P.Treeview",
                         background=CLR_FRAME_BG,
                         foreground=CLR_TEXT,
                         rowheight=26,
                         fieldbackground=CLR_FRAME_BG,
                         font=("Segoe UI", 9),
                         borderwidth=0)
        style.configure("P.Treeview.Heading",
                         background=CLR_HEADER,
                         foreground="white",
                         font=("Segoe UI", 9, "bold"),
                         relief='flat')
        style.map("P.Treeview",
                  background=[('selected', CLR_ACCENT)])

        cols = ('Placeholder', 'Captura encontrada', 'Estado')
        self.tree = ttk.Treeview(frm, columns=cols,
                                 show='headings', height=10,
                                 style="P.Treeview")
        for c in cols:
            self.tree.heading(c, text=c)
        self.tree.column('Placeholder',        width=220, anchor='w')
        self.tree.column('Captura encontrada', width=280, anchor='w')
        self.tree.column('Estado',             width=100, anchor='center')

        self.tree.tag_configure('found',   background=FOUND_COLOR)
        self.tree.tag_configure('missing', background=MISSING_COLOR)
        self.tree.tag_configure('empty',   background=PLACEHOLDER_COLOR)

        sb = ttk.Scrollbar(frm, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True, padx=6, pady=8)
        sb.pack(side='right', fill='y', pady=8, padx=(0, 4))

    def _build_legend(self, parent, row):
        leg = tk.Frame(parent, bg=CLR_BG)
        leg.grid(row=row, column=0, sticky='w', pady=(0, 6))

        for color, label in [
            (FOUND_COLOR,       "✔  Captura encontrada"),
            (MISSING_COLOR,     "⚠  Sin captura"),
            (PLACEHOLDER_COLOR, "—  Placeholder vacío"),
        ]:
            dot = tk.Frame(leg, bg=color, width=14, height=14,
                           relief='solid', bd=1)
            dot.pack(side='left', padx=(0, 4))
            tk.Label(leg, text=label, bg=CLR_BG,
                     font=("Segoe UI", 8), fg=CLR_SUBTEXT).pack(
                side='left', padx=(0, 18))

    def _build_bottom_bar(self, parent, row):
        bar = tk.Frame(parent, bg=CLR_BG)
        bar.grid(row=row, column=0, sticky='ew', pady=(4, 0))
        bar.columnconfigure(0, weight=1)

        # Checkbox PDF
        pdf_box = tk.Frame(bar, bg=CLR_FRAME_BG, relief='solid', bd=1)
        pdf_box.grid(row=0, column=0, sticky='w')
        tk.Checkbutton(pdf_box,
                       text="  📄  También generar PDF (mismo nombre y carpeta)",
                       variable=self.generate_pdf,
                       bg=CLR_FRAME_BG, fg=CLR_TEXT,
                       activebackground=CLR_FRAME_BG,
                       selectcolor=CLR_FRAME_BG,
                       font=("Segoe UI", 9),
                       cursor='hand2').pack(padx=10, pady=8)

        # Botón Generar
        self.btn_gen = RoundedButton(
            bar,
            text="⚙  Generar documento",
            command=self._generate,
            bg=CLR_SUCCESS, hover_bg=CLR_SUCCESS2,
            width=180, height=38,
            font=("Segoe UI", 10, "bold"),
            state='disabled',
        )
        self.btn_gen.grid(row=0, column=1, sticky='e')

    # ── Acciones ────────────────────────────────────────────────────────────

    def _browse_template(self):
        path = filedialog.askopenfilename(
            title="Seleccionar plantilla Word",
            filetypes=[("Documentos Word", "*.docx"),
                       ("Todos los archivos", "*.*")]
        )
        if not path:
            return
        self.template_path.set(path)
        self._load_placeholders(path)

    def _browse_captures(self):
        path = filedialog.askdirectory(title="Carpeta con capturas")
        if not path:
            return
        self.captures_dir.set(path)
        self._load_capture_files(path)
        self._refresh_table()

    def _load_placeholders(self, path):
        self._set_status("🔍  Analizando plantilla…")
        self.update()
        try:
            self.placeholders = detect_placeholders(path)
        except Exception as e:
            messagebox.showerror("Error",
                                 f"No se pudo leer la plantilla:\n{e}")
            self.placeholders = []

        if not self.placeholders:
            self._set_status(
                "⚠  No se encontraron placeholders con texto alternativo."
            )
        else:
            self._set_status(
                f"✔  {len(self.placeholders)} placeholder(s) detectados."
            )
        self._refresh_table()

    def _load_capture_files(self, folder):
        self.capture_files = {}
        for f in Path(folder).iterdir():
            if f.suffix.lower() in CT_IMAGE and f.is_file():
                self.capture_files[f.stem.lower()] = str(f)

    def _refresh_table(self):
        self.tree.delete(*self.tree.get_children())
        for ph in self.placeholders:
            name  = ph['name']
            match = self.capture_files.get(name.lower(), '')
            if match:
                tag, estado, label = 'found', '✔  OK', Path(match).name
            else:
                tag, estado, label = 'missing', '⚠  Falta', '(sin archivo)'
            self.tree.insert('', 'end',
                             values=(name, label, estado), tags=(tag,))

        has_match = any(
            self.capture_files.get(p['name'].lower())
            for p in self.placeholders
        )
        self.btn_gen.config_state(
            'normal' if (self.template_path.get() and has_match) else 'disabled'
        )

    def _generate(self):
        if not self.template_path.get():
            messagebox.showwarning("Aviso",
                                   "Selecciona primero una plantilla.")
            return

        output_path = filedialog.asksaveasfilename(
            title="Guardar documento generado",
            defaultextension=".docx",
            filetypes=[("Documentos Word", "*.docx")]
        )
        if not output_path:
            return
        output_path = str(Path(output_path).with_suffix('.docx'))

        mappings = {
            ph['name']: self.capture_files[ph['name'].lower()]
            for ph in self.placeholders
            if ph['name'].lower() in self.capture_files
        }

        self._set_status("⚙  Generando documento…")
        self.btn_gen.config_state('disabled')
        self.update()

        try:
            replaced, skipped = replace_images(
                self.template_path.get(), mappings, output_path
            )
        except Exception as e:
            messagebox.showerror("Error",
                                 f"No se pudo generar el documento:\n{e}")
            self._set_status("❌  Error al generar.")
            self.btn_gen.config_state('normal')
            return

        msg  = f"✅  Documento guardado en:\n{output_path}\n\n"
        msg += f"Imágenes reemplazadas: {len(replaced)}\n"
        if replaced:
            msg += "  • " + "\n  • ".join(replaced) + "\n"
        if skipped:
            msg += f"\nNo encontrados ({len(skipped)}): " + ", ".join(skipped)

        # ── Conversión a PDF ──────────────────────────────────────────────
        if self.generate_pdf.get():
            pdf_path = str(Path(output_path).with_suffix('.pdf'))
            self._set_status("📄  Convirtiendo a PDF…")
            self.update()
            try:
                method = convert_to_pdf(output_path, pdf_path)
                label  = "Microsoft Word" if method == 'word' else "LibreOffice"
                msg   += f"\n\n📄  PDF generado con {label}:\n{pdf_path}"
                self._set_status(
                    f"✔  Documentos generados · "
                    f"{len(replaced)} imagen(es) · DOCX + PDF"
                )
            except Exception as e:
                msg += f"\n\n⚠  No se pudo generar el PDF:\n{e}"
                self._set_status("✔  DOCX generado · ⚠  PDF fallido")
        else:
            self._set_status(
                f"✔  Documento generado · "
                f"{len(replaced)} imagen(es) reemplazada(s)."
            )

        self.btn_gen.config_state('normal')
        messagebox.showinfo("¡Listo!", msg)

    # ── Utilidades ───────────────────────────────────────────────────────────

    def _set_status(self, msg: str):
        self.status.set(f"  {msg}")

    def _center(self):
        self.update_idletasks()
        w, h = 780, 600
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")


# ════════════════════════════════════════════════════════════════════════════

def check_deps():
    missing = []
    if not HAVE_DOCX:
        missing.append("python-docx  →  pip install python-docx")
    try:
        from lxml import etree
    except ImportError:
        missing.append("lxml         →  pip install lxml")
    return missing


if __name__ == '__main__':
    missing = check_deps()
    if missing:
        print("Faltan dependencias:\n  " + "\n  ".join(missing))
        sys.exit(1)

    app = App()
    app.mainloop()
