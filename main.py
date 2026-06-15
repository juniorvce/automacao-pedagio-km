
import os, shutil, threading
from pathlib import Path

from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.popup import Popup
from kivy.uix.progressbar import ProgressBar
from kivy.clock import Clock
from kivy.metrics import dp
from kivy.utils import platform

# ── Android: permissões de armazenamento ────────────────────────────
if platform == "android":
    from android.permissions import request_permissions, Permission
    request_permissions([
        Permission.READ_EXTERNAL_STORAGE,
        Permission.WRITE_EXTERNAL_STORAGE,
    ])

# ── Importações de processamento ─────────────────────────────────────
from pdf2image import convert_from_path
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
from openpyxl.utils.units import pixels_to_EMU
from PIL import Image as PILImage
import tempfile

# ═══════════════════════════════════════════════════════════════════
# CONSTANTES DE POSICIONAMENTO (idênticas ao gabarito perfeito)
# ═══════════════════════════════════════════════════════════════════
# Cada slot: (col_from, colOff_from, row_from, rowOff_from,
#              col_to,   colOff_to,   row_to,   rowOff_to)
SLOTS_ROW1 = [
    (1, 198000,  79, 83520,   3, 363960,  99, 69480),   # Slot 1
    (4, 173880,  79, 83520,   6, 469440,  99, 69480),   # Slot 2
    (7, 149760,  79, 83520,   8, 531000,  99, 69480),   # Slot 3
    (9, 210960,  79, 83520,  10, 1239480, 99, 69480),   # Slot 4
]
SLOTS_ROW2 = [
    (1, 198000, 100, 59760,   3, 363960, 120, 45720),   # Slot 5
    (4, 173880, 100, 59760,   6, 469440, 120, 45720),   # Slot 6
    (7, 149760, 100, 59760,   8, 531000, 120, 45720),   # Slot 7
    (9, 210960, 100, 59760,  10, 1239480,120, 45720),   # Slot 8
]
ALL_SLOTS = SLOTS_ROW1 + SLOTS_ROW2   # 8 slots

# ═══════════════════════════════════════════════════════════════════
# PROCESSAMENTO CORE
# ═══════════════════════════════════════════════════════════════════
def pdf_to_png(pdf_path: str, tmp_dir: str) -> str:
    pages = convert_from_path(pdf_path, dpi=200, first_page=1, last_page=1)
    out = os.path.join(tmp_dir, Path(pdf_path).stem + ".png")
    pages[0].save(out, "PNG")
    return out

def fit_image_in_slot(img_path: str, slot_w_px: int, slot_h_px: int, tmp_dir: str) -> str:
    img = PILImage.open(img_path).convert("RGB")
    img.thumbnail((slot_w_px, slot_h_px), PILImage.LANCZOS)
    canvas = PILImage.new("RGB", (slot_w_px, slot_h_px), (255, 255, 255))
    ox = (slot_w_px - img.width) // 2
    oy = (slot_h_px - img.height) // 2
    canvas.paste(img, (ox, oy))
    out = os.path.join(tmp_dir, "fit_" + os.path.basename(img_path))
    canvas.save(out, "PNG")
    return out

def emu_to_px(emu: int) -> int:
    return int(emu / 9525)

def insert_images(xlsx_path: str, png_list: list, output_path: str):
    wb = load_workbook(xlsx_path)
    ws = wb.active
    ws._images = []

    for i, png_path in enumerate(png_list):
        if i >= len(ALL_SLOTS):
            break
        cf, cof, rf, rof, ct, cot, rt, rot = ALL_SLOTS[i]
        slot_w = emu_to_px(cot - cof + 100000)
        slot_h = emu_to_px(rot - rof + 100000)

        with tempfile.TemporaryDirectory() as td:
            fitted = fit_image_in_slot(png_path, max(slot_w,100), max(slot_h,100), td)
            img = XLImage(fitted)
            anchor = TwoCellAnchor()
            anchor.editAs = "twoCell"
            anchor._from = AnchorMarker(col=cf, colOff=cof, row=rf, rowOff=rof)
            anchor.to    = AnchorMarker(col=ct, colOff=cot, row=rt, rowOff=rot)
            img.anchor   = anchor
            ws.add_image(img)

    wb.save(output_path)

# ═══════════════════════════════════════════════════════════════════
# UI KIVY
# ═══════════════════════════════════════════════════════════════════
class MainLayout(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(orientation="vertical", padding=dp(16), spacing=dp(12), **kwargs)
        self.pdf_files = []
        self.template_path = ""
        self._build_ui()

    def _build_ui(self):
        self.add_widget(Label(
            text="[b]🚛 Automação Pedágio - KM[/b]\n[size=14]Gerador de Planilha[/size]",
            markup=True, size_hint_y=None, height=dp(70),
            font_size=dp(22), halign="center"
        ))

        btn_tpl = Button(text="📂 Selecionar Template (.xlsx)",
                         size_hint_y=None, height=dp(52),
                         background_color=(0.2, 0.5, 0.9, 1))
        btn_tpl.bind(on_press=self.pick_template)
        self.add_widget(btn_tpl)

        self.lbl_template = Label(text="Nenhum template selecionado",
                                  size_hint_y=None, height=dp(28),
                                  font_size=dp(12), color=(0.6,0.6,0.6,1))
        self.add_widget(self.lbl_template)

        btn_pdf = Button(text="📄 Selecionar PDFs (até 8)",
                         size_hint_y=None, height=dp(52),
                         background_color=(0.2, 0.7, 0.4, 1))
        btn_pdf.bind(on_press=self.pick_pdfs)
        self.add_widget(btn_pdf)

        scroll = ScrollView(size_hint_y=0.35)
        self.pdf_grid = GridLayout(cols=1, size_hint_y=None, spacing=dp(4))
        self.pdf_grid.bind(minimum_height=self.pdf_grid.setter("height"))
        scroll.add_widget(self.pdf_grid)
        self.add_widget(scroll)

        self.progress = ProgressBar(max=100, value=0,
                                    size_hint_y=None, height=dp(24))
        self.add_widget(self.progress)

        self.lbl_status = Label(text="Aguardando...",
                                size_hint_y=None, height=dp(32),
                                font_size=dp(13))
        self.add_widget(self.lbl_status)

        btn_gen = Button(text="⚡ Gerar Planilha",
                         size_hint_y=None, height=dp(60),
                         background_color=(0.9, 0.4, 0.1, 1),
                         font_size=dp(18))
        btn_gen.bind(on_press=self.gerar_planilha)
        self.add_widget(btn_gen)

    def pick_template(self, *_):
        self._show_path_input("Template (.xlsx)", self._set_template)

    def pick_pdfs(self, *_):
        self._show_path_input("Pasta com PDFs", self._load_pdfs_from_folder)

    def _show_path_input(self, title, callback):
        from kivy.uix.textinput import TextInput
        content = BoxLayout(orientation="vertical", spacing=dp(8), padding=dp(12))
        ti = TextInput(hint_text="Cole o caminho aqui...", multiline=False,
                       size_hint_y=None, height=dp(44))
        content.add_widget(Label(text=title))
        content.add_widget(ti)
        btn = Button(text="OK", size_hint_y=None, height=dp(44))
        content.add_widget(btn)
        popup = Popup(title=title, content=content,
                      size_hint=(0.9, 0.5))
        def _ok(*_):
            popup.dismiss()
            callback(ti.text.strip())
        btn.bind(on_press=_ok)
        popup.open()

    def _set_template(self, path):
        if os.path.isfile(path):
            self.template_path = path
            self.lbl_template.text = f"✅ {os.path.basename(path)}"
        else:
            self.lbl_template.text = "❌ Arquivo não encontrado"

    def _load_pdfs_from_folder(self, folder):
        if os.path.isdir(folder):
            pdfs = sorted([
                os.path.join(folder, f)
                for f in os.listdir(folder)
                if f.lower().endswith(".pdf")
            ])[:8]
            self.pdf_files = pdfs
            self.pdf_grid.clear_widgets()
            for p in pdfs:
                self.pdf_grid.add_widget(
                    Label(text=f"  📄 {os.path.basename(p)}",
                          size_hint_y=None, height=dp(28),
                          font_size=dp(11), halign="left")
                )
            self.lbl_status.text = f"{len(pdfs)} PDF(s) carregados"
        else:
            self.lbl_status.text = "❌ Pasta não encontrada"

    def gerar_planilha(self, *_):
        if not self.template_path:
            self.lbl_status.text = "⚠️  Selecione o template primeiro!"
            return
        if not self.pdf_files:
            self.lbl_status.text = "⚠️  Selecione os PDFs primeiro!"
            return

        self.lbl_status.text = "⏳ Processando..."
        self.progress.value = 0
        t = threading.Thread(target=self._process, daemon=True)
        t.start()

    def _process(self):
        try:
            total = len(self.pdf_files)
            png_list = []

            with tempfile.TemporaryDirectory() as tmp:
                for i, pdf in enumerate(self.pdf_files):
                    Clock.schedule_once(lambda dt, i=i: self._upd(
                        int((i/total)*70), f"Convertendo PDF {i+1}/{total}..."))
                    png = pdf_to_png(pdf, tmp)
                    png_list.append(png)

                Clock.schedule_once(lambda dt: self._upd(80, "Inserindo imagens na planilha..."))

                out_dir = os.path.dirname(self.template_path)
                out_path = os.path.join(out_dir, "Planilha_Gerada.xlsx")
                insert_images(self.template_path, png_list, out_path)

            Clock.schedule_once(lambda dt: self._done(out_path))

        except Exception as e:
            Clock.schedule_once(lambda dt: self._error(str(e)))

    def _upd(self, val, txt):
        self.progress.value = val
        self.lbl_status.text = txt

    def _done(self, path):
        self.progress.value = 100
        self.lbl_status.text = f"✅ Salvo em:\n{path}"

    def _error(self, msg):
        self.lbl_status.text = f"❌ Erro: {msg}"

class PedagioApp(App):
    def build(self):
        self.title = "Automação Pedágio - KM"
        return MainLayout()

if __name__ == "__main__":
    PedagioApp().run()
