# -*- coding: utf-8 -*-
"""
isogd.py — QGIS/NextGIS плагин формирования PDF-отчёта ИСОГД
Совместим с Python 3.8, QGIS 3.32 (NextGIS)

Зависимости (лежать внутри папки плагина):
- extlibs/fpdf2 (пакет fpdf)
- extlibs/PIL (Pillow — опционально: для вставки изображений)
- fonts/NotoSans-*.ttf (кириллица)
"""

import os
import re
import sys
import json
import datetime as dt

# --- extlibs в sys.path до импортов внешних пакетов ---
_EXTLIBS = os.path.join(os.path.dirname(__file__), 'extlibs')
if os.path.isdir(_EXTLIBS) and _EXTLIBS not in sys.path:
    sys.path.insert(0, _EXTLIBS)
# ------------------------------------------------------

from fpdf import FPDF  # fpdf2

# Qt / QGIS
from qgis.PyQt.QtCore import QSettings, QTranslator, QCoreApplication, Qt
from qgis.PyQt.QtGui import QIcon
from qgis.PyQt.QtWidgets import QAction, QMessageBox, QProgressDialog, QFileDialog
from qgis.core import (
    QgsVectorLayer, QgsFeatureRequest,
    QgsSpatialIndex, QgsMessageLog, Qgis,
    QgsWkbTypes
)
from qgis.utils import iface

from .resources import *
from .isogd_dockwidget import InfoDockWidget


# ========================= ЛОГ =========================
def log_info(msg: str):
    QgsMessageLog.logMessage(str(msg), "ISOGD", Qgis.Info)

def log_warn(msg: str):
    QgsMessageLog.logMessage(str(msg), "ISOGD", Qgis.Warning)

def log_err(msg: str):
    QgsMessageLog.logMessage(str(msg), "ISOGD", Qgis.Critical)


# ========== ПРАВИЛА ВЫБОРА КОЛОНОК (JSON) ============
def load_column_rules(plugin_dir: str):
    cfg_path = os.path.join(plugin_dir, "columns_config.json")
    if not os.path.exists(cfg_path):
        log_warn("Не найден columns_config.json — используем правило по умолчанию только с %пересеч")
        return [], ["%пересеч"]
    try:
        with open(cfg_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        rules = data.get("rules", [])
        default_rule = data.get("default", ["%пересеч"])
        valid_rules = []
        for r in rules:
            if isinstance(r.get("pattern"), str) and isinstance(r.get("columns"), list):
                valid_rules.append({"pattern": r["pattern"], "columns": r["columns"]})
        return valid_rules, default_rule
    except Exception as e:
        log_err(f"Ошибка чтения columns_config.json: {e}")
        return [], ["%пересеч"]


def pick_rule_for_layer(rules, default_rule, path_layer: str):
    for r in rules:
        try:
            if re.search(r["pattern"], path_layer, flags=re.IGNORECASE):
                return r["columns"]
        except re.error:
            pass
    return default_rule


def resolve_columns(fields_intersect, selector_list):
    """
    headers  — имена полей пересекаемого слоя (строго как в слое),
    idxs     — индексы этих полей,
    add_percent — добавлять ли столбец '% пересеч'.
    Если idxs пуст — берём первые три поля слоя.
    """
    add_percent = False
    inter_name_to_idx = {f.name(): i for i, f in enumerate(fields_intersect)}

    if selector_list == ["*"]:
        headers = [f.name() for f in fields_intersect]
        idxs = list(range(len(fields_intersect)))
        return headers, idxs, False

    headers, idxs = [], []
    for item in selector_list:
        if isinstance(item, str) and item == "%пересеч":
            add_percent = True
            continue
        if isinstance(item, int):
            if 0 <= item < len(fields_intersect):
                headers.append(fields_intersect[item].name())
                idxs.append(item)
            else:
                log_warn(f"Индекс колонки вне диапазона пересекаемого слоя: {item}")
            continue
        if isinstance(item, str):
            if item in inter_name_to_idx:
                idx = inter_name_to_idx[item]
                headers.append(item)
                idxs.append(idx)
            else:
                log_warn(f"Поле '{item}' отсутствует в пересекаемом слое; пропущено")
            continue
        log_warn(f"Некорректный селектор колонки: {item}")

    if not idxs:
        fallback_count = min(3, len(fields_intersect))
        idxs = list(range(fallback_count))
        headers = [fields_intersect[i].name() for i in idxs]

    return headers, idxs, add_percent


def make_table_filtered(fields_intersect, feat_intersect, procs, areas, path_layer, rules, default_rule):
    """
    Формирует заголовки и строки таблицы согласно rules/default.
    Дополнительно: для слоя ACTUAL_ZOUIT.TAB добавляет колонку 'Площадь пересеч, м²'.
    """
    selector = pick_rule_for_layer(rules, default_rule, path_layer)
    headers, idxs, add_percent = resolve_columns(fields_intersect, selector)

    if not headers and len(fields_intersect) == 0 and add_percent:
        headers = ["% пересеч"]

    if add_percent:
        headers = list(headers)
        if "% пересеч" not in headers:
            headers.append("% пересеч")

    # Нужен ли столбец площади? (только для ACTUAL_ZOUIT.TAB)
    add_area_col = bool(re.search(r'ACTUAL_ZOUIT\.TAB$', path_layer, re.IGNORECASE))
    if add_area_col:
        headers = list(headers)
        if "Площадь пересеч, м²" not in headers:
            headers.append("Площадь пересеч, м²")

    rows = []
    if feat_intersect:
        for feat, proc, area in zip(feat_intersect, procs, areas):
            row = []
            for i in idxs:
                val = feat.attribute(fields_intersect[i].name())
                row.append("" if val is None else str(val))
            if add_percent:
                row.append(proc)
            if add_area_col:
                row.append("" if area is None else str(round(float(area), 2)))
            rows.append(row)
    else:
        if not headers:
            headers = ["Сообщение"]
        rows = [["Нет пересечений"]]

    return headers, rows


# ============== ШИРИНЫ/ПЕРЕНОС/ТАБЛИЦЫ ==============
def _natural_text_width(pdf: FPDF, text: str, padding=2.0):
    return pdf.get_string_width(text) + 2 * padding

def _distribute_delta(widths, target_sum):
    s = sum(widths)
    if s == 0:
        return widths
    delta = target_sum - s
    if abs(delta) < 0.05:
        return widths
    total = sum(widths)
    if total <= 0:
        step = delta / len(widths)
        return [w + step for w in widths]
    return [w + delta * (w / total) for w in widths]


def measure_col_widths(pdf: FPDF, headers, rows, max_total_width, font_name="Sans", font_size=7):
    pdf.set_font(font_name, "", font_size)
    MIN_W = 14.0
    header_font_size = max(7, font_size + 1)

    natural = []
    for j, head in enumerate(headers):
        pdf.set_font(font_name, "B", header_font_size)
        w_head = _natural_text_width(pdf, "" if head is None else str(head))
        pdf.set_font(font_name, "", font_size)
        w_data = 0.0
        for row in rows:
            if j < len(row):
                w_data = max(w_data, _natural_text_width(pdf, "" if row[j] is None else str(row[j])))
        natural.append(max(w_head, w_data, MIN_W))

    total = sum(natural)
    if total <= 0:
        return [max_total_width / max(1, len(headers))] * len(headers)

    if total > max_total_width:
        k = max_total_width / total
        widths = [max(MIN_W, w * k) for w in natural]
    else:
        widths = natural[:]
        widths = _distribute_delta(widths, max_total_width)

    s = sum(widths)
    if s != 0 and abs(s - max_total_width) > 0.05:
        k = max_total_width / s
        widths = [w * k for w in widths]

    return widths


def estimate_row_height(pdf: FPDF, texts, col_widths, line_h=5.0, font_name="Sans", font_size=7):
    pdf.set_font(font_name, "", font_size)
    max_lines = 1
    avg_char_w = max(1.0, pdf.get_string_width("W"))
    for text, w in zip(texts, col_widths):
        s = str(text) if text is not None else ""
        chars_per_line = max(1, int((w - 2.0) / avg_char_w))
        lines = 1 + (len(s) // max(1, chars_per_line))
        if lines > max_lines:
            max_lines = lines
    return max_lines * line_h


def draw_row(pdf: FPDF, texts, col_widths, line_h=5.0, font_name="Sans", font_size=7, align="L", fill_bg=None, text_color=(0,0,0)):
    y_start = pdf.get_y()
    x = pdf.l_margin
    pdf.set_x(x)
    h = estimate_row_height(pdf, texts, col_widths, line_h=line_h, font_name=font_name, font_size=font_size)

    if fill_bg:
        pdf.set_fill_color(*fill_bg)
        pdf.rect(x, y_start, sum(col_widths), h, style="F")

    pdf.set_draw_color(200, 200, 200)
    for w in col_widths:
        pdf.rect(x, y_start, w, h)
        x += w

    x = pdf.l_margin
    pdf.set_text_color(*text_color)
    pdf.set_font(font_name, "", font_size)
    for text, w in zip(texts, col_widths):
        pdf.set_xy(x, y_start)
        pdf.multi_cell(w, line_h, "" if text is None else str(text), border=0, align=align)
        x += w
    pdf.set_text_color(0,0,0)
    pdf.set_xy(pdf.l_margin, y_start + h)


def draw_header_row_single_line(pdf: FPDF, headers, col_widths, line_h=6.0, font_name="Sans", font_size=8, fill_bg=None, text_color=(0,0,0)):
    y_start = pdf.get_y()
    x = pdf.l_margin
    pdf.set_x(x)
    h = line_h

    if fill_bg:
        pdf.set_fill_color(*fill_bg)
        pdf.rect(x, y_start, sum(col_widths), h, style="F")

    pdf.set_draw_color(200, 200, 200)
    pdf.set_font(font_name, "B", font_size)
    pdf.set_text_color(*text_color)

    for w, head in zip(col_widths, headers):
        pdf.rect(x, y_start, w, h)
        pdf.set_xy(x, y_start)
        pdf.cell(w, h, "" if head is None else str(head), border=0, align="C")
        x += w

    pdf.set_text_color(0,0,0)
    pdf.set_xy(pdf.l_margin, y_start + h)


def _is_empty_row(row):
    if not row:
        return True
    for c in row:
        if c is None:
            continue
        if str(c).strip() != "":
            return False
    return True


def draw_table(pdf: FPDF, title, headers, rows, max_total_width,
               font_name="Sans", font_size=7, header_font_size=8,
               colored=True):
    """
    Секция: полоса-заголовок, шапка (в одну строку), строки (с переносом).
    Если colored=False (нет пересечений) — ч/б с СЕРЫМИ заливками (не чёрными).
    """
    rows = list(rows)
    while rows and _is_empty_row(rows[0]):
        rows.pop(0)
    if not rows:
        rows = [["Нет данных"]]

    col_widths = measure_col_widths(pdf, headers, rows, max_total_width, font_name, font_size)

    section_bar_h = 7
    header_line_h = 6.0

    if (pdf.get_y() + section_bar_h + header_line_h) > pdf.page_break_trigger:
        pdf.add_page()

    if colored:
        accent = (0, 51, 102)
        title_text = (255, 255, 255)
        header_fill = (245, 245, 245)
        zebra_color = (248, 248, 248)
    else:
        accent = (180, 180, 180)
        title_text = (0, 0, 0)
        header_fill = (235, 235, 235)
        zebra_color = None

    pdf.set_font(font_name, "B", 9)
    pdf.set_fill_color(*accent)
    pdf.set_text_color(*title_text)
    pdf.cell(sum(col_widths), section_bar_h, title, ln=1, align="C", fill=True)
    pdf.set_text_color(0,0,0)

    draw_header_row_single_line(
        pdf, headers, col_widths,
        line_h=header_line_h,
        font_name=font_name, font_size=header_font_size,
        fill_bg=header_fill, text_color=(0,0,0)
    )

    fill = False
    for r in rows:
        rh = estimate_row_height(pdf, r, col_widths, line_h=5.0, font_name=font_name, font_size=font_size)
        if pdf.get_y() + rh > pdf.page_break_trigger:
            pdf.add_page()
            draw_header_row_single_line(
                pdf, headers, col_widths,
                line_h=header_line_h,
                font_name=font_name, font_size=header_font_size,
                fill_bg=header_fill, text_color=(0,0,0)
            )
        bg = (zebra_color if (colored and fill) else None)
        draw_row(pdf, r, col_widths, line_h=5.0, font_name=font_name, font_size=font_size,
                 align="L", fill_bg=bg, text_color=(0,0,0))
        fill = not fill if colored else False


# ===================== PDF-КОНТЕЙНЕР =====================
def safe_filename(name: str) -> str:
    s = re.sub(r'[<>:"/\\|?*\x00-\x1F]+', '_', name).strip()
    while s.endswith(".") or s.endswith(" "):
        s = s[:-1]
    return s or "report.pdf"


def try_place_logo(pdf: FPDF, path, x=15, y=10, target_w=12):
    try:
        try:
            from PIL import Image
            with Image.open(path) as im:
                w, h = im.size
                ratio = (target_w / float(w)) if w else 1.0
                target_h = h * ratio if h else target_w
            pdf.image(path, x=x, y=y, w=target_w, h=target_h)
        except Exception:
            pdf.image(path, x=x, y=y, w=target_w, h=0)
    except Exception:
        pdf.set_draw_color(150, 150, 150)
        pdf.rect(x, y, target_w, target_w)  # плейсхолдер


def register_fonts(pdf: FPDF, plugin_dir: str) -> bool:
    fonts_dir = os.path.join(plugin_dir, "fonts")
    reg = os.path.join(fonts_dir, "NotoSans-Regular.ttf")
    bold = os.path.join(fonts_dir, "NotoSans-Bold.ttf")
    italic = os.path.join(fonts_dir, "NotoSans-Italic.ttf")
    bolditalic = os.path.join(fonts_dir, "NotoSans-BoldItalic.ttf")

    if os.path.exists(reg) and os.path.exists(bold):
        try:
            pdf.add_font("Sans", style="", fname=reg)
            pdf.add_font("Sans", style="B", fname=bold)
            if os.path.exists(italic):
                pdf.add_font("Sans", style="I", fname=italic)
            if os.path.exists(bolditalic):
                pdf.add_font("Sans", style="BI", fname=bolditalic)
            pdf.set_font("Sans", size=7)
            return True
        except Exception as e:
            log_warn(f"Не удалось зарегистрировать TTF-шрифты: {e}")

    pdf.set_font("helvetica", size=7)
    return False


class PDF(FPDF):
    def __init__(self, *args, **kwargs):
        self.plugin_dir = kwargs.pop("plugin_dir", os.path.dirname(__file__))
        font_subsetting = kwargs.pop("font_subsetting", None)  # для новых версий fpdf2
        super().__init__(*args, **kwargs)

        self._doc_header_shown = False

        if font_subsetting:
            setter = getattr(self, "set_font_subsetting", None)
            if callable(setter):
                try:
                    setter(True)
                except Exception:
                    pass

        self.set_compression(True)
        self.set_title("Сведения ИСОГД")
        self.set_author("Комитет градостроительства и ЗР")
        self.set_subject("Отчёт по проверке пересечений")
        self.set_keywords("ИСОГД, отчёт, QGIS")
        self.set_creator("isogd QGIS plugin")
        self.alias_nb_pages()

        # Поля — чтобы таблица была «во всю ширину»
        self.set_auto_page_break(auto=True, margin=15)  # нижнее поле
        self.set_left_margin(10)
        self.set_right_margin(10)

        self.has_fonts = register_fonts(self, self.plugin_dir)

    def header(self):
        if self._doc_header_shown:
            return

        logo_path = os.path.join("T:/NOVOKUZ/Герб/Герб.jpg")
        try_place_logo(self, logo_path, x=15, y=10, target_w=12)

        self.set_xy(self.l_margin, 10)
        self.set_font("Sans" if self.has_fonts else "helvetica", "B", 14)
        self.cell(self.w - self.l_margin - self.r_margin, 6, "СВЕДЕНИЯ ИСОГД", ln=1, align="R")

        self.set_font("Sans" if self.has_fonts else "helvetica", "", 7)
        self.cell(self.w - self.l_margin - self.r_margin, 5,
                  dt.datetime.now().strftime("Дата/время формирования: %d.%m.%Y %H:%M:%S"),
                  ln=1, align="R")

        self.ln(18)
        self.set_draw_color(200, 200, 200)
        self.line(self.l_margin, self.get_y(), self.w - self.r_margin, self.get_y())
        self.ln(2)

        self._doc_header_shown = True

    def footer(self):
        self.set_y(-15)
        self.set_draw_color(200, 200, 200)
        self.line(self.l_margin, self.get_y(), self.w - self.r_margin, self.get_y())
        self.set_y(-12)
        self.set_font("Sans" if self.has_fonts else "helvetica", "I", 8)
        self.set_text_color(90, 90, 90)
        self.cell(0, 8, f"Стр. {self.page_no()} / {{nb}}", align="C")


# ========================= ПОИСК ПЕРЕСЕЧЕНИЙ =========================
class CheckLayers:
    """
    Для слоя pathLayer находит фичи, пересекающиеся с geom_sel (ожидаем выбранный ПОЛИГОН).
    Возвращает:
      self.fields         — QgsFields пересекаемого слоя,
      self.feat_intersect — список фич с пересечением,
      self.procs          — проценты пересечения (строки, 5 знаков),
      self.areas          — площадь пересечения (м²) для полигонов, иначе None.
    %
      - полигоны: area(intersection) / area(geom_sel) * 100
      - линии:    length(intersection) / length(фичи) * 100
      - точки:    100, если точка попала внутрь geom_sel
    """
    def __init__(self, pathLayer, geom_sel):
        self.fields = []
        self.feat_intersect = []
        self.procs = []
        self.areas = []

        layer_check = QgsVectorLayer(pathLayer, pathLayer, "ogr")
        if not layer_check or not layer_check.isValid():
            log_warn(f"Слой невалиден: {pathLayer}")
            return

        self.fields = layer_check.fields()

        s1_area = float(geom_sel.area()) if geom_sel and not geom_sel.isEmpty() else 0.0

        try:
            idx = QgsSpatialIndex(layer_check.getFeatures())
        except Exception:
            idx = None

        if idx:
            ids = idx.intersects(geom_sel.boundingBox())
            req = QgsFeatureRequest().setFilterFids(ids) if ids else QgsFeatureRequest()
            features = layer_check.getFeatures(req)
        else:
            features = layer_check.getFeatures()

        for feat in features:
            geom = feat.geometry()
            if not geom or geom.isEmpty():
                continue
            try:
                if not geom_sel.intersects(geom):
                    continue

                inter = geom_sel.intersection(geom)
                if not inter or inter.isEmpty():
                    continue

                gtype = QgsWkbTypes.geometryType(geom.wkbType())

                if gtype == QgsWkbTypes.PolygonGeometry:
                    if s1_area <= 0:
                        continue
                    inter_area = float(inter.area())
                    if inter_area <= 0:
                        continue
                    percent = inter_area * 100.0 / s1_area
                    area_val = inter_area  # м²

                elif gtype == QgsWkbTypes.LineGeometry:
                    line_len = float(geom.length())
                    if line_len <= 0:
                        continue
                    inter_len = float(inter.length())
                    if inter_len <= 0:
                        continue
                    percent = inter_len * 100.0 / line_len
                    area_val = None

                elif gtype == QgsWkbTypes.PointGeometry:
                    percent = 100.0
                    area_val = None

                else:
                    continue

                self.feat_intersect.append(feat)
                self.procs.append(str(round(percent, 5)))
                self.areas.append(area_val)

            except Exception:
                continue


# ================================ ПЛАГИН ================================
class Info:
    """QGIS Plugin Implementation."""

    def __init__(self, iface_):
        self.iface = iface_
        self.plugin_dir = os.path.dirname(__file__)

        val = QSettings().value('locale/userLocale', '', type=str)
        locale = val[:2] if isinstance(val, str) and len(val) >= 2 else 'en'
        locale_path = os.path.join(self.plugin_dir, 'i18n', f'Info_{locale}.qm')
        if os.path.exists(locale_path):
            self.translator = QTranslator()
            self.translator.load(locale_path)
            QCoreApplication.installTranslator(self.translator)

        self.actions = []
        self.menu = self.tr(u'&ISOGD')
        self.toolbar = self.iface.addToolBar(u'Info')
        self.toolbar.setObjectName(u'Info')

        self.pluginIsActive = False
        self.dockwidget = None

        self.column_rules, self.default_rule = load_column_rules(self.plugin_dir)

    def tr(self, message):
        return QCoreApplication.translate('Info', message)

    def add_action(self, icon_path, text, callback,
                   enabled_flag=True, add_to_menu=True, add_to_toolbar=True,
                   status_tip=None, whats_this=None, parent=None):
        icon = QIcon(icon_path)
        action = QAction(icon, text, parent)
        action.triggered.connect(callback)
        action.setEnabled(enabled_flag)
        if status_tip is not None:
            action.setStatusTip(status_tip)
        if whats_this is not None:
            action.setWhatsThis(whats_this)
        if add_to_toolbar:
            self.toolbar.addAction(action)
        if add_to_menu:
            self.iface.addPluginToMenu(self.menu, action)
        self.actions.append(action)
        return action

    def initGui(self):
        icon_path = ':/plugins/isogd/icon.png'
        self.add_action(icon_path, text=self.tr(u''), callback=self.run, parent=self.iface.mainWindow())

    def onClosePlugin(self):
        if self.dockwidget:
            try:
                self.dockwidget.closingPlugin.disconnect(self.onClosePlugin)
            except Exception:
                pass
        self.pluginIsActive = False

    def unload(self):
        for action in self.actions:
            self.iface.removePluginMenu(self.tr(u'&ISOGD'), action)
        try:
            del self.toolbar
        except Exception:
            pass

    def run(self):
        if not self.pluginIsActive:
            self.pluginIsActive = True
            if self.dockwidget is None:
                self.dockwidget = InfoDockWidget()
            self.dockwidget.closingPlugin.connect(self.onClosePlugin)
            self.iface.addDockWidget(Qt.TopDockWidgetArea, self.dockwidget)
            self.dockwidget.show()
            self.dockwidget.pushButton.clicked.connect(self.generate_report)

    # ----------------------- Генерация отчёта -----------------------
    def window_error(self, text="Не выбран объект для проверки! Выберите полигон."):
        QMessageBox.warning(None, "ISOGD", text)

    def get_dir_name(self):
        return QFileDialog.getExistingDirectory(None, 'Выберите папку для сохранения Справки')

    def _print_header_block(self, pdf: FPDF, layer, SelectedFeature):
        pdf.set_font("Sans" if pdf.has_fonts else "helvetica", "", 7)
        pdf.set_draw_color(0, 0, 0)
        pdf.ln(2)
        pdf.set_font("Sans" if pdf.has_fonts else "helvetica", "B", 9)
        pdf.cell(0, 6, f"Исходный слой: {layer.sourceName()}", ln=1)
        pdf.set_font("Sans" if pdf.has_fonts else "helvetica", "", 7)

        fields_source = SelectedFeature.fields()
        max_show = min(5, len(fields_source))
        for i in range(max_show):
            name = fields_source[i].name()
            val = SelectedFeature.attribute(name)
            pdf.cell(60, 5, name)
            pdf.cell(0, 5, "" if val is None else str(val), ln=1)

        pdf.ln(2)
        pdf.set_draw_color(200, 200, 200)
        pdf.cell(0, 0.5, "", ln=1)
        pdf.ln(2)

    def generate_report(self):
        Layer = iface.activeLayer()
        if Layer is None:
            self.window_error("Нет активного слоя.")
            return
        if Layer.selectedFeatureCount() == 0:
            self.window_error("Не выбрано ни одного объекта.")
            return

        dir_name = self.get_dir_name()
        if not dir_name:
            QMessageBox.information(None, "ISOGD", "Сохранение отменено.")
            return

        SelectedFeatures = list(Layer.getSelectedFeatures())
        if not SelectedFeatures:
            self.window_error("Не удалось получить выбранные объекты.")
            return

        # === ОБНОВЛЁННЫЙ СПИСОК СЛОЁВ ===
        DATA_LAYERS = [
            'T:/Номенклатура/Лист_500.TAB',
            'T:/NOVOKUZ/Красные_линии_полигон.TAB',
            'T:/NOVOKUZ/_Правила землепользования и застройки/Территориальные_зоны_пр.TAB',
            'T:/NOVOKUZ/ФГУ участки/KK.TAB',
            'T:/NOVOKUZ/Участки.TAB',
            'T:/NOVOKUZ/ФГУ участки/ACTUAL_LAND.TAB',
            'T:/ADRES/Строения.TAB',
            'T:/NOVOKUZ/ФГУ участки/ACTUAL_OKSN.TAB',
            'T:/NOVOKUZ/Схема расположения.TAB',
            'T:/NOVOKUZ/Предварительно согласованные.TAB',
            'T:/NOVOKUZ/КУМИ_схема_расположения.TAB',
            'T:/NOVOKUZ/КУМИ_предварительно_согласованные_с_кадастровым_номером.TAB',
            'T:/NOVOKUZ/Банк ЗУ многодетные.TAB',
            'T:/NOVOKUZ/Проекты планировок и межеваний.TAB',
            'T:/NOVOKUZ/ФГУ участки/ACTUAL_ZOUIT.TAB',
            'T:/NOVOKUZ/ФГУ участки/ЗОУИТ.TAB',
            'T:/NOVOKUZ/Водоохранные зоны Томь, Аба, Горбуниха, Бунгур, Кондома (ЦИТ Барнаул)/Береговая_Прибрежная_Водоохранная.TAB',
            'T:/NOVOKUZ/Приаэродромная территория МСК42зона2.tab',
            'T:/NOVOKUZ/Санзоны новые.TAB',
            'T:/NOVOKUZ/Охранная зона.TAB',
            'T:/NOVOKUZ/ЗОНЫ КУЛЬТУРНОГО НАСЛЕДИЯ/Границы территорий объектов Культурного наследия.TAB',
            'T:/NOVOKUZ/ЗОНЫ КУЛЬТУРНОГО НАСЛЕДИЯ/Зоны охраны объектов культурного наследия.TAB',
            'T:/NOVOKUZ/ЗОНЫ КУЛЬТУРНОГО НАСЛЕДИЯ/Объекты культурного наследия на ГКУ.TAB',
            'T:/NOVOKUZ/ЗОНЫ КУЛЬТУРНОГО НАСЛЕДИЯ/Объекты культурного наследия.TAB',
            'T:/NOVOKUZ/Разрешение на использование ЗУ.TAB',
            'T:/NOVOKUZ/Схема НТО.TAB',
            'T:/NOVOKUZ/Схема НТО/Договоры НТО.TAB',
            'T:/NOVOKUZ/Металлические_гаражи.TAB',
            'T:/NOVOKUZ/Сервитуты.TAB',
            'T:/NOVOKUZ/Ветхое и аварийное жильё.TAB',
            'T:/NOVOKUZ/Запреты/Переселение 185 ФЗ/Лэнд_аварий.TAB',
            'T:/NOVOKUZ/Запреты/Арест судебные приставы/Запрет судебных приставов.TAB',
            'T:/NOVOKUZ/Запреты/Реконструкция южного въезда/Южная автомагистраль.TAB',
            'T:/NOVOKUZ/Запреты/Строительство Байдаевского моста/Байдаевский мост изъятие.TAB',
            'T:/NOVOKUZ/Запреты/Ликвидация последствий горных выработок/ЛиквидацияОстатковШахт_не предоставлять.TAB',
            'T:/Темы/Сибиреязвенное захоронение.TAB',
            'T:/NOVOKUZ/Месторождения полезных ископаемых/Месторождения полезных ископаемых.TAB',
            'T:/Темы/Реки города.TAB',
            'T:/NOVOKUZ/ПРОЕКТЫ/План лесонасаждений городских лесов/Отцифрованный А-3_горНовокузнецк_КатЗащ/Городские леса.TAB',
            'T:/NOVOKUZ/Граница Новокузнецкого городского округа.TAB',
            'T:/NOVOKUZ/Граница населенного пункта город Новокузнецк.TAB'
        ]

        pdf = PDF(orientation="L", unit="mm", format="A3", plugin_dir=self.plugin_dir)
        pdf.add_page()

        SelectedFeature = SelectedFeatures[0]
        try:
            self._print_header_block(pdf, Layer, SelectedFeature)
        except Exception as e:
            log_warn(f"Не удалось вывести шапку с атрибутами исходного слоя: {e}")

        progress = QProgressDialog("ИСОГД формирует отчёт...", "Отмена", 0, len(DATA_LAYERS))
        progress.setWindowModality(Qt.WindowModal)

        SelectedFeatureGeometry = SelectedFeature.geometry()
        max_table_width = pdf.w - pdf.l_margin - pdf.r_margin  # на всю ширину страницы

        for i, pathLayer in enumerate(DATA_LAYERS, 1):
            if progress.wasCanceled():
                QMessageBox.information(None, "ISOGD", "Операция отменена.")
                return

            a = CheckLayers(pathLayer, SelectedFeatureGeometry)

            headers, rows = make_table_filtered(
                a.fields, a.feat_intersect, a.procs, a.areas, pathLayer,
                self.column_rules, self.default_rule
            )

            if not headers:
                headers = ["Нет данных"]
                rows = [["—"]]

            # ч/б стиль, если нет пересечений
            is_bw = (len(rows) == 1 and len(rows[0]) == 1 and str(rows[0][0]).strip().lower() == "нет пересечений")

            draw_table(
                pdf, title=pathLayer, headers=headers, rows=rows,
                max_total_width=max_table_width,
                font_name="Sans" if pdf.has_fonts else "helvetica",
                colored=not is_bw
            )

            pdf.ln(4)
            progress.setValue(i)

        time_now = dt.datetime.now().strftime("%H%M%S")
        fields_source = SelectedFeature.fields()
        f0_name = fields_source[0].name() if len(fields_source) > 0 else ""
        f4_name = fields_source[4].name() if len(fields_source) > 4 else (fields_source[max(0, len(fields_source)-1)].name() if len(fields_source) else "")

        field_0 = str(SelectedFeature.attribute(f0_name)) if f0_name else ""
        field_4 = str(SelectedFeature.attribute(f4_name)) if f4_name else ""

        base_name = f"Вх№{field_0} {field_4} {time_now}.pdf"
        file_name = os.path.join(dir_name, safe_filename(base_name))

        try:
            pdf.output(file_name)
            os.startfile(file_name)
        except Exception as e:
            log_err(f"Ошибка сохранения PDF: {e}")
            QMessageBox.critical(None, "ISOGD", f"Не удалось сохранить PDF:\n{e}")
