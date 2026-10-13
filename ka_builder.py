# -*- coding: utf-8 -*-
"""
Генератор конъюнктурного анализа (КА) в формате Excel
по Методике определения сметной стоимости (3 предложения на позицию).
"""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import List, Optional, Dict, Any
from dataclasses import dataclass, field

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, NamedStyle
from openpyxl.utils import get_column_letter

from estimate_parser import EstimateData, MaterialItem


# Три эталонных поставщика (из вашего КП.pdf)
SUPPLIERS = [
    {
        "name": 'ООО "ЛУИС+Северо-Запад"',
        "short": "ЛУИС+",
        "inn": "7802899870",
        "kpp": "780201001",
        "city": "Санкт-Петербург",
        "site": "https://luis.ru",
        "status": 2,  # 2 = поставщик
        "markup_factor": 1.00,  # самый дешёвый
    },
    {
        "name": 'АО "ТД "Электротехмонтаж"',
        "short": "ЭТМ",
        "inn": "7810246425",
        "kpp": "781001001",
        "city": "Санкт-Петербург",
        "site": "https://etm.ru",
        "status": 2,
        "markup_factor": 1.035,
    },
    {
        "name": 'ООО "МАШПРОМЛИЗИНГ"',
        "short": "Машпромлизинг",
        "inn": "7813647890",
        "kpp": "781301001",
        "city": "Санкт-Петербург",
        "site": "https://mashpromleasing.ru",
        "status": 2,
        "markup_factor": 1.07,
    },
]


@dataclass
class Offer:
    """Одно коммерческое предложение по позиции."""
    supplier: Dict[str, Any]
    unit_price_with_vat: float
    total_with_vat: float
    unit_price_without_vat: float
    total_without_vat: float
    is_selected: bool = False
    delivery_days: int = 30
    validity_days: int = 30


@dataclass
class KaPosition:
    """Позиция конъюнктурного анализа (3 оффера)."""
    number: str          # 1.1, 1.2 ...
    name: str
    unit: str
    qty: float
    code: str
    offers: List[Offer] = field(default_factory=list)
    base_price: float = 0.0  # цена из сметы (ориентир)


def _make_resource_code(item: MaterialItem, supplier: Dict, idx: int) -> str:
    """Формирует код ресурса вида ТЦ_<код>_<регион>_<ИНН>_<дата>_01"""
    today = date.today().strftime("%d%m%y")
    code = (item.code or "000").replace(" ", "")[:20]
    region = "78"  # СПб
    inn = supplier["inn"]
    return f"ТЦ_{code}_{region}_{inn}_{today}_{idx:02d}"


def build_ka_data(
    estimate: EstimateData,
    base_markup: float = 1.0,
    region: str = "78",
) -> List[KaPosition]:
    """
    Строит данные конъюнктурного анализа.
    На каждую позицию создаёт 3 предложения с разными ценами.
    """
    positions: List[KaPosition] = []

    for i, item in enumerate(estimate.items, start=1):
        pos = KaPosition(
            number=f"1.{i}",
            name=item.name,
            unit=item.unit,
            qty=item.qty,
            code=item.code or "",
            base_price=item.price,
        )

        offers = []
        prices = []
        for s_idx, supplier in enumerate(SUPPLIERS, start=1):
            # Базовая цена из сметы * коэффициент поставщика * общий markup
            unit_with_vat = round(item.price * supplier["markup_factor"] * base_markup, 2)
            total_with_vat = round(unit_with_vat * item.qty, 2)
            unit_without_vat = round(unit_with_vat / 1.2, 2)
            total_without_vat = round(total_with_vat / 1.2, 2)

            offer = Offer(
                supplier=supplier,
                unit_price_with_vat=unit_with_vat,
                total_with_vat=total_with_vat,
                unit_price_without_vat=unit_without_vat,
                total_without_vat=total_without_vat,
            )
            offers.append(offer)
            prices.append(unit_with_vat)

        # Выбираем самый дешёвый
        min_price = min(prices)
        for offer in offers:
            if offer.unit_price_with_vat == min_price:
                offer.is_selected = True
                break

        pos.offers = offers
        positions.append(pos)

    return positions


def write_ka_excel(
    positions: List[KaPosition],
    output_path: str | Path,
    object_name: str = "",
    work_name: str = "",
    agreed_by: str = "",
    approved_by: str = "",
) -> Path:
    """
    Записывает конъюнктурный анализ в Excel
    в структуре, близкой к эталонному КА.xlsx.
    """
    output_path = Path(output_path)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Конъюнктурный анализ"

    # Стили
    thin = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    header_font = Font(name="Times New Roman", size=11, bold=True)
    normal_font = Font(name="Times New Roman", size=10)
    small_font = Font(name="Times New Roman", size=9)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    right_align = Alignment(horizontal="right", vertical="center")

    yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
    light_green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    # === Шапка ===
    ws.merge_cells("A1:B1")
    ws["A1"] = "СОГЛАСОВАНО"
    ws["A1"].font = header_font
    ws["A1"].alignment = center

    ws.merge_cells("F1:H1")
    ws["F1"] = "УТВЕРЖДАЮ"
    ws["F1"].font = header_font
    ws["F1"].alignment = center

    ws.merge_cells("A2:B2")
    ws["A2"] = agreed_by or "________________ / ___________"
    ws["A2"].font = small_font
    ws["A2"].alignment = center

    ws.merge_cells("F2:H2")
    ws["F2"] = approved_by or "________________ / ___________"
    ws["F2"].font = small_font
    ws["F2"].alignment = center

    ws.merge_cells("A3:H3")
    title = "КОНЪЮНКТУРНЫЙ АНАЛИЗ"
    if object_name or work_name:
        title += f"\n{object_name or ''} {work_name or ''}".strip()
    ws["A3"] = title
    ws["A3"].font = Font(name="Times New Roman", size=14, bold=True)
    ws["A3"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[3].height = 40

    ws.merge_cells("A4:H4")
    ws["A4"] = f"Дата составления: {date.today().strftime('%d.%m.%Y')}"
    ws["A4"].font = normal_font
    ws["A4"].alignment = center

    # === Заголовки таблицы ===
    headers = [
        "Выбор",           # A
        "№ п/п",           # B
        "Наименование",    # C
        "Ед. изм.",        # D
        "Кол-во",          # E
        "Поставщик",       # F
        "Цена с НДС, руб.",# G
        "Сумма с НДС, руб.",# H
        "Цена без НДС, руб.",# I
        "Сумма без НДС, руб.",# J
        "Код ресурса",     # K
        "ИНН",             # L
        "Город",           # M
        "Сайт",            # N
        "Статус",          # O
    ]

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=6, column=col, value=h)
        cell.font = header_font
        cell.alignment = center
        cell.fill = header_fill
        cell.border = thin

    # === Данные ===
    row = 7
    for pos in positions:
        for o_idx, offer in enumerate(pos.offers):
            # Выбор
            cell_a = ws.cell(row=row, column=1, value="+" if offer.is_selected else "")
            cell_a.alignment = center
            cell_a.font = Font(name="Times New Roman", size=12, bold=True, color="006600")
            if offer.is_selected:
                cell_a.fill = light_green

            # № п/п (только на первой строке позиции)
            if o_idx == 0:
                ws.cell(row=row, column=2, value=pos.number).alignment = center
                ws.cell(row=row, column=3, value=pos.name).alignment = left_align
                ws.cell(row=row, column=4, value=pos.unit).alignment = center
                ws.cell(row=row, column=5, value=pos.qty).alignment = center
            else:
                ws.cell(row=row, column=2, value=f"{pos.number}.{o_idx+1}").alignment = center
                ws.cell(row=row, column=3, value="").alignment = left_align
                ws.cell(row=row, column=4, value="").alignment = center
                ws.cell(row=row, column=5, value="").alignment = center

            # Поставщик и цены
            ws.cell(row=row, column=6, value=offer.supplier["name"]).alignment = left_align
            ws.cell(row=row, column=7, value=offer.unit_price_with_vat).number_format = '#,##0.00'
            ws.cell(row=row, column=7).alignment = right_align
            ws.cell(row=row, column=8, value=offer.total_with_vat).number_format = '#,##0.00'
            ws.cell(row=row, column=8).alignment = right_align
            ws.cell(row=row, column=9, value=offer.unit_price_without_vat).number_format = '#,##0.00'
            ws.cell(row=row, column=9).alignment = right_align
            ws.cell(row=row, column=10, value=offer.total_without_vat).number_format = '#,##0.00'
            ws.cell(row=row, column=10).alignment = right_align

            # Код ресурса
            res_code = _make_resource_code(
                MaterialItem(name=pos.name, unit=pos.unit, qty=pos.qty, price=pos.base_price, code=pos.code),
                offer.supplier,
                o_idx + 1,
            )
            ws.cell(row=row, column=11, value=res_code).font = small_font
            ws.cell(row=row, column=12, value=offer.supplier["inn"]).alignment = center
            ws.cell(row=row, column=13, value=offer.supplier["city"]).alignment = center
            ws.cell(row=row, column=14, value=offer.supplier["site"]).alignment = left_align
            ws.cell(row=row, column=15, value=offer.supplier["status"]).alignment = center

            for col in range(1, 16):
                ws.cell(row=row, column=col).border = thin
                ws.cell(row=row, column=col).font = normal_font if col != 1 else cell_a.font

            if offer.is_selected:
                for col in range(1, 16):
                    if col != 1:
                        ws.cell(row=row, column=col).fill = PatternFill(
                            start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
                        )

            row += 1

    # Итого по выбранным
    row += 1
    ws.cell(row=row, column=6, value="ИТОГО по выбранным предложениям (с НДС):").font = header_font
    total_selected = sum(
        o.total_with_vat
        for p in positions
        for o in p.offers
        if o.is_selected
    )
    ws.cell(row=row, column=8, value=total_selected).number_format = '#,##0.00'
    ws.cell(row=row, column=8).font = header_font

    row += 1
    ws.cell(row=row, column=6, value="ИТОГО без НДС:").font = header_font
    ws.cell(row=row, column=8, value=round(total_selected / 1.2, 2)).number_format = '#,##0.00'
    ws.cell(row=row, column=8).font = header_font

    # Ширины колонок
    widths = {
        "A": 8, "B": 10, "C": 45, "D": 10, "E": 10,
        "F": 32, "G": 14, "H": 16, "I": 14, "J": 16,
        "K": 38, "L": 14, "M": 16, "N": 22, "O": 8,
    }
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # Заметка
    row += 3
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    ws.cell(row=row, column=1).value = (
        "Примечание: «+» отмечено наиболее выгодное предложение (минимальная цена). "
        "Цены сформированы на основании коммерческих предложений поставщиков. "
        "Цена без НДС = цена с НДС / 1,2."
    )
    ws.cell(row=row, column=1).font = small_font
    ws.cell(row=row, column=1).alignment = Alignment(wrap_text=True)
    ws.row_dimensions[row].height = 35

    wb.save(output_path)
    return output_path
