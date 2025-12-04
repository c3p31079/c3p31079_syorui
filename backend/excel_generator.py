# -*- coding: utf-8 -*-
"""
Excel 生成ロジック
・template.xlsx を読み込み
・指定座標に図形を貼り付け
・最終的に更新した Excel バイナリを返す
"""

import io
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import json
import os

# ------------ 設定 -------------
TEMPLATE_PATH = "backend/template.xlsx"
ICON_DIR = "backend/icons"
CHECK_MAP_PATH = "docs/check_coord_map.json"

# Excel のキャンバス上 xy 座標 → openpyxl では Cell に貼る形になるため
# 微調整用
DEFAULT_ANCHOR_CELL = "A1"


# ---------------------------
# JSON 読み込み
# ---------------------------
def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


# ---------------------------
# Excel 生成メイン
# ---------------------------
def generate_excel_file(rows, checks):
    # ① テンプレート読み込み
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # ② 各図形アイコンを openpyxl Image として読み込む
    icon_triangle = Image(os.path.join(ICON_DIR, "triangle.png"))
    icon_cross = Image(os.path.join(ICON_DIR, "cross.png"))
    icon_circle = Image(os.path.join(ICON_DIR, "circle.png"))
    icon_check = Image(os.path.join(ICON_DIR, "check.png"))

    # ③ 座標マップ読み込み
    check_coord_map = load_json(CHECK_MAP_PATH)

    # ④ 点検部位＋項目の記号配置
    #    → rows = [{part, item, symbol}]
    #    → 座標は map.json にリンクできるよう後で拡張可能
    base_y = 100           # 1行目の図形の基準位置 Y（px）
    offset_y = 40          # 行ごとにずらす距離（px）
    x_symbol = 120         # 記号を置く X座標（px）

    for idx, r in enumerate(rows):
        y = base_y + idx * offset_y

        if r["symbol"] == "triangle":
            icon = icon_triangle
        elif r["symbol"] == "cross":
            icon = icon_cross
        elif r["symbol"] == "circle":
            icon = icon_circle
        else:
            continue

        # openpyxl は Cell 固定なので、A1 に貼って position だけ変える
        img = Image(icon.ref)
        img.width = icon.width
        img.height = icon.height
        img.anchor = "A1"
        img.drawing.top = y
        img.drawing.left = x_symbol

        ws.add_image(img)

    # ⑤ チェック項目（✓）の配置
    for label, coord in check_coord_map.items():
        if label not in checks:
            continue

        x = coord["x"]
        y = coord["y"]

        img = Image(os.path.join(ICON_DIR, "check.png"))
        img.anchor = DEFAULT_ANCHOR_CELL
        img.drawing.left = x
        img.drawing.top = y

        ws.add_image(img)

    # ⑥ バイナリ返却
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
