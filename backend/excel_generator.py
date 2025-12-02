import os
import io
import uuid
import json
from PIL import Image, ImageDraw, ImageFont
from openpyxl import load_workbook

BASE_XLSX = "templates/base.xlsx"
TEMPLATE_PNG = "templates/base_sheet.png"  # optional; あれば背景に使う
PREVIEW_DIR = "preview"

# マップの読み込む
def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)

COORD_MAP = load_json("coord_map.json")
CHECK_MAP = load_json("check_coord_map.json")

# PIL画像上に図形を描画する
def draw_circle(draw, x, y, size=30, outline="red", width=4):
    r = size
    draw.ellipse((x-r, y-r, x+r, y+r), outline=outline, width=width)

def draw_triangle(draw, x, y, size=30, outline="red", width=4):
    s = size
    points = [(x, y-s), (x-s, y+s), (x+s, y+s)]
    draw.polygon(points, outline=outline, fill=None)
    # PILのポリゴンはストロークのみの描画を容易にサポートしてないらしい。線を描画することでストロークを実現！
    draw.line([points[0], points[1], points[2], points[0]], fill=outline, width=width)

def draw_cross(draw, x, y, size=25, outline="red", width=4):
    s = size
    draw.line((x-s, y-s, x+s, y+s), fill=outline, width=width)
    draw.line((x+s, y-s, x-s, y+s), fill=outline, width=width)

def draw_check(draw, x, y, size=25, outline="green", width=4):
    # おおよそのチェックマーク
    draw.line((x-size, y, x-size/3, y+size), fill=outline, width=width)
    draw.line((x-size/3, y+size, x+size, y-size/2), fill=outline, width=width)

# 正確に0.5 → △；正確に0.2 → ×
def decide_shape(score, part, item):
    if score == 0.5:
        return "triangle"
    if score == 0.2:
        return "cross"
    # アイテムがCOORD_MAPで形状「circle」として定義されている場合のみ円を表示する（デフォルトでは使用されないよ！）
    return None

# プレビュー用PNGを作成
def create_preview_image(part, item, score, checked_list):
    # ベースの背景
    if os.path.exists(TEMPLATE_PNG):
        bg = Image.open(TEMPLATE_PNG).convert("RGBA")
    else:
        # 大きな空白の「シート状」の画像を作成する
        bg = Image.new("RGBA", (1200, 800), "white")
    draw = ImageDraw.Draw(bg)

    # タイトル！
    try:
        font = ImageFont.truetype("DejaVuSans.ttf", 16)
    except Exception:
        font = ImageFont.load_default()

    # 画像上に参照用の部分/項目を記入する
    draw.text((12, 12), f"点検部位: {part}", fill="black", font=font)
    draw.text((12, 32), f"項目: {item}", fill="black", font=font)
    draw.text((12, 52), f"score: {score}", fill="black", font=font)

    key = f"{part}:{item}"
    coords = COORD_MAP.get(key)
    if coords:
        x = coords["x"]
        y = coords["y"]
        shape = decide_shape(score, part, item)
        if shape == "triangle":
            draw_triangle(draw, x, y, size=36)
        elif shape == "cross":
            draw_cross(draw, x, y, size=30)
        else:
            # 一致しない場合、デフォルトで小さな円を表示する
            draw_circle(draw, x, y, size=28)

    # チェック済みチェック項目の文字列の配列
    # チェックマップ座標でチェックを出す
    for chk in checked_list:
        ch_coord = CHECK_MAP.get(chk)
        if ch_coord:
            draw_check(draw, ch_coord["x"], ch_coord["y"], size=20)

    # プレビューディレクトリの存在を確認
    os.makedirs(PREVIEW_DIR, exist_ok=True)
    filename = f"preview_{uuid.uuid4().hex}.png"
    path = os.path.join(PREVIEW_DIR, filename)
    bg.save(path)
    return path

# 実際のExcelを生成！書き込みとシートへの簡易図形の描画
def generate_excel_file(part, item, score, checked_list, coords):
    wb = load_workbook(BASE_XLSX)
    ws = wb.active

    # 簡易テキスト置換：使うかわからん
    ws["B2"] = part
    ws["B3"] = item
    ws["B4"] = score

    # Excelへの図形挿入
    # 簡略化のため、小さなPNGマーカーを生成
    # マーカーフォルダを作成
    markers_dir = "preview/markers"
    os.makedirs(markers_dir, exist_ok=True)

    # PNG 画像を作成するための補助ツールらしい：ここら辺はちょっと理解できてない
    def make_marker_png(kind, path, size=60, color="red"):
        im = Image.new("RGBA", (size, size), (255,255,255,0))
        d = ImageDraw.Draw(im)
        if kind == "triangle":
            draw_triangle(d, size//2, size//2, size//3, outline=color, width=6)
        elif kind == "cross":
            draw_cross(d, size//2, size//2, size//3, outline=color, width=6)
        elif kind == "circle":
            draw_circle(d, size//2, size//2, size//3, outline=color, width=6)
        elif kind == "check":
            draw_check(d, size//2, size//2, size//3, outline="green", width=6)
        im.save(path)

    from openpyxl.drawing.image import Image as XLImage

    # エクセルの主要形状
    if coords:
        x = coords["x"]
        y = coords["y"]
        # プレビューと同じ形状をExcel用に決定する
        s = decide_shape(score, part, item)
        if s is None:
            s = "circle"
        marker_path = os.path.join(markers_dir, f"m_{uuid.uuid4().hex}.png")
        make_marker_png(s, marker_path, size=80)
        img = XLImage(marker_path)

        # 経験則に基づき分割により近似列/行セルへピクセルx,yをマッピング！
        col = max(1, int(x // 80) + 1)  # マッピング
        row = max(1, int(y // 40) + 1)
        cell_anchor = f"{chr(64+min(col,26))}{row}"
        ws.add_image(img, cell_anchor)

    # チェックマークを追加
    for chk in checked_list:
        ch_coord = CHECK_MAP.get(chk)
        if ch_coord:
            marker_path = os.path.join(markers_dir, f"c_{uuid.uuid4().hex}.png")
            make_marker_png("check", marker_path, size=48)
            img = XLImage(marker_path)
            col = max(1, int(ch_coord["x"] // 80) + 1)
            row = max(1, int(ch_coord["y"] // 40) + 1)
            ws.add_image(img, f"{chr(64+min(col,26))}{row}")

    # ファイルの保存
    out_path = f"/tmp/inspection_{uuid.uuid4().hex}.xlsx"
    wb.save(out_path)
    return out_path
