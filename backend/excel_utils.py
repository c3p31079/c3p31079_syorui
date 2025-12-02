import os
import io
import uuid
from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter

ICONS_DIR = "backend/icons"

def ensure_icons_exist():
    """
    Create simple PNG icons if not present (triangle/cross/circle/check).
    32x32 transparent background.
    """
    os.makedirs(ICONS_DIR, exist_ok=True)
    size = 64
    # △
    tri_path = os.path.join(ICONS_DIR, "triangle.png")
    cross_path = os.path.join(ICONS_DIR, "cross.png")
    circ_path = os.path.join(ICONS_DIR, "circle.png")
    check_path = os.path.join(ICONS_DIR, "check.png")

    if not os.path.exists(tri_path):
        im = Image.new("RGBA", (size, size), (0,0,0,0))
        d = ImageDraw.Draw(im)
        d.line([(size*0.5, size*0.12), (size*0.12, size*0.85), (size*0.88, size*0.85), (size*0.5, size*0.12)],
               fill=(220,40,40,255), width=8)
        im.save(tri_path)
    if not os.path.exists(cross_path):
        im = Image.new("RGBA", (size, size), (0,0,0,0))
        d = ImageDraw.Draw(im)
        d.line([(size*0.18,size*0.18),(size*0.82,size*0.82)], fill=(220,40,40,255), width=8)
        d.line([(size*0.82,size*0.18),(size*0.18,size*0.82)], fill=(220,40,40,255), width=8)
        im.save(cross_path)
    if not os.path.exists(circ_path):
        im = Image.new("RGBA", (size, size), (0,0,0,0))
        d = ImageDraw.Draw(im)
        d.ellipse((size*0.12,size*0.12,size*0.88,size*0.88), outline=(220,40,40,255), width=8)
        im.save(circ_path)
    if not os.path.exists(check_path):
        im = Image.new("RGBA", (size, size), (0,0,0,0))
        d = ImageDraw.Draw(im)
        # チェック（レ点）
        d.line([(size*0.15,size*0.55),(size*0.4,size*0.78),(size*0.85,size*0.2)], fill=(16,185,129,255), width=8)
        im.save(check_path)

def _get_column_pixel_width(ws, col_letter):
    # 近似値: Excelの列幅 -> px: 近似幅 * 7 らしい。
    dim = ws.column_dimensions.get(col_letter)
    if dim and dim.width:
        return int(dim.width * 7)
    return 64  # デフォルト

def _get_row_pixel_height(ws, row_idx):
    dim = ws.row_dimensions.get(row_idx)
    if dim and dim.height:
        return int(dim.height)
    return 20  # デフォルト

def render_range_to_image(xlsx_path, sheet_name, cell_range):
    """
    Render the given range (e.g. 'A1:H37') to a PNG image in memory and return BytesIO.
    """
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb[sheet_name]

    # パース範囲
    start_cell, end_cell = cell_range.split(":")
    start_col = ''.join([c for c in start_cell if c.isalpha()])
    start_row = int(''.join([c for c in start_cell if c.isdigit()]))
    end_col = ''.join([c for c in end_cell if c.isalpha()])
    end_row = int(''.join([c for c in end_cell if c.isdigit()]))

    # ピクセルの幅/高さを計算
    cols = []
    total_w = 0
    for col_idx in range(ord(start_col.upper()) - 64, ord(end_col.upper()) - 64 + 1):
        col_letter = get_column_letter(col_idx)
        w = _get_column_pixel_width(ws, col_letter)
        cols.append(w)
        total_w += w

    rows = []
    total_h = 0
    for r in range(start_row, end_row+1):
        h = _get_row_pixel_height(ws, r)
        rows.append(h)
        total_h += h

    # 画像を作成
    img = Image.new("RGBA", (total_w + 2, total_h + 2), "white")
    draw = ImageDraw.Draw(img)
    # 簡易フォント
    try:
        font = ImageFont.truetype("DejaVuSans.ttf", 12)
    except Exception:
        font = ImageFont.load_default()

    # 描画セル
    y = 0
    for r_idx, r_h in enumerate(rows):
        x = 0
        for c_idx, c_w in enumerate(cols):
            # 矯正
            draw.rectangle([x, y, x + c_w, y + r_h], outline=(200,200,200))
            # 実際のセル名へのマッピング
            col_letter = get_column_letter(ord(start_col.upper()) - 64 + c_idx)
            row_num = start_row + r_idx
            val = ws[f"{col_letter}{row_num}"].value
            if val is not None:
                # 文字列に変換して描画
                text = str(val)
                draw.text((x + 4, y + 2), text, font=font, fill=(20,20,20))
            x += c_w
        y += r_h

    # BytesIOとして返す
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf

def _pixel_to_cell_anchor(ws, x_px, y_px):
    """
    Approximate pixel to cell anchor mapping (return cell like 'C5').
    This is heuristic: use the same assumptions as render_range_to_image.
    """
    # 累積幅／高さを計算する
    cumw = []
    total = 0
    max_col = 100
    for i in range(1, max_col+1):
        col_letter = get_column_letter(i)
        w = _get_column_pixel_width(ws, col_letter)
        total += w
        cumw.append((total, col_letter))

    cumh = []
    totalh = 0
    max_row = 1000
    for r in range(1, max_row+1):
        h = _get_row_pixel_height(ws, r)
        totalh += h
        cumh.append((totalh, r))

    # 列を検索
    col_letter = cumw[-1][1]
    for total_w, col in cumw:
        if x_px <= total_w:
            col_letter = col
            break
    row_num = cumh[-1][1]
    for total_h, r in cumh:
        if y_px <= total_h:
            row_num = r
            break
    return f"{col_letter}{row_num}"

def generate_xlsx_with_shapes(template_path, shapes, sheet_name, cell_range, save_markers=False):
    """
    shapes: list of {type: 'triangle'|'cross'|'circle'|'check', x: px, y: px}
    Will create markers and add to a copy of the template xlsx and return path to file.
    """
    wb = load_workbook(template_path)
    ws = wb[sheet_name]

    markers_dir = os.path.join("backend","_markers")
    os.makedirs(markers_dir, exist_ok=True)

    # マーカー用PNGを32-64ピクセルにスケーリングして作成配置
    for s in shapes:
        stype = s.get("type")
        x = int(s.get("x", 0))
        y = int(s.get("y", 0))
        icon_name = f"{stype}.png"
        icon_path = os.path.join(ICONS_DIR, icon_name)
        if not os.path.exists(icon_path):
            # 失敗したときはスキップ
            continue
        # エクセル挿入時のマーカーサイズを変更
        marker_copy = os.path.join(markers_dir, f"mk_{uuid.uuid4().hex}_{icon_name}")
        im = Image.open(icon_path)
        # 幅を48ピクセルに拡大
        im = im.resize((48,48), Image.ANTIALIAS)
        im.save(marker_copy)

        # 基準セル
        anchor_cell = _pixel_to_cell_anchor(ws, x, y)
        # 画像を追加
        img = XLImage(marker_copy)
        try:
            ws.add_image(img, anchor_cell)
        except Exception as e:
            # フォールバック
            print("add_image failed:", e)

    # 保存
    out_path = f"/tmp/inspection_{uuid.uuid4().hex}.xlsx"
    wb.save(out_path)
    if save_markers:
        # バックエンド/ドキュメントにマーカーを任意で保持
        pass
    return out_path
