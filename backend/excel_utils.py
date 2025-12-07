# Excel上にピクセル座標で画像を配置するヘルパー
from openpyxl.utils import get_column_letter
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor, Extent
from openpyxl.drawing.image import Image as XLImage

EMU_PER_PIXEL = 9525  # EMU 単位（OpenXML の仕様でおおよそこの値）

def col_width_to_pixels(width):
    # openpyxlの column_dimensions.width は "Excelの列幅" 単位
    # 簡易換算：幅 * 7 + 5 の目安（完璧ではないが調整可能）
    if width is None:
        width = 8.43
    return int(width * 7 + 5)

def row_height_to_pixels(height):
    # row_dimensions.height はポイント（pt）. 1pt ≒ 1.333px (おおよその換算)
    if height is None:
        # Excelのデフォルト行高は約15ポイント（環境で違う）
        return 15
    return int(height)

def pixel_to_cell(ws, x, y):
    """
    (x,y) px が Excel のどのセル上（左上原点0,0）に入るかを計算し、
    (col_index, row_index, offset_x_emu, offset_y_emu) を返す。
    col_index, row_index は 1-origin
    """
    total_x = 0
    col = 1
    # カラム幅は column_dimensions[col_letter].width
    while True:
        col_letter = get_column_letter(col)
        col_dim = ws.column_dimensions.get(col_letter)
        width = None
        if col_dim is not None and col_dim.width is not None:
            width = col_dim.width
        w_px = col_width_to_pixels(width)
        if total_x + w_px > x:
            offset_x = x - total_x
            break
        total_x += w_px
        col += 1

        # セーフティ
        if col > 200: break

    total_y = 0
    row = 1
    while True:
        row_dim = ws.row_dimensions.get(row)
        height = None
        if row_dim is not None and row_dim.height is not None:
            height = row_dim.height
        h_px = row_height_to_pixels(height)
        if total_y + h_px > y:
            offset_y = y - total_y
            break
        total_y += h_px
        row += 1

        if row > 1000: break

    # EMU に変換
    off_x_emu = int(offset_x * EMU_PER_PIXEL)
    off_y_emu = int(offset_y * EMU_PER_PIXEL)

    return col, row, off_x_emu, off_y_emu

def place_image_xy(ws, img_path, x_px, y_px):
    """
    ws: worksheet
    img_path: 画像パス
    x_px, y_px: 左上基準のピクセル座標
    """
    img = XLImage(img_path)
    # width/height がない場合はデフォルトを設定
    try:
        img_w = img.width
        img_h = img.height
    except Exception:
        img_w, img_h = 32, 32

    col, row, off_x_emu, off_y_emu = pixel_to_cell(ws, x_px, y_px)
    # マーカーを作る
    marker = AnchorMarker(col=col-1, colOff=off_x_emu, row=row-1, rowOff=off_y_emu)
    # 画像の ext（幅・高さ）を EMU に
    ext = Extent(cx=int(img_w * EMU_PER_PIXEL), cy=int(img_h * EMU_PER_PIXEL))
    anchor = OneCellAnchor(_from=marker, ext=ext)
    img.anchor = anchor
    ws.add_image(img)
