import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Font
import os
import io

# =============================================
# 画像ファイルパス
# =============================================
CHECK_IMG = "static/img/check.png"
CIRCLE_IMG = "static/img/circle.png"

# =============================================
# xy座標調整用（セル基準の微調整）
# =============================================
xy_coords = {
    "柱・梁（本体）": {
        "ぐらつき": (0, 0),
        "破損": (0, 0),
        "変形": (0, 0),
        "腐食 (腐朽)": (0, 0),
        "接合部の緩み": (0, 0)
    },
    "接合部(継ぎ手)": {
        "破損": (0,0),
        "変形": (0,0),
        "腐食": (0,0),
        "ボルトの緩み": (0,0),
        "欠落": (0,0)
    },
    # 他の点検部位も同様に追加可能
}

# =============================================
# チェックマーク画像をセルに追加
# =============================================
def add_checkmark(ws, cell, x_offset=0, y_offset=0):
    """
    チェックボックスのチェックをExcelに反映
    画像が存在すればcheck.png、なければセルに●
    """
    if os.path.exists(CHECK_IMG):
        img = Image(CHECK_IMG)
        # ここに追加で座標補正可能
        img.anchor = cell.coordinate
        ws.add_image(img)
    else:
        ws[cell.coordinate] = "●"

# =============================================
# circle画像をセルに追加（対応予定時期など）
# =============================================
def add_circle(ws, cell, x_offset=0, y_offset=0):
    if os.path.exists(CIRCLE_IMG):
        img = Image(CIRCLE_IMG)
        img.anchor = cell.coordinate
        ws.add_image(img)
    else:
        ws[cell.coordinate] = "〇"

# =============================================
# Excel生成メイン関数
# =============================================
def generate_inspection_excel(form_data):
    """
    form_data: HTMLから送信されたJSON
    {
        "search_park": "...",
        "inspection_year": "...",
        "install_year_num": "...",
        "点検部位・項目": {...},
        "action_grease": true,
        "overall_result": "A",
        ...
    }
    """
    # -----------------------------------------
    # ワークブック作成
    # -----------------------------------------
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "点検チェックシート"

    # -----------------------------------------
    # セル幅・フォント調整（必要に応じ追加可能）
    # -----------------------------------------
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 30
    ws.column_dimensions['E'].width = 30
    ws.column_dimensions['F'].width = 15
    ws.column_dimensions['G'].width = 15
    ws.column_dimensions['H'].width = 20

    # -----------------------------------------
    # 公園名・作成者・点検年度・設置年度
    # -----------------------------------------
    ws["C2"] = form_data.get("search_park", "")
    ws["F2"] = form_data.get("search_author", "")
    ws["H2"] = form_data.get("inspection_year", "")
    ws["H3"] = form_data.get("install_year_num", "")

    # -----------------------------------------
    # 点検部位と項目の反映
    # -----------------------------------------
    inspection_parts = {
        "柱・梁（本体）": ["ぐらつき","破損","変形","腐食 (腐朽)","接合部の緩み"],
        "接合部(継ぎ手)": ["破損","変形","腐食","ボルトの緩み","欠落"],
        "吊金具": ["破損","変形","腐食","異音","金具本体のずれ","摩耗","ボルトの緩み","欠落"],
        "揺動部（チェーン・ロープ）": ["ねじれ","変形","破損","ほつれ","断線","摩耗"],
        "揺動部（座板・座面(タイヤ)）": ["ヒビ","割れ","湾曲等変形","破損","腐朽","金具の摩耗","ボルト、袋ナットの緩み","欠落"],
        "安全柵": ["ぐらつき","破損","変形","腐食","〔接合部・ボルト〕緩み","欠落"],
        "その他": ["異物","落書き"],
        "基礎": ["基礎が露出","亀裂","破損"],
        "地表部・安全柵内": ["大きな凹凸","石や根の露出","異物","マットのめくれ","破損","樹木の枝"]
    }

    start_row = 6
    for part_name, items in inspection_parts.items():
        for idx, item in enumerate(items):
            row = start_row
            ws[f"B{row}"] = part_name if idx==0 else ""  # 点検部位は最初だけ
            ws[f"C{row}"] = ""  # 必要に応じ編集可
            ws[f"D{row}"] = item
            ws[f"E{row}"] = ""  # 項目詳細
            start_row += 1

    # -----------------------------------------
    # チェックボックス反映（例）
    # -----------------------------------------
    for key, val in form_data.items():
        if key.startswith("action_") and val:
            # TODO: 実際のセル位置に対応させる
            row = 6  # 仮: F列にマーク
            add_checkmark(ws, ws.cell(row=row, column=6))

    # -----------------------------------------
    # 総合結果
    # -----------------------------------------
    overall = form_data.get("overall_result")
    if overall:
        # TODO: 実際セル位置にマーク
        add_checkmark(ws, ws["F13"])

    # -----------------------------------------
    # 対応予定時期（circle.png）
    # -----------------------------------------
    month = form_data.get("response_month")
    period = form_data.get("period")
    if month:
        ws["H6"] = f"{month}月"
        if period:
            add_circle(ws, ws["H6"])

    # -----------------------------------------
    # 備考・所見
    # -----------------------------------------
    ws["F10"] = form_data.get("observations","")
    ws["H12"] = form_data.get("remarks","")

    # -----------------------------------------
    # ここに追加可能: 他のセルに文字列・画像を追加
    # ws["D6"] = "例: 自由入力"
    # add_checkmark(ws, ws["F7"])
    # -----------------------------------------

    # -----------------------------------------
    # Excelをバイトに変換して返す
    # -----------------------------------------
    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream.getvalue()
