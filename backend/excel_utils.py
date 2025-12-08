import io
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage

# 画像アイコンのディレクトリ
ICON_DIR = "backend/icons"

# アイコンマッピング（選択値 → アイコンファイル）
ICON_MAP = {
    "B": "triangle.png",
    "C": "none.png",
    "check": "check.png",
    "circle": "circle.png"
}

# 点検項目ごとの座標マップ（自由に調整可能）
COORD_MAP = {
    "揺動部（チェーン・ロープ）_ねじれ": (2, 3),
    "揺動部（チェーン・ロープ）_変形": (4, 3),
    "揺動部（チェーン・ロープ）_破損": (6, 3),
    "揺動部（チェーン・ロープ）_ほつれ": (8, 3),
    "揺動部（チェーン・ロープ）_断線": (10, 3),
    "揺動部（チェーン・ロープ）_摩耗": (12, 3),
    # 座板・座面（タイヤ）
    "揺動部（座板・座面(タイヤ)）_ヒビ": (2, 5),
    "揺動部（座板・座面(タイヤ)）_割れ": (4, 5),
    "揺動部（座板・座面(タイヤ)）_湾曲等変形": (6, 5),
    "揺動部（座板・座面(タイヤ)）_破損": (8, 5),
    "揺動部（座板・座面(タイヤ)）_腐朽": (10, 5),
    "揺動部（座板・座面(タイヤ)）_金具の摩耗": (12, 5),
    "揺動部（座板・座面(タイヤ)）_ボルト、袋ナットの緩み": (14, 5),
    "揺動部（座板・座面(タイヤ)）_破損、腐食、欠落": (16, 5),
    # 安全柵
    "安全柵_ぐらつき": (2, 7),
    "安全柵_破損": (4, 7),
    "安全柵_変形": (6, 7),
    "安全柵_腐食": (8, 7),
    "安全柵_〔接合部・ボルト〕緩み": (10, 7),
    "安全柵_欠落": (12, 7),
    # その他
    "その他_異物": (2, 9),
    "その他_落書き": (4, 9),
    # 基礎
    "基礎_基礎が露出": (2, 11),
    "基礎_亀裂": (4, 11),
    "基礎_破損": (6, 11),
    # 地表部・安全柵内
    "地表部・安全柵内_大きな凹凸": (2, 13),
    "地表部・安全柵内_石や根の露出": (4, 13),
    "地表部・安全柵内_異物": (6, 13),
    "地表部・安全柵内_マットのめくれ": (8, 13),
    "地表部・安全柵内_破損": (10, 13),
    "地表部・安全柵内_樹木の枝": (12, 13),
    # チェック系（例：対応方針、実施済/予定など）は座標別に追加可能（今は試しなので後でやるよ）
}

# Excel生成関数
def generate_excel(data: dict) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "点検結果"

    # ヘッダ行
    ws.append(["点検部位", "項目", "選択"])

    # データ書き込み＋アイコン配置
    for key, value in data.items():
        ws.append([key.split("_")[0], key.split("_")[1], value])

        if value in ICON_MAP:
            icon_file = os.path.join(ICON_DIR, ICON_MAP[value])
            if os.path.exists(icon_file):
                img = Image(icon_file)
                coord = COORD_MAP.get(key)
                if coord:
                    col_letter = get_column_letter(coord[0])
                    ws.add_image(img, f"{col_letter}{coord[1]}")

    # バイト配列に書き出し
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()