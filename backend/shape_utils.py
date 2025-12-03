# JSON の座標情報に基づいて Excel に図形を配置する
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage


def place_shapes_on_excel(template_path, output_path,
                          items, coord_map, check_coord_map):

    wb = load_workbook(template_path)
    ws = wb.active

    for item in items:
        part = item["part"]          # 点検部位
        category = item["category"]  # 項目
        mark = item["mark"]          # △ × ○ ✓
        checks = item["checks"]      # 複数のチェック項目

        key = f"{part}:{category}"
        if key not in coord_map:
            continue

        base = coord_map[key]
        x = base["x"]
        y = base["y"]

        # 図形画像のパス
        if mark == "△":
            img_path = "backend/shapes/triangle.png"
        elif mark == "×":
            img_path = "backend/shapes/cross.png"
        elif mark == "○":
            img_path = "backend/shapes/circle.png"
        elif mark == "✓":
            img_path = "backend/shapes/check.png"
        else:
            continue

        img = XLImage(img_path)
        img.width = 22
        img.height = 22

        img.anchor = f"A1"
        img.drawing.left = x
        img.drawing.top = y

        ws.add_image(img)

        # 追加チェック（✓ × の別の座標へ配置）
        for c in checks:
            if c in check_coord_map:
                chk = check_coord_map[c]
                cimg = XLImage("backend/shapes/check.png")
                cimg.width = 18
                cimg.height = 18
                cimg.anchor = "A1"
                cimg.drawing.left = chk["x"]
                cimg.drawing.top = chk["y"]
                ws.add_image(cimg)

    wb.save(output_path)
