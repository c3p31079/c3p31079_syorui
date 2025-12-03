# Excel の指定範囲を PNG として保存
import win32com.client as win32


def render_range_to_png(excel_path, sheet_name, cell_range, out_png):

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False

    wb = excel.Workbooks.Open(excel_path)
    ws = wb.Worksheets(sheet_name)

    # 範囲を画像としてコピー
    ws.Range(cell_range).CopyPicture()

    # チャートに貼り付け → PNG 保存
    chart = excel.Charts.Add()
    chart.Paste()
    chart.Export(out_png)

    chart.Delete()
    wb.Close(SaveChanges=False)
    excel.Quit()
