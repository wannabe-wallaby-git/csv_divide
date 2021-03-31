# -*- coding: utf-8 -*-

import wx
import os
import win32com.client

# 固定値
WIN_NAME = '[v1]ツール名称'
DIVIDE_NUM = 1000000

class FileDropTarget(wx.FileDropTarget):
    """
    FileDropTarget
    """
    def __init__(self, window):
        wx.FileDropTarget.__init__(self)
        self.window = window

    def OnDropFiles(self, x, y, files):
        self.window.ProcTargetFile(files[0])
        return 0

class App(wx.Frame):
    """
    App(wx.Frame)
    """
    def __init__(self, parent, id, title):
        wx.Frame.__init__(self, parent, id, title, size=(300,300), style=wx.CAPTION|wx.CLOSE_BOX|wx.CLIP_CHILDREN)

        self.pnl_main = wx.Panel(self, wx.ID_ANY)

        # テキストボックス（csvファイルをDnDする）
        self.txt_drop = wx.TextCtrl(self.pnl_main, wx.ID_ANY, 'ここにcsvファイルをドロップしてください。', style=wx.TE_READONLY|wx.TE_MULTILINE)
        self.txt_drop.SetBackgroundColour('#CCCCCC')

        # 閉じるボタン
        self.btn_close = wx.Button(self.pnl_main, wx.ID_CLOSE)

        # Excelを作成し次第開く
        self.chk_open = wx.CheckBox(self.pnl_main, wx.ID_ANY, 'データ分割後、ファイルを開く')

        # Excelファイル作成場所
        self.pnl_excel_file_folder = wx.Panel(self.pnl_main, wx.ID_ANY)
        self.lbl_excel_file_folder = wx.StaticText(self.pnl_excel_file_folder, wx.ID_ANY, 'Excelファイル作成場所')
        self.txt_excel_file_folder = wx.TextCtrl(self.pnl_excel_file_folder, wx.ID_ANY, style=wx.TE_READONLY)
        self.txt_excel_file_folder.SetBackgroundColour('#CCCCCC')
        self.szr_excel_file_folder = wx.BoxSizer(orient=wx.HORIZONTAL)
        self.szr_excel_file_folder.Add(self.lbl_excel_file_folder, flag=wx.ALIGN_CENTER_VERTICAL|wx.RIGHT, border=10)
        self.szr_excel_file_folder.Add(self.txt_excel_file_folder)
        self.pnl_excel_file_folder.SetSizer(self.szr_excel_file_folder)

        self.szr_main = wx.BoxSizer(orient=wx.VERTICAL)
        self.szr_main.Add(self.txt_drop, proportion=1, flag=wx.EXPAND|wx.ALL, border=2)
        self.szr_main.Add(self.pnl_excel_file_folder, flag=wx.GROW|wx.ALL, border=2)
        self.szr_main.Add(self.chk_open, flag=wx.ALL, border=2)
        self.szr_main.Add(self.btn_close, flag=wx.ALIGN_RIGHT|wx.ALL, border=2)

        self.pnl_main.SetSizer(self.szr_main)

        # DnD対象の設定
        self.dnd = FileDropTarget(self)
        self.txt_drop.SetDropTarget(self.dnd)

        # 閉じるボタンイベント設定
        self.btn_close.Bind(wx.EVT_BUTTON, self.close_win)

    # 画面終了
    def close_win(self, event):
        self.Destroy()

    # 行列入れ替え
    def convert_1d_to_2d(self, l, cols):
        return [l[i:i + cols] for i in range(0, len(l), cols)]

    # DnD発生時に実行される関数
    def ProcTargetFile(self, path):
        if os.path.splitext(path)[1] == '.csv':
            self.txt_drop.SetValue('以下のファイルを処理しています。' + os.linesep + '{0}'.format(path))
            self.Refresh()
            
            # csvファイル読み込み（newlineを省略⇒改行コード：\n）
            with open(path, mode='r') as csv_file:
                content = csv_file.read()

            lines = [s.split(',') for s in content.split('\n')]
            line_num = len(lines)
            if type(lines[0]) == str:
                cols = 1
                lines = self.convert_1d_to_2d(lines, cols)
            else:
                cols = len(lines[0])

            # Excelアプリケーションを起動しcsvデータを貼り付ける
            xl = win32com.client.Dispatch("Excel.Application")
            xl.Visible = True
            wb = xl.Workbooks.Add()
            while wb.Worksheets.count > 1:
                wb.Worksheets[0].Delete()
            ws = wb.Worksheets[0]
            ws.Name = '1'

            div_num = line_num // DIVIDE_NUM

            add_num = 1 if line_num % DIVIDE_NUM == 0 else 2

            for i in range(1, div_num + add_num):


                data_offset_from = (i - 1) * DIVIDE_NUM
                if i == div_num + add_num :
                    rows = line_num % DIVIDE_NUM
                else:
                    rows = DIVIDE_NUM
                data_offset_to = data_offset_from + rows

                if len(lines[data_offset_from : data_offset_to]) > 0 and lines[data_offset_from : data_offset_to]:

                    # 次のシートを追加
                    if i != 1:
                        ws = wb.Worksheets.Add(Before=None, After=ws)
                        ws.Name = str(i)

                    try:
                        ws.Range(ws.Cells(1,1), ws.Cells(rows, cols)).Value = lines[data_offset_from : data_offset_to]
                    except Exception as e:
                        print(e)

            # 最初のシートを表示
            wb.Worksheets('1').Activate()

            self.txt_drop.SetValue('ここにcsvファイルをドロップしてください。')

        else:
            self.txt_drop.SetValue('ドロップされたファイルは、csvファイルではありません。')

if __name__ == '__main__':
    app = wx.App()
    frm = App(None, wx.ID_ANY, WIN_NAME)
    frm.Show()
    app.MainLoop()

