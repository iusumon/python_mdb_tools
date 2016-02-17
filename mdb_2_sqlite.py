#import adodbapi
#database = r"d:\workspace\wxpython\Increment.mdb"
#constr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=%s;" % database

import wx
import os
import pyodbc
import sqlite3


def ExportData(src_db, dst_db):
    conn_src = pyodbc.connect("DSN=incr")
    cur_src = conn_src.cursor()

    conn_dst = sqlite3.connect(r'%s' % dst_db)
    cur_dst = conn_dst.cursor()

    table_list = []
    for row in cur_src.tables():
        if row.table_type == 'TABLE':
            table_list.append(row.table_name)

    #print table_list

    for tbl in table_list[0:1]:
        sql = "SELECT * FROM %s" % tbl
        cur_src.execute(sql)
        names = [f[0] for f in cur_src.description]
        row_list = []
        for row in list(cur_src.fetchall()):
            cur_row = []
            for i in range(len(names)):
                cur_row.append(row[i])
                #print row[i]
            row_list.append(cur_row)

        query = 'DELETE FROM %s' % tbl
        cur_dst.execute(query)
        query = 'INSERT INTO %s VALUES (?,?,?,?,?,?,?,?,?,?)' % tbl
        for line in row_list:
            cur_dst.execute(query, line)

        conn_dst.commit()
        conn_dst.close()
    cur_src.close()
    conn_src.close()


class MyFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, -1, "Export Access DB to MySql", size=(530, 200))
        bkg = wx.Panel(self)

        self.STextSrcFile = wx.StaticText(bkg, label="Source DB:", size=(100, 30))
        self.TxtSrcFile = wx.TextCtrl(bkg, size=(400, 30))
        self.BtnSrc = wx.Button(bkg, label="...", size=(30, 25))
        self.STextDstFile = wx.StaticText(bkg, label="Destination DB:")
        self.TxtDstFile = wx.TextCtrl(bkg, size=(400, 30))
        self.BtnDst = wx.Button(bkg, label="...", size=(30, 25))
        self.BtnExp = wx.Button(bkg, label="Export")
        self.DlgSrc = wx.FileDialog(bkg, "Input File Name", os.getcwd(), style=wx.OPEN, wildcard="*.*")

        self.BtnSrc.Bind(wx.EVT_BUTTON, self.BtnSrc_click)
        self.BtnDst.Bind(wx.EVT_BUTTON, self.BtnDst_click)
        self.BtnExp.Bind(wx.EVT_BUTTON, self.BtnExp_click)

        box = wx.BoxSizer()
        box.Add(self.STextSrcFile)
        box.Add(self.TxtSrcFile)
        box.Add(self.BtnSrc)

        box1 = wx.BoxSizer()
        box1.Add(self.STextDstFile)
        box1.Add(self.TxtDstFile)
        box1.Add(self.BtnDst)

        box2 = wx.BoxSizer()
        box2.Add(self.BtnExp)

        vbox = wx.BoxSizer(wx.VERTICAL)
        vbox.Add(box)
        vbox.Add(box1)
        vbox.Add(box2)

        # bkg.SetSizer(vbox)
        bkg.SetSizerAndFit(vbox)

        self.Show()
        self.CenterOnScreen()

    def BtnSrc_click(self, event):
        dlg = wx.FileDialog(self, "Open Access database", os.getcwd(), style=wx.OPEN, wildcard="*.mdb")
        if dlg.ShowModal() == wx.ID_OK:
            self.src_db = dlg.GetPath()
            self.TxtSrcFile.SetValue(dlg.GetPath())
        dlg.Destroy()

    def BtnDst_click(self, event):
        dlg = wx.FileDialog(self, "Open SQLite database", os.getcwd(), style=wx.OPEN, wildcard="*.*")
        if dlg.ShowModal() == wx.ID_OK:
            self.dst_db = dlg.GetPath()
            self.TxtDstFile.SetValue(dlg.GetPath())
        dlg.Destroy()

    def BtnExp_click(self, event):
        ExportData(self.src_db, self.dst_db)

app = wx.App()
frame = MyFrame()
app.MainLoop()
