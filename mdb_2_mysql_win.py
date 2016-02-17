import wx
import os
import pyodbc
import sqlite3
import pymysql
import adodbapi


def ExportData(idsn='', ihost='localhost', iuser='root', ipasswd='', idb=''):
    #conn_src = adodbapi.connect("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='d:\workspace\wxpython\Increment.mdb'")
    conn_src = pyodbc.connect(idsn)
    #conn_src = pyodbc.connect("DSN=incr")
    cur_src = conn_src.cursor()

    conn_dst = pymysql.connect(host=ihost, port=3306, user=iuser, passwd=ipasswd, db=idb)
    cur_dst = conn_dst.cursor()

    table_list = []
    for row in cur_src.tables():
        if row.table_type == 'TABLE':
            table_list.append(row.table_name)

    #print table_list

    for tbl in table_list:
        sql = "SELECT * FROM %s" % tbl
        cur_src.execute(sql)
        names = [f[0] for f in cur_src.description]
        field_no = len(names)
        row_list = []
        for row in list(cur_src.fetchall()):
            cur_row = []
            for i in range(len(names)):
                cur_row.append(str(row[i]))
                #print row[i]
            row_list.append(cur_row)
        #print row_list

        query = 'DELETE FROM %s' % tbl
        cur_dst.execute(query)
        qry_str = "%s," * (field_no - 1) + '%s' 
        query = 'INSERT INTO %s VALUES (%s)' % (tbl, qry_str)
        #print query 
        for line in row_list:
            #print str(line)
            cur_dst.execute(query, line)

    conn_dst.commit()
    conn_dst.close()

    cur_src.close()
    conn_src.close()
    wx.MessageBox("Data Upload Completed", "Sucess!", style=wx.OK)


class MyFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, -1, "Export Access DB to MySql", size=(530, 200))
        bkg = wx.Panel(self)

        self.STxtSrcFile = wx.StaticText(bkg, label="Source DB/DSN:", size=(85, 25))
        self.TxtSrcFile = wx.TextCtrl(bkg, size=(400, 25))
        self.BtnSrc = wx.Button(bkg, label="...", size=(30, 25))
        self.STxtDstDB = wx.StaticText(bkg, label="Mysql DB Name", size=(85, 25))
        self.TxtDstDB = wx.TextCtrl(bkg, size=(70, 25))
        self.STxtHost = wx.StaticText(bkg, label="Host")
        self.TxtHost = wx.TextCtrl(bkg, size=(120, 25), value="localhost")
        self.STxtUser = wx.StaticText(bkg, label="User")
        self.TxtUser = wx.TextCtrl(bkg, size=(70, 25), value="root")
        self.STxtPasswd = wx.StaticText(bkg, label="Password")
        self.TxtPasswd = wx.TextCtrl(bkg, size=(70, 25), style=wx.TE_PASSWORD)
        self.BtnExp = wx.Button(bkg, label="Export")
        self.DlgSrc = wx.FileDialog(bkg, "Input File Name", os.getcwd(), style=wx.OPEN, wildcard="*.*")

        self.BtnSrc.Bind(wx.EVT_BUTTON, self.BtnSrc_click)
        self.BtnExp.Bind(wx.EVT_BUTTON, self.BtnExp_click)

        box = wx.BoxSizer()
        box.Add(self.STxtSrcFile)
        box.Add(self.TxtSrcFile)
        box.Add(self.BtnSrc)

        box1 = wx.BoxSizer()
        box1.Add(self.STxtDstDB)
        box1.Add(self.TxtDstDB)
        box1.Add(self.STxtHost)
        box1.Add(self.TxtHost)
        box1.Add(self.STxtUser)
        box1.Add(self.TxtUser)
        box1.Add(self.STxtPasswd)
        box1.Add(self.TxtPasswd)

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
        host = self.TxtHost.GetValue()
        user = self.TxtUser.GetValue()
        passwd = self.TxtPasswd.GetValue()
        db = self.TxtDstDB.GetValue()
        dsn = "DSN=" + self.TxtSrcFile.GetValue()
        #print dsn
        ExportData(idsn=dsn, ihost=host, iuser=user, ipasswd=passwd, idb=db)

app = wx.App()
frame = MyFrame()
app.MainLoop()
