import wx
import datetime
import os
import pyodbc
import sqlite3
import pymysql
import adodbapi


class MyFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, -1, "Export Access DB to MySql", size=(530, 200))
        bkg = wx.Panel(self)

        self.STxtSrcFile = wx.StaticText(bkg, label="Source DSN:", size=(85, 25))
        self.TxtSrcFile = wx.TextCtrl(bkg, size=(170, 25))
        self.STxtSrcPWD = wx.StaticText(bkg, label="Source Password:", size=(115, 25))
        self.TxtSrcPWD = wx.TextCtrl(bkg, size=(130, 25), style=wx.TE_PASSWORD)
        self.STxtDstDB = wx.StaticText(bkg, label="Mysql DB Name", size=(85, 25))
        self.TxtDstDB = wx.TextCtrl(bkg, size=(70, 25))
        self.STxtHost = wx.StaticText(bkg, label="Host")
        self.TxtHost = wx.TextCtrl(bkg, size=(120, 25), value="localhost")
        self.STxtUser = wx.StaticText(bkg, label="User")
        self.TxtUser = wx.TextCtrl(bkg, size=(70, 25), value="root")
        self.STxtPasswd = wx.StaticText(bkg, label="Password")
        self.TxtPasswd = wx.TextCtrl(bkg, size=(90, 25), style=wx.TE_PASSWORD)
        self.BtnExp = wx.Button(bkg, label="Export", size=(90, 35))
        self.BtnExit = wx.Button(bkg, label="Exit", size=(90, 35))
        self.DlgSrc = wx.FileDialog(bkg, "Input File Name", os.getcwd(), style=wx.OPEN, wildcard="*.*")
        self.STxtCurProcessLabel = wx.StaticText(bkg, label="Current Process:")
        self.STxtCurProcess = wx.StaticText(bkg, label="......................................................", size=(200, 25))

        #self.BtnSrc.Bind(wx.EVT_BUTTON, self.BtnSrc_click)
        self.BtnExp.Bind(wx.EVT_BUTTON, self.BtnExp_click)
        self.BtnExit.Bind(wx.EVT_BUTTON, self.BtnExit_click)

        box = wx.BoxSizer()
        box.Add(self.STxtSrcFile, 0, wx.LEFT, 0)
        box.Add(self.TxtSrcFile, 0, wx.LEFT, 2)
        box.Add(self.STxtSrcPWD, 0, wx.LEFT, 29)
        box.Add(self.TxtSrcPWD)

        box1 = wx.BoxSizer()
        box1.Add(self.STxtDstDB, 0, wx.LEFT, 0)
        box1.Add(self.TxtDstDB, 0, wx.LEFT, 0)
        box1.Add(self.STxtHost, 0, wx.LEFT, 2)
        box1.Add(self.TxtHost, 0, wx.LEFT, 0)
        box1.Add(self.STxtUser, 0, wx.LEFT, 2)
        box1.Add(self.TxtUser, 0, wx.LEFT, 0)
        box1.Add(self.STxtPasswd, 0, wx.LEFT, 2)
        box1.Add(self.TxtPasswd, 0, wx.LEFT, 0)

        box2 = wx.BoxSizer()
        box2.Add(self.BtnExp, 0, wx.LEFT, 0)
        box2.Add(self.STxtCurProcessLabel, 0, wx.LEFT, 10)
        box2.Add(self.STxtCurProcess, 0, wx.LEFT, 2)
        box2.Add(self.BtnExit, 0, wx.LEFT, 40)

        vbox = wx.BoxSizer(wx.VERTICAL)
        vbox.Add(box, 0, wx.ALL, 7)
        vbox.Add(box1, 0, wx.ALL, 7)
        vbox.Add(box2, 0, wx.ALL, 7)

        # bkg.SetSizer(vbox)
        bkg.SetSizerAndFit(vbox)

        self.Show()
        self.CenterOnScreen()

    def ExportData(self, idsn='', ihost='localhost', iuser='root', ipasswd='', idb=''):
        #conn_src = adodbapi.connect("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='d:\workspace\wxpython\Increment.mdb'")
        conn_src = pyodbc.connect(idsn)
        #conn_src = pyodbc.connect(idsn+";PWD=hyparabola")
        #conn_src = pyodbc.connect("DSN=incr")
        cur_src = conn_src.cursor()

        conn_dst = pymysql.connect(host=ihost, port=3306, user=iuser, passwd=ipasswd)
        cur_dst = conn_dst.cursor()
        cur_dst.execute("CREATE DATABASE IF NOT EXISTS " + idb)
        cur_dst.execute("USE " + idb)

        table_list = [cur_table.table_name for cur_table in cur_src.tables() if cur_table.table_type == 'TABLE']
        #print table_list

        for tbl in table_list:
            sql_count = "SELECT COUNT(*) FROM %s" % tbl
            cur_src.execute(sql_count)
            row_number = cur_src.fetchone()[0]
            self.STxtCurProcess.SetLabel(tbl + "Table Name with Rows..............." + str(row_number))
            self.STxtCurProcess.SetLabel(tbl + "-Rows-" + str(row_number))

            sql = "SELECT * FROM %s" % tbl
            cur_src.execute(sql)

            #Get Field Names and types for Table Creation
            field_names = [f[0] for f in cur_src.description]
            field_types = [f[1] for f in cur_src.description]
            field_lengths = [f[3] for f in cur_src.description]
            field_description = zip(field_names, field_types, field_lengths)
            #print field_description

            #Create TABLE
            create_sql = "CREATE TABLE " + tbl + " ("

            for field, type, length in field_description:
                if type == int:
                    create_sql += "`" + field + "` DOUBLE, " 
                elif type == float:
                    create_sql += "`" + field + "` DOUBLE, " 
                elif type == unicode:
                    create_sql += "`" + field + "` VARCHAR(" + str(length) + "), " 
                elif type == datetime.datetime:
                    create_sql += "`" + field + "` DATETIME, " 
                elif type == bool:
                    create_sql += "`" + field + "` BOOLEAN, "
                else:
                    create_sql += "`" + field + "` VARCHAR(255), " 


            cur_dst.execute(create_sql[:-2] + ")")

            #INSERT DATA into Table
            if row_number > 0:
                field_no = len(field_names)
                row_list = []
                for row in list(cur_src.fetchall()):
                    cur_row = []
                    for i in range(len(field_names)):
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
        dsn = "DSN=" + self.TxtSrcFile.GetValue() + ";PWD=" + self.TxtSrcPWD.GetValue()
        #print dsn
        self.ExportData(idsn=dsn, ihost=host, iuser=user, ipasswd=passwd, idb=db)

    def BtnExit_click(self, event):
        self.Close()

app = wx.App()
frame = MyFrame()
app.MainLoop()
