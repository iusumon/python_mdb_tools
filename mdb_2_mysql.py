#!/usr/bin/env python
import wx
import sys, getpass, shlex, subprocess, re, os


def raw_default(prompt, dflt=None):
    prompt = "%s [%s]: " % (prompt, dflt)
    res = raw_input(prompt)
    if not res and dflt:
        return dflt
    return res


def get_external_command_output(command):
    args = shlex.split(command)
    ret = subprocess.check_output(args)  # this needs Python 2.7 or higher
    return ret


def get_pipe_series_output(commands, stdinput=None):
    # Python arrays indexes are zero-based, i.e. an array is indexed from 0 to len(array)-1.
    # The range/xrange commands, by default, start at 0 and go to one less than the maximum specified.
    # print commands
    processes = []
    for i in xrange(len(commands)):
        if (i == 0):  # first processes
            processes.append(subprocess.Popen(shlex.split(commands[i]), stdin=subprocess.PIPE, stdout=subprocess.PIPE))
        else:  # subsequent ones
            processes.append(subprocess.Popen(shlex.split(commands[i]), stdin=processes[i-1].stdout, stdout=subprocess.PIPE))
    return processes[len(processes)-1].communicate(stdinput)[0]  # communicate() returns a tuple; 0=stdout, 1=stderr; so this returns stdout


def replace_type_in_sql(sql, fromstr, tostr):
    whitespaceregroup = "([\ \t\n]+)"
    whitespaceorcommaregroup = "([\ \t\),\n]+)"
    rg1 = "\g<1>"
    rg2 = "\g<2>"
    return re.sub(whitespaceregroup + fromstr + whitespaceorcommaregroup, rg1 + tostr + rg2, sql, 0, re.MULTILINE | re.IGNORECASE)


def ExportData(src_db, dst_db, hostname="127.0.0.1", user="root", passwd="torva"):
    mdbfile = src_db
    tempfile = "TEMP.txt"
    host = hostname  # not "localhost" but 127.0.0.1
    port = 3306
    user = user
    password = passwd
    mysqldb = dst_db

    # print "Getting list of tables"
    tablecmd = "mdb-tables -1 " + mdbfile
    tables = get_external_command_output(tablecmd).splitlines()
    print tables

    # print "Creating new database"
    createmysqldbcmd = "mysqladmin create %s --host=%s --port=%s --user=%s --password=%s" % (mysqldb, host, port, user, password)
    print get_external_command_output(createmysqldbcmd)

    # print "Shipping table definitions (sanitized), converted to MySQL types, through some syntax filters, to MySQL"
    schemacmd = "mdb-schema " + mdbfile + " mysql"

    schemasyntax = get_external_command_output(schemacmd)
    schemasyntax = re.sub("^COMMENT ON COLUMN.*$", "", schemasyntax, 0, re.MULTILINE)

    # print "-----------------"
    # print schemasyntax
    # print "-----------------"

    mysqlcmd = "mysql --host=%s --port=%s --database=%s --user=%s --password=%s" % (host, port, mysqldb, user, password)
    print get_pipe_series_output([mysqlcmd], schemasyntax)

    for t in tables:
        print "Processing table", t
        #exportcmd = "mdb-export -I mysql -D \"%Y-%m-%d %H:%M:%S\" " + mdbfile + " " + t
            # -I backend: INSERT statements, not CSV
            # -D: date format
            #     MySQL's DATETIME field has this format: "YYYY-MM-DD HH:mm:SS"
            #     so we want this from the export
        #print get_pipe_series_output( [exportcmd, semicolonfilter, groupfilter, mysqlcmd] )
        #print get_pipe_series_output( [exportcmd, mysqlcmd] )

        os.system('echo "SET autocommit=0;" > ' + tempfile)
        exportcmd = 'mdb-export -I mysql -D "%Y-%m-%d" ' + mdbfile + ' "' + t + '" >> ' + tempfile
        os.system(exportcmd)
        os.system('echo "COMMIT;" >> ' + tempfile)
        importcmd = mysqlcmd + " < " + tempfile
        os.system(importcmd)

    # print "Finished."


class MyFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, -1, "Export Access DB to MySql", size=(530, 200))
        bkg = wx.Panel(self)

        self.STextSrcDB = wx.StaticText(bkg, label="Source MDB:", size=(110, 30))
        self.TxtSrcDB = wx.TextCtrl(bkg, size=(390, 30))
        self.BtnSrc = wx.Button(bkg, label="...", size=(30, 25))

        self.STextDstDB = wx.StaticText(bkg, label="mysql Database:")
        self.TxtDstDB = wx.TextCtrl(bkg, size=(70, 30))
        self.STextHost = wx.StaticText(bkg, label="IP", size=(30, 30))
        self.TxtHost = wx.TextCtrl(bkg, size=(70, 30), value="127.0.0.1")
        self.STextUser = wx.StaticText(bkg, label="User", size=(45, 30))
        self.TxtUser = wx.TextCtrl(bkg, size=(70, 30), value="root")
        self.STextPasswd = wx.StaticText(bkg, label="Password", size=(65, 30))
        self.TxtPasswd = wx.TextCtrl(bkg, size=(70, 30), style=wx.TE_PASSWORD)

        self.BtnExp = wx.Button(bkg, label="Export")
        self.DlgSrc = wx.FileDialog(bkg, "Input File Name", os.getcwd(), style=wx.OPEN, wildcard="*.*")

        self.BtnSrc.Bind(wx.EVT_BUTTON, self.BtnSrc_click)
        self.BtnExp.Bind(wx.EVT_BUTTON, self.BtnExp_click)

        box = wx.BoxSizer()
        box.Add(self.STextSrcDB)
        box.Add(self.TxtSrcDB)
        box.Add(self.BtnSrc)

        box1 = wx.BoxSizer()
        box1.Add(self.STextDstDB)
        box1.Add(self.TxtDstDB)
        box1.Add(self.STextHost)
        box1.Add(self.TxtHost)
        box1.Add(self.STextUser)
        box1.Add(self.TxtUser)
        box1.Add(self.STextPasswd)
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
            self.TxtSrcDB.SetValue(dlg.GetPath())
        dlg.Destroy()

    def ValidateData(self):
        if not self.TxtSrcDB.GetValue().strip():
            return "Source DB is required"
        elif not self.TxtDstDB.GetValue().strip():
            return "mysql DB Name is required"
        elif not self.TxtUser.GetValue().strip():
            return "User is required"
        elif not self.TxtPasswd.GetValue().strip():
            return "Password is required"
        else:
            return "valid"

    def BtnExp_click(self, event):
        valid_msg = self.ValidateData()
        try:
            if valid_msg == "valid":
                ExportData(self.src_db, self.TxtDstDB.GetValue(), self.TxtHost.GetValue(), self.TxtUser.GetValue(), self.TxtPasswd.GetValue())
                wx.MessageBox("Database exported successfully", "Completed", style=wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)
            else:
                wx.MessageBox(valid_msg, "Information Missing!", style=wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)
        except Exception as e:
            print e
            if "exists" in str(e):
                wx.MessageBox("Database already exists", "Error!", style=wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

app = wx.App()
frame = MyFrame()
app.MainLoop()
