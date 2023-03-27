
import wx


class LoginFrame(wx.Frame):
    def __init__(self, parent, title):
        super(LoginFrame, self).__init__(parent, title=title, size=(300, 200))
        self.InitUI()

    def InitUI(self):
        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        st1 = wx.StaticText(panel, label='Username')
        hbox1.Add(st1, flag=wx.RIGHT, border=8)
        self.tc1 = wx.TextCtrl(panel)
        hbox1.Add(self.tc1, proportion=1)
        vbox.Add(hbox1, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, border=10)

        hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        st2 = wx.StaticText(panel, label='Password')
        hbox2.Add(st2, flag=wx.RIGHT, border=8)
        self.tc2 = wx.TextCtrl(panel, style=wx.TE_PASSWORD)
        hbox2.Add(self.tc2, proportion=1)
        vbox.Add(hbox2, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, border=10)

        hbox3 = wx.BoxSizer(wx.HORIZONTAL)
        sb = wx.StaticText(panel, label='')
        hbox3.Add(sb, proportion=1)
        btn1 = wx.Button(panel, label='Log in', size=(70, 30))
        hbox3.Add(btn1, flag=wx.RIGHT, border=10)
        vbox.Add(hbox3, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, border=10)

        btn1.Bind(wx.EVT_BUTTON, self.OnLogin)
        panel.SetSizer(vbox)

    def OnLogin(self, event):
        username = self.tc1.GetValue()
        password = self.tc2.GetValue()
        if username == 'admin' and password == 'admin':
            self.Destroy()
            self.parent = FunctionFrame(parent=None, title='Function')
            self.parent.Show(True)
        else:
            dlg = wx.MessageDialog(None, 'The username or password you entered is incorrect.', 'Error',
                                   wx.OK | wx.ICON_ERROR)
            dlg.ShowModal()
            dlg.Destroy()


class FunctionFrame(wx.Frame):
    def __init__(self, parent, title):
        super(FunctionFrame, self).__init__(parent, title=title, size=(300, 200))
        self.InitUI()

    def InitUI(self):
        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        st1 = wx.StaticText(panel, label='Welcome!')
        hbox1.Add(st1,
                  flag=wx.RIGHT | wx.ALIGN_CENTER_VERTICAL,
                  border=8)
        vbox.Add(hbox1, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP,
                 border=10)

        hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        btn1 = wx.Button(panel, label='Function 1', size=(70, 30))
        hbox2.Add(btn1,
                  flag=wx.RIGHT,
                  border=10)
        btn2 = wx.Button(panel, label='Function 2', size=(70, 30))
        hbox2.Add(btn2,
                  flag=wx.RIGHT,
                  border=10)
        btn3 = wx.Button(panel, label='Function 3', size=(70, 30))
        hbox2.Add(btn3)
        vbox.Add(hbox2, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP,
                 border=10)

        hbox3 = wx.BoxSizer(wx.HORIZONTAL)
        sb = wx.StaticText(panel, label='')
        hbox3.Add(sb, proportion=1)
        btn4 = wx.Button(panel, label='Log out', size=(70, 30))
        hbox3.Add(btn4, flag=wx.RIGHT, border=10)
        vbox.Add(hbox3, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, border=10)

        btn1.Bind(wx.EVT_BUTTON, self.OnFunction1)
        btn2.Bind(wx.EVT_BUTTON, self.OnFunction2)
        btn3.Bind(wx.EVT_BUTTON, self.OnFunction3)
        btn4.Bind(wx.EVT_BUTTON, self.OnLogOut)
        panel.SetSizer(vbox)

    def OnFunction1(self, event):
        dlg = wx.MessageDialog(None, 'Function 1 is clicked!', 'Information', wx.OK | wx.ICON_INFORMATION)
        dlg.ShowModal()
        dlg.Destroy()
        self.Destroy()

    def OnFunction2(self, event):
        dlg = wx.MessageDialog(None, 'Function 2 is clicked!', 'Information', wx.OK | wx.ICON_INFORMATION)
        dlg.ShowModal()
        dlg.Destroy()

    def OnFunction3(self, event):
        dlg = wx.MessageDialog(None, 'Function 3 is clicked!', 'Information', wx.OK | wx.ICON_INFORMATION)
        dlg.ShowModal()
        dlg.Destroy()

    def OnLogOut(self, event):
        self.Destroy()
        self.parent = LoginFrame(parent=None, title='Login')
        self.parent.Show(True)


if __name__ == '__main__':
    app = wx.App()
    login = LoginFrame(parent=None, title='Login')
    login.Show()
    app.MainLoop()

