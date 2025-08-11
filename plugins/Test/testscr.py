import wx
import wx.lib.scrolledpanel as sp

class Notebooks(wx.Notebook):
    def __init__(self, parent):
        wx.Notebook.__init__(self, parent, wx.ID_ANY, style=wx.NB_TOP)

        tab1 = wx.Panel(self, wx.ID_ANY)
        tab2 = wx.Panel(self, wx.ID_ANY)

        btn1 = wx.Button(tab1, wx.ID_ANY, "Button1", size=(200,-1))
        btn2 = wx.Button(tab2, wx.ID_ANY, "Button2", size=(200,-1))
        self.AddPage(tab1, 'Tab1')
        self.AddPage(tab2, 'Tab2')

class MyApp(wx.App):
    def OnInit(self):
        self.frame = MyFrame(None, title="Test frame")
        self.SetTopWindow(self.frame)
        self.frame.Show()

        return True

class MyFrame(wx.Frame):
    def __init__(self, parent, title):
        wx.Frame.__init__(self, parent, wx.ID_ANY, title,
                          pos=wx.DefaultPosition, size=wx.DefaultSize,
                          style=wx.DEFAULT_FRAME_STYLE)

        self.panel = sp.ScrolledPanel(self)
        vsizer = wx.BoxSizer(wx.VERTICAL)

        gd01 = wx.TextCtrl(self.panel, wx.ID_ANY, " ")
        gd02 = wx.TextCtrl(self.panel, wx.ID_ANY, " ")
        gd03 = wx.TextCtrl(self.panel, wx.ID_ANY, " ")
        gd04 = wx.TextCtrl(self.panel, wx.ID_ANY, " ")
        gd05 = wx.TextCtrl(self.panel, wx.ID_ANY, " ")

        gd10 = Notebooks(self.panel)

        vsizer.Add(gd01, 0, wx.ALIGN_LEFT|wx.ALL, 5)
        vsizer.Add(gd02, 0, wx.ALIGN_LEFT|wx.ALL, 5)
        vsizer.Add(gd03, 0, wx.ALIGN_LEFT|wx.ALL, 5)
        vsizer.Add(gd04, 0, wx.ALIGN_LEFT|wx.ALL, 5)
        vsizer.Add(gd05, 0, wx.ALIGN_LEFT|wx.ALL, 5)
        vsizer.Add(gd10, 0, wx.ALIGN_LEFT|wx.ALL, 5)

        self.panel.SetSizer(vsizer)
        vsizer.Fit(self.panel)
        self.panel.SetupScrolling()

if __name__ == "__main__":
    app = MyApp(False)
    app.MainLoop()
    