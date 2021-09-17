import wx
from frame import Frame

version = "1.0"


class App(wx.App):

    def OnInit(self):
        self.main = Frame(version=version)
        self.main.Show()
        self.SetTopWindow(self.main)
        return True
