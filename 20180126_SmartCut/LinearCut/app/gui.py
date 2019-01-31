# pylint: disable=C0103, W0613
"""
GUI for SmartCut.
"""
import os
import wx

from .app import first_full_run, second_run


class SmartCutPanel(wx.Panel):
    """This Panel hold two simple buttons, but doesn't really do anything."""

    def __init__(self, parent, *args, **kwargs):
        """Create the SmartCutPanel."""
        wx.Panel.__init__(self, parent, *args, **kwargs)

        self.parent = parent  # Sometimes one can use inline Comments
        self.excel = ''
        self.e2k = ''

        vbox = wx.BoxSizer(orient=wx.VERTICAL)

        # set 9 rows 2 cols
        fgs_1 = wx.FlexGridSizer(2, 3, 20, 20)
        fgs_1.AddGrowableCol(idx=1)

        fgs_2 = wx.FlexGridSizer(5, 2, 20, 20)

        fgs_3 = wx.FlexGridSizer(rows=1, cols=5, vgap=20, hgap=20)

        fgs_4 = wx.FlexGridSizer(1, 2, 20, 20)
        fgs_4.AddGrowableCol(idx=0)
        fgs_4.AddGrowableCol(idx=1)

        self.BEAM_DESIGN = wx.TextCtrl(
            self, style=wx.TE_CENTRE, value=self.excel)
        self.excel_btn = wx.Button(self, label="Browser Excel")
        self.excel_btn.Bind(wx.EVT_BUTTON, self.OnOpenExcel)
        fgs_1.AddMany([wx.StaticText(self, label="Beam Design Excel"),
                       (self.BEAM_DESIGN, 1, wx.EXPAND), self.excel_btn])

        self.E2K = wx.TextCtrl(self, style=wx.TE_CENTRE, value=self.e2k)
        self.e2k_btn = wx.Button(self, label="Browser E2k")
        self.e2k_btn.Bind(wx.EVT_BUTTON, self.OnOpenE2k)
        fgs_1.AddMany([wx.StaticText(self, label="E2k"),
                       (self.E2K, 1, wx.EXPAND), self.e2k_btn])

        self.BARTOP = wx.TextCtrl(
            self, value='#7, #8, #10, #11, #14', size=(250, -1))
        fgs_2.AddMany([wx.StaticText(self, label="Top Rebar"),
                       (self.BARTOP, 1, wx.EXPAND | wx.RIGHT | wx.LEFT, 20)])

        self.BARBOT = wx.TextCtrl(
            self, value='#7, #8, #10, #11, #14')
        fgs_2.AddMany([wx.StaticText(self, label="Bot Rebar"),
                       (self.BARBOT, 1, wx.EXPAND | wx.RIGHT | wx.LEFT, 20)])

        self.DB_SPACING = wx.TextCtrl(self, value='1.5')
        fgs_2.AddMany([wx.StaticText(self, label="Db Spacing"),
                       (self.DB_SPACING, 1, wx.EXPAND | wx.RIGHT | wx.LEFT, 20)])

        self.STIRRUP_REBAR = wx.TextCtrl(
            self, value='#4, 2#4, 2#5, 2#6')
        fgs_2.AddMany([wx.StaticText(self, label="Stirrup Rebar"),
                       (self.STIRRUP_REBAR, 1, wx.EXPAND | wx.RIGHT | wx.LEFT, 20)])

        self.STIRRUP_SPACING = wx.TextCtrl(
            self, value='10, 12, 15, 18, 20, 22, 25, 30')
        fgs_2.AddMany([wx.StaticText(self, label="Stirrup Spacing"),
                       (self.STIRRUP_SPACING, 1, wx.EXPAND | wx.RIGHT | wx.LEFT, 20)])

        self.left = wx.TextCtrl(self, value='0.15', style=wx.TE_CENTRE)
        self.leftmid = wx.TextCtrl(self, value='0.45', style=wx.TE_CENTRE)
        self.rightmid = wx.TextCtrl(self, value='0.55', style=wx.TE_CENTRE)
        self.right = wx.TextCtrl(self, value='0.85', style=wx.TE_CENTRE)
        fgs_3.AddMany([wx.StaticText(self, label="Boundry Condition"),
                       self.left, self.leftmid, self.rightmid, self.right])

        first_run_btn = wx.Button(self, label="First Run")
        first_run_btn.Bind(wx.EVT_BUTTON, self.FirstRun)

        second_run_btn = wx.Button(self, label="Second Run")
        second_run_btn.Bind(wx.EVT_BUTTON, self.SecondRun)

        fgs_4.AddMany([(first_run_btn, 1, wx.EXPAND),
                       (second_run_btn, 1, wx.EXPAND)])

        vbox.Add(fgs_1, flag=wx.LEFT | wx.RIGHT |
                 wx.TOP | wx.EXPAND, border=30)
        vbox.Add(fgs_2, flag=wx.LEFT | wx.RIGHT |
                 wx.TOP | wx.EXPAND, border=30)
        vbox.Add(fgs_3, flag=wx.LEFT | wx.RIGHT |
                 wx.TOP | wx.EXPAND, border=30)
        vbox.Add(wx.StaticLine(self), flag=wx.LEFT | wx.RIGHT |
                 wx.TOP | wx.EXPAND, border=30)
        vbox.Add(fgs_4, flag=wx.LEFT | wx.RIGHT |
                 wx.TOP | wx.EXPAND, border=30)

        self.SetSizer(vbox)

    def GET_BAR(self):
        """ BAR """
        return {
            'Top': self.BARTOP.GetValue(),
            'Bot': self.BARBOT.GetValue()
        }

    def FirstRun(self, event):
        """ first run by beam"""
        # first_full_run()
        print(self.GET_BAR())

    def SecondRun(self, event):
        """ second run by frame"""
        second_run()

    def OnOpenExcel(self, event):
        """ Open a file"""
        dlg = wx.FileDialog(self, message="Choose a file",
                            wildcard="*.xlsx", style=wx.FD_OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.excel = os.path.join(dlg.GetDirectory(), dlg.GetFilename())
            # f = open(os.path.join(self.dirname, self.filename), 'r')
            # self.control.SetValue(f.read())
            # f.close()
        dlg.Destroy()

    def OnOpenE2k(self, event):
        """ Open a file"""
        dlg = wx.FileDialog(self, message="Choose a file",
                            wildcard="*.e2k", style=wx.FD_OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.e2k = os.path.join(dlg.GetDirectory(), dlg.GetFilename())
            # f = open(os.path.join(self.dirname, self.filename), 'r')
            # self.control.SetValue(f.read())
            # f.close()
        dlg.Destroy()


class SmartCutFrame(wx.Frame):
    """ We simply derive a new class of Frame. """

    def __init__(self, *args, **kwargs):
        # ensure the parent's __init__ is called
        wx.Frame.__init__(self, *args, **kwargs)

        # create a menu bar
        self.makeMenuBar()

        # and a status bar
        self.CreateStatusBar()
        self.SetStatusText("Welcome to SmartCut!")

        # create a panel in the frame
        self.PANEL = SmartCutPanel(self)

    def makeMenuBar(self):
        """
        A menu bar is composed of menus, which are composed of menu items.
        This method builds a set of menus and binds handlers to be called
        when the menu item is selected.
        """

        # Setting up the menu.
        fileMenu = wx.Menu()

        # When using a stock ID we don't need to specify the menu item's
        # label
        exitItem = fileMenu.Append(wx.ID_EXIT)

        # Now a help menu for the about item
        helpMenu = wx.Menu()
        aboutItem = helpMenu.Append(wx.ID_ABOUT)

        # Make the menu bar and add the two menus to it. The '&' defines
        # that the next letter is the "mnemonic" for the menu item. On the
        # platforms that support it those letters are underlined and can be
        # triggered from the keyboard.
        menuBar = wx.MenuBar()
        menuBar.Append(fileMenu, "&File")
        menuBar.Append(helpMenu, "&Help")

        # Give the menu bar to the frame
        self.SetMenuBar(menuBar)

        # Finally, associate a handler function with the EVT_MENU event for
        # each of the menu items. That means that when that menu item is
        # activated then the associated handler function will be called.
        self.Bind(wx.EVT_MENU, self.OnExit, exitItem)
        self.Bind(wx.EVT_MENU, self.OnAbout, aboutItem)

    def OnExit(self, event):
        """Close the frame, terminating the application."""
        self.Close(True)

    def OnAbout(self, event):
        """Display an About Dialog"""
        wx.MessageBox("Copyright © 2019 RCBIMX Team. Powered by Paul.",
                      "About Smart Cut",
                      wx.OK | wx.ICON_INFORMATION)


APP = wx.App()
FRAME = SmartCutFrame(None, title='Smart Cut', size=(800, 600))
FRAME.Show()
APP.MainLoop()
