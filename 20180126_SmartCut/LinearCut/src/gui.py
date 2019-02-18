"""
GUI for SmartCut.
"""
import os
import wx

from app import cut_by_beam, cut_by_frame


# def print_path():
#     print(FRAME.PANEL.etabs_design_path)


class SmartCutPanel(wx.Panel):
    """This Panel hold two simple buttons, but doesn't really do anything."""

    def __init__(self, parent, *args, **kwargs):
        """Create the SmartCutPanel."""
        wx.Panel.__init__(self, parent, *args, **kwargs)

        self.parent = parent  # Sometimes one can use inline Comments
        self.etabs_design_path = ''
        self.e2k_path = ''
        self.beam_name_path = ''
        self.output_dir = ''

        vbox = wx.BoxSizer(orient=wx.VERTICAL)

        # set 9 rows 2 cols
        fgs_1 = wx.FlexGridSizer(4, 3, 20, 20)
        fgs_1.AddGrowableCol(idx=1)

        fgs_2 = wx.FlexGridSizer(5, 2, 20, 20)

        fgs_3 = wx.FlexGridSizer(rows=1, cols=5, vgap=20, hgap=20)

        fgs_4 = wx.FlexGridSizer(1, 2, 20, 20)
        fgs_4.AddGrowableCol(idx=0)
        fgs_4.AddGrowableCol(idx=1)

        self.etabs_design = wx.TextCtrl(
            self, style=wx.TE_CENTRE)
        self.etabs_design_btn = wx.Button(self, label="Browser Excel")
        self.etabs_design_btn.Bind(
            wx.EVT_BUTTON, self.on_click_etabs_dsign_btn)
        fgs_1.AddMany([wx.StaticText(self, label="ETBAS Beam Design Excel"),
                       (self.etabs_design, 1, wx.EXPAND), self.etabs_design_btn])

        self.e2k = wx.TextCtrl(self, style=wx.TE_CENTRE)
        self.e2k_btn = wx.Button(self, label="Browser E2k")
        self.e2k_btn.Bind(wx.EVT_BUTTON, self.on_click_e2k_btn)
        fgs_1.AddMany([wx.StaticText(self, label="E2k"),
                       (self.e2k, 1, wx.EXPAND), self.e2k_btn])

        self.beam_name = wx.TextCtrl(
            self, style=wx.TE_CENTRE)
        self.beam_name_btn = wx.Button(self, label="Browser Excel")
        self.beam_name_btn.Bind(wx.EVT_BUTTON, self.on_click_beam_name_btn)
        fgs_1.AddMany([wx.StaticText(self, label="Beam Name Excel"),
                       (self.beam_name, 1, wx.EXPAND), self.beam_name_btn])

        self.output = wx.TextCtrl(
            self, style=wx.TE_CENTRE)
        self.output_btn = wx.Button(self, label="Browser Folder")
        self.output_btn.Bind(wx.EVT_BUTTON, self.on_click_output_btn)
        fgs_1.AddMany([wx.StaticText(self, label="Output Folder"),
                       (self.output, 1, wx.EXPAND), self.output_btn])

        self.bartop = wx.TextCtrl(
            self, value='#7, #8, #10, #11, #14', size=(250, -1))
        fgs_2.AddMany([wx.StaticText(self, label="Top Rebar"),
                       (self.bartop, 1, wx.EXPAND | wx.RIGHT | wx.LEFT, 20)])

        self.barbot = wx.TextCtrl(
            self, value='#7, #8, #10, #11, #14')
        fgs_2.AddMany([wx.StaticText(self, label="Bot Rebar"),
                       (self.barbot, 1, wx.EXPAND | wx.RIGHT | wx.LEFT, 20)])

        self.db_spacing = wx.TextCtrl(self, value='1.5')
        fgs_2.AddMany([wx.StaticText(self, label="Db Spacing"),
                       (self.db_spacing, 1, wx.EXPAND | wx.RIGHT | wx.LEFT, 20)])

        self.stirrup_rebar = wx.TextCtrl(
            self, value='#4, 2#4, 2#5, 2#6')
        fgs_2.AddMany([wx.StaticText(self, label="Stirrup Rebar"),
                       (self.stirrup_rebar, 1, wx.EXPAND | wx.RIGHT | wx.LEFT, 20)])

        self.stirrup_spacing = wx.TextCtrl(
            self, value='10, 12, 15, 18, 20, 22, 25, 30')
        fgs_2.AddMany([wx.StaticText(self, label="Stirrup Spacing"),
                       (self.stirrup_spacing, 1, wx.EXPAND | wx.RIGHT | wx.LEFT, 20)])

        self.left = wx.TextCtrl(self, value='0.15', style=wx.TE_CENTRE)
        self.leftmid = wx.TextCtrl(self, value='0.45', style=wx.TE_CENTRE)
        self.rightmid = wx.TextCtrl(self, value='0.55', style=wx.TE_CENTRE)
        self.right = wx.TextCtrl(self, value='0.85', style=wx.TE_CENTRE)
        fgs_3.AddMany([wx.StaticText(self, label="Boundry Condition"),
                       self.left, self.leftmid, self.rightmid, self.right])

        first_run_btn = wx.Button(self, label="Run by Beam")
        first_run_btn.Bind(wx.EVT_BUTTON, lambda x: cut_by_beam())

        second_run_btn = wx.Button(self, label="Run by Frame")
        second_run_btn.Bind(wx.EVT_BUTTON, lambda x: cut_by_frame())

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
            'Top': self.bartop.GetValue(),
            'Bot': self.barbot.GetValue()
        }

    def get_etabs_design_path(self):
        print(self.etabs_design_path)
        return self.etabs_design_path

    def on_click_beam_name_btn(self, event):
        """ Open a file"""
        dlg = wx.FileDialog(self, message="Choose a file",
                            wildcard="*.xlsx", style=wx.FD_OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.beam_name_path = os.path.join(
                dlg.GetDirectory(), dlg.GetFilename())
            self.beam_name.SetValue(self.beam_name_path)

        dlg.Destroy()

    def on_click_output_btn(self, event):
        """ Open a file"""
        dlg = wx.DirDialog(self, message="Choose output directory",
                           style=wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST)
        if dlg.ShowModal() == wx.ID_OK:
            self.output_dir = dlg.GetPath()
            self.output.SetValue(self.output_dir)

        dlg.Destroy()

    def on_click_etabs_dsign_btn(self, event):
        """ Open a file"""
        dlg = wx.FileDialog(self, message="Choose a file",
                            wildcard="*.xlsx", style=wx.FD_OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.etabs_design_path = os.path.join(
                dlg.GetDirectory(), dlg.GetFilename())
            # f = open(os.path.join(self.dirname, self.filename), 'r')
            # self.control.SetValue(f.read())
            # f.close()
            self.etabs_design.SetValue(self.etabs_design_path)
            # print_path()

        dlg.Destroy()

    def on_click_e2k_btn(self, event):
        """ Open a file"""
        dlg = wx.FileDialog(self, message="Choose a file",
                            wildcard="*.e2k", style=wx.FD_OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.e2k_path = os.path.join(dlg.GetDirectory(), dlg.GetFilename())
            self.e2k.SetValue(self.e2k_path)

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
        wx.MessageBox("Copyright 2019 RCBIMX Team. Powered by Paul.",
                      "About Smart Cut",
                      wx.OK | wx.ICON_INFORMATION)


APP = wx.App()
FRAME = SmartCutFrame(None, title='Smart Cut', size=(800, 600))
FRAME.Show()
APP.MainLoop()
