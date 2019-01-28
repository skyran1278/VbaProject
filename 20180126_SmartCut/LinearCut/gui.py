# pylint: disable=C0103, W0613
"""
GUI for SmartCut.
"""
import os
import wx


class SmartCutPanel(wx.Panel):
    """This Panel hold two simple buttons, but doesn't really do anything."""

    def __init__(self, parent, *args, **kwargs):
        """Create the SmartCutPanel."""
        wx.Panel.__init__(self, parent, *args, **kwargs)

        self.parent = parent  # Sometimes one can use inline Comments

        vbox = wx.BoxSizer(orient=wx.VERTICAL)

        hbox_design = wx.BoxSizer(wx.HORIZONTAL)
        hbox_design.Add(wx.StaticText(self, label="Beam Design Excel:"),
                        flag=wx.ALL, border=10)
        hbox_design.Add(wx.TextCtrl(self), flag=wx.ALL, border=10)
        hbox_design.Add(wx.Button(self, label='Browser File'),
                        flag=wx.ALL, border=10)

        hbox_e2k = wx.BoxSizer(wx.HORIZONTAL)
        hbox_e2k.Add(wx.StaticText(self, label="E2k:"),
                     flag=wx.ALL, border=10)
        hbox_e2k.Add(wx.TextCtrl(self), flag=wx.ALL, border=10)
        hbox_e2k.Add(wx.Button(self, label='Browser File'),
                     flag=wx.ALL, border=10)

        hbox_rebar_top = wx.BoxSizer(wx.HORIZONTAL)
        hbox_rebar_top.Add(wx.StaticText(self, label="Top Rebar:"),
                           flag=wx.ALL, border=10)
        hbox_rebar_top.Add(wx.TextCtrl(self, value='#7, #8, #10, #11, #14'),
                           flag=wx.ALL, border=10)

        hbox_rebar_bot = wx.BoxSizer(wx.HORIZONTAL)
        hbox_rebar_bot.Add(wx.StaticText(self, label="Bot Rebar:"),
                           flag=wx.ALL, border=10)
        hbox_rebar_bot.Add(wx.TextCtrl(self, value='#7, #8, #10, #11, #14'),
                           flag=wx.ALL, border=10)

        hbox_db_spacing = wx.BoxSizer(wx.HORIZONTAL)
        hbox_db_spacing.Add(wx.StaticText(self, label="Db Spacing:"),
                            flag=wx.ALL, border=10)
        hbox_db_spacing.Add(wx.TextCtrl(self, value='1.5'),
                            flag=wx.ALL, border=10)

        hbox_stirrup = wx.BoxSizer(wx.HORIZONTAL)
        hbox_stirrup.Add(wx.StaticText(self, label="Stirrup Rebar:"),
                         flag=wx.ALL, border=10)
        hbox_stirrup.Add(wx.TextCtrl(self, value='#4, 2#4, 2#5, 2#6'),
                         flag=wx.ALL, border=10)

        hbox_dh_spacing = wx.BoxSizer(wx.HORIZONTAL)
        hbox_dh_spacing.Add(wx.StaticText(self, label="Stirrup Spacing:"),
                            flag=wx.ALL, border=10)
        hbox_dh_spacing.Add(wx.TextCtrl(self, value='10, 12, 15, 18, 20, 22, 25, 30'),
                            flag=wx.ALL, border=10)

        hbox_boundry = wx.BoxSizer(wx.HORIZONTAL)
        hbox_boundry.Add(wx.StaticText(self, label="Boundry Condition:"),
                         flag=wx.ALL, border=10)
        hbox_boundry.Add(wx.TextCtrl(self, value='0.15'),
                         flag=wx.ALL, border=10)
        hbox_boundry.Add(wx.TextCtrl(self, value='0.45'),
                         flag=wx.ALL, border=10)
        hbox_boundry.Add(wx.TextCtrl(self, value='0.55'),
                         flag=wx.ALL, border=10)
        hbox_boundry.Add(wx.TextCtrl(self, value='0.85'),
                         flag=wx.ALL, border=10)

        # NothingBtn = wx.Button(self, label="Do Nothing with a long label")
        # NothingBtn.Bind(wx.EVT_BUTTON, self.DoNothing)

        # MsgBtn = wx.Button(self, label="Send Message")
        # MsgBtn.Bind(wx.EVT_BUTTON, self.OnMsgBtn)

        vbox.Add(hbox_design, proportion=0,
                 flag=wx.ALIGN_CENTER | wx.ALL, border=5)
        vbox.Add(hbox_e2k, proportion=0,
                 flag=wx.ALIGN_CENTER | wx.ALL, border=5)
        vbox.Add(hbox_rebar_top, proportion=0,
                 flag=wx.ALIGN_CENTER | wx.ALL, border=5)
        vbox.Add(hbox_rebar_bot, proportion=0,
                 flag=wx.ALIGN_CENTER | wx.ALL, border=5)
        vbox.Add(hbox_db_spacing, proportion=0,
                 flag=wx.ALIGN_CENTER | wx.ALL, border=5)
        vbox.Add(hbox_stirrup, proportion=0,
                 flag=wx.ALIGN_CENTER | wx.ALL, border=5)
        vbox.Add(hbox_dh_spacing, proportion=0,
                 flag=wx.ALIGN_CENTER | wx.ALL, border=5)
        vbox.Add(hbox_boundry, proportion=0,
                 flag=wx.ALIGN_CENTER | wx.ALL, border=5)

        self.SetSizer(vbox)

    def OnOpen(self, event):
        """ Open a file"""
        self.dirname = ''
        dlg = wx.FileDialog(self, "Choose a file",
                            self.dirname, "", "*.*", wx.FD_OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.filename = dlg.GetFilename()
            self.dirname = dlg.GetDirectory()
            f = open(os.path.join(self.dirname, self.filename), 'r')
            self.control.SetValue(f.read())
            f.close()
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
        self.Panel = SmartCutPanel(self)

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


if __name__ == '__main__':
    # When this module is run (not imported) then create the app, the
    # frame, show it, and start the event loop.
    app = wx.App()
    frm = SmartCutFrame(None, title='Smart Cut', size=(800, 600))
    frm.Show()
    app.MainLoop()
