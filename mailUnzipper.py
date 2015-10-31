import os
import os.path
import win32com.client
import pythoncom
import re
import sys
import wx
import ConfigParser
import time
import zipfile

class MUFrame(wx.Frame):
    def __init__(self, parent, id=wx.ID_ANY, title="", pos=wx.DefaultPosition,
            size=wx.DefaultSize, style=wx.DEFAULT_FRAME_STYLE, 
            name="MailUnzipper"):
        super(MUFrame, self).__init__(parent, id, title, pos, size, style,
                name)

        self.panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)
        vbox.AddSpacer(10)

        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        folderLbl = wx.StaticText(self.panel, label="Folder: ") 
        self.folderTxt = wx.TextCtrl(self.panel, size=(350, 25))
        folderBtn = wx.Button(self.panel, id=wx.ID_ANY, label="Browse",
                size=(-1, 26))
        hbox1.Add(folderLbl, 0, wx.LEFT|wx.RIGHT, 5)
        hbox1.Add(self.folderTxt, 0, wx.LEFT|wx.RIGHT, 5)
        hbox1.Add(folderBtn, 0, wx.LEFT|wx.RIGHT, 5)
        vbox.Add(hbox1, 0, wx.LEFT|wx.RIGHT|wx.BOTTOM, 10) 

        hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        keywordsLbl = wx.StaticText(self.panel, label="Subject (keywords): ") 
        self.keywordsTxt = wx.TextCtrl(self.panel, size=(350, 25))
        hbox2.Add(keywordsLbl, 0, wx.LEFT|wx.RIGHT, 5)
        hbox2.Add(self.keywordsTxt, 0, wx.LEFT|wx.RIGHT, 5)
        vbox.Add(hbox2, 0, wx.LEFT|wx.RIGHT|wx.BOTTOM, 10) 

        hbox3 = wx.BoxSizer(wx.HORIZONTAL)
        startBtn = wx.Button(self.panel, id=wx.ID_ANY, label="Start")
        #stopBtn = wx.Button(self.panel, id=wx.ID_ANY, label="Stop")
        saveBtn = wx.Button(self.panel, id=wx.ID_ANY, label="Save options")
        loadBtn = wx.Button(self.panel, id=wx.ID_ANY, label="Load options")
        hbox3.Add(startBtn, 0, wx.LEFT|wx.RIGHT, 5)
        #hbox3.Add(stopBtn, 0, wx.LEFT|wx.RIGHT, 5)
        hbox3.Add(saveBtn, 0, wx.LEFT|wx.RIGHT, 5)
        hbox3.Add(loadBtn, 0, wx.LEFT|wx.RIGHT, 5)
        vbox.Add(hbox3, 0, wx.CENTER|wx.LEFT|wx.RIGHT|wx.BOTTOM, 10) 

        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)
        self.Bind(wx.EVT_BUTTON, self.OnStartButton, startBtn)
        #self.Bind(wx.EVT_BUTTON, self.OnStopButton, stopBtn)
        self.Bind(wx.EVT_BUTTON, self.OnSaveButton, saveBtn)
        self.Bind(wx.EVT_BUTTON, self.OnLoadButton, loadBtn)
        self.Bind(wx.EVT_BUTTON, self.OnFolderButton, folderBtn)

        self.statusBar = self.CreateStatusBar()

        self.panel.SetSizer(vbox)
        self.panel.SetAutoLayout(True)
        vbox.Fit(self)
        self.Show()

        self.statusBar.SetStatusText("Done!")

        self.config = ConfigParser.RawConfigParser()
        self.loadSettings()

    def OnCloseWindow(self, event):
        self.Destroy()

    def OnStartButton(self, event):
        global outlook, folder, keywords
        # TODO find more elegant solution than global keywords.

        # Check if Outlook is already running, if not then start it. If we 
        # don't do this here, PyWin32 will start outlook minimized in the 
        # notification area, and it will not respond to new emails. It will 
        # also not close properly.
        try:
            # TODO this is really bad form. Here, we basically just check if we
            # get an error when trying to retrieve an Outlook instance. It
            # would be better if we could use the returned object directly
            # (e.g. outlook = win32com.client.GetActiveObject(...), but this
            # will not return any events. There should be a way to get an
            # active object AND return it's event, but I have yet to find it. 
            win32com.client.GetActiveObject("Outlook.Application") 
        except: 
            os.startfile("Outlook")

        outlook = win32com.client.DispatchWithEvents("Outlook.Application", 
                MailHandler) 

        folder = self.folderTxt.GetValue()
        keywords = self.keywordsTxt.GetValue()
        
        pythoncom.PumpWaitingMessages() # infinite loop that waits for events
        self.statusBar.SetStatusText("Waiting for email...")

    def OnSaveButton(self, event):
        if (self.folderTxt.GetValue() == ""):
            self.throwWarning("Please choose a folder to save the attachments.")
        if (self.keywordsTxt.GetValue() == ""):
            self.throwWarning("Please fill in a subject/keywords.")

        self.saveSettings()

        
    def OnLoadButton(self, event):
        if not (os.path.isfile("settings.cfg")):
            self.throwError("Settings not found (settings.cfg).")
            self.statusBar.SetStatusText("Settings could not be loaded.")
            return

        self.loadSettings()

    def OnFolderButton(self, event):
        dlg = wx.DirDialog(self, message="Choose a folder to extract "
            "attachments", style=wx.DD_DEFAULT_STYLE|wx.DD_NEW_DIR_BUTTON)

        if dlg.ShowModal() == wx.ID_OK:
            self.folderTxt.SetValue(dlg.GetPath())
            dlg.Destroy()

    #TODO Implement stop button after a better understanding of COM.
    #def OnStopButton(self, event):
    #    #TODO this makes everything crash
    #    ctypes.windll.user32.PostQuitMessage(0) # send WM_QUIT to message pump

    #    self.statusBar.SetStatusText("MailUnzipper stopped.")

    def loadSettings(self):
        global folder, keywords

        try:
            self.config.read("settings.cfg")

            folder = self.config.get('settings', 'folder')
            keywords = self.config.get('settings', 'subject')

        except Exception, e:
            self.throwError(
                    "Error loading settings: {0}".format(e))
            self.statusBar.SetStatusText("Could not load settings.")
            return

        self.folderTxt.SetValue(folder)
        self.keywordsTxt.SetValue(keywords)

    def saveSettings(self):
        try:
            self.config.set("settings", "folder", self.folderTxt.GetValue())
            self.config.set("settings", "subject", self.keywordsTxt.GetValue())

            with open("settings.cfg", "w") as settingsFile:
                self.config.write(settingsFile)

        except Exception, e:
            self.throwError(
                    "Error saving settings: {0}".format(e))
            self.statusBar.SetStatusText("Settings could not be saved.")
            return

    def throwError(self, message):
        dlg = wx.MessageDialog(None, message, "Error",
                wx.OK|wx.ICON_ERROR)
        if (dlg.ShowModal() == wx.ID_OK):
            dlg.Destroy()
        return

    def throwWarning(self, message):
        dlg = wx.MessageDialog(None, message, "Warning",
                wx.OK|wx.ICON_WARNING)
        if (dlg.ShowModal() == wx.ID_OK):
            dlg.Destroy()
        return

class MailHandler(object):
    def __init__(self):
        # Some variables to help with the double event call
        self.prevSubject = None
        self.prevCount = None
        self.prevTime = None

    def OnNewMailEx(self, receivedItemsIDs):
        global outlook, keywords, folder

        # ReceivedItemIDs is a collection of mail IDs separated by a ",".
        # (Sometimes more than 1 mail is received at the same moment.)
        for ID in receivedItemsIDs.split(","):
            mail = outlook.Session.GetItemFromID(ID)

            app.frame.statusBar.SetStatusText("Mail ontvangen.")
            
            subject = mail.Subject
            attachments = mail.Attachments

            # Search for keywords in email subject.
            # . matches any character except newline 
            # * matches 0 or more times
            match = bool(re.search(".*{}.*".format(keywords), subject)) 
            
            if (match):
                # VBA is apparently 1-indexed!
                app.frame.statusBar.SetStatusText("Received attachment.") 
                for i in range(1, attachments.Count+1):
                    attachment = attachments.Item(i)
                    fileName = attachment.FileName
                    # If folder does not exist, create it.
                    if not (os.path.isdir(folder)):
                        os.mkdir(folder)
                    filePath = folder + "\\" +  fileName
                    attachment.SaveAsFile(filePath)
                    unzip(filePath)
                    os.remove(filePath)

def unzip(filePath):
    app.frame.statusBar.SetStatusText(
            "Unzipping file: {}".format(filePath))
    zfile = zipfile.ZipFile(filePath)

    for name in zfile.namelist():
        zfile.extract(name, path=folder)
    
    time.sleep(10)
    app.frame.statusBar.SetStatusText("Waiting for email...")

if __name__ == "__main__":
    app = wx.App(redirect=True)
    app.frame = MUFrame(None, -1, "MailUnzipper")
    app.MainLoop()
