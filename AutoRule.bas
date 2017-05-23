Attribute VB_Name = "AutoRule"
'*****  AutoRule
'*****  Outlook VBA Macro to automatically create a rule from the selected email
'*****  Based on the MS product team article "Best practices for Outlook 2010"
'*****  By:  Sarah Pierce

Sub AutoRule()
    Dim myOlExp     As Outlook.Explorer
    Dim myOlSel     As Outlook.Selection
    Dim oMail       As Outlook.MailItem
    Dim strSender   As String
    Dim oInbox      As Outlook.Folder
    Dim oGrpFolder  As Outlook.Folder

    
    'get the currently selected email
    Set myOlExp = Application.ActiveExplorer
    Set myOlSel = myOlExp.Selection
    Set oMail = myOlSel.Item(1)
    
    'setup move folder
    Set oInbox = Application.Session.GetDefaultFolder(olFolderInbox)
    Set oGrpFolder = oInbox.Folders("Contact Groups")
    'For i = 1 To oGrpFolder.Folders.Count
        
    
    'for testing
    strSender = oMail.Sender
    MsgBox (UCase(strSender))
    MsgBox (oGrpFolder.Folders.Count)
End Sub
