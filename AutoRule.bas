Attribute VB_Name = "AutoRule"
'*****  AutoRule
'*****  Outlook VBA Macro to automatically create a rule from the selected email
'*****  Based on the MS product team article "Best practices for Outlook 2010"
'*****  Uses "Contact Groups" folder under Inbox for storing group email
'*****  By:  Sarah Pierce

Sub AutoRule()
    Dim myOlExp         As Outlook.Explorer
    Dim myOlSel         As Outlook.Selection
    Dim oMail           As Outlook.MailItem
    Dim strSender       As String
    Dim oInbox          As Outlook.Folder
    Dim oGrpFolder      As Outlook.Folder
    Dim strNote         As String
    Dim colRules        As Outlook.Rules
    Dim oRule           As Outlook.Rule
    Dim blnFound        As Boolean

    
    'get the currently selected email
    Set myOlExp = Application.ActiveExplorer
    Set myOlSel = myOlExp.Selection
    Set oMail = myOlSel.Item(1) 'the selected email
    strSender = oMail.Sender
    strNote = "Current email from: " + strSender
    
    'check for existing rule
    Set colRules = Application.Session.DefaultStore.GetRules()
    For Each oRule In colRules
        If UCase(oRule.Name) = UCase(strSender) Then
            blnFound = True
            strNote = strNote + vbNewLine + "Existing rule found"
            
            'is this a new email address?
            For j = 0 To oRule.Conditions.From.Recipients.Count - 1
                strNote = strNote + vbNewLine + oRule.Conditions.From.Recipients.Item(j + 1).Address
            Next j
            
            Exit For
        End If
    Next
    
    'skip if existing rule found
    If blnFound = False Then
        
        'setup move folder
        Set oInbox = Application.Session.GetDefaultFolder(olFolderInbox)
        Set oGrpFolder = oInbox.Folders("Contact Groups")
        'For i = 1 To oGrpFolder.Folders.Count
    End If
    
    'for testing
    strNote = strNote + vbNewLine + "note"
    UserForm1.Label1.Caption = strNote
    UserForm1.Show
    'MsgBox (UCase(strSender))
    'MsgBox (oGrpFolder.Folders.Count)
End Sub
