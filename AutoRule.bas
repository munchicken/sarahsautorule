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
    Dim blnFoundRule    As Boolean
    Dim blnFoundAdd     As Boolean
    
    blnFoundAdd = False
    
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
            blnFoundRule = True
            strNote = strNote + vbNewLine + "Existing rule found"
            
            'is this a new email address?
            For j = 0 To oRule.Conditions.From.Recipients.Count - 1
                strNote = strNote + vbNewLine + "Address(es) from Rules: " + oRule.Conditions.From.Recipients.Item(j + 1).Address
                strNote = strNote + vbNewLine + "Address from Email: " + oMail.SenderEmailAddress
                If oRule.Conditions.From.Recipients.Item(j + 1).Address = oMail.SenderEmailAddress Then
                    blnFoundAdd = True
                    strNote = strNote + vbNewLine + "Not a new email address"
                    Exit For
                End If
            Next j
                        
            'add new email address
            If blnFoundAdd = False Then
                oRule.Conditions.From.Recipients.Add (oMail.SenderEmailAddress)
                oRule.Conditions.From.Recipients.ResolveAll
                colRules.Save
                strNote = strNote + vbNewLine + "This is a new email address"
            End If
            Exit For
        End If
    Next
    
    'skip if existing rule found
    If blnFoundRule = False Then
        
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
