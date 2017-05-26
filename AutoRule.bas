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
    Dim blnFoundFolder  As Boolean
    Dim blnFoundTarget  As Boolean
    
    blnFoundAdd = False
    blnFoundFolder = False
    blnFoundTarget = False
    
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
        
        'does group folder exist?
        Set oInbox = Application.Session.GetDefaultFolder(olFolderInbox)
        For i = 1 To oInbox.Folders.Count
            If oInbox.Folders(i) = "Contact Groups" Then
                blnFoundFolder = True
                strNote = strNote + vbNewLine + "Group folder exists"
                Exit For
            End If
        Next i
        'create group folder
        If blnFoundFolder = False Then
            oInbox.Folders.Add ("Contact Groups")
            strNote = strNote + vbNewLine + "Group folder doesn't exist"
        End If
        'does target folder exist?
        Set oGrpFolder = oInbox.Folders("Contact Groups")
        For i = 1 To oGrpFolder.Folders.Count
            If oGrpFolder.Folders(i) = strSender Or oGrpFolder.Folders(i) = oMail.SenderEmailAddress Then
                blnFoundTarget = True
                strNote = strNote + vbNewLine + "Target folder exists"
                Exit For
            End If
        Next i
        'create target folder
        If blnFoundTarget = False Then
            oGrpFolder.Folders.Add (strSender)
            strNote = strNote + vbNewLine + "Target folder doesn't exist"
        End If
    End If
    
    'for testing
    strNote = strNote + vbNewLine + "note"
    UserForm1.Label1.Caption = strNote
    UserForm1.Show
    'MsgBox (UCase(strSender))
    'MsgBox (oGrpFolder.Folders.Count)
    'MsgBox (oInbox.Folders.Count)
End Sub
