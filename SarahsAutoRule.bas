Attribute VB_Name = "SarahsAutoRule"
'*****  Sarah's AutoRule
'*****  Description:
'*****    Outlook VBA Macro to automatically create a "Contact Group" rule from the selected email
'*****  By:  Sarah Pierce

Option Explicit

'globals
Dim m_strNote               As String
Dim m_strChanges            As String
Public g_strUserGrpFolder   As String
Public g_blnToCc            As Boolean
Public g_arrWords           As Variant
Public g_blnWords           As Boolean

'main subroutine
Sub AutoRule()
    Dim strSender       As String
    Dim blnFoundRule    As Boolean
    Dim blnFoundAdd     As Boolean
    Dim blnFoundFolder  As Boolean
    Dim blnFoundTarget  As Boolean
    Dim i               As Integer
    Dim oRule           As Outlook.Rule
    Dim myOlExp         As Outlook.Explorer
    Dim myOlSel         As Outlook.Selection
    Dim oGrpFolder      As Outlook.Folder
    Dim colRules        As Outlook.Rules
    Dim oTargetFolder   As Outlook.Folder
    Dim oMail           As Outlook.MailItem
    Dim oInbox          As Outlook.Folder
    Dim oFromCond       As Outlook.ToOrFromRuleCondition
    Dim oMoveAction     As Outlook.MoveOrCopyRuleAction
    Dim oStopAction     As Outlook.RuleAction
    Dim oExceptMe       As Outlook.RuleCondition
    Dim oExceptWords    As Outlook.TextRuleCondition
    Dim oCurrentFolder  As Outlook.Folder
    Dim strGrpFolder    As String
    Dim blnNewRule      As Boolean
    Dim blnButton       As Boolean
    Dim blnMove         As Boolean
    Dim blnAddAddy      As Boolean
    Dim intProceed      As Integer
    
    'initialize
    blnFoundAdd = False
    blnFoundFolder = False
    blnFoundTarget = False
    m_strNote = ""
    blnFoundRule = False
    strGrpFolder = "Contact Groups"
    blnNewRule = False
    blnButton = False
    blnMove = False
    blnAddAddy = False
    intProceed = vbNo
    m_strChanges = ""
    g_arrWords = Array()
    
    'show notification box & propsed changes
    frmStatus.Show
    Call Notify("Sarah's AutoRule starting...")
    Call AddChange("Proposed Changes:")
    
    'get the currently selected email
    Set myOlExp = Application.ActiveExplorer
    Set myOlSel = myOlExp.Selection
    Set oMail = myOlSel.Item(1) 'the selected email
    strSender = oMail.Sender
    Set oInbox = Application.Session.GetDefaultFolder(olFolderInbox)
    
    'check to see if we are in the Inbox
    Set oCurrentFolder = oMail.Parent
    If oCurrentFolder <> "Inbox" Then
        Call Notify("  ERROR: Selected email is not in the Inbox.  AutoRule must be run on items in Inbox.")
        frmStatus.txtStatus.ForeColor = vbRed
        Call Notify("Sarah's AutoRule finished")
        frmStatus.cmdClose.Visible = True
        Exit Sub
    End If
    
    'check for existing rule
    Set colRules = Application.Session.DefaultStore.GetRules()
    For Each oRule In colRules
        
        'rule exists
        If UCase(oRule.Name) = UCase(strSender) Then
            blnFoundRule = True
            Call Notify("  Existing rule found for: " + strSender)
            
            'is this an external email?
            If oMail.SenderEmailType = "SMTP" Then
                Call Notify("  This is an external email address")
                
                'is this a new email address?
                For i = 0 To oRule.Conditions.From.Recipients.Count - 1
                    If oRule.Conditions.From.Recipients.Item(i + 1).Address = oMail.SenderEmailAddress Then
                        blnFoundAdd = True
                        'show status
                        Call Notify("  This is not a new email address for: " + strSender)
                        Call AddChange("Move email to folder")
                        'set flag for action
                        blnMove = True
                        Exit For
                    End If
                Next i
                            
                'add new email address
                If blnFoundAdd = False Then
                    'show status
                    Call AddChange("Add new email address to existing rule")
                    Call AddChange("Re-run existing rule")
                    'set flag for action
                    blnAddAddy = True
                End If
                Exit For
            
            'this is an internal email
            Else
                'show status
                Call Notify("  This is an internal email address")
                Call AddChange("Move email to folder")
                'set flag for action
                blnMove = True
            End If
        End If
    Next
    
    'rule not found, skip if existing rule found
    If blnFoundRule = False Then
    
        'show status
        Call Notify("  Rule not found for: " + strSender)
        Call AddChange("Create new rule using settings dispalyed")
        Call AddChange("Create group folder, if not present")
        Call AddChange("Create folder for sender, if not present")
        Call AddChange("Run new rule")
        
        'set flag for action
        blnNewRule = True
    End If
    
    'ask permission & show settings
    If blnNewRule = True Then
        frmSettings.Show
    End If
    
    intProceed = MsgBox("Please read Proposed Changes.  Proceed?", vbYesNo + vbQuestion + vbDefaultButton2 + vbApplicationModal, "Confirm Changes")
    If intProceed = vbYes Then
        
        'do flagged actions
        If blnMove Then
            'move it to the correct folder (assumes user looked at email, which was an exception, and wants it out of inbox, otherwise they would delete it)
            Call MoveIt(oRule, oMail)
        End If
        
        If blnAddAddy Then
            'add new address
            oRule.Conditions.From.Recipients.Add (oMail.SenderEmailAddress)
            oRule.Conditions.From.Recipients.ResolveAll
            colRules.Save
            Call Notify("  Added new email address for " + strSender + " to existing rule")
            
            're-run rule with new address
            Call Notify("  Re-Running rule for: " + strSender + " with new email address, please stand by")
            oRule.Execute ShowProgress:=True
        End If
        
        If blnNewRule = True Then
            'get settings
            If g_strUserGrpFolder <> strGrpFolder And g_strUserGrpFolder <> "" Then
                strGrpFolder = g_strUserGrpFolder
            End If
            
            'does group folder exist?
            For i = 1 To oInbox.Folders.Count
                If oInbox.Folders(i) = strGrpFolder Then
                    blnFoundFolder = True
                    Call Notify("  Group folder exists")
                    Exit For
                End If
            Next i
            
            'create group folder
            If blnFoundFolder = False Then
                oInbox.Folders.Add (strGrpFolder)
                Call Notify("  Group folder doesn't exist")
            End If
            
            'does target folder exist?
            Set oGrpFolder = oInbox.Folders(strGrpFolder)
            For i = 1 To oGrpFolder.Folders.Count
                If oGrpFolder.Folders(i) = strSender Or oGrpFolder.Folders(i) = oMail.SenderEmailAddress Then
                    blnFoundTarget = True
                    Set oTargetFolder = oGrpFolder.Folders(i)
                    Call Notify("  " + strSender + " folder exists")
                    Exit For
                End If
            Next i
            
            'create target folder
            If blnFoundTarget = False Then
                oGrpFolder.Folders.Add (strSender)
                Set oTargetFolder = oGrpFolder.Folders(strSender)
                Call Notify("  " + strSender + " folder doesn't exist, creating")
            End If
            
            'add new rule
            Set oRule = colRules.Create(strSender, olRuleReceive)
            
            'set condition
            Set oFromCond = oRule.Conditions.From
            With oFromCond
                .Enabled = True
                .Recipients.Add (strSender)
                .Recipients.ResolveAll
            End With
            
            'set action
            Set oMoveAction = oRule.Actions.MoveToFolder
            With oMoveAction
                .Enabled = True
                .Folder = oTargetFolder
            End With
            
            'set exception
            If g_blnToCc = True Then
                Set oExceptMe = oRule.Exceptions.ToOrCc
                With oExceptMe
                    .Enabled = True
                End With
            End If
            If g_blnWords = True Then
                Set oExceptWords = oRule.Exceptions.BodyOrSubject
                With oExceptWords
                    .Enabled = True
                    .Text = g_arrWords
                End With
            End If
            
            'stop processing rules
            Set oStopAction = oRule.Actions.Stop
            With oStopAction
                .Enabled = True
            End With
            
            'move to bottom of list
            oRule.ExecutionOrder = colRules.Count
            
            'save rules
            colRules.Save
            Call Notify("  Creating new rule for: " + strSender + ", please stand by")
            
            'run new rule
            Call Notify("  Running new rule for: " + strSender + ", please stand by")
            oRule.Execute ShowProgress:=True
        End If
    End If
    
    'when complete
    Call Notify("Sarah's AutoRule finished")
    frmStatus.cmdClose.Visible = True
    
    'free objects
    Set oTargetFolder = Nothing
    Set oMail = Nothing
    Set oRule = Nothing
    Set oInbox = Nothing
    Set myOlExp = Nothing
    Set myOlSel = Nothing
    Set oGrpFolder = Nothing
    Set colRules = Nothing
    Set oFromCond = Nothing
    Set oMoveAction = Nothing
    Set oStopAction = Nothing
    Set oExceptMe = Nothing
    Set oExceptWords = Nothing
    Set oCurrentFolder = Nothing
    
End Sub

'creates notification messages on form
Sub Notify(message As String)
    m_strNote = m_strNote + vbNewLine + message
    frmStatus.txtStatus.Text = m_strNote
End Sub

'moves email to proper folder
Sub MoveIt(oRule As Outlook.Rule, oMail As Outlook.MailItem)
    Dim oMovedMail As Outlook.MailItem
    Dim oTargetFolder   As Outlook.Folder
    
    Set oTargetFolder = oRule.Actions.MoveToFolder.Folder
    Set oMovedMail = oMail.Move(oTargetFolder)
    Call Notify("  Email moved to " + oMovedMail.Parent + " folder")
    
    Set oMovedMail = Nothing
    Set oTargetFolder = Nothing
End Sub

'adds proposed changes to list
Sub AddChange(change As String)
    m_strChanges = m_strChanges + vbNewLine + change
    frmStatus.txtChanges.Text = m_strChanges
End Sub

