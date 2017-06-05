Attribute VB_Name = "SarahsAutoRule"
'*****  Sarah's AutoRule
'*****  Description:
'*****    Outlook VBA Macro to automatically create a "Contact Group" rule from the selected email
'*****    Based on the MS product team article "Best practices for Outlook 2010" (https://support.office.com/en-us/article/Best-practices-for-Outlook-2010-f90e5f69-8832-4d89-95b3-bfdf76c82ef8)
'*****  Instructions:
'*****    Run the macro on an email in the Inbox which is selected in the Outlook Explorer that you want automatically moved to a "Contact Group" folder
'*****  Actions:
'*****    It checks to see if there is an existing rule for this sender
'*****    If there is not an exising rule, then it creates one with the following settings:
'*****      Move messages from "Sender" to "Folder"
'*****      It checks for a "Contact Groups" folder, and creates one if necessary
'*****      It then checks for a folder in Contact Groups named "Sender", and creates one if necessary
'*****      Except if users name is in the To or Cc box
'*****      Except if "specific words" are in the subject or body (see the array below marked with "+++++" if you would like to change these)
'*****      Stop processing more rules
'*****      It moves the new rule to the bottom of the rule list
'*****      It then runs the new rule
'*****    If there is an existing rule, it checks to see if this is a new email address, if so it adds it to the existing rule and re-runs the rule
'*****    If the rule exists and has the correct email addresses, then this email is in your Inbox due to one of the exceptions
'*****      If you choose not to delete the email, but rather run AutoRule on it, then it assumes you just want to move it to the proper folder and does so
'*****  Notes:
'*****    The notification box indicates all actions taken
'*****    You can check & modify any created rules in the Outlook Rules & Alerts Manager
'*****  Installation:
'*****    Download the module (SarahsAutoRule.bas) & forms (SarahsAutoRuleUserForm .frm & .frx)
'*****    Enable the "Developer" tab on the Outlook ribbon
'*****    From the Developer tab, click Visual Basic to open the editor
'*****    Under "Froms" in the Project Explorer (left), import the two forms
'*****    Under "Modules", import the module
'*****    The macro can be run from from the Developer tab or can be placed in a menu like the ribbon, etc.
'*****    You may need to adjust your Macro security settings (macro can also be self-signed with SelfCert, if desired)
'*****  By:  Sarah Pierce

'globals
Dim strNote         As String
Dim oTargetFolder   As Outlook.Folder
Dim oMail           As Outlook.MailItem
Dim oRule           As Outlook.Rule

'main subroutine
Sub AutoRule()
    Dim myOlExp         As Outlook.Explorer
    Dim myOlSel         As Outlook.Selection
    Dim strSender       As String
    Dim oInbox          As Outlook.Folder
    Dim oGrpFolder      As Outlook.Folder
    Dim colRules        As Outlook.Rules
    Dim blnFoundRule    As Boolean
    Dim blnFoundAdd     As Boolean
    Dim blnFoundFolder  As Boolean
    Dim blnFoundTarget  As Boolean
    Dim oFromCond       As Outlook.ToOrFromRuleCondition
    Dim oMoveAction     As Outlook.MoveOrCopyRuleAction
    Dim oStopAction     As Outlook.RuleAction
    Dim oExceptMe       As Outlook.RuleCondition
    Dim oExceptWords    As Outlook.TextRuleCondition
    Dim oCurrentFolder  As Outlook.Folder
    
    'initialize
    blnFoundAdd = False
    blnFoundFolder = False
    blnFoundTarget = False
    strNote = ""
    
    'show notification box
    SarahsAutoRuleUserForm.Show
    Notify ("Sarah's AutoRule starting...")
    
    'get the currently selected email
    Set myOlExp = Application.ActiveExplorer
    Set myOlSel = myOlExp.Selection
    Set oMail = myOlSel.Item(1) 'the selected email
    strSender = oMail.Sender
    
    'check to see if we are in the Inbox
    Set oCurrentFolder = oMail.Parent
    MsgBox (oCurrentFolder)
    If oCurrentFolder <> "Inbox" Then
        Notify ("Selected email is not in the Inbox.  AutoRule must be run on item in Inbox.")
        SarahsAutoRuleUserForm.Label1.ForeColor = vbRed
        Exit Sub
    End If
    
    'check for existing rule
    Set colRules = Application.Session.DefaultStore.GetRules()
    For Each oRule In colRules
        
        'rule exists
        If UCase(oRule.Name) = UCase(strSender) Then
            blnFoundRule = True
            Notify ("Existing rule found for: " + strSender)
            
            'is this an external email?
            If oMail.SenderEmailType = "SMTP" Then
                Notify ("This is an external email address")
                
                'is this a new email address?
                For j = 0 To oRule.Conditions.From.Recipients.Count - 1
                    If oRule.Conditions.From.Recipients.Item(j + 1).Address = oMail.SenderEmailAddress Then
                        blnFoundAdd = True
                        Notify ("This is not a new email address for: " + strSender)
                        
                        'move it to the correct folder (assumes user looked at email, which was an exception, and wants it out of inbox, otherwise they would delete it)
                        MoveIt
                        Exit For
                    End If
                Next j
                            
                'add new email address
                If blnFoundAdd = False Then
                    oRule.Conditions.From.Recipients.Add (oMail.SenderEmailAddress)
                    oRule.Conditions.From.Recipients.ResolveAll
                    colRules.Save
                    Notify ("Added new email address for " + strSender + " to existing rule")
                    
                    're-run rule with new address
                    Notify ("Re-Running rule for: " + strSender + " with new email address, please stand by")
                    oRule.Execute ShowProgress:=True
                End If
                Exit For
            
            'this is an internal email
            Else
                Notify ("This is an internal email address")
                'move it to the correct folder (assumes user looked at email, which was an exception, and wants it out of inbox, otherwise they would delete it)
                MoveIt
            End If
        End If
    Next
    
    'rule not found, skip if existing rule found
    If blnFoundRule = False Then
    
        Notify ("Creating new rule for: " + strSender)
        
        'does group folder exist?
        Set oInbox = Application.Session.GetDefaultFolder(olFolderInbox)
        For i = 1 To oInbox.Folders.Count
            If oInbox.Folders(i) = "Contact Groups" Then
                blnFoundFolder = True
                Notify ("Group folder exists")
                Exit For
            End If
        Next i
        
        'create group folder
        If blnFoundFolder = False Then
            oInbox.Folders.Add ("Contact Groups")
            Notify ("Group folder doesn't exist, creating")
        End If
        
        'does target folder exist?
        Set oGrpFolder = oInbox.Folders("Contact Groups")
        For i = 1 To oGrpFolder.Folders.Count
            If oGrpFolder.Folders(i) = strSender Or oGrpFolder.Folders(i) = oMail.SenderEmailAddress Then
                blnFoundTarget = True
                Set oTargetFolder = oGrpFolder.Folders(i)
                Notify (strSender + " folder exists")
                Exit For
            End If
        Next i
        
        'create target folder
        If blnFoundTarget = False Then
            oGrpFolder.Folders.Add (strSender)
            Set oTargetFolder = oGrpFolder.Folders(strSender)
            Notify (strSender + " folder doesn't exist, creating")
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
        Set oExceptMe = oRule.Exceptions.ToOrCc
        With oExceptMe
            .Enabled = True
        End With
        Set oExceptWords = oRule.Exceptions.BodyOrSubject
        With oExceptWords
            .Enabled = True
            '+++++ change these if you would like
            .Text = Array("deadline", "urgent", "renew", "important", "quote", "respond", "waiting", "enroll", "fair", "submit", "meeting", "register", "expire", "expiration", "schedule", "remind")
        End With
        
        'stop processing rules
        Set oStopAction = oRule.Actions.Stop
        With oStopAction
            .Enabled = True
        End With
        
        'move to bottom of list
        oRule.ExecutionOrder = colRules.Count
        
        'save rules
        colRules.Save
        Notify ("Creating new rule for: " + strSender + ", please stand by")
        
        'run new rule
        Notify ("Running new rule for: " + strSender + ", please stand by")
        oRule.Execute ShowProgress:=True
    End If
    
    'for testing
    Notify ("Sarah's AutoRule finished")
End Sub

'creates notification messages on form
Sub Notify(message)
    strNote = strNote + vbNewLine + message
    SarahsAutoRuleUserForm.Label1.Caption = strNote
End Sub

'moves email to proper folder
Sub MoveIt()
    Set oTargetFolder = oRule.Actions.MoveToFolder.Folder
    oMail.Move oTargetFolder
    Notify ("Email moved to " + strSender + " folder")
End Sub
