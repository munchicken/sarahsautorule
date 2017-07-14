VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "Settings for New Rule"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3705
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkWords_Click()
    If Me.chkWords.Value = True Then
        Me.txtWords.Enabled = True
    End If
End Sub

Private Sub cmdLoadDefault_Click()
    Me.txtFolder.Text = "Contact Groups"
    Me.chkToCc.Value = True
    Me.chkWords.Value = True
    Me.txtWords.Text = "urgent"
    Me.txtWords.Enabled = True
End Sub

Private Sub cmdSave_Click()
    g_strUserGrpFolder = Me.txtFolder.Text
    g_blnToCc = Me.chkToCc.Value
    g_blnWords = Me.chkWords.Value
    If g_blnWords = True Then
        g_arrWords = Split(Me.txtWords.Text, vbCrLf)
    End If
    Unload Me
End Sub

