VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SarahsAutoRuleUserForm 
   Caption         =   "Sarah's AutoRule"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5160
   OleObjectBlob   =   "SarahsAutoRuleUserForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SarahsAutoRuleUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OkButton_Click()
    blnDone = True
    Unload Me
End Sub

