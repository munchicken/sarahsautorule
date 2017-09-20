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

Private Sub cmdLoadUser_Click()
    'look at modify in save_click
    'may need some variables from save_click
End Sub

Private Sub cmdSave_Click()
    Dim strExists As String
    Dim strBasePath     As String
    Dim strMyPath       As String
    Dim strFilename     As String
    Dim strAppPath      As String
    Dim intInput        As Integer
    Dim xmlDoc          As MSXML2.DOMDocument60
    Dim objRoot         As MSXML2.IXMLDOMNode
    Dim objChildFldr    As MSXML2.IXMLDOMElement
    Dim objChildTo      As MSXML2.IXMLDOMElement
    Dim objChildWords   As MSXML2.IXMLDOMElement
    Dim objChildList    As MSXML2.IXMLDOMNode
    Dim objChildWord    As MSXML2.IXMLDOMNode
    Dim strWord         As Variant
    
    strBasePath = Environ("AppData")
    strMyPath = "\Munchicken\"
    strFilename = "Settings.xml"
    strAppPath = "SarahsAutoRule"
    blnFound = False
    g_strUserGrpFolder = Me.txtFolder.Text
    g_blnToCc = Me.chkToCc.Value
    g_blnWords = Me.chkWords.Value
    
    'catch exceptions on filesystem operations
    On Error GoTo cmdSave_Click_Err
    
    'set exception word array
    If g_blnWords = True Then
        g_arrWords = Split(Me.txtWords.Text, vbCrLf)
    End If
    
    'create dir for saving settings, if necessary
    strExists = Dir(strBasePath & strMyPath & strAppPath, vbDirectory)
    If StrComp(strExists, strAppPath, vbTextCompare) <> 0 Then
                MkDir (strBasePath & strMyPath)
                MkDir (strBasePath & strMyPath & strAppPath)  'can only make top level dir, so have to do it twice
    End If
    
    'create file if necesary
    strExists = Dir(strBasePath + strMyPath + strAppPath + strFilename, vbNormal)
    If StrComp(strExists, strFilename, vbTextCompare) <> 0 Then
        'create
        Set xmlDoc = New DOMDocument60
        'create root node
        Set objRoot = xmlDoc.createElement("Settings")
        xmlDoc.appendChild objRoot
        'create folder element
        Set objChildFldr = xmlDoc.createElement("Folder")
        objRoot.appendChild objChildFldr
        Call objChildFldr.setAttribute("Name", Me.txtFolder.Text)
        'create ToCC element
        Set objChildTo = xmlDoc.createElement("ToCC")
        objRoot.appendChild objChildTo
        Call objChildTo.setAttribute("Setting", Me.chkToCc.Value)
        'create Words element
        Set objChildWords = xmlDoc.createElement("Words")
        objRoot.appendChild objChildWords
        Call objChildWords.setAttribute("Setting", Me.chkWords.Value)
        'create word list node
        Set objChildList = xmlDoc.createElement("List")
        objChildWords.appendChild objChildList
        'create word elements
        'Array.Sort(g_arrWords)
        For Each strWord In g_arrWords
            Set objChildWord = xmlDoc.createElement("Word")
            objChildWord.Text = strWord
            objChildList.appendChild objChildWord
        Next
        'save file
        xmlDoc.Save (strBasePath & strMyPath & strAppPath & "\config.xml")
    Else
        'modify
        Set xmlDoc = New DOMDocument60
        xmlDoc.Load (strBasePath & strMyPath & strAppPath & "\config.xml")
        'change folder element
        Set objChildFldr = xmlDoc.getElementsByTagName("Folder")
        If StrComp(objChildFldr.getAttribute("Name"), Me.txtFolder.Text, vbTextCompare) <> 0 Then
            Call objChildFlder.setAttribute("Name", Me.txtFolder.Text)
        End If
        'change ToCC element
        Set objChildTo = xmlDoc.getElementsByTagName("ToCC")
        If objChildTo.getAttribute("ToCC") <> Me.chkToCc.Value Then
            Call objChildTo.setAttribute("Setting", Me.chkToCc.Value)
        End If
        'change Words element
        Set objChildWords = xmlDoc.getElementsByTagName("Words")
        If objChildWords.getAttribute("Words") <> Me.chkWords.Value Then
            Call objChildWords.setAttribute("Setting", Me.chkWords.Value)
        End If
        'change word list node
        Set objChildList = xmlDoc.getElementsByTagName("List")
        
        ' *** should sort array alphebaticaly, then compare, then erase/replace
        ' ** how to read each element and then comapre & erase
        ' * temp using same method as before
        For Each strWord In g_arrWords
            Set objChildWord = xmlDoc.createElement("Word")
            objChildWord.Text = strWord
            objChildList.appendChild objChildWord
        Next
        'save file
        xmlDoc.Save (strBasePath & strMyPath & strAppPath & "\config.xml")
    End If
    
    'end
    GoTo cmdSave_Click_Exit
    
'error handler
cmdSave_Click_Err:
    intInput = MsgBox("Unable to save settings " & strBasePath & strMyPath & strAppPath, vbOKOnly + vbCritical + vbDefaultButton1 + vbApplicationModal, "Error Encountered (" & Err.Number & ")")
    If intInput = vbOK Then
        Unload Me
    End If
'exit
cmdSave_Click_Exit:
    Set xmlDoc = Nothing
    Set objRoot = Nothing
    Set objChildFldr = Nothing
    Set objChildTo = Nothing
    Set objChildWords = Nothing
    Set objChildList = Nothing
    Unload Me
End Sub

