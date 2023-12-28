VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMFTSearch 
   Caption         =   "MFTSearch"
   ClientHeight    =   5520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8370
   OleObjectBlob   =   "frmMFTSearch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMFTSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboDrive_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyCode = vbNull
End Sub

Private Sub cmdSearch_Click()
    Dim filesFound() As String, i As Long
    Dim search1 As clsMFTSearch
    Set search1 = New clsMFTSearch
    
    lstResults.Clear
    filesFound = search1.Find(txtFilename.Text, modFormFuncs.DriveLetters(cboDrive.ListIndex), chkExact.Value, chkCase.Value)
    
    If (Not Not filesFound) <> 0 Then
        For i = 0 To UBound(filesFound)
            lstResults.AddItem filesFound(i)
        Next i
    Else
        lstResults.AddItem "No results found."
    End If
End Sub

Private Sub lstResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Shell "explorer /select," & lstResults.Text, vbNormalFocus 'open the file explorer to the directory and select the file
End Sub

Private Sub UserForm_Initialize()
    'get drives on system
    Dim fs() As String, i As Long
    fs = GetNTFSDrives
    
    For i = 0 To UBound(fs)
        cboDrive.AddItem fs(i)
    Next i
    
    cboDrive.ListIndex = 0
End Sub
