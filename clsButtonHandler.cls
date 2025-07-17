VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsButtonHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' === clsButtonHandler ===
Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1
Public payee As String
Public denpyo As String
Public frameID As String ' Å© í«â¡
Public actionType As String
Public ParentForm As Object


Private Sub Btn_Click()
    Select Case actionType
        Case "Save"
            Call SaveSingleEntry(ParentForm, frameID, denpyo)
        Case "Filter"
            Call FilterDataByPayee(frameID)
    End Select
End Sub


