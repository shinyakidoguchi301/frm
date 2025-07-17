VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsData5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'   プロパティ
'
Public Data As Collection
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Sub Store()  '[d]でインスタンス化

    Set Data = New Collection
    
    Dim r As Range
    With Range("A4").CurrentRegion.Offset(1, 0)
        For Each r In .Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Rows
            Dim d As Data: Set d = New Data
            d.Init Range(Cells(r.row, 1), Cells(r.row, 18))
                Data.Add d
        Next r
    End With
    
End Sub
'----科目群ボタンクリック
Private Sub CommandButton1_Click()
    frmKamokuGroup.Show (vbModeless)
End Sub
'----検索ボタンクリック
Private Sub CommandButton2_Click()
    frmSearch.Show (vbModeless)
End Sub
'----得意先ボタンクリック
Private Sub CommandButton3_Click()
    frmClientGroup.Show (vbModeless)
End Sub

Sub ShowRegularPaymentForm()
    RegularPaymentForm.Show (vbModeless)
End Sub

'Private Sub Worksheet_Change(ByVal Target As Range)
'    Dim objCustom As Object
'    Dim myRange As Range    'マスタのセル範囲
'    Dim varList As Variant  'プロパティFormula1の文字列作成用
'    Dim strAdr  As String   '最初にヒットしたセル
'
'    'マスタリストの範囲をセット
'    Set myRange = wsClient.Range("B3:B500") 'Worksheets("得意先マスタ").Range("B1:B400")
'
'    With ActiveSheet
'        If Target.Count > 1 Then Exit Sub
'            If Target = .Range("C1") Then
'                On Error Resume Next
'                Set objCustom = myRange.Find(what:=Target.Value, LookAt:=xlPart)
'
'                Application.EnableEvents = False
'
'                If objCustom Is Nothing Then
'                    Target.Value = Target.Value
'                Else
'                    '1件目を文字列にセット
'                    varList = objCustom
'                    strAdr = objCustom.Address
'
'                    Do
'                        Set objCustom = myRange.FindNext(objCustom)
'                        If objCustom Is Nothing Then
'                            Exit Do
'                        Else
'                            If strAdr <> objCustom.Address Then
'                                varList = varList & "," & objCustom
'                            End If
'                        End If
'                    Loop While Not objCustom Is Nothing And objCustom.Address <> strAdr
'
'                    With Target.Validation
'                        .Delete
'                        .Add Type:=xlValidateList, _
'                        Formula1:=varList
'                        .ShowError = False
'                    End With
'
'                End If
'                Application.EnableEvents = True
'            End If
'    End With
'
'End Sub

Sub 支払先選択()
    Dim selectedRange As Range
    Set selectedRange = Selection

    Dim cell As Range
    For Each cell In selectedRange
        If cell.Column = 2 Then
            Range("C1").Value = cell.Value
            Exit For
        End If
    Next cell


End Sub



Public Sub AddRow()
    Dim n As Long
'    Range("A4").ListObject.ListRows.Add
    n = Range("A4").ListObject.ListRows.Count

    Dim i As Long: i = n + 5
    With Range("A4").ListObject.ListRows(n)
    Range(Cells(i, 1), Cells(i, 2)) = Array( _
        Range("C2").Value, _
        Range("C1").Value _
        )
    End With
'------------------------------------------------------
'    Dim n As Long
'    Range("A4").ListObject.ListRows.Add
'    n = Range("A4").ListObject.ListRows.Count
'
'    With Range("A4").ListObject.ListRows(n)
'        .Range(1) = Range("C2").Value
'        .Range(2) = Range("C1").Value
'        .Range(3) = "北國BK"
'        .Range(16) = 1.1
'        .Range(17) = .Range(17).Offset(-1, 0).FormulaR1C1
'        .Range(18) = .Range(18).Offset(-1, 0).FormulaR1C1
'    End With
'------------------------------------------------------


End Sub
Public Sub AddRows(ByVal targetData As Collection)

    Dim d As Object
    For Each d In targetData
        Dim n As Long
        n = Range("A4").ListObject.ListRows.Count
    
        Dim i As Long: i = n + 5
        With Range("A4").ListObject.ListRows(n)
        Range(Cells(i, 1), Cells(i, 20)) = Array( _
            Range("C2").Value, _
            Range("C1").Value, _
            d.BkAccount, _
            d.SiteId, _
            , _
            d.PeresonIC, _
            d.AccountID, _
            , _
             , _
             d.Content1, _
             d.Content2, _
             d.Content3, _
             d.Quantity, _
             d.unit, _
             d.UPWTax, _
             d.tax, _
            , _
            , , _
             d.hasuu _
            )
        End With
        
'        Range(Cells(i, 2), Cells(i, 11)) = Array(d.SiteName, d.SiteId, d.Account, d.Content1, d.Content2, d.Content3, d.Quantity, d.unit, d.UnitPrice, d.Price)
'        Cells(i, 13) = Array(d.PeresonIC)
'        Range(Cells(i, 15), Cells(i, 16)) = Array(d.UPWTax, d.Tax)
        
        i = i + 1
    Next d

End Sub

