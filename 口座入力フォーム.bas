Attribute VB_Name = "口座入力フォーム"
Option Explicit
Public Sub AddBasicControls(frmPayee As MSForms.Frame, payee As String, topOffset As Integer)
    ' 金額
    frmPayee.Controls.Add "Forms.Label.1", "Label_Amount_" & payee
    frmPayee.Controls.Add "Forms.TextBox.1", "TextBox_Amount_" & payee
    With frmPayee.Controls("Label_Amount_" & payee)
        .caption = "金額"
        .Left = 10: .Top = topOffset: .Width = 40
    End With
    With frmPayee.Controls("TextBox_Amount_" & payee)
        .Left = 60: .Top = topOffset - 2: .Width = 80
    End With
    ' 摘要1?3
    Dim j As Integer
    For j = 1 To 3
        frmPayee.Controls.Add "Forms.Label.1", "Label_Tekiyo" & j & "_" & payee
        frmPayee.Controls.Add "Forms.TextBox.1", "TextBox_Tekiyo" & j & "_" & payee
        With frmPayee.Controls("Label_Tekiyo" & j & "_" & payee)
            .caption = "摘要" & j
            .Left = 10: .Top = topOffset + 20 * j: .Width = 40
        End With
        With frmPayee.Controls("TextBox_Tekiyo" & j & "_" & payee)
            .Left = 60: .Top = topOffset + 20 * j - 2: .Width = 120
        End With
    Next j

    ' 支払日
    frmPayee.Controls.Add "Forms.Label.1", "Label_Date_" & payee
    frmPayee.Controls.Add "Forms.TextBox.1", "TextBox_Date_" & payee
    With frmPayee.Controls("Label_Date_" & payee)
        .caption = "支払日"
        .Left = 10: .Top = topOffset + 80: .Width = 50
    End With
    With frmPayee.Controls("TextBox_Date_" & payee)
        .Left = 70: .Top = topOffset + 78: .Width = 100
        .Value = Format(DateSerial(year(Date), month(Date) + 1, 0), "yyyy/mm/dd")
    End With

    ' 科目
    frmPayee.Controls.Add "Forms.Label.1", "Label_Kamoku_" & payee
    frmPayee.Controls.Add "Forms.ComboBox.1", "ComboBox_Kamoku_" & payee
    With frmPayee.Controls("Label_Kamoku_" & payee)
        .caption = "口座（科目）"
        .Left = 10: .Top = topOffset + 100: .Width = 80
    End With
    With frmPayee.Controls("ComboBox_Kamoku_" & payee)
        .Left = 100: .Top = topOffset + 98: .Width = 100
    End With

    ' 補助
    frmPayee.Controls.Add "Forms.Label.1", "Label_Hojo_" & payee
    frmPayee.Controls.Add "Forms.ComboBox.1", "ComboBox_Hojo_" & payee
    With frmPayee.Controls("Label_Hojo_" & payee)
        .caption = "補助"
        .Left = 10: .Top = topOffset + 120: .Width = 80
    End With
    With frmPayee.Controls("ComboBox_Hojo_" & payee)
        .Left = 100: .Top = topOffset + 118: .Width = 100
    End With
End Sub

Public Sub ApplyInitialValues(frmPayee As MSForms.Frame, uniqueID As String, selectedMonth As Integer, selectedYear As Integer, category As String, denpyo As String)
    Dim payee As String
    payee = Split(uniqueID, "_")(0) ' ユニークIDから取引先名を抽出

    Dim dataRow As Range
    Set dataRow = GetMonthlyDataRow(payee, category, denpyo, selectedYear, selectedMonth)

    If Not dataRow Is Nothing Then
        Call ApplyMonthlyHistoryToFrame(frmPayee, uniqueID, dataRow)
    Else
        Dim latestRow As Range
        Set latestRow = GetLatestHistoryRow(payee, category, denpyo)

        If Not latestRow Is Nothing Then
            Dim tbl As ListObject
            Set tbl = latestRow.ListObject

            frmPayee.Controls("TextBox_Amount_" & uniqueID).Value = latestRow.Cells(tbl.ListColumns("借方金額").index).Value
            frmPayee.Controls("TextBox_Amount_" & uniqueID).ForeColor = RGB(150, 150, 150)

            frmPayee.Controls("TextBox_Tekiyo1_" & uniqueID).Value = latestRow.Cells(tbl.ListColumns("摘要1").index).Value
            frmPayee.Controls("TextBox_Tekiyo1_" & uniqueID).ForeColor = RGB(150, 150, 150)

            frmPayee.Controls("TextBox_Tekiyo3_" & uniqueID).Value = latestRow.Cells(tbl.ListColumns("摘要3").index).Value
            frmPayee.Controls("TextBox_Tekiyo3_" & uniqueID).ForeColor = RGB(150, 150, 150)

            frmPayee.Controls("ComboBox_Kamoku_" & uniqueID).Value = latestRow.Cells(tbl.ListColumns("貸方科目").index).Value
            frmPayee.Controls("ComboBox_Kamoku_" & uniqueID).ForeColor = RGB(150, 150, 150)

            frmPayee.Controls("ComboBox_Hojo_" & uniqueID).Value = latestRow.Cells(tbl.ListColumns("貸方補助").index).Value
            frmPayee.Controls("ComboBox_Hojo_" & uniqueID).ForeColor = RGB(150, 150, 150)
        End If

        ' 摘要2補完
        Dim wsReg As Worksheet
        Set wsReg = ThisWorkbook.Sheets("口　　座")
        Dim tblReg As ListObject
        Set tblReg = wsReg.ListObjects("レギュラーリスト")

        Dim tekiyoMonthOffset As Integer: tekiyoMonthOffset = 0
        Dim r As ListRow
        For Each r In tblReg.ListRows
            With r.Range
                If .Cells(tblReg.ListColumns("取引先").index).Value = payee And _
                   .Cells(tblReg.ListColumns("カテゴリ").index).Value = category Then
                    tekiyoMonthOffset = Val(.Cells(tblReg.ListColumns("摘要月").index).Value)
                    Exit For
                End If
            End With
        Next r

        Dim tekiyoMonth As Integer
        tekiyoMonth = selectedMonth + tekiyoMonthOffset
        If tekiyoMonth < 1 Then tekiyoMonth = tekiyoMonth + 12
        If tekiyoMonth > 12 Then tekiyoMonth = tekiyoMonth - 12

        frmPayee.Controls("TextBox_Tekiyo2_" & uniqueID).Value = tekiyoMonth & "月分"
        frmPayee.Controls("TextBox_Tekiyo2_" & uniqueID).ForeColor = RGB(150, 150, 150)

        frmPayee.BackColor = RGB(255, 230, 230)
    End If
End Sub



Public Function GetMonthlyDataRow(payee As String, category As String, denpyo As String, selectedYear As Integer, selectedMonth As Integer) As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("口　　座")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("Data口座")

    Dim latestDate As Date: latestDate = #1/1/1900#
    Dim latestRow As Range

    Dim r As ListRow
    For Each r In tbl.ListRows
        With r.Range
            If .Cells(tbl.ListColumns("取引先").index).Value = payee And _
               .Cells(tbl.ListColumns("カテゴリ").index).Value = category And _
               .Cells(tbl.ListColumns("伝票").index).Value = denpyo And _
               .Cells(tbl.ListColumns("年").index).Value = selectedYear And _
               .Cells(tbl.ListColumns("月").index).Value = selectedMonth Then

                Dim d As Variant
                d = .Cells(tbl.ListColumns("支払日").index).Value
                If IsDate(d) And d > latestDate Then
                    latestDate = d
                    Set latestRow = r.Range
                End If
            End If
        End With
    Next r

    Set GetMonthlyDataRow = latestRow
End Function


Public Sub ApplyMonthlyHistoryToFrame(frmPayee As MSForms.Frame, payee As String, dataRow As Range)
    Dim tbl As ListObject
    Set tbl = dataRow.ListObject

    ' 初期値反映
    frmPayee.Controls("TextBox_Amount_" & payee).Value = dataRow.Cells(tbl.ListColumns("借方金額").index).Value
    frmPayee.Controls("TextBox_Tekiyo1_" & payee).Value = dataRow.Cells(tbl.ListColumns("摘要1").index).Value
    frmPayee.Controls("TextBox_Tekiyo2_" & payee).Value = dataRow.Cells(tbl.ListColumns("摘要2").index).Value
    frmPayee.Controls("TextBox_Tekiyo3_" & payee).Value = dataRow.Cells(tbl.ListColumns("摘要3").index).Value
    frmPayee.Controls("TextBox_Date_" & payee).Value = Format(dataRow.Cells(tbl.ListColumns("支払日").index).Value, "yyyy/mm/dd")
    frmPayee.Controls("ComboBox_Kamoku_" & payee).Value = dataRow.Cells(tbl.ListColumns("貸方科目").index).Value
    frmPayee.Controls("ComboBox_Hojo_" & payee).Value = dataRow.Cells(tbl.ListColumns("貸方補助").index).Value
    ' 背景色：保存済み（緑）
    frmPayee.BackColor = RGB(230, 255, 230)
End Sub

Public Sub AddSaveButtonToFrame(frmPayee As MSForms.Frame, payee As String)
    frmPayee.Controls.Add "Forms.CommandButton.1", "CommandButton_Save_" & payee
    With frmPayee.Controls("CommandButton_Save_" & payee)
        .caption = "保存"
        .Left = 10
        .Top = frmPayee.Height - 30
        .Width = 60
    End With
End Sub

Public Sub SetKamokuDropdown(cmb As MSForms.ComboBox)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("勘定科目")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("勘定科目")

    Dim r As ListRow
    cmb.Clear

    For Each r In tbl.ListRows
        With r.Range
            If .Cells(tbl.ListColumns("大カテゴリ").index).Value = "資産の部" Then
                cmb.AddItem .Cells(tbl.ListColumns("勘定科目").index).Value
            End If
        End With
    Next r
End Sub

Public Sub SetHojoDropdown(cmb As MSForms.ComboBox)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("BK")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("BK")

    Dim r As ListRow
    cmb.Clear

    For Each r In tbl.ListRows
        With r.Range
            cmb.AddItem .Cells(tbl.ListColumns("補助科目").index).Value
        End With
    Next r
End Sub
Private Function GetPageIndexByCaption(mp As MSForms.MultiPage, caption As String) As Integer
    Dim i As Integer
    For i = 0 To mp.Pages.Count - 1
        If mp.Pages(i).caption = caption Then
            GetPageIndexByCaption = i
            Exit Function
        End If
    Next i
    GetPageIndexByCaption = -1
End Function


Public Sub SaveSingleEntry(frm As RegularPaymentForm, frameID As String, denpyo As String)
    Dim payee As String
    payee = Split(frameID, "_")(0)
    Dim category As String
    category = frm.ComboBox_Category.Value

    ' === 該当伝票ページを取得 ===
    Dim pageIndex As Integer
    pageIndex = GetPageIndexByCaption(frm.MultiPage_Denpyo, denpyo)

    If pageIndex = -1 Then
        MsgBox "伝票ページが見つかりません：" & denpyo, vbExclamation
        Exit Sub
    End If

    Dim pg As Object
    Set pg = frm.MultiPage_Denpyo.Pages(pageIndex)

    ' === 該当フレームを取得 ===
    Dim ctrlFrame As MSForms.Frame
    On Error Resume Next
    Set ctrlFrame = pg.Controls("Frame_Payee_" & frameID)
    On Error GoTo 0

    If ctrlFrame Is Nothing Then
        MsgBox "取引先フレームが見つかりません：" & frameID, vbExclamation
        Exit Sub
    End If

    ' === 入力値取得 ===
    Dim amount As Double
    amount = Val(ctrlFrame.Controls("TextBox_Amount_" & frameID).Value)

    Dim tekiyo2 As String
    tekiyo2 = ctrlFrame.Controls("TextBox_Tekiyo2_" & frameID).Value

    Dim shiharaiDate As Date
    shiharaiDate = DateValue(ctrlFrame.Controls("TextBox_Date_" & frameID).Value)

    Dim kamoku As String
    kamoku = ctrlFrame.Controls("ComboBox_Kamoku_" & frameID).Value
    Dim hojo As String
    hojo = ctrlFrame.Controls("ComboBox_Hojo_" & frameID).Value

    ' === 履歴取得 ===
    Dim prevData As Variant
    prevData = GetPreviousMonthData(payee, category, denpyo, frm.SpinButton_Month.Value)

    Dim tekiyo1 As String, tekiyo3 As String
    tekiyo1 = prevData(2)
    tekiyo3 = prevData(4)

    ' === 保存処理 ===
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Sheets("口　　座")
    Dim tblData As ListObject
    Set tblData = wsData.ListObjects("Data口座")

    Dim newRow As ListRow
    Set newRow = tblData.ListRows.Add

    With newRow.Range
        .Cells(tblData.ListColumns("支払日").index).Value = shiharaiDate
        .Cells(tblData.ListColumns("取引先").index).Value = payee
        .Cells(tblData.ListColumns("カテゴリ").index).Value = category
        .Cells(tblData.ListColumns("伝票").index).Value = denpyo
        .Cells(tblData.ListColumns("No").index).Value = tblData.ListRows.Count
        .Cells(tblData.ListColumns("行番").index).Value = 0 ' 単一行として保存

        .Cells(tblData.ListColumns("借方金額").index).Value = amount
        .Cells(tblData.ListColumns("貸方金額").index).Value = amount
        .Cells(tblData.ListColumns("摘要1").index).Value = tekiyo1
        .Cells(tblData.ListColumns("摘要2").index).Value = tekiyo2
        .Cells(tblData.ListColumns("摘要3").index).Value = tekiyo3

        .Cells(tblData.ListColumns("貸方補助").index).Value = hojo
    End With

    ' === 視覚的フィードバック ===
    Dim ctrlList As New Collection
    ctrlList.Add ctrlFrame.Controls("TextBox_Amount_" & frameID)
    ctrlList.Add ctrlFrame.Controls("TextBox_Tekiyo1_" & frameID)
    ctrlList.Add ctrlFrame.Controls("TextBox_Tekiyo2_" & frameID)
    ctrlList.Add ctrlFrame.Controls("TextBox_Tekiyo3_" & frameID)
    Call FinalizeSave(ctrlList, ctrlFrame)

    MsgBox "保存しました：" & payee, vbInformation
End Sub




Public Sub FilterDataByPayee(frameID As String)
    Dim payee As String
    payee = Split(frameID, "_")(0)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("口　　座")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("Data口座")

    If ws.AutoFilterMode Then ws.ShowAllData

    Dim colIndex As Long
    colIndex = tbl.ListColumns("取引先").index

    tbl.Range.AutoFilter Field:=colIndex, Criteria1:=payee
End Sub


Public Function GetLatestHistoryRow(payee As String, category As String, denpyo As String) As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("口　　座")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("Data口座")

    Dim latestDate As Date: latestDate = #1/1/1900#
    Dim latestRow As Range

    Dim r As ListRow
    For Each r In tbl.ListRows
        With r.Range
            If .Cells(tbl.ListColumns("取引先").index).Value = payee And _
               .Cells(tbl.ListColumns("カテゴリ").index).Value = category And _
               .Cells(tbl.ListColumns("伝票").index).Value = denpyo Then

                Dim d As Variant
                d = .Cells(tbl.ListColumns("支払日").index).Value
                If IsDate(d) And d > latestDate Then
                    latestDate = d
                    Set latestRow = r.Range
                End If
            End If
        End With
    Next r

    Set GetLatestHistoryRow = latestRow
End Function



' === 前月データ取得 ===

Public Function GetPreviousMonthData(payee As String, category As String, denpyo As String, selectedMonth As Integer) As Variant

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("口　　座").ListObjects("Data口座")

    Dim prevMonth As Integer, prevYear As Integer
    If selectedMonth = 1 Then
        prevMonth = 12
        prevYear = year(Date) - 1
    Else
        prevMonth = selectedMonth - 1
        prevYear = year(Date)
    End If
    Dim r As ListRow
    Dim result(1 To 4) As Variant
    Dim totalAmount As Double: totalAmount = 0
    Dim tekiyo1List As Object: Set tekiyo1List = CreateObject("Scripting.Dictionary")
    Dim tekiyo2List As Object: Set tekiyo2List = CreateObject("Scripting.Dictionary")
    Dim tekiyo3List As Object: Set tekiyo3List = CreateObject("Scripting.Dictionary")

    For Each r In tbl.ListRows
        With r.Range
            If .Cells(tbl.ListColumns("取引先").index).Value = payee And _
               .Cells(tbl.ListColumns("カテゴリ").index).Value = category And _
               .Cells(tbl.ListColumns("年").index).Value = prevYear And _
               .Cells(tbl.ListColumns("月").index).Value = prevMonth Then

                totalAmount = totalAmount + Val(.Cells(tbl.ListColumns("借方金額").index).Value)

                Dim t1 As String: t1 = Trim(.Cells(tbl.ListColumns("摘要1").index).Value)
                Dim t2 As String: t2 = Trim(.Cells(tbl.ListColumns("摘要2").index).Value)
                Dim t3 As String: t3 = Trim(.Cells(tbl.ListColumns("摘要3").index).Value)

                If Len(t1) > 0 Then tekiyo1List(t1) = 1
                If Len(t2) > 0 Then tekiyo2List(t2) = 1
                If Len(t3) > 0 Then tekiyo3List(t3) = 1
            End If
        End With
    Next r

    result(1) = totalAmount
    result(2) = Join(tekiyo1List.Keys, "／")
    result(3) = Join(tekiyo2List.Keys, "／")
    result(4) = Join(tekiyo3List.Keys, "／")

    GetPreviousMonthData = result
End Function

' === 保存後に色を黒に変更し、背景色を緑に ===
Public Sub FinalizeSave(ctrls As Collection, frameCtrl As Object)
    Dim ctrl As Object
    For Each ctrl In ctrls
        ctrl.ForeColor = RGB(0, 0, 0)
    Next ctrl
    frameCtrl.BackColor = RGB(230, 255, 230) ' 保存済み（緑）
End Sub
