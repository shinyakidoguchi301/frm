Attribute VB_Name = "�������̓t�H�[��"
Option Explicit
Public Sub AddBasicControls(frmPayee As MSForms.Frame, payee As String, topOffset As Integer)
    ' ���z
    frmPayee.Controls.Add "Forms.Label.1", "Label_Amount_" & payee
    frmPayee.Controls.Add "Forms.TextBox.1", "TextBox_Amount_" & payee
    With frmPayee.Controls("Label_Amount_" & payee)
        .caption = "���z"
        .Left = 10: .Top = topOffset: .Width = 40
    End With
    With frmPayee.Controls("TextBox_Amount_" & payee)
        .Left = 60: .Top = topOffset - 2: .Width = 80
    End With
    ' �E�v1?3
    Dim j As Integer
    For j = 1 To 3
        frmPayee.Controls.Add "Forms.Label.1", "Label_Tekiyo" & j & "_" & payee
        frmPayee.Controls.Add "Forms.TextBox.1", "TextBox_Tekiyo" & j & "_" & payee
        With frmPayee.Controls("Label_Tekiyo" & j & "_" & payee)
            .caption = "�E�v" & j
            .Left = 10: .Top = topOffset + 20 * j: .Width = 40
        End With
        With frmPayee.Controls("TextBox_Tekiyo" & j & "_" & payee)
            .Left = 60: .Top = topOffset + 20 * j - 2: .Width = 120
        End With
    Next j

    ' �x����
    frmPayee.Controls.Add "Forms.Label.1", "Label_Date_" & payee
    frmPayee.Controls.Add "Forms.TextBox.1", "TextBox_Date_" & payee
    With frmPayee.Controls("Label_Date_" & payee)
        .caption = "�x����"
        .Left = 10: .Top = topOffset + 80: .Width = 50
    End With
    With frmPayee.Controls("TextBox_Date_" & payee)
        .Left = 70: .Top = topOffset + 78: .Width = 100
        .Value = Format(DateSerial(year(Date), month(Date) + 1, 0), "yyyy/mm/dd")
    End With

    ' �Ȗ�
    frmPayee.Controls.Add "Forms.Label.1", "Label_Kamoku_" & payee
    frmPayee.Controls.Add "Forms.ComboBox.1", "ComboBox_Kamoku_" & payee
    With frmPayee.Controls("Label_Kamoku_" & payee)
        .caption = "�����i�Ȗځj"
        .Left = 10: .Top = topOffset + 100: .Width = 80
    End With
    With frmPayee.Controls("ComboBox_Kamoku_" & payee)
        .Left = 100: .Top = topOffset + 98: .Width = 100
    End With

    ' �⏕
    frmPayee.Controls.Add "Forms.Label.1", "Label_Hojo_" & payee
    frmPayee.Controls.Add "Forms.ComboBox.1", "ComboBox_Hojo_" & payee
    With frmPayee.Controls("Label_Hojo_" & payee)
        .caption = "�⏕"
        .Left = 10: .Top = topOffset + 120: .Width = 80
    End With
    With frmPayee.Controls("ComboBox_Hojo_" & payee)
        .Left = 100: .Top = topOffset + 118: .Width = 100
    End With
End Sub

Public Sub ApplyInitialValues(frmPayee As MSForms.Frame, uniqueID As String, selectedMonth As Integer, selectedYear As Integer, category As String, denpyo As String)
    Dim payee As String
    payee = Split(uniqueID, "_")(0) ' ���j�[�NID�������於�𒊏o

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

            frmPayee.Controls("TextBox_Amount_" & uniqueID).Value = latestRow.Cells(tbl.ListColumns("�ؕ����z").index).Value
            frmPayee.Controls("TextBox_Amount_" & uniqueID).ForeColor = RGB(150, 150, 150)

            frmPayee.Controls("TextBox_Tekiyo1_" & uniqueID).Value = latestRow.Cells(tbl.ListColumns("�E�v1").index).Value
            frmPayee.Controls("TextBox_Tekiyo1_" & uniqueID).ForeColor = RGB(150, 150, 150)

            frmPayee.Controls("TextBox_Tekiyo3_" & uniqueID).Value = latestRow.Cells(tbl.ListColumns("�E�v3").index).Value
            frmPayee.Controls("TextBox_Tekiyo3_" & uniqueID).ForeColor = RGB(150, 150, 150)

            frmPayee.Controls("ComboBox_Kamoku_" & uniqueID).Value = latestRow.Cells(tbl.ListColumns("�ݕ��Ȗ�").index).Value
            frmPayee.Controls("ComboBox_Kamoku_" & uniqueID).ForeColor = RGB(150, 150, 150)

            frmPayee.Controls("ComboBox_Hojo_" & uniqueID).Value = latestRow.Cells(tbl.ListColumns("�ݕ��⏕").index).Value
            frmPayee.Controls("ComboBox_Hojo_" & uniqueID).ForeColor = RGB(150, 150, 150)
        End If

        ' �E�v2�⊮
        Dim wsReg As Worksheet
        Set wsReg = ThisWorkbook.Sheets("���@�@��")
        Dim tblReg As ListObject
        Set tblReg = wsReg.ListObjects("���M�����[���X�g")

        Dim tekiyoMonthOffset As Integer: tekiyoMonthOffset = 0
        Dim r As ListRow
        For Each r In tblReg.ListRows
            With r.Range
                If .Cells(tblReg.ListColumns("�����").index).Value = payee And _
                   .Cells(tblReg.ListColumns("�J�e�S��").index).Value = category Then
                    tekiyoMonthOffset = Val(.Cells(tblReg.ListColumns("�E�v��").index).Value)
                    Exit For
                End If
            End With
        Next r

        Dim tekiyoMonth As Integer
        tekiyoMonth = selectedMonth + tekiyoMonthOffset
        If tekiyoMonth < 1 Then tekiyoMonth = tekiyoMonth + 12
        If tekiyoMonth > 12 Then tekiyoMonth = tekiyoMonth - 12

        frmPayee.Controls("TextBox_Tekiyo2_" & uniqueID).Value = tekiyoMonth & "����"
        frmPayee.Controls("TextBox_Tekiyo2_" & uniqueID).ForeColor = RGB(150, 150, 150)

        frmPayee.BackColor = RGB(255, 230, 230)
    End If
End Sub



Public Function GetMonthlyDataRow(payee As String, category As String, denpyo As String, selectedYear As Integer, selectedMonth As Integer) As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("���@�@��")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("Data����")

    Dim latestDate As Date: latestDate = #1/1/1900#
    Dim latestRow As Range

    Dim r As ListRow
    For Each r In tbl.ListRows
        With r.Range
            If .Cells(tbl.ListColumns("�����").index).Value = payee And _
               .Cells(tbl.ListColumns("�J�e�S��").index).Value = category And _
               .Cells(tbl.ListColumns("�`�[").index).Value = denpyo And _
               .Cells(tbl.ListColumns("�N").index).Value = selectedYear And _
               .Cells(tbl.ListColumns("��").index).Value = selectedMonth Then

                Dim d As Variant
                d = .Cells(tbl.ListColumns("�x����").index).Value
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

    ' �����l���f
    frmPayee.Controls("TextBox_Amount_" & payee).Value = dataRow.Cells(tbl.ListColumns("�ؕ����z").index).Value
    frmPayee.Controls("TextBox_Tekiyo1_" & payee).Value = dataRow.Cells(tbl.ListColumns("�E�v1").index).Value
    frmPayee.Controls("TextBox_Tekiyo2_" & payee).Value = dataRow.Cells(tbl.ListColumns("�E�v2").index).Value
    frmPayee.Controls("TextBox_Tekiyo3_" & payee).Value = dataRow.Cells(tbl.ListColumns("�E�v3").index).Value
    frmPayee.Controls("TextBox_Date_" & payee).Value = Format(dataRow.Cells(tbl.ListColumns("�x����").index).Value, "yyyy/mm/dd")
    frmPayee.Controls("ComboBox_Kamoku_" & payee).Value = dataRow.Cells(tbl.ListColumns("�ݕ��Ȗ�").index).Value
    frmPayee.Controls("ComboBox_Hojo_" & payee).Value = dataRow.Cells(tbl.ListColumns("�ݕ��⏕").index).Value
    ' �w�i�F�F�ۑ��ς݁i�΁j
    frmPayee.BackColor = RGB(230, 255, 230)
End Sub

Public Sub AddSaveButtonToFrame(frmPayee As MSForms.Frame, payee As String)
    frmPayee.Controls.Add "Forms.CommandButton.1", "CommandButton_Save_" & payee
    With frmPayee.Controls("CommandButton_Save_" & payee)
        .caption = "�ۑ�"
        .Left = 10
        .Top = frmPayee.Height - 30
        .Width = 60
    End With
End Sub

Public Sub SetKamokuDropdown(cmb As MSForms.ComboBox)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("����Ȗ�")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("����Ȗ�")

    Dim r As ListRow
    cmb.Clear

    For Each r In tbl.ListRows
        With r.Range
            If .Cells(tbl.ListColumns("��J�e�S��").index).Value = "���Y�̕�" Then
                cmb.AddItem .Cells(tbl.ListColumns("����Ȗ�").index).Value
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
            cmb.AddItem .Cells(tbl.ListColumns("�⏕�Ȗ�").index).Value
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

    ' === �Y���`�[�y�[�W���擾 ===
    Dim pageIndex As Integer
    pageIndex = GetPageIndexByCaption(frm.MultiPage_Denpyo, denpyo)

    If pageIndex = -1 Then
        MsgBox "�`�[�y�[�W��������܂���F" & denpyo, vbExclamation
        Exit Sub
    End If

    Dim pg As Object
    Set pg = frm.MultiPage_Denpyo.Pages(pageIndex)

    ' === �Y���t���[�����擾 ===
    Dim ctrlFrame As MSForms.Frame
    On Error Resume Next
    Set ctrlFrame = pg.Controls("Frame_Payee_" & frameID)
    On Error GoTo 0

    If ctrlFrame Is Nothing Then
        MsgBox "�����t���[����������܂���F" & frameID, vbExclamation
        Exit Sub
    End If

    ' === ���͒l�擾 ===
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

    ' === �����擾 ===
    Dim prevData As Variant
    prevData = GetPreviousMonthData(payee, category, denpyo, frm.SpinButton_Month.Value)

    Dim tekiyo1 As String, tekiyo3 As String
    tekiyo1 = prevData(2)
    tekiyo3 = prevData(4)

    ' === �ۑ����� ===
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Sheets("���@�@��")
    Dim tblData As ListObject
    Set tblData = wsData.ListObjects("Data����")

    Dim newRow As ListRow
    Set newRow = tblData.ListRows.Add

    With newRow.Range
        .Cells(tblData.ListColumns("�x����").index).Value = shiharaiDate
        .Cells(tblData.ListColumns("�����").index).Value = payee
        .Cells(tblData.ListColumns("�J�e�S��").index).Value = category
        .Cells(tblData.ListColumns("�`�[").index).Value = denpyo
        .Cells(tblData.ListColumns("No").index).Value = tblData.ListRows.Count
        .Cells(tblData.ListColumns("�s��").index).Value = 0 ' �P��s�Ƃ��ĕۑ�

        .Cells(tblData.ListColumns("�ؕ����z").index).Value = amount
        .Cells(tblData.ListColumns("�ݕ����z").index).Value = amount
        .Cells(tblData.ListColumns("�E�v1").index).Value = tekiyo1
        .Cells(tblData.ListColumns("�E�v2").index).Value = tekiyo2
        .Cells(tblData.ListColumns("�E�v3").index).Value = tekiyo3

        .Cells(tblData.ListColumns("�ݕ��⏕").index).Value = hojo
    End With

    ' === ���o�I�t�B�[�h�o�b�N ===
    Dim ctrlList As New Collection
    ctrlList.Add ctrlFrame.Controls("TextBox_Amount_" & frameID)
    ctrlList.Add ctrlFrame.Controls("TextBox_Tekiyo1_" & frameID)
    ctrlList.Add ctrlFrame.Controls("TextBox_Tekiyo2_" & frameID)
    ctrlList.Add ctrlFrame.Controls("TextBox_Tekiyo3_" & frameID)
    Call FinalizeSave(ctrlList, ctrlFrame)

    MsgBox "�ۑ����܂����F" & payee, vbInformation
End Sub




Public Sub FilterDataByPayee(frameID As String)
    Dim payee As String
    payee = Split(frameID, "_")(0)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("���@�@��")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("Data����")

    If ws.AutoFilterMode Then ws.ShowAllData

    Dim colIndex As Long
    colIndex = tbl.ListColumns("�����").index

    tbl.Range.AutoFilter Field:=colIndex, Criteria1:=payee
End Sub


Public Function GetLatestHistoryRow(payee As String, category As String, denpyo As String) As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("���@�@��")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("Data����")

    Dim latestDate As Date: latestDate = #1/1/1900#
    Dim latestRow As Range

    Dim r As ListRow
    For Each r In tbl.ListRows
        With r.Range
            If .Cells(tbl.ListColumns("�����").index).Value = payee And _
               .Cells(tbl.ListColumns("�J�e�S��").index).Value = category And _
               .Cells(tbl.ListColumns("�`�[").index).Value = denpyo Then

                Dim d As Variant
                d = .Cells(tbl.ListColumns("�x����").index).Value
                If IsDate(d) And d > latestDate Then
                    latestDate = d
                    Set latestRow = r.Range
                End If
            End If
        End With
    Next r

    Set GetLatestHistoryRow = latestRow
End Function



' === �O���f�[�^�擾 ===

Public Function GetPreviousMonthData(payee As String, category As String, denpyo As String, selectedMonth As Integer) As Variant

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("���@�@��").ListObjects("Data����")

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
            If .Cells(tbl.ListColumns("�����").index).Value = payee And _
               .Cells(tbl.ListColumns("�J�e�S��").index).Value = category And _
               .Cells(tbl.ListColumns("�N").index).Value = prevYear And _
               .Cells(tbl.ListColumns("��").index).Value = prevMonth Then

                totalAmount = totalAmount + Val(.Cells(tbl.ListColumns("�ؕ����z").index).Value)

                Dim t1 As String: t1 = Trim(.Cells(tbl.ListColumns("�E�v1").index).Value)
                Dim t2 As String: t2 = Trim(.Cells(tbl.ListColumns("�E�v2").index).Value)
                Dim t3 As String: t3 = Trim(.Cells(tbl.ListColumns("�E�v3").index).Value)

                If Len(t1) > 0 Then tekiyo1List(t1) = 1
                If Len(t2) > 0 Then tekiyo2List(t2) = 1
                If Len(t3) > 0 Then tekiyo3List(t3) = 1
            End If
        End With
    Next r

    result(1) = totalAmount
    result(2) = Join(tekiyo1List.Keys, "�^")
    result(3) = Join(tekiyo2List.Keys, "�^")
    result(4) = Join(tekiyo3List.Keys, "�^")

    GetPreviousMonthData = result
End Function

' === �ۑ���ɐF�����ɕύX���A�w�i�F��΂� ===
Public Sub FinalizeSave(ctrls As Collection, frameCtrl As Object)
    Dim ctrl As Object
    For Each ctrl In ctrls
        ctrl.ForeColor = RGB(0, 0, 0)
    Next ctrl
    frameCtrl.BackColor = RGB(230, 255, 230) ' �ۑ��ς݁i�΁j
End Sub
