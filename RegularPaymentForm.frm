VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegularPaymentForm 
   Caption         =   "RegularPaymentForm"
   ClientHeight    =   9555.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14250
   OleObjectBlob   =   "RegularPaymentForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "RegularPaymentForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ButtonHandlers As Collection

Private Sub UserForm_Initialize()
    Set ButtonHandlers = New Collection
    InitializeForm Me
End Sub

Private Sub SpinButton_Month_Change()
    '���ύX���̏���
    MonthChanged Me
End Sub

Private Sub MultiPage_Category_Change()
    ' �y�[�W�؂�ւ����ɃX�N���[���ʒu�����Z�b�g
    With Me.MultiPage_Category
        If .Pages.Count > 0 And .Value >= 0 Then
            With .Pages(.Value)
                .ScrollBars = fmScrollBarsVertical
                .ScrollTop = 0
            End With
        End If
    End With
End Sub

Private Sub UserForm_Click()
    ' �ۑ��E�i�荞�݃{�^���̋��ʏ����i�e�t���[���̃{�^���ɑΉ��j
    Dim ctrl As MSForms.Control
    Set ctrl = Me.ActiveControl

    If TypeName(ctrl) = "CommandButton" Then
        Dim btnName As String: btnName = ctrl.Name

        ' �ۑ��{�^���i�ʁj
        If Left(btnName, 17) = "CommandButton_Save_" Then
            Dim payee As String
            payee = Mid(btnName, 18)

            Call SaveSingleEntry(Me, payee, Me.MultiPage_Denpyo.Pages(Me.MultiPage_Denpyo.Value).caption)


        ' �i�荞�݃{�^���i�ʁj
        ElseIf Left(btnName, 20) = "CommandButton_Filter_" Then
'            Dim payee As String
            payee = Mid(btnName, 21)

            Call FilterDataByPayee(payee)
        End If
    End If
End Sub


' === ���[�U�[�t�H�[�������� ===
Public Sub InitializeForm(frm As Object)
    Dim currentMonth As Integer: currentMonth = month(Date)
    Dim currentYear As Integer: currentYear = year(Date)

    frm.SpinButton_Month.Value = currentMonth
    frm.Label_Month.caption = currentMonth & "��"

    With frm.SpinButton_Year
        .Min = 2020
        .Max = 2030
    
        If currentYear < .Min Then
            .Value = .Min
        ElseIf currentYear > .Max Then
            .Value = .Max
        Else
            .Value = currentYear
        End If
    End With

    frm.Label_Year.caption = currentYear & "�N"

    ' �J�e�S��ComboBox������
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("���@�@��")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("���M�����[���X�g")

    Dim categoryDict As Object
    Set categoryDict = CreateObject("Scripting.Dictionary")

    Dim r As ListRow
    For Each r In tbl.ListRows
        Dim cat As String
        cat = r.Range(1, tbl.ListColumns("�J�e�S��").index).Value
        If Not categoryDict.exists(cat) Then
            categoryDict.Add cat, 1
            frm.ComboBox_Category.AddItem cat
        End If
    Next r

    ' �����I���Ɠ`�[�y�[�W����
    If frm.ComboBox_Category.ListCount > 0 Then
        frm.ComboBox_Category.ListIndex = 0
        Call LoadRegularListByCategory(frm, frm.ComboBox_Category.Value, frm.SpinButton_Year.Value, frm.SpinButton_Month.Value)
    End If
End Sub



' === ���ύX���̏��� ===
Public Sub MonthChanged(frm As Object)
    Dim selectedMonth As Integer: selectedMonth = frm.SpinButton_Month.Value
    Dim selectedYear As Integer: selectedYear = frm.SpinButton_Year.Value

    frm.Label_Month.caption = selectedMonth & "��"
    frm.Label_Year.caption = selectedYear & "�N"

    ' �V�\���ɑΉ�
    Call LoadRegularListByCategory(frm, frm.ComboBox_Category.Value, selectedYear, selectedMonth)
End Sub



' === ���M�����[���X�g�ǂݍ��݂ƃt�H�[������ ===
Public Sub LoadRegularListByCategory(frm As Object, selectedCategory As String, selectedYear As Integer, selectedMonth As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("���@�@��")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("���M�����[���X�g")

    ' MultiPage_Denpyo �̏�����
    With frm.MultiPage_Denpyo
        Do While .Pages.Count > 0
            .Pages.Remove 0
        Loop
    End With

    Dim denpyoDict As Object
    Set denpyoDict = CreateObject("Scripting.Dictionary")

    Dim r As ListRow
    For Each r In tbl.ListRows
        With r.Range
            If .Cells(tbl.ListColumns("�J�e�S��").index).Value = selectedCategory Then
                Dim denpyo As String
                denpyo = .Cells(tbl.ListColumns("�`�[").index).Value
                Dim payee As String
                payee = .Cells(tbl.ListColumns("�����").index).Value

                ' �`�[�y�[�W���Ȃ���Βǉ�
                If Not denpyoDict.exists(denpyo) Then
                    Dim pg As Object
                    Set pg = frm.MultiPage_Denpyo.Pages.Add
                    pg.caption = denpyo
                    pg.ScrollBars = fmScrollBarsVertical
                    pg.ScrollHeight = 0
                    denpyoDict.Add denpyo, pg
                End If

                ' ���j�[�NID�����ipayee + �s�ԍ��j
                Dim uniqueID As String
                uniqueID = payee & "_" & r.index

                ' ���C�A�E�g�v�Z�i3��z�u�j
                Dim index As Integer
                index = 0
                Dim ctrl As Object
                For Each ctrl In denpyoDict(denpyo).Controls
                    If TypeName(ctrl) = "Frame" Then index = index + 1
                Next ctrl

                Dim colCount As Integer: colCount = 3
                Dim frameWidth As Integer: frameWidth = 220
                Dim frameHeight As Integer: frameHeight = 180
                Dim margin As Integer: margin = 10
                Dim xOffset As Integer: xOffset = (index Mod colCount) * (frameWidth + margin)
                Dim yOffset As Integer: yOffset = (index \ colCount) * (frameHeight + margin)

                ' �t���[���ǉ�
                Dim frmPayee As MSForms.Frame
                Set frmPayee = denpyoDict(denpyo).Controls.Add("Forms.Frame.1", "Frame_Payee_" & uniqueID)
                With frmPayee
                    .caption = payee
                    .Left = xOffset
                    .Top = yOffset
                    .Width = frameWidth
                    .Height = frameHeight
                    .BackColor = RGB(255, 230, 230)
                End With

                ' ���͗��ǉ��idenpyo ��n���j
                Call AddInputFieldsToFrame(frmPayee, payee, selectedMonth, selectedYear, selectedCategory, denpyo)
            End If
        End With
    Next r

    ' �y�[�W���Ƃ̍�������
    Dim denpyoKey As Variant
    For Each denpyoKey In denpyoDict.Keys
'        Dim pg As Object
        Set pg = denpyoDict(denpyoKey)

        Dim totalFrames As Integer
        totalFrames = 0
        For Each ctrl In pg.Controls
            If TypeName(ctrl) = "Frame" Then totalFrames = totalFrames + 1
        Next ctrl

        Dim rowsNeeded As Integer: rowsNeeded = (totalFrames + 2) \ 3
        Dim pageHeight As Integer: pageHeight = rowsNeeded * (frameHeight + margin) + 60

        pg.ScrollHeight = pageHeight
        pg.ScrollTop = 0
    Next denpyoKey
End Sub


Private Sub ComboBox_Category_Change()
    Call LoadRegularListByCategory(Me, Me.ComboBox_Category.Value, Me.SpinButton_Year.Value, Me.SpinButton_Month.Value)
End Sub


Public Sub AddInputFieldsToFrame(frmPayee As MSForms.Frame, payee As String, selectedMonth As Integer, selectedYear As Integer, category As String, denpyo As String)
    Dim topOffset As Integer: topOffset = 10

    ' === ���j�[�NID���� ===
    Dim uniqueID As String
    uniqueID = payee & "_" & denpyo

    ' === ���z ===
    frmPayee.Controls.Add "Forms.Label.1", "Label_Amount_" & uniqueID
    frmPayee.Controls.Add "Forms.TextBox.1", "TextBox_Amount_" & uniqueID
    With frmPayee.Controls("Label_Amount_" & uniqueID)
        .caption = "���z"
        .Left = 10: .Top = topOffset: .Width = 40
    End With
    With frmPayee.Controls("TextBox_Amount_" & uniqueID)
        .Left = 60: .Top = topOffset - 2: .Width = 80
    End With

    ' === �E�v1?3 ===
    Dim j As Integer
    For j = 1 To 3
        frmPayee.Controls.Add "Forms.Label.1", "Label_Tekiyo" & j & "_" & uniqueID
        frmPayee.Controls.Add "Forms.TextBox.1", "TextBox_Tekiyo" & j & "_" & uniqueID
        With frmPayee.Controls("Label_Tekiyo" & j & "_" & uniqueID)
            .caption = "�E�v" & j
            .Left = 10: .Top = topOffset + 20 * j: .Width = 40
        End With
        With frmPayee.Controls("TextBox_Tekiyo" & j & "_" & uniqueID)
            .Left = 60: .Top = topOffset + 20 * j - 2: .Width = 120
        End With
    Next j

    ' === �x���� ===
    frmPayee.Controls.Add "Forms.Label.1", "Label_Date_" & uniqueID
    frmPayee.Controls.Add "Forms.TextBox.1", "TextBox_Date_" & uniqueID
    With frmPayee.Controls("Label_Date_" & uniqueID)
        .caption = "�x����"
        .Left = 10: .Top = topOffset + 80: .Width = 50
    End With
    With frmPayee.Controls("TextBox_Date_" & uniqueID)
        .Left = 70: .Top = topOffset + 78: .Width = 100
        .Value = Format(DateSerial(year(Date), month(Date) + 1, 0), "yyyy/mm/dd")
    End With

    ' === �Ȗ� ===
    frmPayee.Controls.Add "Forms.Label.1", "Label_Kamoku_" & uniqueID
    frmPayee.Controls.Add "Forms.ComboBox.1", "ComboBox_Kamoku_" & uniqueID
    With frmPayee.Controls("Label_Kamoku_" & uniqueID)
        .caption = "�����i�Ȗځj"
        .Left = 10: .Top = topOffset + 100: .Width = 80
    End With
    With frmPayee.Controls("ComboBox_Kamoku_" & uniqueID)
        .Left = 100: .Top = topOffset + 98: .Width = 100
    End With

    ' === �⏕ ===
    frmPayee.Controls.Add "Forms.Label.1", "Label_Hojo_" & uniqueID
    frmPayee.Controls.Add "Forms.ComboBox.1", "ComboBox_Hojo_" & uniqueID
    With frmPayee.Controls("Label_Hojo_" & uniqueID)
        .caption = "�⏕"
        .Left = 10: .Top = topOffset + 120: .Width = 80
    End With
    With frmPayee.Controls("ComboBox_Hojo_" & uniqueID)
        .Left = 100: .Top = topOffset + 118: .Width = 100
    End With

    ' === �����l�̔��f ===
    Call ApplyInitialValues(frmPayee, uniqueID, selectedMonth, selectedYear, category, denpyo)

    ' === �ۑ��{�^�� ===
    frmPayee.Controls.Add "Forms.CommandButton.1", "CommandButton_Save_" & uniqueID
    With frmPayee.Controls("CommandButton_Save_" & uniqueID)
        .caption = "�ۑ�"
        .Left = 10
        .Top = frmPayee.Height - 30
        .Width = 60
    End With
    Call RegisterButtonHandler(frmPayee.Controls("CommandButton_Save_" & uniqueID), uniqueID, denpyo, "Save")

    ' === �i�荞�݃{�^�� ===
    frmPayee.Controls.Add "Forms.CommandButton.1", "CommandButton_Filter_" & uniqueID
    With frmPayee.Controls("CommandButton_Filter_" & uniqueID)
        .caption = "�i�荞��"
        .Left = 80
        .Top = frmPayee.Height - 30
        .Width = 60
    End With
    Call RegisterButtonHandler(frmPayee.Controls("CommandButton_Filter_" & uniqueID), uniqueID, denpyo, "Filter")

    ' === �h���b�v�_�E���ݒ� ===
    Call SetKamokuDropdown(frmPayee.Controls("ComboBox_Kamoku_" & uniqueID))
    Call SetHojoDropdown(frmPayee.Controls("ComboBox_Hojo_" & uniqueID))
End Sub

Public Sub RegisterButtonHandler(btn As MSForms.CommandButton, frameID As String, denpyo As String, actionType As String)
    Dim handler As New clsButtonHandler
    Set handler.btn = btn
    handler.frameID = frameID
    handler.payee = Split(frameID, "_")(0)
    handler.denpyo = denpyo
    handler.actionType = actionType
    Set handler.ParentForm = Me
    ButtonHandlers.Add handler
End Sub



