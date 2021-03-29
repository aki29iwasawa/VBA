VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExpenseForm 
   Caption         =   "�x�o�f�[�^��ǉ�"
   ClientHeight    =   9359.001
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   6412
   OleObjectBlob   =   "ExpenseForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "ExpenseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
    
    Unload ExpenseForm
    
End Sub

Private Sub addExpense_Click()
   
    '�f�[�^���͍s���擾
    Dim NewRow As Long
    NewRow = Worksheets("�x�o").Cells(Rows.Count, 2).End(xlUp).Row + 1
    
    '�f�[�^���͍s�̃f�U�C����ύX
    Call RowDesign(Range(Cells(NewRow, 2), Cells(NewRow, 12)))
    
    Cells(NewRow, 2).Value = BDate
    Cells(NewRow, 3).Value = ECatergory
    Cells(NewRow, 4).Value = ESubcatergory
    Cells(NewRow, 5).Value = Item
    
    '�u���ʁv�𐔒l�ϊ�
    Dim qua As Long
    qua = Val(Quantity)
    Cells(NewRow, 6).Value = qua
    
    '�u�P���v�𐔒l�ϊ�
    Dim pri As Long
    pri = Val(Price)
    Cells(NewRow, 7).Value = pri
    
    '�u�Ŕ����v�v
    Dim priSum As Long
    priSum = qua * pri
    Cells(NewRow, 8).Value = priSum
    
    '�u�ō����v�v
    Dim addTax As Long
    addTax = qua * pri * 1.1
    Cells(NewRow, 9).Value = addTax
    
    Cells(NewRow, 10).Value = PMethod
    Cells(NewRow, 11).Value = PDetail
    Cells(NewRow, 12).Value = Memo
   
    Unload ExpenseForm
    
End Sub

Private Sub DateSpin_SpinUp()
'���t
    Dim nDate As Date
    nDate = ExpenseForm.BDate.Value
    ExpenseForm.BDate.Value = DateAdd("d", 1, nDate)
    
End Sub

Private Sub DateSpin_SpinDown()
'���t
    Dim nDate As Date
    nDate = ExpenseForm.BDate.Value
    ExpenseForm.BDate.Value = DateAdd("d", -1, nDate)
    
End Sub

Private Sub Memo_Change()

End Sub

Private Sub UserForm_Initialize()

    '�J�e�S�������󔒂łȂ���Ώ��Ɏ擾�A�I�����ɒǉ�
    If Worksheets("�x�o�J�e�S��").Range("E10") <> "" Then
        '�u�J�e�S���v�ŏI�s���擾
        Dim CateRowCnt As Long
        CateRowCnt = Worksheets("�x�o�J�e�S��").Range("E10").End(xlDown).Row
        
        '�u�J�e�S���v�ꗗ���ڂP�ɓ����
        Dim i As Long, Cate As String
        '�u�J�e�S���v�̂͂��߂���ŏI�s�܂ŌJ��Ԃ�
        For i = 10 To CateRowCnt
            Cate = Worksheets("�x�o�J�e�S��").Cells(i, 5)
            Me.ECatergory.AddItem Cate
        Next
    End If
        
    Me.BDate.Value = Date

    '���ϕ��@�����󔒂łȂ���Ώ��Ɏ擾�A�I�����ɒǉ�
    If Worksheets("���ϕ��@").Range("B10") <> "" Then
        '�u���ϕ��@�v�ŏI�s���擾
        Dim MethodRowCnt As Long
        MethodRowCnt = Worksheets("���ϕ��@").Range("B10").End(xlDown).Row
        
        '�u���ϕ��@�v�ꗗ���ڂP�ɓ����
        Dim j As Long, Method As String
        '�u���ϕ��@�v�̂͂��߂���ŏI�s�܂ŌJ��Ԃ�
        For i = 10 To MethodRowCnt
            Method = Worksheets("���ϕ��@").Cells(i, 2)
            Me.PMethod.AddItem Method
        Next
    End If

End Sub
Private Sub ECatergory_Change()

    '��ڂQ���N���A
    ESubcatergory.Clear
    
    '��ڂP�őI�����ꂽ���̂��擾
    Dim Cate As String
    Cate = ECatergory.Text
    
    With Worksheets("�x�o�J�e�S��")
        '�������ʂ�\������B1������N���A
        .Range("B14:C" & Rows.Count).Clear

        '���1�őI�����ꂽ���̂ŃI�[�g�t�B���^�[
        .Range("G9").AutoFilter Field:=1, Criteria1:="*" & Cate & "*"
        '�R�s��
        .Range("G9").CurrentRegion.Copy
        .Range("B13").PasteSpecial Paste:=xlPasteAll
        '�I�[�g�t�B���^�[����
        .Range("G9").AutoFilter

    End With
    
    '�J�e�S�������󔒂łȂ���Ώ��Ɏ擾�A�I�����ɒǉ�
    If Worksheets("�x�o�J�e�S��").Range("C14") <> "" Then
        '�u�J�e�S���v�ŏI�s���擾
        Dim CateRowCnt As Long
        CateRowCnt = Worksheets("�x�o�J�e�S��").Range("C13").End(xlDown).Row
        
'        MsgBox CateRowCnt
      
        '�u�J�e�S���v�ꗗ����2�ɓ����
        Dim i As Long, subCate As String
        '�u�J�e�S���v�̂͂��߂���ŏI�s�܂ŌJ��Ԃ�
        For i = 14 To CateRowCnt
            subCate = Worksheets("�x�o�J�e�S��").Cells(i, 3).Value
            Me.ESubcatergory.AddItem subCate
        Next
    End If
    
    '�������ʂ�\������B1������N���A
    Worksheets("�x�o�J�e�S��").Range("B14:C" & Rows.Count).Clear

    Application.CutCopyMode = False
    
End Sub

Private Sub PMethod_Change()

    '�ڍׂ��N���A
    PDetail.Clear
    
    '���ϕ��@�őI�����ꂽ���̂��擾
    Dim Meth As String
    Meth = PMethod.Text
    
    With Worksheets("���ϕ��@")
'        '�������ʂ�\������G1������N���A
'        .Range("B14:C" & Rows.Count).Clear

        '���ϕ��@�őI�����ꂽ���̂ŃI�[�g�t�B���^�[
        .Range("D9").AutoFilter Field:=1, Criteria1:="*" & Meth & "*"
        '�R�s��
        .Range("D9").CurrentRegion.Copy
        .Range("G9").PasteSpecial Paste:=xlPasteAll
        '�I�[�g�t�B���^�[����
        .Range("D9").AutoFilter

    End With
    
    '�ڍח����󔒂łȂ���Ώ��Ɏ擾�A�I�����ɒǉ�
    If Worksheets("���ϕ��@").Range("H10") <> "" Then
        '���ϕ��@�ŏI�s���擾
        Dim MethRowCnt As Long
        MethRowCnt = Worksheets("���ϕ��@").Range("H9").End(xlDown).Row
        
        '�u���ϕ��@�v�ꗗ���ڍׂɓ����
        Dim i As Long, Det As String
        '�u�J�e�S���v�̂͂��߂���ŏI�s�܂ŌJ��Ԃ�
        For i = 10 To MethRowCnt
            Det = Worksheets("���ϕ��@").Cells(i, 8).Value
            Me.PDetail.AddItem Det
        Next
    End If
    
    '�I�[�g�t�B���^�[�̌��ʂ��y�[�X�g�����Z�����N���A
    Worksheets("���ϕ��@").Range("G9:H" & Rows.Count).Clear
    
    Application.CutCopyMode = False
        
End Sub



