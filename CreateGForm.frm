VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateGForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3248
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   5292
   OleObjectBlob   =   "CreateGForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "CreateGForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
'�t�H�[�������
    
    Unload CreateGForm
    
End Sub

Private Sub CommandButton1_Click()
    
    '���[�N�V�[�g��ϐ��ɒu��
    Dim gra, cat As Worksheet
    Set gra = ThisWorkbook.Worksheets("�O���t")
    Set cat = ThisWorkbook.Worksheets("�x�o�J�e�S��")

    '�u�x�o�J�e�S���v�V�[�g�̔�ڂP�ŏI�s���擾
    Dim CateRowCnt, title As Long
    CateRowCnt = cat.range("E10").End(xlDown).Row
    title = CateRowCnt + 4

    '�J�e�S�����ŕ\�^�C�g�����쐬
    Rows(title).RowHeight = 40
    With Cells(title, 2)
        .Value = ECatergory & "�̏ڍ׎x�o"
        .Font.Bold = True
        .Font.Size = 26
    End With
    With range(Cells(title, 2), Cells(title, 3))
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    
    '�\���쐬����Z��,�J������ݒ�i�^�C�g���������2�s���j
    title = title + 2
    Cells(title, 2).Value = ECatergory & "�̕i��"
    Cells(title, 3).Value = "�x�o"
    Call ColDesign(range(Cells(title, 2), Cells(title, 3)))
    
    '�I�����ꂽ���1��ϐ��ɒu��
    Dim Cate As String
    Cate = ECatergory.Value
    
'�u�x�o�J�e�S���v�V�[�g�̔��2���I�[�g�t�B���^�[�ōi��
    With cat
        '�������ʂ�\������B1������N���A
        .range("B13:C" & Rows.Count).Clear
        'G10����̂P��ڂ�Cate�ōi�荞��
        .range("G9").AutoFilter Field:=1, Criteria1:="*" & Cate & "*"
        '�R�s�[
        .range("G9").CurrentRegion.Copy
        .range("B13").PasteSpecial Paste:=xlPasteAll
        '�I�[�g�t�B���^������
        .AutoFilterMode = False
    End With
        
    Application.CutCopyMode = False
        
'�J�e�S�����Ƃ̍��v�x�o�\���쐬
    Dim PSum, AllRow As Long
    Dim CCol, ExCol As range
    Dim Cate2 As String
    
    '�x�o�̍ŏI�s���擾
    AllRow = Worksheets("�x�o").range("C9").End(xlDown).Row
    
    '���2�̗��9�s�ڂ���ŏI�s�܂ŃZ�b�g
    Set CCol = Worksheets("�x�o").range("D9:D" & AllRow)
    '�x�o�z�̗��9�s�ڂ���ŏI�s�܂ŃZ�b�g
    Set ExCol = Worksheets("�x�o").range("I9:I" & AllRow)
    
    '�u�J�e�S���v�ŏI�s���擾
    Dim Cate2RowCnt As Long
    Cate2RowCnt = cat.range("C14").End(xlDown).Row
        
    '
    Dim i, j As Long
    j = title + 1
    '�u�J�e�S���v�̂͂��߂���ŏI�s�܂ŌJ��Ԃ�
    For i = 14 To Cate2RowCnt
        '���2������Cate2�ɃZ�b�g����
        Cate2 = cat.Cells(i, 3)
        
        'Cate�Ō���
        PSum = WorksheetFunction.SumIf(CCol, Cate2, ExCol)
        
        MsgBox Cate2 & "�̍��v�x�o��" & PSum & "�~�ł�"
        
        '���݂̃J�e�S�������O���t�V�[�g�̕\�ɓ����
        With gra
            .Cells(j, 2).Value = Cate2
            .Cells(j, 3).Value = PSum
        End With
        '�ǉ������s�̃f�U�C���𐮂���
        Call RowDesign(range(Cells(j, 2), Cells(j, 3)))
        j = j + 1
    Next i

'��L�̕\���g���ăO���t���쐬
    '�\�̍ŏI�s���擾
    Dim tableL As Long
    tableL = gra.Cells(title, 2).End(xlDown).Row
    
    '�\��t�������Z�����`
    Dim pasteRng As range
    Set pasteRng = gra.Cells(title, 5)
    
    '�\�ɂ���f�[�^�͈͂��Z�b�g
    Dim dataRng As range
    Set dataRng = range(Cells(title + 1, 2), Cells(tableL, 3))
         
    '�O���t�쐬
    With gra.Shapes.AddChart.Chart
        '�O���t�̎�ގw��
        .ChartType = xlColumnClustered
        '�Ώۃf�[�^�͈͂��w��
        .SetSourceData dataRng
        '�O���t�^�C�g���ݒ�
        .HasTitle = True
        .ChartTitle.Text = ECatergory & "�̏ڍ׎x�o"
         
        '�O���t�̈ʒu���w��
        .Parent.Top = pasteRng.Top
        .Parent.Left = pasteRng.Left
    End With

    '�t�H�[�������
    Unload CreateGForm

End Sub

Private Sub UserForm_Initialize()
'�t�H�[���̏����ݒ�

    '�J�e�S�������󔒂łȂ���Ώ��Ɏ擾�A�I�����ɒǉ�
    If Worksheets("�x�o�J�e�S��").range("E10") <> "" Then
        '�u�J�e�S���v�ŏI�s���擾
        Dim CateRowCnt As Long
        CateRowCnt = Worksheets("�x�o�J�e�S��").range("E10").End(xlDown).Row
        
        '�u�J�e�S���v�ꗗ���R���{�{�b�N�X�ɓ����
        Dim i As Long, Cate As String
        '�u�J�e�S���v�̂͂��߂���ŏI�s�܂ŌJ��Ԃ�
        For i = 10 To CateRowCnt
            Cate = Worksheets("�x�o�J�e�S��").Cells(i, 5)
            Me.ECatergory.AddItem Cate
        Next
    End If
    
End Sub

