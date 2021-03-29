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

    '�J�e�S�����Ń^�C�g�����쐬
    Rows(title).RowHeight = 40
    With Cells(title, 2)
        .Value = ECatergory & "�̏ڍ׎x�o"
        .Font.Bold = True
        .Font.Size = 26
    End With
    With Range(Cells(title, 2), Cells(title, 3))
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    
    '�u�J�e�S���v�ŏI�s���擾
    Dim CateRowCnt, title As Long
    CateRowCnt = Worksheets("�x�o�J�e�S��").Range("E10").End(xlDown).Row
    title = CateRowCnt + 4

'�ڍה�ڕ\���쐬
    Dim PSum, AllRow As Long
    Dim CCol, ExCol As Range
    Dim Cate As String
    
    '�x�o�̍ŏI�s���擾
    AllRow = Worksheets("�x�o").Range("C9").End(xlDown).Row
    
    '��ڗ��9�s�ڂ���ŏI�s�܂ŃZ�b�g
    Set CCol = Worksheets("�x�o").Range("C9:C" & AllRow)
    '�x�o�z�̗��9�s�ڂ���ŏI�s�܂ŃZ�b�g
    Set ExCol = Worksheets("�x�o").Range("I9:I" & AllRow)
    
        
    '
    Dim i As Long
    '�u�J�e�S���v�̂͂��߂���ŏI�s�܂ŌJ��Ԃ�
    For i = 10 To CateRowCnt
        '���1������Cate�ɃZ�b�g����
        Cate = Worksheets("�x�o�J�e�S��").Cells(i, 5)
        
        'Cate�Ō���
        PSum = WorksheetFunction.SumIf(CCol, Cate, ExCol)
        
'        MsgBox Cate & "�̍��v�x�o��" & PSum & "�~�ł�"
        
        '���݂̃J�e�S�������O���t�V�[�g�̕\�ɓ����
        With Worksheets("�O���t")
            .Cells(i, 2).Value = Cate
            .Cells(i, 3).Value = PSum
        End With
        '�ǉ������s�̃f�U�C���𐮂���
        Call RowDesign(Range(Cells(i, 2), Cells(i, 3)))
    Next i

'�O���t��L�̕\���g���ăO���t���쐬
    Dim trgtSh As Worksheet
    Set trgtSh = ThisWorkbook.Worksheets("�O���t")
    
    Dim dataRng As Range
    Set dataRng = Range(Cells(10, 2), Cells(CateRowCnt, 3))
    
    '�\��t�������Z�����`
    Dim pasteRng As Range
    Set pasteRng = trgtSh.Range("E4")
     
    '�O���t�쐬
    With trgtSh.Shapes.AddChart.Chart
        '�O���t�̎�ގw��
        .ChartType = xlColumnClustered
        '�Ώۃf�[�^�͈͂��w��
        .SetSourceData dataRng
        '�O���t�^�C�g���ݒ�
        .HasTitle = True
        .ChartTitle.Text = "��ڕʎx�o"
         
        '�O���t�̈ʒu���w��
        .Parent.Top = pasteRng.Top
        .Parent.Left = pasteRng.Left
    End With




End Sub

Private Sub UserForm_Initialize()
'�t�H�[���̏����ݒ�

    '�J�e�S�������󔒂łȂ���Ώ��Ɏ擾�A�I�����ɒǉ�
    If Worksheets("�x�o�J�e�S��").Range("E10") <> "" Then
        '�u�J�e�S���v�ŏI�s���擾
        Dim CateRowCnt As Long
        CateRowCnt = Worksheets("�x�o�J�e�S��").Range("E10").End(xlDown).Row
        
        '�u�J�e�S���v�ꗗ���R���{�{�b�N�X�ɓ����
        Dim i As Long, Cate As String
        '�u�J�e�S���v�̂͂��߂���ŏI�s�܂ŌJ��Ԃ�
        For i = 10 To CateRowCnt
            Cate = Worksheets("�x�o�J�e�S��").Cells(i, 5)
            Me.ECatergory.AddItem Cate
        Next
    End If
    
End Sub

