VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CreateDetailGraph()
'�ڍ׃O���t�̍쐬

    CreateGForm.Show
    
End Sub

Sub graph()

    '�O���t�V�[�g��̃O���t��S�č폜����
    Worksheets("�O���t").ChartObjects.Delete

'�J�e�S�����Ƃ̍��v�x�o�\���쐬
    Dim PSum, AllRow As Long
    Dim CCol, ExCol As Range
    Dim Cate As String
    
    '�x�o�̍ŏI�s���擾
    AllRow = Worksheets("�x�o").Range("C9").End(xlDown).Row
    
    '��ڗ��9�s�ڂ���ŏI�s�܂ŃZ�b�g
    Set CCol = Worksheets("�x�o").Range("C9:C" & AllRow)
    '�x�o�z�̗��9�s�ڂ���ŏI�s�܂ŃZ�b�g
    Set ExCol = Worksheets("�x�o").Range("I9:I" & AllRow)
    
    '�u�J�e�S���v�ŏI�s���擾
    Dim CateRowCnt As Long
    CateRowCnt = Worksheets("�x�o�J�e�S��").Range("E10").End(xlDown).Row
        
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
