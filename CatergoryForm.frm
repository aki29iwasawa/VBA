VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CatergoryForm 
   Caption         =   "��ڒǉ�"
   ClientHeight    =   4389
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   5740
   OleObjectBlob   =   "CatergoryForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "CatergoryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addCatergory_Click()
'�ǉ��{�^���N���b�N���̏���

'�傫���J�e�S���ɓ��͂��ꂽ���̂��u�J�e�S���v�ɂ��邩�m�F�A
'����������ǉ��A��������ǉ��͂��Ȃ�
'�T�u�J�e�S���ɓ��͂��ꂽ���̂���������A���[�g���o���H

    '��ڂP�������͂̏ꍇ
    If ECatergory = "" Then
        MsgBox "���1�������͂ł�"
        
    Else
        '�u��ڂP�v�f�[�^���͍s���擾
        Dim NewRow As Long
        NewRow = Worksheets("�x�o�J�e�S��").Cells(Rows.Count, 5).End(xlUp).Row + 1
        '�u��ڂQ�v�f�[�^���͍s���擾
        Dim NewRowSub As Long
        NewRowSub = Worksheets("�x�o�J�e�S��").Cells(Rows.Count, 8).End(xlUp).Row + 1
            
        '�J�e�S���o�^����Ă��Ȃ���ڂP�����͂��ꂽ�ꍇ�A�o�^����
        Dim Cate1 As range
        Set Cate1 = Worksheets("�x�o�J�e�S��").range(Cells(10, 5), Cells(NewRow, 5)).Find(What:=ECatergory, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
        
        If Cate1 Is Nothing Then
        '�V�K��ڂP�ƂQ��o�^
            With Cells(NewRow, 5)
                .Value = ECatergory
                .Interior.Color = RGB(221, 235, 247)
                .Borders(xlEdgeTop).LineStyle = xlDash
                .Borders(xlEdgeTop).Color = RGB(47, 117, 181)
            End With
            Cells(NewRowSub, 7).Value = ECatergory
            Cells(NewRowSub, 8).Value = ESubcatergory
            
            '�f�[�^���͍s�̃f�U�C����ύX
            Call RowDesign(range(Cells(NewRowSub, 7), Cells(NewRowSub, 8)))
          
        ElseIf Not Cate1 Is Nothing And ESubcatergory = "" Then
        '���͂��ꂽ��ڂP�����ɑ��݂��A���A��ڂQ�̓��͂��Ȃ��ꍇ
            MsgBox ECatergory & "�͊��ɑ��݂����ڂł��B"
        
        ElseIf Not Cate1 Is Nothing Then
        '��ڂP�����ɑ��݂���ꍇ
        '��ڂQ�̑��݃`�F�b�N
            Dim Cate2 As range
            Set Cate2 = Worksheets("�x�o�J�e�S��").range(Cells(10, 8), Cells(NewRowSub, 8)).Find(What:=ESubcatergory, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
            
            If Cate2 Is Nothing Then
                '�����f�[�^�����݂��Ȃ��ꍇ�͓���
                Cells(NewRowSub, 7).Value = ECatergory
                Cells(NewRowSub, 8).Value = ESubcatergory
                '�f�[�^���͍s�̃f�U�C����ύX
                Call RowDesign(range(Cells(NewRowSub, 7), Cells(NewRowSub, 8)))
        
            ElseIf Not Cate2 Is Nothing Then
                '�����̔�ڂQ�̌���findnext���g���Ă݂�
                
                MsgBox "���L�̔��1�ɂ�" & ESubcatergory & "�����݂��܂�"
                
                
            End If
        End If
    End If
    
    '�t�H�[�������
    Unload CatergoryForm
    
End Sub

Private Sub Cancel_Click()
'�t�H�[�������
    
    Unload CatergoryForm
    
End Sub

Private Sub UserForm_Initialize()
'�t�H�[���̏����ݒ�

    '�J�e�S�������󔒂łȂ���Ώ��Ɏ擾�A�I�����ɒǉ�
    If Worksheets("�x�o�J�e�S��").range("E10") <> "" Then
        '�u�J�e�S���v�ŏI�s���擾
        Dim CateRowCnt As Long
        CateRowCnt = Worksheets("�x�o�J�e�S��").range("E10").End(xlDown).Row
        
        '�u�J�e�S���v�ꗗ���ڂP�ɓ����
        Dim i As Long, Cate As String
        '�u�J�e�S���v�̂͂��߂���ŏI�s�܂ŌJ��Ԃ�
        For i = 10 To CateRowCnt
            Cate = Worksheets("�x�o�J�e�S��").Cells(i, 5)
            Me.ECatergory.AddItem Cate
        Next
    End If
    
End Sub


