VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PaymentMethodForm 
   Caption         =   "���ϕ��@"
   ClientHeight    =   4291
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   5768
   OleObjectBlob   =   "PaymentMethodForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "PaymentMethodForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addPMethod_Click()
'�ǉ��{�^���N���b�N���̏���


    '���ϕ��@�������͂̏ꍇ
    If PMethod = "" Then
        MsgBox "���ϕ��@�������͂ł�"
    Else
        '�u���ϕ��@�v�f�[�^���͍s���擾
        Dim NewRow As Long, NewDetRow As Long
        NewRow = Worksheets("���ϕ��@").Cells(Rows.Count, 2).End(xlUp).Row + 1
        NewDetRow = Worksheets("���ϕ��@").Cells(Rows.Count, 5).End(xlUp).Row + 1
        
        '�J�e�S���o�^����Ă��Ȃ���ڂP�����͂��ꂽ�ꍇ�A�o�^����
        Dim Method As Range
        Set Method = Worksheets("���ϕ��@").Range(Cells(10, 2), Cells(NewRow, 2)).Find(What:=PMethod, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
        If Method Is Nothing Then
        '�V�K���ϕ��@�Əڍׂ�o�^
            With Cells(NewRow, 2)
                .Value = PMethod
                .Interior.Color = RGB(221, 235, 247)
                .Borders(xlEdgeTop).LineStyle = xlDash
                .Borders(xlEdgeTop).Color = RGB(47, 117, 181)
            End With
            
            If PDetail <> "" Then
                Cells(NewDetRow, 4).Value = PMethod
                Cells(NewDetRow, 5).Value = PDetail
                '�f�[�^���͍s�̃f�U�C����ύX
                Call RowDesign(Range(Cells(NewDetRow, 4), Cells(NewDetRow, 5)))
            End If
          
        ElseIf Not Method Is Nothing And PDetail = "" Then
        '���͂��ꂽ���ϕ��@�����ɑ��݂��A���A�ڍׂ̓��͂��Ȃ��ꍇ
            MsgBox Method & "�͓o�^�ς݂̌��ϕ��@�ł��B"
        
        ElseIf Not Method Is Nothing Then
        '���ϕ��@����o�^�ς݂̏ꍇ
        '�ڍׂ̑��݃`�F�b�N
            Dim dtl As Range
            Set dtl = Worksheets("���ϕ��@").Range(Cells(10, 5), Cells(NewRow, 5)).Find(What:=PDetail, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
            
            If dtl Is Nothing Then
                '�����f�[�^�����݂��Ȃ��ꍇ�͓���
                Cells(NewDetRow, 4).Value = PMethod
                Cells(NewDetRow, 5).Value = PDetail
                '�f�[�^���͍s�̃f�U�C����ύX
                Call RowDesign(Range(Cells(NewDetRow, 4), Cells(NewDetRow, 5)))
        
            ElseIf Not dtl Is Nothing Then
                '���
                
                MsgBox "�o�^�ς݂ł�"
                
                
            End If
        End If
    End If
    
    '�t�H�[�������
    Unload PaymentMethodForm

End Sub


Private Sub Cancel_Click()
    
    Unload PaymentMethodForm
    
End Sub

Private Sub UserForm_Initialize()

    '�J�e�S�������󔒂łȂ���Ώ��Ɏ擾�A�I�����ɒǉ�
    If Worksheets("���ϕ��@").Range("B10") <> "" Then
        '�u�J�e�S���v�ŏI�s���擾
        Dim RowCnt As Long
        RowCnt = Worksheets("���ϕ��@").Range("B10").End(xlDown).Row
        
        '�u���ϕ��@�v�ꗗ���R���{�{�b�N�X�ɓ����
        Dim i As Long, Cate As String
        '�u���ϕ��@�v�̂͂��߂���ŏI�s�܂ŌJ��Ԃ�
        For i = 10 To RowCnt
            mth = Worksheets("���ϕ��@").Cells(i, 2)
            Me.PMethod.AddItem mth
        Next
    End If
    
End Sub
