VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub addCatergory()
'�J�e�S���ǉ�

    CatergoryForm.Show
    
End Sub

Sub deleteCatergory()
        
'�L�����Z�����ꂽ��G���[����
    On Error GoTo myErr

    '�W�����W���[�����Ăяo��
    Dim delRng As Range
'���[�N�V�[�g��n��
'    Dim sheet As Worksheet
'    Set sheet = Sheets("�x�o�J�e�S��")
    Set delRng = selectData()
    
    delRng.Select
    
    '�I�����ꂽ�Z���̒l�ƍs�ԍ����擾
    Dim RngVal As String, RngRow As Long, RngColumn As Long
    RngVal = delRng.Value
    RngRow = delRng.Row
    RngColumn = delRng.Column
    
'    MsgBox "�l" & RngVal & "�s" & RngRow & "��" & RngColumn

    delRng.Select
        
    '�J�������̗񂪑I�����ꂽ�ꍇ
    If RngRow = 9 Then
        MsgBox "�����ȃZ���ł�"
        GoTo Label4
    End If
        
    '�x�o�J�e�S���̒��ŁA���1���I�����ꂽ���A2���I�����ꂽ��
    Select Case RngColumn
        '���1�̍s���I�����ꂽ�ꍇ
        Case 5, 7
            Dim result1 As Long
            result1 = MsgBox("���1����" & RngVal & "���폜���܂����H" & vbCrLf & "��" & RngVal & "�ɑ�������2���S�č폜����܂�", vbYesNo)
            If result1 = vbYes Then
                'E����폜
                Dim ColumnE As Range, LastE As Long
                Set ColumnE = Range("E10")
                LastE = Range("E10").End(xlDown).Row

                Dim i As Long
                For i = 10 To LastE
                    ColumnE.Select
                    If ColumnE.Value = RngVal Then
                        ColumnE.Delete xlShiftUp
                        Exit For
                    End If
                    Set ColumnE = ColumnE.Offset(1)
                Next
                
                'GH����폜
                '�Ώۂ̍ŏI�s�擾
                Dim LastGH As Long
                LastGH = Worksheets("�x�o�J�e�S��").Range("G10").End(xlDown).Row
                
                Dim j As Long
                '10�s�ڂ���DE�̍ŏI�s�܂ŌJ��Ԃ�
                For j = 10 To LastGH
                Cells(j, 7).Select
Label1:
                    If Worksheets("�x�o�J�e�S��").Cells(j, 7).Value = RngVal Then
                        Cells(j, 7).Delete xlShiftUp
                        Cells(j, 8).Delete xlShiftUp
                        '1�s�폜�������ɂ����̂ŁA������x�����s������
                        GoTo Label1
                    End If
                Next j
                GoTo Label2
            Else
                GoTo Label3
            End If
            
        '���2�̍s���I�����ꂽ�ꍇ
        Case 8
            Dim Cate As String
            Cate = Cells(RngRow, (RngColumn - 1)).Value
            
            Dim result2 As Long
            result2 = MsgBox(Cate & "�ɑ�����" & RngVal & "���폜���܂����H", vbYesNo)
            
            If result2 = vbYes Then
                Range(Cells(RngRow, (RngColumn - 1)), Cells(RngRow, RngColumn)).Delete xlShiftUp
                GoTo Label2
            Else
                GoTo Label3
            End If
        '����ȊO�̃P�[�X
        Case Else
            MsgBox "�����ȃZ���ł�"
            GoTo Label4
        End Select
'    End If

'�傫���J�e�S������S�č폜
Label2:
    MsgBox "�폜���܂���"
    GoTo Label4

'�L�����Z���̏ꍇ
Label3:
    MsgBox "�L�����Z�����܂���"
Label4:

myErr: Exit Sub

End Sub


Sub CateAutoFilter()

    '�������ʂ�\������B1������N���A
    Range("B13:C" & Rows.Count).Clear
    
    With Sheets("�x�o�J�e�S��")
        'G10����̂P��ڂ�C9�ōi�荞��
        .Range("H9").AutoFilter Field:=1, Criteria1:="*" & Range("C9").Value & "*"
        '�R�s�[
        .Range("G9").CurrentRegion.Copy
        .Range("B13").PasteSpecial Paste:=xlPasteAll
        '�I�[�g�t�B���^������
        .AutoFilterMode = False
    End With
    
    Application.CutCopyMode = False
    Range("A1").Select
    
End Sub



