VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub addExpense()
'�x�o�ǉ�

    ExpenseForm.Show
    
End Sub

Sub deleteExpense()
        
    Dim delRng As Range
    Set delRng = selectData()
    
    delRng.Select
    
    '�I�����ꂽ�Z���̒l�ƍs�ԍ����擾
    Dim RngVal As String, RngRow As Long, RngColumn As Long
    RngVal = delRng.Value
    RngRow = delRng.Row
    RngColumn = delRng.Column
   
'    '���̓`�F�b�N
'    MsgBox "�l" & RngVal & "�s" & RngRow & "��" & RngColumn

    '�J�������̗񂪑I�����ꂽ�ꍇ
    If RngRow = 8 Or RngRow = 2 Then
        MsgBox "�����ȃZ���ł�"
        GoTo Label4
    End If
    
    Dim thisRow As Range
    Set thisRow = Worksheets("�x�o").UsedRange.Rows(RngRow - 1)
    thisRow.Select
'    Worksheets("�x�o").UsedRange.Rows(RngRow - 1).Select
    
    Dim result As Long
    result = MsgBox("�I�����ꂽ�Z���̍s���폜���܂����H", vbYesNo)
    If result = vbYes Then
        thisRow.Delete xlShiftUp
        GoTo Label2
    Else
        GoTo Label3
    End If

'�I�����ꂽ�s��S�č폜
Label2:
    MsgBox "�폜���܂���"
    GoTo Label4
'�L�����Z���̏ꍇ
Label3:
    MsgBox "�L�����Z������܂���"

Label4:


            

    
End Sub