VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub addPMethod()
'���ϕ��@�ǉ�

    PaymentMethodForm.Show
    
End Sub

Sub deletePMethod()

'�L�����Z�����ꂽ��G���[����
    On Error GoTo myErr
        
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
    If RngRow = 9 Or RngRow = 2 Then
        MsgBox "�����ȃZ���ł�"
        GoTo Label4
    End If
        
    '���ϕ��@�̒��ŁA���@���I�����ꂽ���A�ڍׂ��I�����ꂽ��
    Select Case RngColumn
        '���@�̍s���I�����ꂽ�ꍇ
        Case 2, 4
            Dim result1 As Long
            result1 = MsgBox("���ϕ��@����" & RngVal & "���폜���܂����H" & vbCrLf & "��" & RngVal & "�ɑ�����ڍׂ��S�č폜����܂�", vbYesNo)
            If result1 = vbYes Then
                'B����폜
                Dim ColumnB As Range, LastB As Long
                Set ColumnB = Range("B10")
                LastB = Range("B10").End(xlDown).Row

                Dim i As Long
                For i = 10 To LastB
                    If ColumnB.Value = RngVal Then
                        ColumnB.Delete xlShiftUp
                        Exit For
                    End If
                    Set ColumnB = ColumnB.Offset(1)
                Next
                
'���s����̎c��
'                'DH����폜
'                Dim ColumnD As Range, LastD As Long
'                Set ColumnD = Range("D10")
'                LastD = Range("D10").End(xlDown).Row
'
'                Dim j As Long
'                For j = 10 To LastD
'                    If ColumnD.Value = RngVal Then
'                        ColumnD.Delete xlShiftUp
'                        ColumnD.Offset(0, 1).Select
'                        ColumnD.Offset(0, 1).Delete xlShiftUp
'                        ColumnD.Offset(1).Delete xlShiftUp
'                    End If
'                    Set ColumnD = ColumnD.Offset(1)
'                Next
'                GoTo Label2
'            Else
'                GoTo Label3
'            End If
'�Z������̌J��Ԃ��ɂȂ��Ă��܂�
'                Dim hani As Range
'                Dim thisRow As Range
'                Set hani = Range("D10:E10")
'                For Each thisRow In hani
'                    thisRow.Select
'                    MsgBox thisRow.Value
'                Next thisRow


                'DH����폜
                '�Ώۂ̍ŏI�s�擾
                Dim LastDE As Long
                LastDE = Worksheets("���ϕ��@").Range("D10").End(xlDown).Row
                
                Dim j As Long
                '10�s�ڂ���DE�̍ŏI�s�܂ŌJ��Ԃ�
                For j = 10 To LastDE
                Cells(j, 4).Select
Label1:
                    If Worksheets("���ϕ��@").Cells(j, 4).Value = RngVal Then
                        Cells(j, 4).Delete xlShiftUp
                        Cells(j, 5).Delete xlShiftUp
                        GoTo Label1
                    End If
                Next j
                GoTo Label2
            Else
                GoTo Label3
            End If
            
        '�ڍׂ̍s���I�����ꂽ�ꍇ
        Case 5
            Dim Meth As String
            Meth = Cells(RngRow, (RngColumn - 1)).Value
            Dim result2 As Long
            result2 = MsgBox(Meth & "�ɑ�����" & RngVal & "���폜���܂����H", vbYesNo)
            If result2 = vbYes Then
                Range(Cells(RngRow, (RngColumn - 1)), Cells(RngRow, RngColumn)).Delete xlShiftUp
                GoTo Label2
            Else
                GoTo Label3
            End If
        Case Else
            MsgBox "�����ȃZ���ł�"
            GoTo Label4
        End Select

'�I�����ꂽ�s��S�č폜
Label2:
    MsgBox "�폜���܂���"
    GoTo Label4

'�L�����Z���̏ꍇ
Label3:
    MsgBox "�L�����Z������܂���"

Label4:
    
myErr: Exit Sub
    
End Sub
