Attribute VB_Name = "Module1"
Function selectData() As range
    
'�L�����Z�����ꂽ��G���[����
    On Error GoTo myErr
   
    '�Z���͈͂�I�����Ă��炤
Label1:
    Dim delRng As range
    Set delRng = Application.InputBox(Prompt:="�Z����I�����Ă��������B", Type:=8)
    
'    '�I�����ꂽ�Z���̒l�ƍs�ԍ����擾
    Dim RngVal As String
', RngRow As Long, RngColumn As Long
    RngVal = delRng.Value
'    RngRow = delRng.Row
'    RngColumn = delRng.Column
'
'    MsgBox "�l" & RngVal & "�s" & RngRow & "��" & RngColumn
    
    '�󔒃Z�����I�����ꂽ�ꍇ�ALabel1�Z���I���ɖ߂�
    If RngVal = "" Then
        MsgBox "�󔒂̃Z�����I������܂���"
        GoTo Label1
    End If

    Set selectData = delRng

myErr: Exit Function
    
End Function

Sub RowDesign(range)
'�V�K�ǉ��s�̃f�U�C����ύX

    With range
        .Interior.Color = RGB(221, 235, 247)
        .Borders(xlEdgeTop).LineStyle = xlDash
        .Borders(xlEdgeTop).Color = RGB(47, 117, 181)
    End With

End Sub


Sub ColDesign(range)
'�J�����̃f�U�C����ύX
    
    With range
        .Font.Bold = True
        .Interior.ThemeColor = msoThemeColorAccent5
        .Interior.TintAndShade = 0.4
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Color = RGB(47, 117, 181)
    End With
        
End Sub

Sub sample()

End Sub

