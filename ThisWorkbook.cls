VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

    '�������擾
    Dim m As Integer
    m = Month(Now)
'    MsgBox "������" & m & "���ł�"
    '�挎���V�[�g���ɂ���
    Dim prevM As String
    prevM = m - 1 & "��"
'    MsgBox prevM
    
    Dim SName As String, ws As Worksheet
    On Error Resume Next
    SName = Sheets(prevM).Name
    If Err.Number > 0 Then
        MsgBox "�挎�̎x�o��ʃV�[�g�Ɉړ����܂�"
        
        '�V�K�V�[�g�̒ǉ�
        Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = prevM
        
        '�f�[�^���ڂ�
        Sheets("�x�o").range("B8:L16").Copy
        
        ws.range("A1").PasteSpecial Paste:=xlPasteColumnWidths
        ws.range("A1").PasteSpecial Paste:=xlPasteAll
'        Sheets("�x�o").range("B8:L16").PasteSpecial Paste:=xlPasteColumnWidths
'        Sheets("�x�o").range("B8:L16").Copy Destination:=ws.range("A1")
        
    Else
        Exit Sub
    End If

End Sub