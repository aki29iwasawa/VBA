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

    '今月を取得
    Dim m As Integer
    m = Month(Now)
'    MsgBox "今月は" & m & "月です"
    '先月をシート名にする
    Dim prevM As String
    prevM = m - 1 & "月"
'    MsgBox prevM
    
    Dim SName As String, ws As Worksheet
    On Error Resume Next
    SName = Sheets(prevM).Name
    If Err.Number > 0 Then
        MsgBox "先月の支出を別シートに移動します"
        
        '新規シートの追加
        Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = prevM
        
        'データを移す
        Sheets("支出").range("B8:L16").Copy
        
        ws.range("A1").PasteSpecial Paste:=xlPasteColumnWidths
        ws.range("A1").PasteSpecial Paste:=xlPasteAll
'        Sheets("支出").range("B8:L16").PasteSpecial Paste:=xlPasteColumnWidths
'        Sheets("支出").range("B8:L16").Copy Destination:=ws.range("A1")
        
    Else
        Exit Sub
    End If

End Sub
