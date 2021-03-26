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
'支出追加

    ExpenseForm.Show
    
End Sub

Sub deleteExpense()
        
    Dim delRng As Range
    Set delRng = selectData()
    
    delRng.Select
    
    '選択されたセルの値と行番号を取得
    Dim RngVal As String, RngRow As Long, RngColumn As Long
    RngVal = delRng.Value
    RngRow = delRng.Row
    RngColumn = delRng.Column
   
'    '入力チェック
'    MsgBox "値" & RngVal & "行" & RngRow & "列" & RngColumn

    'カラム名の列が選択された場合
    If RngRow = 8 Or RngRow = 2 Then
        MsgBox "無効なセルです"
        GoTo Label4
    End If
    
    Dim thisRow As Range
    Set thisRow = Worksheets("支出").UsedRange.Rows(RngRow - 1)
    thisRow.Select
'    Worksheets("支出").UsedRange.Rows(RngRow - 1).Select
    
    Dim result As Long
    result = MsgBox("選択されたセルの行を削除しますか？", vbYesNo)
    If result = vbYes Then
        thisRow.Delete xlShiftUp
        GoTo Label2
    Else
        GoTo Label3
    End If

'選択された行を全て削除
Label2:
    MsgBox "削除しました"
    GoTo Label4
'キャンセルの場合
Label3:
    MsgBox "キャンセルされました"

Label4:


            

    
End Sub
