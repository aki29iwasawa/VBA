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

    '現在の年月を取得
    Dim y As Long, m As Long
    y = Year(Now)
    m = Month(Now)
'    MsgBox "年" & y & vbCrLf & "月" & m
    
    '「支出」シート上の前回の日付、年月をそれぞれ取得
    Dim prevY As Long, prevM As Long, pDate As Variant
    pDate = range("B1")
    prevY = Year(pDate)
    prevM = Month(pDate)
    
'    'サンプル値
'    prevY = 2020
'    prevM = 2
    
    '前回の利用から年をまたいだかどうか
    If y <> prevY Then
        Call newYear(prevY, prevM)
        GoTo Label1
    '前回の利用から月をまたいだかどうか
    ElseIf m <> prevM Then
        Call newMonth(prevM)
    End If

Label1:
    '「支出」シートの今日の日付を更新
    Sheets("支出").range("B1").Value = Date

End Sub
Sub newYear(prevY As Long, prevM As Long)
'年替わりにする処理
    '月替わりの処理を実行
    Call newMonth(prevM)

    MsgBox "年が変わりました" & vbCrLf & "昨年の支出シートを別ファイルに移動します。"
        
    Dim nbName As String, Path As String
    Dim nb As Workbook, main As Workbook
    
    '元の家計簿ブックを変数に入れる
    Set main = ThisWorkbook
    '新規ブックの名前
    nbName = prevY & "年家計簿.xlsx"
    '現在のファイルパス
    Path = ThisWorkbook.Path & "\" & nbName
'    MsgBox Path
    
    '現在のフォルダに同名のファイルがないか確認する
    If Dir(Path) = "" Then
        'ワークブックを新規追加、名前を変更
        Set nb = Workbooks.Add
        nb.SaveAs Path
        
        '元の家計簿ファイルからシートをもってくる
        Dim sh As Worksheet
        Dim cnt As Long
        
        '現在のブックの月ごとの支出ワークシートを切り取り
        For i = 5 To main.Worksheets.Count
            Set sh = main.Worksheets(i)
            'MsgBox sh
            cnt = nb.Sheets.Count
            main.Worksheets(i).Move after:=nb.Sheets(cnt)
            Columns.AutoFit
        Next i
        
        'ブック作成時にできたシートを削除
        nb.Sheets(1).Delete
        
        '新しく作ったファイルを保存して閉じる
        nb.Save
        nb.Close

    Else
        MsgBox prevY & "年のファイル(" & prevY & "年家計簿" & ")が既に存在します。"
    End If
    

End Sub

Sub newMonth(m As Long)
'月替わりにする処理

    '先月をシート名にする
    Dim prevM As String
    prevM = m - 1 & "月"
    
    'シート名をチェック
    Dim ws As Worksheet
    On Error Resume Next
    sName = Sheets(prevM).Name
    If Err.Number > 0 Then
        MsgBox "先月の支出を別シートに移動します"
        
        '新規シートの追加
        Set ws = Sheets.Add(after:=Sheets(Sheets.Count))
        ws.Name = prevM
        
        '支出シートの行数を取得
        Dim LastRow As Long
        '支出の最終行を取得
        LastRow = Worksheets("支出").range("B9").End(xlDown).Row
        
        'データを移す
        Sheets("支出").range(Cells(9, 2), Cells(LastRow, 12)).Cut ws.range("A1")
        Columns.AutoFit
        
        '支出シートのカラムを作成
        ' 配列を生成
        Dim strArray(10) As String
        strArray(0) = "日付"
        strArray(1) = "費目１"
        strArray(2) = "費目２"
        strArray(3) = "品名"
        strArray(4) = "個数"
        strArray(5) = "単価"
        strArray(6) = "税抜小計"
        strArray(7) = "税込小計"
        strArray(8) = "決済方法"
        strArray(9) = "決済の詳細"
        strArray(10) = "メモ"
        '横方向に書き込み
        Sheets("支出").range("B8:L8") = strArray
               
        Call ColDesign(Sheets("支出").range("B8:L8"))
        
    End If

End Sub
