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
'カテゴリ追加

    CatergoryForm.Show
    
End Sub

Sub deleteCatergory()
        
'キャンセルされたらエラー処理
    On Error GoTo myErr

    '標準モジュールを呼び出す
    Dim delRng As range
'ワークシートを渡す
'    Dim sheet As Worksheet
'    Set sheet = Sheets("支出カテゴリ")
    Set delRng = selectData()
    
    delRng.Select
    
    '選択されたセルの値と行番号を取得
    Dim RngVal As String, RngRow As Long, RngColumn As Long
    RngVal = delRng.Value
    RngRow = delRng.Row
    RngColumn = delRng.Column
    
'    MsgBox "値" & RngVal & "行" & RngRow & "列" & RngColumn

    delRng.Select
        
    'カラム名の列が選択された場合
    If RngRow = 9 Then
        MsgBox "無効なセルです"
        GoTo Label4
    End If
        
    '支出カテゴリの中で、費目1が選択されたか、2が選択されたか
    Select Case RngColumn
        '費目1の行が選択された場合
        Case 5, 7
            Dim result1 As Long
            result1 = MsgBox("費目1から" & RngVal & "を削除しますか？" & vbCrLf & "※" & RngVal & "に属する費目2も全て削除されます", vbYesNo)
            If result1 = vbYes Then
                'E列を削除
                Dim ColumnE As range, LastE As Long
                Set ColumnE = range("E10")
                LastE = range("E10").End(xlDown).Row

                Dim i As Long
                For i = 10 To LastE
                    ColumnE.Select
                    If ColumnE.Value = RngVal Then
                        ColumnE.Delete xlShiftUp
                        Exit For
                    End If
                    Set ColumnE = ColumnE.Offset(1)
                Next
                
                'GH列を削除
                '対象の最終行取得
                Dim LastGH As Long
                LastGH = Worksheets("支出カテゴリ").range("G10").End(xlDown).Row
                
                Dim j As Long
                '10行目からDEの最終行まで繰り返し
                For j = 10 To LastGH
                Cells(j, 7).Select
Label1:
                    If Worksheets("支出カテゴリ").Cells(j, 7).Value = RngVal Then
                        Cells(j, 7).Delete xlShiftUp
                        Cells(j, 8).Delete xlShiftUp
                        '1行削除したら上にずれるので、もう一度同じ行を見る
                        GoTo Label1
                    End If
                Next j
                GoTo Label2
            Else
                GoTo Label3
            End If
            
        '費目2の行が選択された場合
        Case 8
            Dim Cate As String
            Cate = Cells(RngRow, (RngColumn - 1)).Value
            
            Dim result2 As Long
            result2 = MsgBox(Cate & "に属する" & RngVal & "を削除しますか？", vbYesNo)
            
            If result2 = vbYes Then
                range(Cells(RngRow, (RngColumn - 1)), Cells(RngRow, RngColumn)).Delete xlShiftUp
                GoTo Label2
            Else
                GoTo Label3
            End If
        'それ以外のケース
        Case Else
            MsgBox "無効なセルです"
            GoTo Label4
        End Select
'    End If

'大きいカテゴリから全て削除
Label2:
    MsgBox "削除しました"
    GoTo Label4

'キャンセルの場合
Label3:
    MsgBox "キャンセルしました"
Label4:

myErr: Exit Sub

End Sub


Sub CateAutoFilter()
'きちんと動かないので、明日作り直す

    '検索結果を表示するB1からをクリア
    range("B13:C" & Rows.Count).Clear
    
    With Sheets("支出カテゴリ")
        'G10からの１列目をC9で絞り込み
        .range("G9").AutoFilter Field:=1, Criteria1:="*" & range("C9").Value & "*"
        'コピー
        .range("G9").CurrentRegion.Copy
        .range("B13").PasteSpecial Paste:=xlPasteAll
        'オートフィルタを解除
        .AutoFilterMode = False
    End With
    
    Application.CutCopyMode = False
    range("A1").Select
    
End Sub




