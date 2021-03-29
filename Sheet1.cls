VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CreateDetailGraph()
'詳細グラフの作成

    CreateGForm.Show
    
End Sub

Sub graph()

    'グラフシート上のグラフを全て削除する
    Worksheets("グラフ").ChartObjects.Delete

'カテゴリごとの合計支出表を作成
    Dim PSum, AllRow As Long
    Dim CCol, ExCol As Range
    Dim Cate As String
    
    '支出の最終行を取得
    AllRow = Worksheets("支出").Range("C9").End(xlDown).Row
    
    '費目列を9行目から最終行までセット
    Set CCol = Worksheets("支出").Range("C9:C" & AllRow)
    '支出額の列を9行目から最終行までセット
    Set ExCol = Worksheets("支出").Range("I9:I" & AllRow)
    
    '「カテゴリ」最終行を取得
    Dim CateRowCnt As Long
    CateRowCnt = Worksheets("支出カテゴリ").Range("E10").End(xlDown).Row
        
    '
    Dim i As Long
    '「カテゴリ」のはじめから最終行まで繰り返す
    For i = 10 To CateRowCnt
        '費目1を順にCateにセットする
        Cate = Worksheets("支出カテゴリ").Cells(i, 5)
        
        'Cateで検索
        PSum = WorksheetFunction.SumIf(CCol, Cate, ExCol)
        
'        MsgBox Cate & "の合計支出は" & PSum & "円です"
        
        '現在のカテゴリ名をグラフシートの表に入れる
        With Worksheets("グラフ")
            .Cells(i, 2).Value = Cate
            .Cells(i, 3).Value = PSum
        End With
        '追加した行のデザインを整える
        Call RowDesign(Range(Cells(i, 2), Cells(i, 3)))
    Next i

'グラフ上記の表を使ってグラフを作成
    Dim trgtSh As Worksheet
    Set trgtSh = ThisWorkbook.Worksheets("グラフ")
    
    Dim dataRng As Range
    Set dataRng = Range(Cells(10, 2), Cells(CateRowCnt, 3))
    
    '貼り付けたいセルを定義
    Dim pasteRng As Range
    Set pasteRng = trgtSh.Range("E4")
     
    'グラフ作成
    With trgtSh.Shapes.AddChart.Chart
        'グラフの種類指定
        .ChartType = xlColumnClustered
        '対象データ範囲を指定
        .SetSourceData dataRng
        'グラフタイトル設定
        .HasTitle = True
        .ChartTitle.Text = "費目別支出"
         
        'グラフの位置を指定
        .Parent.Top = pasteRng.Top
        .Parent.Left = pasteRng.Left
    End With

End Sub

