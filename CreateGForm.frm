VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateGForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3248
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   5292
   OleObjectBlob   =   "CreateGForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "CreateGForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
'フォームを閉じる
    
    Unload CreateGForm
    
End Sub

Private Sub CommandButton1_Click()
    
    'ワークシートを変数に置く
    Dim gra, cat As Worksheet
    Set gra = ThisWorkbook.Worksheets("グラフ")
    Set cat = ThisWorkbook.Worksheets("支出カテゴリ")

    '「支出カテゴリ」シートの費目１最終行を取得
    Dim CateRowCnt, title As Long
    CateRowCnt = cat.range("E10").End(xlDown).Row
    title = CateRowCnt + 4

    'カテゴリ名で表タイトルを作成
    Rows(title).RowHeight = 40
    With Cells(title, 2)
        .Value = ECatergory & "の詳細支出"
        .Font.Bold = True
        .Font.Size = 26
    End With
    With range(Cells(title, 2), Cells(title, 3))
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    
    '表を作成するセル,カラムを設定（タイトルを作った2行下）
    title = title + 2
    Cells(title, 2).Value = ECatergory & "の品目"
    Cells(title, 3).Value = "支出"
    Call ColDesign(range(Cells(title, 2), Cells(title, 3)))
    
    '選択された費目1を変数に置く
    Dim Cate As String
    Cate = ECatergory.Value
    
'「支出カテゴリ」シートの費目2をオートフィルターで絞る
    With cat
        '検索結果を表示するB1からをクリア
        .range("B13:C" & Rows.Count).Clear
        'G10からの１列目をCateで絞り込み
        .range("G9").AutoFilter Field:=1, Criteria1:="*" & Cate & "*"
        'コピー
        .range("G9").CurrentRegion.Copy
        .range("B13").PasteSpecial Paste:=xlPasteAll
        'オートフィルタを解除
        .AutoFilterMode = False
    End With
        
    Application.CutCopyMode = False
        
'カテゴリごとの合計支出表を作成
    Dim PSum, AllRow As Long
    Dim CCol, ExCol As range
    Dim Cate2 As String
    
    '支出の最終行を取得
    AllRow = Worksheets("支出").range("C9").End(xlDown).Row
    
    '費目2の列を9行目から最終行までセット
    Set CCol = Worksheets("支出").range("D9:D" & AllRow)
    '支出額の列を9行目から最終行までセット
    Set ExCol = Worksheets("支出").range("I9:I" & AllRow)
    
    '「カテゴリ」最終行を取得
    Dim Cate2RowCnt As Long
    Cate2RowCnt = cat.range("C14").End(xlDown).Row
        
    '
    Dim i, j As Long
    j = title + 1
    '「カテゴリ」のはじめから最終行まで繰り返す
    For i = 14 To Cate2RowCnt
        '費目2を順にCate2にセットする
        Cate2 = cat.Cells(i, 3)
        
        'Cateで検索
        PSum = WorksheetFunction.SumIf(CCol, Cate2, ExCol)
        
        MsgBox Cate2 & "の合計支出は" & PSum & "円です"
        
        '現在のカテゴリ名をグラフシートの表に入れる
        With gra
            .Cells(j, 2).Value = Cate2
            .Cells(j, 3).Value = PSum
        End With
        '追加した行のデザインを整える
        Call RowDesign(range(Cells(j, 2), Cells(j, 3)))
        j = j + 1
    Next i

'上記の表を使ってグラフを作成
    '表の最終行を取得
    Dim tableL As Long
    tableL = gra.Cells(title, 2).End(xlDown).Row
    
    '貼り付けたいセルを定義
    Dim pasteRng As range
    Set pasteRng = gra.Cells(title, 5)
    
    '表にするデータ範囲をセット
    Dim dataRng As range
    Set dataRng = range(Cells(title + 1, 2), Cells(tableL, 3))
         
    'グラフ作成
    With gra.Shapes.AddChart.Chart
        'グラフの種類指定
        .ChartType = xlColumnClustered
        '対象データ範囲を指定
        .SetSourceData dataRng
        'グラフタイトル設定
        .HasTitle = True
        .ChartTitle.Text = ECatergory & "の詳細支出"
         
        'グラフの位置を指定
        .Parent.Top = pasteRng.Top
        .Parent.Left = pasteRng.Left
    End With

    'フォームを閉じる
    Unload CreateGForm

End Sub

Private Sub UserForm_Initialize()
'フォームの初期設定

    'カテゴリ欄が空白でなければ順に取得、選択肢に追加
    If Worksheets("支出カテゴリ").range("E10") <> "" Then
        '「カテゴリ」最終行を取得
        Dim CateRowCnt As Long
        CateRowCnt = Worksheets("支出カテゴリ").range("E10").End(xlDown).Row
        
        '「カテゴリ」一覧をコンボボックスに入れる
        Dim i As Long, Cate As String
        '「カテゴリ」のはじめから最終行まで繰り返す
        For i = 10 To CateRowCnt
            Cate = Worksheets("支出カテゴリ").Cells(i, 5)
            Me.ECatergory.AddItem Cate
        Next
    End If
    
End Sub

