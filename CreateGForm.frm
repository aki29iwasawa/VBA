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

    'カテゴリ名でタイトルを作成
    Rows(title).RowHeight = 40
    With Cells(title, 2)
        .Value = ECatergory & "の詳細支出"
        .Font.Bold = True
        .Font.Size = 26
    End With
    With Range(Cells(title, 2), Cells(title, 3))
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    
    '「カテゴリ」最終行を取得
    Dim CateRowCnt, title As Long
    CateRowCnt = Worksheets("支出カテゴリ").Range("E10").End(xlDown).Row
    title = CateRowCnt + 4

'詳細費目表を作成
    Dim PSum, AllRow As Long
    Dim CCol, ExCol As Range
    Dim Cate As String
    
    '支出の最終行を取得
    AllRow = Worksheets("支出").Range("C9").End(xlDown).Row
    
    '費目列を9行目から最終行までセット
    Set CCol = Worksheets("支出").Range("C9:C" & AllRow)
    '支出額の列を9行目から最終行までセット
    Set ExCol = Worksheets("支出").Range("I9:I" & AllRow)
    
        
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

Private Sub UserForm_Initialize()
'フォームの初期設定

    'カテゴリ欄が空白でなければ順に取得、選択肢に追加
    If Worksheets("支出カテゴリ").Range("E10") <> "" Then
        '「カテゴリ」最終行を取得
        Dim CateRowCnt As Long
        CateRowCnt = Worksheets("支出カテゴリ").Range("E10").End(xlDown).Row
        
        '「カテゴリ」一覧をコンボボックスに入れる
        Dim i As Long, Cate As String
        '「カテゴリ」のはじめから最終行まで繰り返す
        For i = 10 To CateRowCnt
            Cate = Worksheets("支出カテゴリ").Cells(i, 5)
            Me.ECatergory.AddItem Cate
        Next
    End If
    
End Sub

