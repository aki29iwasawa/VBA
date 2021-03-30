VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CatergoryForm 
   Caption         =   "費目追加"
   ClientHeight    =   4389
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   5740
   OleObjectBlob   =   "CatergoryForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "CatergoryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addCatergory_Click()
'追加ボタンクリック時の処理

'大きいカテゴリに入力されたものが「カテゴリ」にあるか確認、
'無かったら追加、あったら追加はしない
'サブカテゴリに入力されたものがあったらアラートを出す？

    '費目１が未入力の場合
    If ECatergory = "" Then
        MsgBox "費目1が未入力です"
        
    Else
        '「費目１」データ入力行を取得
        Dim NewRow As Long
        NewRow = Worksheets("支出カテゴリ").Cells(Rows.Count, 5).End(xlUp).Row + 1
        '「費目２」データ入力行を取得
        Dim NewRowSub As Long
        NewRowSub = Worksheets("支出カテゴリ").Cells(Rows.Count, 8).End(xlUp).Row + 1
            
        'カテゴリ登録されていない費目１が入力された場合、登録する
        Dim Cate1 As range
        Set Cate1 = Worksheets("支出カテゴリ").range(Cells(10, 5), Cells(NewRow, 5)).Find(What:=ECatergory, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
        
        If Cate1 Is Nothing Then
        '新規費目１と２を登録
            With Cells(NewRow, 5)
                .Value = ECatergory
                .Interior.Color = RGB(221, 235, 247)
                .Borders(xlEdgeTop).LineStyle = xlDash
                .Borders(xlEdgeTop).Color = RGB(47, 117, 181)
            End With
            Cells(NewRowSub, 7).Value = ECatergory
            Cells(NewRowSub, 8).Value = ESubcatergory
            
            'データ入力行のデザインを変更
            Call RowDesign(range(Cells(NewRowSub, 7), Cells(NewRowSub, 8)))
          
        ElseIf Not Cate1 Is Nothing And ESubcatergory = "" Then
        '入力された費目１が既に存在し、かつ、費目２の入力がない場合
            MsgBox ECatergory & "は既に存在する費目です。"
        
        ElseIf Not Cate1 Is Nothing Then
        '費目１が既に存在する場合
        '費目２の存在チェック
            Dim Cate2 As range
            Set Cate2 = Worksheets("支出カテゴリ").range(Cells(10, 8), Cells(NewRowSub, 8)).Find(What:=ESubcatergory, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
            
            If Cate2 Is Nothing Then
                '既存データが存在しない場合は入力
                Cells(NewRowSub, 7).Value = ECatergory
                Cells(NewRowSub, 8).Value = ESubcatergory
                'データ入力行のデザインを変更
                Call RowDesign(range(Cells(NewRowSub, 7), Cells(NewRowSub, 8)))
        
            ElseIf Not Cate2 Is Nothing Then
                '既存の費目２の検索findnextを使ってみる
                
                MsgBox "下記の費目1には" & ESubcatergory & "が存在します"
                
                
            End If
        End If
    End If
    
    'フォームを閉じる
    Unload CatergoryForm
    
End Sub

Private Sub Cancel_Click()
'フォームを閉じる
    
    Unload CatergoryForm
    
End Sub

Private Sub UserForm_Initialize()
'フォームの初期設定

    'カテゴリ欄が空白でなければ順に取得、選択肢に追加
    If Worksheets("支出カテゴリ").range("E10") <> "" Then
        '「カテゴリ」最終行を取得
        Dim CateRowCnt As Long
        CateRowCnt = Worksheets("支出カテゴリ").range("E10").End(xlDown).Row
        
        '「カテゴリ」一覧を費目１に入れる
        Dim i As Long, Cate As String
        '「カテゴリ」のはじめから最終行まで繰り返す
        For i = 10 To CateRowCnt
            Cate = Worksheets("支出カテゴリ").Cells(i, 5)
            Me.ECatergory.AddItem Cate
        Next
    End If
    
End Sub


