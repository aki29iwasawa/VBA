VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExpenseForm 
   Caption         =   "支出データを追加"
   ClientHeight    =   9359.001
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   6412
   OleObjectBlob   =   "ExpenseForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ExpenseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
    
    Unload ExpenseForm
    
End Sub

Private Sub addExpense_Click()
   
    'データ入力行を取得
    Dim NewRow As Long
    NewRow = Worksheets("支出").Cells(Rows.Count, 2).End(xlUp).Row + 1
    
    'データ入力行のデザインを変更
    Call RowDesign(Range(Cells(NewRow, 2), Cells(NewRow, 12)))
    
    Cells(NewRow, 2).Value = BDate
    Cells(NewRow, 3).Value = ECatergory
    Cells(NewRow, 4).Value = ESubcatergory
    Cells(NewRow, 5).Value = Item
    
    '「数量」を数値変換
    Dim qua As Long
    qua = Val(Quantity)
    Cells(NewRow, 6).Value = qua
    
    '「単価」を数値変換
    Dim pri As Long
    pri = Val(Price)
    Cells(NewRow, 7).Value = pri
    
    '「税抜小計」
    Dim priSum As Long
    priSum = qua * pri
    Cells(NewRow, 8).Value = priSum
    
    '「税込小計」
    Dim addTax As Long
    addTax = qua * pri * 1.1
    Cells(NewRow, 9).Value = addTax
    
    Cells(NewRow, 10).Value = PMethod
    Cells(NewRow, 11).Value = PDetail
    Cells(NewRow, 12).Value = Memo
   
    Unload ExpenseForm
    
End Sub

Private Sub DateSpin_SpinUp()
'日付
    Dim nDate As Date
    nDate = ExpenseForm.BDate.Value
    ExpenseForm.BDate.Value = DateAdd("d", 1, nDate)
    
End Sub

Private Sub DateSpin_SpinDown()
'日付
    Dim nDate As Date
    nDate = ExpenseForm.BDate.Value
    ExpenseForm.BDate.Value = DateAdd("d", -1, nDate)
    
End Sub

Private Sub Memo_Change()

End Sub

Private Sub UserForm_Initialize()

    'カテゴリ欄が空白でなければ順に取得、選択肢に追加
    If Worksheets("支出カテゴリ").Range("E10") <> "" Then
        '「カテゴリ」最終行を取得
        Dim CateRowCnt As Long
        CateRowCnt = Worksheets("支出カテゴリ").Range("E10").End(xlDown).Row
        
        '「カテゴリ」一覧を費目１に入れる
        Dim i As Long, Cate As String
        '「カテゴリ」のはじめから最終行まで繰り返す
        For i = 10 To CateRowCnt
            Cate = Worksheets("支出カテゴリ").Cells(i, 5)
            Me.ECatergory.AddItem Cate
        Next
    End If
        
    Me.BDate.Value = Date

    '決済方法欄が空白でなければ順に取得、選択肢に追加
    If Worksheets("決済方法").Range("B10") <> "" Then
        '「決済方法」最終行を取得
        Dim MethodRowCnt As Long
        MethodRowCnt = Worksheets("決済方法").Range("B10").End(xlDown).Row
        
        '「決済方法」一覧を費目１に入れる
        Dim j As Long, Method As String
        '「決済方法」のはじめから最終行まで繰り返す
        For i = 10 To MethodRowCnt
            Method = Worksheets("決済方法").Cells(i, 2)
            Me.PMethod.AddItem Method
        Next
    End If

End Sub
Private Sub ECatergory_Change()

    '費目２をクリア
    ESubcatergory.Clear
    
    '費目１で選択されたものを取得
    Dim Cate As String
    Cate = ECatergory.Text
    
    With Worksheets("支出カテゴリ")
        '検索結果を表示するB1からをクリア
        .Range("B14:C" & Rows.Count).Clear

        '費目1で選択されたものでオートフィルター
        .Range("G9").AutoFilter Field:=1, Criteria1:="*" & Cate & "*"
        'コピぺ
        .Range("G9").CurrentRegion.Copy
        .Range("B13").PasteSpecial Paste:=xlPasteAll
        'オートフィルター解除
        .Range("G9").AutoFilter

    End With
    
    'カテゴリ欄が空白でなければ順に取得、選択肢に追加
    If Worksheets("支出カテゴリ").Range("C14") <> "" Then
        '「カテゴリ」最終行を取得
        Dim CateRowCnt As Long
        CateRowCnt = Worksheets("支出カテゴリ").Range("C13").End(xlDown).Row
        
'        MsgBox CateRowCnt
      
        '「カテゴリ」一覧を費目2に入れる
        Dim i As Long, subCate As String
        '「カテゴリ」のはじめから最終行まで繰り返す
        For i = 14 To CateRowCnt
            subCate = Worksheets("支出カテゴリ").Cells(i, 3).Value
            Me.ESubcatergory.AddItem subCate
        Next
    End If
    
    '検索結果を表示するB1からをクリア
    Worksheets("支出カテゴリ").Range("B14:C" & Rows.Count).Clear

    Application.CutCopyMode = False
    
End Sub

Private Sub PMethod_Change()

    '詳細をクリア
    PDetail.Clear
    
    '決済方法で選択されたものを取得
    Dim Meth As String
    Meth = PMethod.Text
    
    With Worksheets("決済方法")
'        '検索結果を表示するG1からをクリア
'        .Range("B14:C" & Rows.Count).Clear

        '決済方法で選択されたものでオートフィルター
        .Range("D9").AutoFilter Field:=1, Criteria1:="*" & Meth & "*"
        'コピぺ
        .Range("D9").CurrentRegion.Copy
        .Range("G9").PasteSpecial Paste:=xlPasteAll
        'オートフィルター解除
        .Range("D9").AutoFilter

    End With
    
    '詳細欄が空白でなければ順に取得、選択肢に追加
    If Worksheets("決済方法").Range("H10") <> "" Then
        '決済方法最終行を取得
        Dim MethRowCnt As Long
        MethRowCnt = Worksheets("決済方法").Range("H9").End(xlDown).Row
        
        '「決済方法」一覧を詳細に入れる
        Dim i As Long, Det As String
        '「カテゴリ」のはじめから最終行まで繰り返す
        For i = 10 To MethRowCnt
            Det = Worksheets("決済方法").Cells(i, 8).Value
            Me.PDetail.AddItem Det
        Next
    End If
    
    'オートフィルターの結果をペーストしたセルをクリア
    Worksheets("決済方法").Range("G9:H" & Rows.Count).Clear
    
    Application.CutCopyMode = False
        
End Sub



