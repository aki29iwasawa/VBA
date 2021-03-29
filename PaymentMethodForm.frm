VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PaymentMethodForm 
   Caption         =   "決済方法"
   ClientHeight    =   4291
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   5768
   OleObjectBlob   =   "PaymentMethodForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "PaymentMethodForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addPMethod_Click()
'追加ボタンクリック時の処理


    '決済方法が未入力の場合
    If PMethod = "" Then
        MsgBox "決済方法が未入力です"
    Else
        '「決済方法」データ入力行を取得
        Dim NewRow As Long, NewDetRow As Long
        NewRow = Worksheets("決済方法").Cells(Rows.Count, 2).End(xlUp).Row + 1
        NewDetRow = Worksheets("決済方法").Cells(Rows.Count, 5).End(xlUp).Row + 1
        
        'カテゴリ登録されていない費目１が入力された場合、登録する
        Dim Method As Range
        Set Method = Worksheets("決済方法").Range(Cells(10, 2), Cells(NewRow, 2)).Find(What:=PMethod, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
        If Method Is Nothing Then
        '新規決済方法と詳細を登録
            With Cells(NewRow, 2)
                .Value = PMethod
                .Interior.Color = RGB(221, 235, 247)
                .Borders(xlEdgeTop).LineStyle = xlDash
                .Borders(xlEdgeTop).Color = RGB(47, 117, 181)
            End With
            
            If PDetail <> "" Then
                Cells(NewDetRow, 4).Value = PMethod
                Cells(NewDetRow, 5).Value = PDetail
                'データ入力行のデザインを変更
                Call RowDesign(Range(Cells(NewDetRow, 4), Cells(NewDetRow, 5)))
            End If
          
        ElseIf Not Method Is Nothing And PDetail = "" Then
        '入力された決済方法が既に存在し、かつ、詳細の入力がない場合
            MsgBox Method & "は登録済みの決済方法です。"
        
        ElseIf Not Method Is Nothing Then
        '決済方法が場登録済みの場合
        '詳細の存在チェック
            Dim dtl As Range
            Set dtl = Worksheets("決済方法").Range(Cells(10, 5), Cells(NewRow, 5)).Find(What:=PDetail, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
            
            If dtl Is Nothing Then
                '既存データが存在しない場合は入力
                Cells(NewDetRow, 4).Value = PMethod
                Cells(NewDetRow, 5).Value = PDetail
                'データ入力行のデザインを変更
                Call RowDesign(Range(Cells(NewDetRow, 4), Cells(NewDetRow, 5)))
        
            ElseIf Not dtl Is Nothing Then
                '後で
                
                MsgBox "登録済みです"
                
                
            End If
        End If
    End If
    
    'フォームを閉じる
    Unload PaymentMethodForm

End Sub


Private Sub Cancel_Click()
    
    Unload PaymentMethodForm
    
End Sub

Private Sub UserForm_Initialize()

    'カテゴリ欄が空白でなければ順に取得、選択肢に追加
    If Worksheets("決済方法").Range("B10") <> "" Then
        '「カテゴリ」最終行を取得
        Dim RowCnt As Long
        RowCnt = Worksheets("決済方法").Range("B10").End(xlDown).Row
        
        '「決済方法」一覧をコンボボックスに入れる
        Dim i As Long, Cate As String
        '「決済方法」のはじめから最終行まで繰り返す
        For i = 10 To RowCnt
            mth = Worksheets("決済方法").Cells(i, 2)
            Me.PMethod.AddItem mth
        Next
    End If
    
End Sub
