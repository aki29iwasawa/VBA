Attribute VB_Name = "Module1"
Function selectData() As Range
    
'キャンセルされたらエラー処理
    On Error GoTo myErr
   
    'セル範囲を選択してもらう
Label1:
    Dim delRng As Range
    Set delRng = Application.InputBox(Prompt:="セルを選択してください。", Type:=8)
    
'    '選択されたセルの値と行番号を取得
    Dim RngVal As String
', RngRow As Long, RngColumn As Long
    RngVal = delRng.Value
'    RngRow = delRng.Row
'    RngColumn = delRng.Column
'
'    MsgBox "値" & RngVal & "行" & RngRow & "列" & RngColumn
    
    '空白セルが選択された場合、Label1セル選択に戻る
    If RngVal = "" Then
        MsgBox "空白のセルが選択されました"
        GoTo Label1
    End If

    Set selectData = delRng

myErr: Exit Function
    
End Function

Sub RowDesign(Range)
'新規追加行のデザインを変更

    With Range
        .Interior.Color = RGB(221, 235, 247)
        .Borders(xlEdgeTop).LineStyle = xlDash
        .Borders(xlEdgeTop).Color = RGB(47, 117, 181)
    End With

End Sub

Sub graph()

    'グラフの対象データ範囲を定義
    Dim trgtSh As Worksheet
    Set trgtSh = ThisWorkbook.Worksheets("グラフ")
    
    Dim dataRng As Range
    Set dataRng = Union(Range("C8:C16"), Range("I8:I16"))
    
    '貼り付けたいセルを定義
    Dim pasteRng As Range
    Set pasteRng = trgtSh.Range("B2")
     
    'グラフ作成
    With trgtSh.Shapes.AddChart.Chart
        'グラフの種類を指定
        .ChartType = xlColumnClustered
        'グラフの対象データ範囲を指定
        .SetSourceData dataRng
        'グラフタイトルを設定
        .HasTitle = True
        .ChartTitle.Text = "費目別支出"
         
        'グラフの位置を指定
        .Parent.Top = pasteRng.Top
        .Parent.Left = pasteRng.Left
    End With

End Sub
