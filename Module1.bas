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

