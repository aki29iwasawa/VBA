VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub addPMethod()
'決済方法追加

    PaymentMethodForm.Show
    
End Sub

Sub deletePMethod()

'キャンセルされたらエラー処理
    On Error GoTo myErr
        
    Dim delRng As Range
    Set delRng = selectData()
    
    delRng.Select
    
    '選択されたセルの値と行番号を取得
    Dim RngVal As String, RngRow As Long, RngColumn As Long
    RngVal = delRng.Value
    RngRow = delRng.Row
    RngColumn = delRng.Column
   
'    '入力チェック
'    MsgBox "値" & RngVal & "行" & RngRow & "列" & RngColumn

    'カラム名の列が選択された場合
    If RngRow = 9 Or RngRow = 2 Then
        MsgBox "無効なセルです"
        GoTo Label4
    End If
        
    '決済方法の中で、方法が選択されたか、詳細が選択されたか
    Select Case RngColumn
        '方法の行が選択された場合
        Case 2, 4
            Dim result1 As Long
            result1 = MsgBox("決済方法から" & RngVal & "を削除しますか？" & vbCrLf & "※" & RngVal & "に属する詳細も全て削除されます", vbYesNo)
            If result1 = vbYes Then
                'B列を削除
                Dim ColumnB As Range, LastB As Long
                Set ColumnB = Range("B10")
                LastB = Range("B10").End(xlDown).Row

                Dim i As Long
                For i = 10 To LastB
                    If ColumnB.Value = RngVal Then
                        ColumnB.Delete xlShiftUp
                        Exit For
                    End If
                    Set ColumnB = ColumnB.Offset(1)
                Next
                
'試行錯誤の残り
'                'DH列を削除
'                Dim ColumnD As Range, LastD As Long
'                Set ColumnD = Range("D10")
'                LastD = Range("D10").End(xlDown).Row
'
'                Dim j As Long
'                For j = 10 To LastD
'                    If ColumnD.Value = RngVal Then
'                        ColumnD.Delete xlShiftUp
'                        ColumnD.Offset(0, 1).Select
'                        ColumnD.Offset(0, 1).Delete xlShiftUp
'                        ColumnD.Offset(1).Delete xlShiftUp
'                    End If
'                    Set ColumnD = ColumnD.Offset(1)
'                Next
'                GoTo Label2
'            Else
'                GoTo Label3
'            End If
'セル一つずつの繰り返しになってしまう
'                Dim hani As Range
'                Dim thisRow As Range
'                Set hani = Range("D10:E10")
'                For Each thisRow In hani
'                    thisRow.Select
'                    MsgBox thisRow.Value
'                Next thisRow


                'DH列を削除
                '対象の最終行取得
                Dim LastDE As Long
                LastDE = Worksheets("決済方法").Range("D10").End(xlDown).Row
                
                Dim j As Long
                '10行目からDEの最終行まで繰り返し
                For j = 10 To LastDE
                Cells(j, 4).Select
Label1:
                    If Worksheets("決済方法").Cells(j, 4).Value = RngVal Then
                        Cells(j, 4).Delete xlShiftUp
                        Cells(j, 5).Delete xlShiftUp
                        GoTo Label1
                    End If
                Next j
                GoTo Label2
            Else
                GoTo Label3
            End If
            
        '詳細の行が選択された場合
        Case 5
            Dim Meth As String
            Meth = Cells(RngRow, (RngColumn - 1)).Value
            Dim result2 As Long
            result2 = MsgBox(Meth & "に属する" & RngVal & "を削除しますか？", vbYesNo)
            If result2 = vbYes Then
                Range(Cells(RngRow, (RngColumn - 1)), Cells(RngRow, RngColumn)).Delete xlShiftUp
                GoTo Label2
            Else
                GoTo Label3
            End If
        Case Else
            MsgBox "無効なセルです"
            GoTo Label4
        End Select

'選択された行を全て削除
Label2:
    MsgBox "削除しました"
    GoTo Label4

'キャンセルの場合
Label3:
    MsgBox "キャンセルされました"

Label4:
    
myErr: Exit Sub
    
End Sub

