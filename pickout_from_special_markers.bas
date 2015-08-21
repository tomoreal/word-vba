' 蛍光ペンでハイライトにしたところだけを抜き出す。
Sub cutHighlight()
    Call cutMain(1, 0, 0, 0)
End Sub

' Underlineにしたところだけを抜き出す。
Sub cutUnderline()
    Call cutMain(0, 0, 0, 1)
End Sub

' boldにしたところだけを抜き出す。
Sub cutBold()
    Call cutMain(0, 1, 0, 0)
End Sub

' イタリックにしたところだけを抜き出す。
Sub cutItalic()
    Call cutMain(0, 0, 1, 0)
End Sub

' 強調したところだけを抜き出す。メイン。
' 記録日 00/03/28 記録者 Tomo Makoto
Sub cutMain(ctHighlight As Boolean, ctBold As Boolean, ctItalic As Boolean, ctUnderline As Boolean)
'
    Dim stFound As String   '見つかった文字列を格納する変数
    stFound = ""            '変数の初期化
    Dim ctUnderlineSt
        
    Selection.HomeKey Unit:=wdStory '文章の先頭にジャンプ
    
    '検索の開始
    Selection.Find.ClearFormatting
    
    '下線
    If ctUnderline = True Then
        Selection.Find.Font.Underline = wdUnderlineSingle
    End If
    
    With Selection.Find
        .Highlight = ctHighlight           'ハイライト
        .Font.Bold = ctBold                '太字
        .Font.Italic = ctItalic            'イタリック
        .Text = ""
        .Forward = True
        .Format = True
    End With
        
    Do While Selection.Find.Execute = True
        '見つかったところを変数に追加格納。改行を末尾に付ける。
        stFound = stFound & Selection.Text & vbCrLf
    Loop
    
    '新しい文章を開き、
    Documents.Add Template:="Normal", NewTemplate:=False
    '格納した変数を書き出す。
    Selection.InsertAfter (stFound)
    
    Selection.HomeKey Unit:=wdStory '文章の先頭にジャンプ
    
End Sub
