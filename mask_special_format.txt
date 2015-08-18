' 強調したところだけをマスクする。メイン。
' 2005/02/06 Tomo Makoto
Sub maskingMain(ctHighlight As Boolean, ctBold As Boolean, ctItalic As Boolean, ctUnderline As Boolean)
'
    Dim stFound As String   '見つかった文字列を格納する変数
    stFound = ""            '変数の初期化
    Dim ctUnderlineSt
    Dim strLen As Integer
    Dim strRep As String
    Dim i As Integer
        
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
        .Replacement.Text = ""
    End With
        
    Do While Selection.Find.Execute = True
            
        strLen = Len(Selection.Text)
        'strRep = ""
        'For i = 1 To strLen
        '   strRep = strRep & " "
        'Next i
'        strRep = Space(strLen)
        strRep = String(strLen, "*")
        
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        
        With Selection.Find
           .Format = True
           .Replacement.Text = strRep
           .Execute Replace:=wdReplaceOne
        End With
        With Selection.Font
            With .Borders(1)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth025pt
                .Color = wdColorBlack
            End With
            .Borders.Shadow = False
        End With
    Loop
        
    Selection.HomeKey Unit:=wdStory '文章の先頭にジャンプ
    
End Sub
' Underlineにしたところだけを置換。
Sub maskUnderline()
    Call maskingMain(0, 0, 0, 1)
    Selection.WholeStory
    Selection.Font.Underline = wdUnderlineNone
End Sub

' 蛍光ペンでハイライトにしたところだけをマスク。
Sub maskHighlight()
    Call maskingMain(1, 0, 0, 0)
'    Selection.WholeStory
'    Selection.Range.HighlightColorIndex = wdNoHighlight
End Sub

' boldにしたところだけをマスク。
Sub maskBold()
    Call maskingMain(0, 1, 0, 0)
    Selection.WholeStory
    Selection.Font.Bold = False
End Sub

' イタリックにしたところだけをマスク。
Sub maskItalic()
    Call maskingMain(0, 0, 1, 0)
    Selection.WholeStory
    Selection.Font.Italic = False
End Sub
