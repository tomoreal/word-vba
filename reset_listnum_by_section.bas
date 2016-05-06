Sub 箇条書きを探索セクションごと振り直し()
' サブ文書にすると、段落番号が連番になるのを、セクションごとに振り直す
' (c) Makoto Tomo 2016.5.6

    Dim i As Long
    Dim para, myListNo
    
    With ActiveDocument.Sections
        For i = 1 To .Count
            'Debug.Print "i="; i
            For Each para In .Item(i).Range.Paragraphs
                myListNo = para.Range.ListFormat.ListString
            
                If (myListNo <> "") Then
                    'Debug.Print "myListno:"; myListNo
                    'Debug.Print para.Range.Text
                    para.Range.Select
                    Selection.HomeKey Unit:=wdLine
                    Call 段落番号リセット
                    Exit For
                End If
            Next
        Next
    End With
End Sub

Sub 段落番号リセット()
'
' マクロ記録で作成
' 番号振り直しがうまく行かない場合は、
' ご自身のワードで、マクロ記録してこれを作り直して下さい。
' 2016.5.6 by word2010
    
    With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = "%1."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = MillimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = MillimetersToPoints(7.4)
        .TabPosition = MillimetersToPoints(7.4)
        .ResetOnHigher = 0
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    ListGalleries(wdNumberGallery).ListTemplates(1).Name = ""
    Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
        ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:= _
        False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
        wdWord10ListBehavior
End Sub
