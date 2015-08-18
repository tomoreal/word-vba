Sub 目次抜粋()
'
' 記録日 00/04/06 記録者 Tomo Makoto
'
    Selection.EndKey Unit:=wdStory '文章の末尾にジャンプ
    With ActiveDocument
        .TablesOfContents.Add Range:=Selection.Range, RightAlignPageNumbers:= _
            True, UseHeadingStyles:=True, UpperHeadingLevel:=1, _
            LowerHeadingLevel:=9, IncludePageNumbers:=False, AddedStyles:=""
        .TablesOfContents(1).TabLeader = wdTabLeaderDots
        .TablesOfContents.Format = wdIndexIndent
    End With
    
    Selection.MoveUp Unit:=wdLine, Count:=1

    Selection.Fields.Unlink
    Selection.Cut
    
    Selection.Find.ClearFormatting
    
    Documents.Add
    Selection.Paste
    Selection.WholeStory
    Selection.Paragraphs.OutlineDemoteToBody
    Selection.HomeKey Unit:=wdStory

    Selection.HomeKey Unit:=wdStory '文章の先頭にジャンプ

End Sub
