Sub 連続改行削除()
'
' Macro1 Macro
' 記録日 00/03/29 記録者 Tomo Makoto
'
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "^p^p"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .MatchByte = False
        .MatchFuzzy = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
End Sub
Sub 連続改行削除2()
'
' Macro1 Macro
' 記録日 00/03/29 記録者 Tomo Makoto
'
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "^p^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .MatchByte = False
        .MatchFuzzy = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
End Sub
