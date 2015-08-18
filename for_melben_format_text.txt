Sub ★択一問題語群位置修正()
    Dim i As Integer
    
    Dim Message, Title, Default, MyNum
    
    Message = "最大の設問番号の値を入力してください。"    ' 入力を求めるメッセージを設定します。
    Title = "設問数指定"                ' タイトルを設定します。
    Default = "10"                            ' 既定値を設定します。
    ' メッセージ、タイトル、既定値を表示します。
    MyNum = InputBox(Message, Title, Default)
    
    Call 段落幅修正
    
    ActiveWindow.ActivePane.View.Type = wdMasterView
    ActiveWindow.View.WrapToWindow = True
    ActiveWindow.ActivePane.View.Zoom.Percentage = 25

    For i = 1 To MyNum
        Call 語群位置移動(i)
        Call 語群括弧削除(i)
    Next i

    ActiveWindow.ActivePane.View.Zoom.Percentage = 100
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    Else
        ActiveWindow.View.Type = wdPrintView
    End If
    
    Call 語群文字削除

End Sub

Sub ★択一問題語群位置修正2()
    Dim i As Integer

    Call 段落幅修正
    
    ActiveWindow.ActivePane.View.Type = wdMasterView

    For i = 1 To 10
        Call 語群位置移動(i)
        Rem Call 語群括弧削除(i)
    Next i

    ActiveWindow.ActivePane.View.Zoom.Percentage = 100
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    Else
        ActiveWindow.View.Type = wdPrintView
    End If
    
    Call 語群文字削除

End Sub
Sub 段落幅修正()
    Selection.WholeStory
    With Selection.ParagraphFormat
        .RightIndent = MillimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .AutoAdjustRightIndent = True
        .DisableLineHeightGrid = False
        .FarEastLineBreakControl = True
        .WordWrap = True
        .HangingPunctuation = True
        .HalfWidthPunctuationOnTopOfLine = False
        .AddSpaceBetweenFarEastAndAlpha = True
        .AddSpaceBetweenFarEastAndDigit = True
        .BaseLineAlignment = wdBaselineAlignAuto
    End With

End Sub

Sub 語群位置移動(MyNum As Integer)
'
' Macro1 Macro
' 記録日 2005/10/21 記録者 tomo
'

    Dim myNum_text1 As String
    Dim myNum_text2 As String
    
    
    myNum_text1 = "【 " & MyNum & " 】(1)"
    myNum_text2 = "【 " & MyNum & " 】"

    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = myNum_text1
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
        .MatchFuzzy = True
    End With
    Selection.Find.Execute
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.Font.Size = 10.5
    
    Selection.Cut
     
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = myNum_text2
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
        .MatchFuzzy = True
    End With
    Selection.Find.Execute
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.HomeKey Unit:=wdLine
    Selection.Paste
    
    
End Sub

Sub 語群文字削除()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "◆語群 ^p^p"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
        .MatchFuzzy = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub 語群括弧削除(MyNum As Integer)
'
    Dim myNum_text1 As String
    Dim myNum_text2 As String
    
    
    myNum_text1 = "【 " & MyNum & " 】(1)"
    myNum_text2 = "(1)"
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = myNum_text1
        .Replacement.Text = myNum_text2
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
        .MatchFuzzy = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub doc形式保存()
'
    Dim myFileName, myFilePath
    myFileName = ActiveDocument.FullName
    myFilePath = ActiveDocument.PATH

    ChangeFileOpenDirectory myFilePath

    ActiveDocument.SaveAs FileName:=myFileName, FileFormat _
        :=wdFormatDocument, LockComments:=False, Password:="", AddToRecentFiles:= _
        True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:= _
        False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False
End Sub
Sub ページ30行()
'
' ページ30行 Macro
' 記録日 2005/12/19 記録者 tomo
'
    With ActiveDocument.PageSetup
        .LinesPage = 30
    End With
End Sub