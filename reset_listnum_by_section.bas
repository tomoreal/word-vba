Sub �ӏ�������T���Z�N�V�������ƐU�蒼��()
' �T�u�����ɂ���ƁA�i���ԍ����A�ԂɂȂ�̂��A�Z�N�V�������ƂɐU�蒼��
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
                    Call �i���ԍ����Z�b�g
                    Exit For
                End If
            Next
        Next
    End With
End Sub

Sub �i���ԍ����Z�b�g()
'
' �}�N���L�^�ō쐬
' �ԍ��U�蒼�������܂��s���Ȃ��ꍇ�́A
' �����g�̃��[�h�ŁA�}�N���L�^���Ă������蒼���ĉ������B
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
