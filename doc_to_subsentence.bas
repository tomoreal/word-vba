Sub 複数文書を単一化()
' ファイルを選択し、サブ文書にして、それを解除。
' サブ文書にすると、段落番号が連番になるのを、セクションごとに振り直す
' (c) Makoto Tomo 2016.5.6

    Call サブ文書化_選択ファイル
    Call 箇条書きの探索セクションごと振り直し
End Sub

Sub サブ文書化_選択ファイル()
' ファイルを選択し、サブ文書にして、それを解除。
'
' (c) Makoto Tomo 2016.5.6
    
    Dim myFile As Variant
    Dim strFiles As String
    Dim pos As Long
    
    ' 新しいファイルの追加
    Documents.Add DocumentType:=wdNewBlankDocument
    ' アウトラインビューにする
    ActiveWindow.ActivePane.View.Type = wdMasterView
    
    ' ファイルダイアログを開き、選択ファイルをサブ文書に追加する
    With Application.FileDialog(msoFileDialogFilePicker)
        'ファイルの複数選択を可能にする
        .AllowMultiSelect = True
        'ファイルフィルタのクリア
        .Filters.Clear
        'ファイルフィルタの追加
        .Filters.Add "word file", "*.doc*"

        If .Show = -1 Then  'ファイルダイアログ表示
            ' [ OK ] ボタンが押された場合
            For i = 1 To .SelectedItems.Count
                ' 選択したファイルを変数に入れる
                strFiles = .SelectedItems(i)
                ' 選択したファイルを、サブ文書として挿入
                Selection.Range.Subdocuments.AddFromFile Name:= _
                    strFiles, ConfirmConversions:=False, ReadOnly:= _
                    False, PasswordDocument:="", PasswordTemplate:="", Revert:=False, _
                    WritePasswordDocument:="", WritePasswordTemplate:=""
            Next i
        Else
            ' [ キャンセル ] ボタンが押された場合
            MsgBox "ファイル選択がキャンセルされました。", vbExclamation
            Exit Sub
        End If
        pos = InStrRev(strFiles, "\")
        myPath = Left(strFiles, pos)
        'Debug.Print myPath

    End With
   
    ' 最後の行をノーマルにする（アウトライン解除）
    Selection.Range.Style = ActiveDocument.Styles(wdStyleNormal)
    ' 文書先頭に移動
    Selection.HomeKey Unit:=wdStory
    ' 先頭行を削除
    Selection.Delete Unit:=wdCharacter, Count:=1
    
    ' アウトラインビューにする
    If ActiveWindow.View = wdMasterView Then
        ActiveWindow.View = wdOutlineView
    Else
        ActiveWindow.View = wdMasterView
    End If
    
    ' 文書全体を選択
    Selection.WholeStory
    ' サブ文書をコピーして、削除（普通のセクション区切り文書にする）
    Selection.Range.Subdocuments.Delete
    ' 文書先頭に移動
    Selection.HomeKey Unit:=wdStory
    ' 印刷ビューにする
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    Else
        ActiveWindow.View.Type = wdPrintView
    End If
    
End Sub

Sub 箇条書きの探索セクションごと振り直し()
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
