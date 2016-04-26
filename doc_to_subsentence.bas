Sub サブ文書化_選択ファイル()
' ダイアログボックスで選択した複数のdoc文書を、サブ文書に変換し、かつ、セクション区切り文書に変換する。
' 各doc文章に、異なる頁スタイルが設定されていても、セクション区切りのため、それが崩れない。
' 変換後、印刷なり、pdf化なりをする際に、便利に使える。
' 印刷する場合は、個別に出力するよりもスピードアップが図れる。

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
        Debug.Print myPath

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
