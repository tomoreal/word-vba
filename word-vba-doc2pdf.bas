Attribute VB_Name = "一括pdf化"
Sub フォルダ内文書の一括pdf化()
    'Makoto Tomo 2015/08/17
    Dim files As String
    Dim PATH As String
    Dim FSO As Object
    Dim strAFN As String
    Dim strAFN2 As String
    Dim strPDFName As String
    
    'フォルダの選択
    '複数のWord文書に連続して処理を施すマクロ http://stabucky.com/wp/archives/3004
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "フォルダを選択"
        .AllowMultiSelect = False
        If .Show = -1 Then
            PATH = .SelectedItems(1) & "\"
        Else
            Exit Sub
        End If
    End With

    '.doc、.docxなどのファイルを検索]
    '［VBA便利技］複数のOffice文書をまとめてPDF化
    ' http://itpro.nikkeibp.co.jp/article/COLUMN/20140325/545792/?ST=develop&P=1
    Dim file As String
    file = Dir(PATH & "*.doc")
    'すべての検索結果を取得するまでループを継続
    Do While file <> ""
        Dim tmp_path As String
        tmp_path = PATH & file
       
        'ファイルを開いて、PDF形式でエクスポート
        Documents.Open tmp_path
        
        'ファイル名を取得
        'Word文書をWordだけで1クリックでPDF化するマクロボタンを作成する
        'http://dzone.sakura.ne.jp/blog/2014/02/word-vba-wordword1pdf.html
        
        strAFN = ActiveDocument.Name
        Set FSO = CreateObject("Scripting.FileSystemObject")
        strAFN2 = FSO.GetBaseName(strAFN)
        strPDFName = PATH & strAFN2 & ".pdf"

        'PDF変換
        ActiveDocument.ExportAsFixedFormat _
            ExportFormat:=wdExportFormatPDF, _
            OutputFileName:=strPDFName
            
        '現在のファイルを閉じる
        ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
        
        '次に合致するファイルを取得
        file = Dir()
    Loop
End Sub

