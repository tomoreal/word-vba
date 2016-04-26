Public myPath As String

Sub サブ文書化して一挙にPDF化_選択ファイル()
'
' 別途アップロードしている、サブ文書化_選択ファイル　プログラムが必要
' これと、このpdf化を別モジュールにした場合は、上記の、Public宣言が必要です。

    Dim strPDFName As String
    
    Call サブ文書化_選択ファイル
    
    strPDFName = myPath & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".pdf"
    Debug.Print strPDFName

    'PDF変換
    ActiveDocument.ExportAsFixedFormat _
        ExportFormat:=wdExportFormatPDF, _
        OutputFileName:=strPDFName
    
End Sub
