Public myPath As String

Sub 複数文書を単一PDF化()
' ファイルを選択し、サブ文書にして、それを解除。
' サブ文書にすると、段落番号が連番になるのを、セクションごとに振り直す
' できあがったものを、pdfにする
' 別途アップロードしている、「複数文書を単一化」　プログラムが必要
' これと、このpdf化を別モジュールにした場合は、上記の、Public宣言が必要です。
' (c) Makoto Tomo 2016.5.6
    
    Dim strPDFName As String
    
    Call 複数文書を単一化
    'ファイル名を年月日に設定する。読み込んだフォルダに保存する
    strPDFName = myPath & Format(Date, "yyyymmdd") & _
                Format(Time, "hhmm") & ".pdf"
    'Debug.Print strPDFName

    'PDF変換
    ActiveDocument.ExportAsFixedFormat _
        ExportFormat:=wdExportFormatPDF, _
        OutputFileName:=strPDFName
    
End Sub
