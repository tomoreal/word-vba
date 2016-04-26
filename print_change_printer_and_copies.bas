Sub プリンタを切替、かつ印刷部数を指定して、頁単位印刷()
' プリンタを切替、かつダイアログボックスで印刷部数を指定して、頁単位印刷
' 印刷したいプリンターを、あらかじめ、myPrinterで指定して下さい。

    Dim Message As String, Title As String
    Dim Default As Integer, MyNum As Integer
    
    Message = "印刷部数を入力してください。"    ' 入力を求めるメッセージを設定します。
    Title = "印刷部数指定"                ' タイトルを設定します。
    Default = "1"                            ' 既定値を設定します。
    ' メッセージ、タイトル、既定値を表示します。
    MyNum = InputBox(Message, Title, Default)
    Debug.Print MyNum

    Dim myPrinter As String
    myPrinter = "NEC MultiWriter 5750C" ' ★　←要設定。ここで、プリンタ名をあらかじめ指定したプリンターに印刷
    Dim defaultPrinter As String
    
    defaultPrinter = Application.ActivePrinter  '現在のプリンタ名をセーブ
    Debug.Print defaultPrinter

    ActivePrinter = myPrinter
    Debug.Print myPrinter
    
    Application.PrintOut FileName:="", Range:=wdPrintAllDocument, Item:= _
        wdPrintDocumentWithMarkup, Copies:=MyNum, Pages:="", PageType:= _
        wdPrintAllPages, Collate:=False, Background:=True, PrintToFile:=False, _
        PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
        PrintZoomPaperHeight:=0
    
    Debug.Print tmp_path
    
    ' もとのデフォルトプリンターに、出力先を戻す。
    Application.ActivePrinter = defaultPrinter

End Sub
