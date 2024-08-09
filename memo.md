## スクレイピングについて
1. ツール、参照設定から以下2項目にチェック
* Microsoft HTML Object Library
* Microsoft Internet Controls

2. コード入力<br>
vb
~~~
Sub test()<br>

    Dim ie As Object<br>
    Dim htmlDoc As Object<br>
    Dim htmlElement As Object<br>
    Dim i As Integer<br>
    Set ie = CreateObject(“InternetExplorer.Application”)<br>

    ‘ スクレイピングしたいウェブページを開く（例：Example.com）<br>
    ie.navigate “http://www.example.com”<br>
    ie.Visible = False<br>

    ‘ ページが完全に読み込まれるまで待機<br>
    Do While ie.readyState <> READYSTATE_COMPLETE<br>
        Application.Wait DateAdd(“s”, 1, Now)<br>
    Loop<br>

    ‘ HTMLドキュメントを取得<br>
    Set htmlDoc = ie.document  

    ‘ HTMLドキュメントから特定の要素を取得（例：タグ名が”h1″のもの）<br>
    Set htmlElement = htmlDoc.getElementsByTagName(“h1”) <br>

    ‘ 取得した要素をExcelシートに転記<br>
    For i = 0 To htmlElement.Length – 1  
        Sheets(“Sheet1”).Cells(i + 1, 1).Value = htmlElement.Item(i).innerText  
    Next i  

    ‘ IEを閉じる　　
    ie.Quit  
    Set ie = Nothing  

End Sub
~~~
## テストあり