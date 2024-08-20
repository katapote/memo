## スクレイピングについて
1. ツール、参照設定から以下2項目にチェック
* Microsoft HTML Object Library
* Microsoft Internet Controls

2. コード入力<br>
vb

サンプル

    Sub test

        Dim ie As Object
        Dim htmlDoc As Object
        Dim htmlElement As Object
        Dim i As Integer
        Set ie = CreateObject(“InternetExplorer.Application”)
    
        ‘ スクレイピングしたいウェブページを開く（例：Example.com）
        ie.navigate “http://www.example.com”
        ie.Visible = False
    
        ‘ ページが完全に読み込まれるまで待機
        Do While ie.readyState <> READYSTATE_COMPLETE
            Application.Wait DateAdd(“s”, 1, Now)
        Loop
    
        ‘ HTMLドキュメントを取得
        Set htmlDoc = ie.document  
    
        ‘ HTMLドキュメントから特定の要素を取得（例：タグ名が”h1″のもの）
        Set htmlElement = htmlDoc.getElementsByTagName(“h1”) 
    
        ‘ 取得した要素をExcelシートに転記
        For i = 0 To htmlElement.Length – 1  
            Sheets(“Sheet1”).Cells(i + 1, 1).Value = htmlElement.Item(i).innerText  
        Next i  
    
        ‘ IEを閉じる　　
        ie.Quit  
        Set ie = Nothing  
    
    End sub

## 応対コードテスト　DB設計<br>

<table>
<caption>問題テーブル
<thead>
<tr><th>問題ID<th>問題文<th>解答<th>解説<th>URL</th>


<table>
<caption>回答テーブル
<thead>
<tr><th>回答者ID<th>問題ID<th>正誤フラグ<th>回答<th>出題順<th>作成日時</th>
<tbody>
<tr><td>スタート時、<br>出題数分の<br>同一レコードを作成<td>問題TBから<br>出題数分の<br>重複しない問題IDを取得<td>0:正 1:誤<br>すべて<br>1で作成<br>正解時<br>0にする<td>スタート時は<br>null<br>各回答<br>確定時<br>回答内容で<br>更新<td>スタート時<br>出題数分の<br>ランダムな<br>数字<td>スタート時刻</td>


    Option Compare Database
    Option Explicit
     
    Private Sub cmd_start_Click()
        
        'On Error GoTo err1
        
        Dim db As DAO.Database
        Set db = CurrentDb
        Dim sql As String
        Dim q As Recordset
        Dim strName As String
        Dim L1 As Integer
        
        
        sql = "SELECT count(*) FROM 解答者 WHERE 回答者名 = '" & txt_Name & "'"
        
        Set q = db.OpenRecordset(sql)
        'L1 = q!A
        If IsNull(q) Then
            MsgBox "null"
            Exit Sub
        End If
        
        'DoCmd.RunSQL "INSERT INTO 解答者 (回答者名) values('" & txt_Name & "')"
        
        db.Close
        Set q = Nothing
        Set db = Nothing
        Exit Sub
        
    err1:
     
        MsgBox ("エラー：" & Err.Number)
     
 
    End Sub

https://qiita.com/chida09/items/d4b33a28b918958f267f

SELECT（どのカラムから？）
FROM（どのテーブルから？）
WHERE（特定のデータを取得）

<br>

## DAO<br>

    dim db as DAO.DataBase
    set db =current database
    今開いてるデータベースをセット！
    
    dim rs as DAO.recordset
    set rs = db.recordset(セットしたいテーブル名とかクエリの結果)
    
    dim fld as DAO.field
    set fld =rs.fields(セットしたいフィールド名)
    
    do until rs.Eof=true←見てるレコードが1番下になるまでつづけるよ
    
        rs.Edit ←レコードの内容を書き換えます
        fld.value=1 
        rs.Update
        rs.movenext←つぎのレコードにうつるよ
    
    loop


## 仕様書：応対コード

フロント（ユーザー）	
バック
<br>
フォームを起動	
名前を入力	
入力された名前を変数に代入
回答者TBから名前を検索
レコード有　
有:"名前　Ｐ番"で続けます　　　　　　　　　　
>OK  >>スタート画面に進む	
回答者IDで履歴TBを検索　
実績を表示　>>スタート画面に進む
>新しいアカウント作成　　>>スタート画面に進む	
回答者TBに新規レコードを作成　>>スタート画面に進む
>キャンセル　>>最初に戻る	
無:新しいアカウントを作成	
レコード無
>OK  >>スタート画面に進む	
回答者TBに新規レコードを作成　>>スタート画面に進む
>キャンセル　>>最初に戻る	
スタート画面	
チュートリアルを選択    ＊後に他のモードを作成する予定	
スタートボタン押下	
問題TBから複数レコードの問題IDを抽出　　　　　　　
回答レコードに挿入　　　　　　　　　　　　　　
出題順フィールドに、rnd関数で1〜出題数までの値を割り振る

問題画面	現在時刻を変数に代入　　　　　　　　
回答レコードの出題順フィールドの値が若いものから抽出
回答を入力	
回答ボタンを押下	
回答を参照レコードの回答フィールドに更新　　　　　　　　
回答を参照レコードの解答フィールドと照らし合わせる　　　
正解！　　　　　　　　　　
あなたの回答　解答　解説　URLなどを表示　　　　	
一致:参照レコードにタイムスタンプを更新　　　　　　　
不正解　　　　　　　　　
あなたの回答　解答　　　　
解説　URLなどを表示　　　　　　　	
不一致:参照レコードにタイムスタンプ　　
正誤フラグを誤である0に更新
次の問題ボタンを押下　　　　
>>問題画面	
検索条件を、出題フィールドの数値+1にして問題画面に戻る。繰り返し
最後の問題の解説画面　　　
終了ボタンを押下	
リザルト画面　　　　　　　
今回解いた問題　　　　　　
あなたの回答　　　　　　
解答　　　　　　　　　　　
正答率　　　　　　　　　
タイムスコア　　　　　　　
などを表示	開始時刻〜現在時刻　かつ　　　　　　　　
回答者IDが一致するレコードを参照　問題、回答、解答を抽出　　　　　　　　
上記のレコードを履歴tbに挿入　　　　　　
レコードを削除





