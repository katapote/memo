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

    Sub CheckAndAddRecord()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim tableName As String
    Dim fieldName As String
    Dim searchValue As Variant
    Dim found As Boolean
    
    ' テーブル名、フィールド名、検索する値を指定
    tableName = "YourTableName"  ' テーブル名
    fieldName = "YourFieldName"  ' フィールド名
    searchValue = "YourSearchValue"  ' 検索する値

    ' データベースを開く
    Set db = CurrentDb()
    
    ' テーブルのレコードセットを開く
    Set rs = db.OpenRecordset(tableName, dbOpenDynaset)
    
    ' 初期値
    found = False
    
    ' レコードセットのループ
    Do While Not rs.EOF
        If rs.Fields(fieldName).Value = searchValue Then
            found = True
            Exit Do
        End If
        rs.MoveNext
    Loop
    
    ' 検索値が見つからなかった場合、新しいレコードを追加
    If Not found Then
        rs.AddNew
        rs.Fields(fieldName).Value = searchValue
        ' 他のフィールドにも値を設定する場合はここに追加
        rs.Update
        MsgBox "レコードを追加しました: " & searchValue
    Else
        MsgBox "レコードはすでに存在します: " & searchValue
    End If
    
    ' リソースを解放
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    End Sub



