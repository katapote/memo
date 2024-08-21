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

## 応対コードテスト　DB
     
    

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

<br>

Microsoft AccessのDAO（Data Access Objects）は、VBA（Visual Basic for Applications）を使用してデータベースを操作するためのオブジェクトモデルです。DAOを使用すると、テーブル、クエリ、フィールド、レコードセット、データベース全体をプログラムで制御できます。

### 主なDAOオブジェクトとそのプロパティ、メソッド

#### 1. **Databaseオブジェクト**
   - **役割**: データベース全体を表します。
   - **主なプロパティ**:
     - `Name`: データベースの名前
     - `TableDefs`: データベース内のすべてのテーブルを表すコレクション
     - `QueryDefs`: データベース内のすべてのクエリを表すコレクション
   - **主なメソッド**:
     - `OpenRecordset`: クエリやテーブルを開く
     - `CreateTableDef`: 新しいテーブルを作成
     - `CreateQueryDef`: 新しいクエリを作成

#### 2. **TableDefオブジェクト**
   - **役割**: データベース内のテーブルを表します。
   - **主なプロパティ**:
     - `Name`: テーブルの名前
     - `Fields`: テーブル内のフィールドを表すコレクション
   - **主なメソッド**:
     - `CreateField`: 新しいフィールドを作成
     - `Append`: フィールドやインデックスをテーブルに追加

#### 3. **Recordsetオブジェクト**
   - **役割**: テーブルやクエリの結果を表します。
   - **主なプロパティ**:
     - `EOF`: レコードセットの最後かどうかを示すブール値
     - `Fields`: レコードのフィールドを表すコレクション
   - **主なメソッド**:
     - `MoveNext`: 次のレコードに移動
     - `AddNew`: 新しいレコードを追加
     - `Update`: 変更を保存

#### 4. **Fieldオブジェクト**
   - **役割**: テーブルやクエリ内の単一のフィールドを表します。
   - **主なプロパティ**:
     - `Name`: フィールドの名前
     - `Value`: フィールドの値
   - **主なメソッド**:
     - `AppendChunk`: バイナリデータをフィールドに追加
     - `GetChunk`: バイナリデータを取得

### DAOの基本的な使い方

以下に、DAOを使用してテーブル内のデータを読み取る簡単な例を示します。

    vba
    Sub ReadTableData()
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Dim fld As DAO.Field
    
        ' 現在のデータベースを開く
        Set db = CurrentDb()
        
        ' テーブルを開く
        Set rs = db.OpenRecordset("TableName")
    
        ' レコードセットを読み込み
        Do Until rs.EOF
            For Each fld In rs.Fields
                Debug.Print fld.Name & ": " & fld.Value
            Next fld
            rs.MoveNext
        Loop
        
        ' リソースを解放
        rs.Close
        Set rs = Nothing
        Set db = Nothing
    End Sub
    ```

### 参考: その他の主要オブジェクト
- **QueryDefオブジェクト**: クエリを定義・実行するためのオブジェクト。
- **Workspaceオブジェクト**: 複数のデータベース接続を管理。
- **Relationオブジェクト**: テーブル間のリレーションを定義。

DAOを使いこなすことで、Microsoft Access内のデータベース操作がプログラム的に柔軟に行えます。各オブジェクトやメソッドについて詳細を確認するには、AccessのVBAリファレンスを参照することをお勧めします。

Microsoft Accessでは、VBAを使って他のフォーム上のコントロール（テキストボックスなど）の値を取得し、それを変数に代入することができます。以下に、DAOを使用して同じデータベース内にある別のフォームの特定のテキストボックスの値を変数に代入するコードの例を示します。

### コード例

```vba
Sub GetTextBoxValueFromAnotherForm()
    Dim txtValue As String
    Dim frm As Form

    ' "YourFormName" を別のフォームの名前に置き換えます
    ' フォームが開かれている必要があります
    Set frm = Forms("YourFormName")
    
    ' "YourTextBoxName" をテキストボックスの名前に置き換えます
    txtValue = frm.Controls("YourTextBoxName").Value

    ' 変数に代入された値を使用します（例：メッセージボックスに表示）
    MsgBox "テキストボックスの値は: " & txtValue
    
    ' 使用が終わったらオブジェクトを解放
    Set frm = Nothing
End Sub
```

### 説明
1. **`Forms("YourFormName")`**:
   - `YourFormName` には、値を取得したいフォームの名前を入れます。このフォームは事前に開かれている必要があります。

2. **`frm.Controls("YourTextBoxName").Value`**:
   - `YourTextBoxName` には、値を取得したいテキストボックスの名前を入れます。この名前は、フォーム上のテキストボックスの「名前」プロパティに対応します。

3. **変数 `txtValue`**:
   - ここに、テキストボックスの値が代入されます。変数の型をテキストボックスのデータ型に合わせて変更してください（例：整数型なら `Dim txtValue As Integer`）。

4. **オブジェクトの解放**:
   - 処理が終わったら、オブジェクトを解放するために `Set frm = Nothing` を使用します。

### 使用例
このコードは、他のフォームのテキストボックスから値を取得し、何か別の処理を行いたいときに便利です。例えば、メインフォームから別のフォームにあるデータを参照したり、その値を計算に使用したりできます。