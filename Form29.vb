Imports System.IO
Imports System.Text
Imports MySql.Data.MySqlClient

Public Class Form29
    ' フォームロード時の処理
    Private Sub Form29_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.AllowUserToAddRows = False ' 追加用レコードを非表示にする
        LoadStockData()
    End Sub

    ' stock表の内容をDataGridViewに表示
    Private Sub LoadStockData()
        Try
            ' クエリの生成
            Dim query As String = "SELECT * FROM stock"

            ' データテーブルの作成
            Dim stockData As New DataTable()

            ' データテーブル接続実行
            Dim connectionString As String = "Database=jyutyuu;Data Source=localhost;User Id=root;Convert Zero Datetime=True"
            Using connection As New MySqlConnection(connectionString)
                connection.Open()
                Using command As New MySqlCommand(query, connection)
                    Using adapter As New MySqlDataAdapter(command)
                        adapter.Fill(stockData)
                    End Using
                End Using
            End Using

            ' DataGridViewにデータをバインド
            DataGridView1.DataSource = stockData

            ' 全カラムを編集不可に設定
            For Each column As DataGridViewColumn In DataGridView1.Columns
                column.ReadOnly = True
            Next

            ' 行ヘッダーを非表示に設定
            DataGridView1.RowHeadersVisible = False

        Catch ex As Exception
            MessageBox.Show($"在庫データの読み込み中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 検索ボタンの処理
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim itemId As String = TextBox1.Text.Trim()

        Try
            ' クエリの生成
            Dim query As String
            If String.IsNullOrEmpty(itemId) Then
                query = "SELECT * FROM stock"
            Else
                query = "SELECT * FROM stock WHERE item_id = @item_id"
            End If

            ' データテーブルの作成
            Dim stockData As New DataTable()

            ' データテーブル接続実行
            Dim connectionString As String = "Database=jyutyuu;Data Source=localhost;User Id=root;Convert Zero Datetime=True"
            Using connection As New MySqlConnection(connectionString)
                connection.Open()
                Using command As New MySqlCommand(query, connection)
                    If Not String.IsNullOrEmpty(itemId) Then
                        command.Parameters.AddWithValue("@item_id", itemId)
                    End If
                    Using adapter As New MySqlDataAdapter(command)
                        adapter.Fill(stockData)
                    End Using
                End Using
            End Using

            ' レコードが存在しない場合
            If stockData.Rows.Count = 0 Then
                MessageBox.Show("表示するレコードがありません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' DataGridViewにデータをバインド
            DataGridView1.DataSource = stockData

            ' 全カラムを編集不可に設定
            For Each column As DataGridViewColumn In DataGridView1.Columns
                column.ReadOnly = True
            Next

            ' 追加用レコードを非表示にする
            DataGridView1.AllowUserToAddRows = False

            ' 行ヘッダーを非表示に設定
            DataGridView1.RowHeadersVisible = False

        Catch ex As Exception
            MessageBox.Show($"在庫データの検索中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    ' 発注ボタンの処理
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' DataGridViewに表示されているレコードを取得
        If DataGridView1.Rows.Count = 0 Then
            MessageBox.Show("定数在庫を下回る商品はありませんでした", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim connectionString As String = "Database=jyutyuu;Data Source=localhost;User Id=root"
        Dim orderMessages As New List(Of String)
        Dim orderCreated As Boolean = False

        ' データベース接続
        Using conn As New MySqlConnection(connectionString)
            conn.Open()

            For Each row As DataGridViewRow In DataGridView1.Rows
                ' 追加用レコードをスキップ
                If row.IsNewRow Then
                    Continue For
                End If

                Dim itemId As String = row.Cells("item_id").Value.ToString()
                Dim currentStock As Integer = GetCurrentStock(conn, itemId)
                Dim constantStock As Integer = GetConstantStock(conn, itemId)

                ' 発注すべき数を計算
                Dim orderQuantity As Integer = constantStock - currentStock

                ' 発注が必要な場合
                If orderQuantity > 0 Then
                    ' 既存の発注レコードを確認し、重複発注を防ぐ
                    If Not IsDuplicateOrder(conn, itemId, DateTime.Now.Date) Then
                        ' 発注レコードの作成
                        Dim partnerId As String = GetPartnerId(conn, itemId)
                        Dim orderSuffix As Integer = GetNextOrderSuffix(conn, itemId)
                        Dim stockInDate As DateTime = DateTime.Now

                        Dim insertQuery As String = "INSERT INTO stock (item_id, partner_id, stock_in_day, stock_in_val, stock_in_suf, create_date) " &
                                                    "VALUES (@item_id, @partner_id, @stock_in_day, @stock_in_val, @stock_in_suf, @create_date)"
                        Using cmd As New MySqlCommand(insertQuery, conn)
                            cmd.Parameters.AddWithValue("@item_id", itemId)
                            cmd.Parameters.AddWithValue("@partner_id", partnerId)
                            cmd.Parameters.AddWithValue("@stock_in_day", stockInDate)
                            cmd.Parameters.AddWithValue("@stock_in_val", orderQuantity)
                            cmd.Parameters.AddWithValue("@stock_in_suf", orderSuffix)
                            cmd.Parameters.AddWithValue("@create_date", stockInDate)
                            cmd.ExecuteNonQuery()
                        End Using

                        orderMessages.Add($"商品番号「{itemId}」を「{orderQuantity}個」発注しました。")
                        orderCreated = True
                    End If
                End If
            Next
        End Using

        If orderCreated Then
            MessageBox.Show(String.Join(Environment.NewLine, orderMessages), "発注完了", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("定数在庫を下回る商品はありませんでした", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    ' 現在の在庫数を取得
    Private Function GetCurrentStock(conn As MySqlConnection, itemId As String) As Integer
        Dim query As String = "SELECT SUM(stock_in_val) - SUM(stock_out_val) AS current_stock FROM stock WHERE item_id = @item_id"
        Using cmd As New MySqlCommand(query, conn)
            cmd.Parameters.AddWithValue("@item_id", itemId)
            Dim result As Object = cmd.ExecuteScalar()
            If result IsNot DBNull.Value Then
                Return Convert.ToInt32(result)
            End If
        End Using
        Return 0
    End Function

    ' 定数在庫を取得
    Private Function GetConstantStock(conn As MySqlConnection, itemId As String) As Integer
        Dim query As String = "SELECT item_constant_stock, category_id FROM item WHERE item_id = @item_id"
        Dim categoryId As String = ""
        Dim itemConstantStock As Integer? = Nothing

        Using cmd As New MySqlCommand(query, conn)
            cmd.Parameters.AddWithValue("@item_id", itemId)
            Using reader As MySqlDataReader = cmd.ExecuteReader()
                If reader.Read() Then
                    If Not IsDBNull(reader("item_constant_stock")) Then
                        itemConstantStock = Convert.ToInt32(reader("item_constant_stock"))
                    End If
                    categoryId = reader("category_id").ToString()
                End If
            End Using
        End Using

        If itemConstantStock.HasValue Then
            Return itemConstantStock.Value
        End If

        If String.IsNullOrEmpty(categoryId) Then
            Return 0
        End If

        query = "SELECT constant_stock FROM category WHERE category_id = @category_id"
        Using cmd As New MySqlCommand(query, conn)
            cmd.Parameters.AddWithValue("@category_id", categoryId)
            Dim result As Object = cmd.ExecuteScalar()
            If result IsNot DBNull.Value Then
                Return Convert.ToInt32(result)
            End If
        End Using

        Return 0
    End Function

    ' パートナーIDを取得
    Private Function GetPartnerId(conn As MySqlConnection, itemId As String) As String
        Dim query As String = "SELECT partner_id FROM multi_partner WHERE item_id = @item_id LIMIT 1"
        Using cmd As New MySqlCommand(query, conn)
            cmd.Parameters.AddWithValue("@item_id", itemId)
            Dim result As Object = cmd.ExecuteScalar()
            If result IsNot DBNull.Value Then
                Return result.ToString()
            End If
        End Using
        Return ""
    End Function

    ' 次の発注枝番を取得
    Private Function GetNextOrderSuffix(conn As MySqlConnection, itemId As String) As Integer
        Dim query As String = "SELECT MAX(stock_in_suf) FROM stock WHERE item_id = @item_id AND stock_in_day = @stock_in_day"
        Using cmd As New MySqlCommand(query, conn)
            cmd.Parameters.AddWithValue("@item_id", itemId)
            cmd.Parameters.AddWithValue("@stock_in_day", DateTime.Now.Date)
            Dim result As Object = cmd.ExecuteScalar()
            If result IsNot DBNull.Value Then
                Return Convert.ToInt32(result) + 1
            End If
        End Using
        Return 1
    End Function

    ' 重複発注を確認
    Private Function IsDuplicateOrder(conn As MySqlConnection, itemId As String, stockInDay As Date) As Boolean
        Dim query As String = "SELECT COUNT(*) FROM stock WHERE item_id = @item_id AND stock_in_day = @stock_in_day"
        Using cmd As New MySqlCommand(query, conn)
            cmd.Parameters.AddWithValue("@item_id", itemId)
            cmd.Parameters.AddWithValue("@stock_in_day", stockInDay)
            Dim result As Object = cmd.ExecuteScalar()
            If result IsNot DBNull.Value AndAlso Convert.ToInt32(result) > 0 Then
                Return True
            End If
        End Using
        Return False
    End Function

    ' エクスポートボタンの処理
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' SaveFileDialogを使用してファイル保存先を指定
        Using sfd As New SaveFileDialog()
            sfd.Filter = "CSVファイル (*.csv)|*.csv"
            sfd.Title = "CSVファイルにエクスポート"
            sfd.FileName = "在庫データ.csv"

            If sfd.ShowDialog() = DialogResult.OK Then
                ' CSVファイルにエクスポート
                ExportToCsv(DataGridView1, sfd.FileName)
            End If
        End Using
    End Sub

    ' DataGridViewの内容をCSVファイルにエクスポートするメソッド
    Private Sub ExportToCsv(dgv As DataGridView, filePath As String)
        Try
            Using writer As New StreamWriter(filePath, False, Encoding.UTF8)
                ' ヘッダーをエクスポート
                For Each column As DataGridViewColumn In dgv.Columns
                    writer.Write(column.HeaderText & ",")
                Next
                writer.WriteLine()

                ' 行をエクスポート
                For Each row As DataGridViewRow In dgv.Rows
                    ' 追加用レコードをスキップ
                    If row.IsNewRow Then
                        Continue For
                    End If
                    For Each cell As DataGridViewCell In row.Cells
                        writer.Write(cell.Value?.ToString() & ",")
                    Next
                    writer.WriteLine()
                Next
            End Using
            MessageBox.Show("CSVファイルへのエクスポートが完了しました。", "エクスポート完了", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show($"エクスポート中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class
