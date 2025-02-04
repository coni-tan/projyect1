Imports MySql.Data.MySqlClient
Public Class Form10
    ' Form25のコードをほぼ流用
    ' ロード時の処理
    Private Sub Form10_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 注文履歴のロード
        LoadOrderLists()
    End Sub

    ' 注文履歴をデータベースから取得
    Private Sub LoadOrderLists(Optional query As String = Nothing)
        'ロード時および検索条件（引数）がない場合はこちらのクエリ
        'Form25(ユーザー向け)と異なりmemberIdでの絞り込み不要なので全件表示
        If query Is Nothing Then
            query = "SELECT order_status, order_id, order_create_date, total_amount, member_name, send_name 
                 FROM orders"
        End If
        Dim orderLists As New DataTable()

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            connection.Open()
            Using command As New MySqlCommand(query, connection)
                Using adapter As New MySqlDataAdapter(command)
                    adapter.Fill(orderLists)
                End Using
            End Using
        End Using

        ' order_statusの値を文字に変換するため仮の列を追加
        orderLists.Columns.Add("order_status_text", GetType(String))
        For Each row As DataRow In orderLists.Rows
            Dim status As Integer = Convert.ToInt32(row("order_status"))
            row("order_status_text") = If(status = 0, "出荷待ち", "出荷済み")
        Next

        ' 見出しを設定
        DataGridView1.DataSource = orderLists
        DataGridView1.Columns("order_status").Visible = False
        DataGridView1.Columns("order_status_text").HeaderText = "注文ステータス"
        DataGridView1.Columns("order_id").HeaderText = "注文番号"
        DataGridView1.Columns("order_create_date").HeaderText = "注文日"
        DataGridView1.Columns("total_amount").HeaderText = "合計金額"
        DataGridView1.Columns("total_amount").DefaultCellStyle.Format = "N0" ' カンマ区切りで表示
        DataGridView1.Columns("total_amount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight ' 数字を右寄せに設定
        DataGridView1.Columns("member_name").HeaderText = "注文者氏名"
        DataGridView1.Columns("send_name").HeaderText = "お届け先氏名"

        ' 編集不可に設定
        DataGridView1.Columns("order_status_text").ReadOnly = True
        DataGridView1.Columns("order_id").ReadOnly = True
        DataGridView1.Columns("order_create_date").ReadOnly = True
        DataGridView1.Columns("total_amount").ReadOnly = True
        DataGridView1.Columns("member_name").ReadOnly = True
        DataGridView1.Columns("send_name").ReadOnly = True

        ' 不要レコード(追加用のブランクのレコード)の非表示
        DataGridView1.AllowUserToAddRows = False

        ' 行ヘッダーを表示し、見出しを「選択」に設定
        DataGridView1.RowHeadersVisible = True
        DataGridView1.TopLeftHeaderCell.Value = "選択"
    End Sub


    ' 検索ボタンがクリックされたときの処理
    ' Form25との違いはWHERE句をクエリに必ずいれるか、検索条件がある場合のみクエリに追加するかの違い
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim statusCondition As String = "" 'ComboBox1
        Dim idCondition As String = "" 'TextBox1

        ' ComboBox1の条件を取得　SQL文の追加部分の生成
        'すべてが選ばれてる場合は何もしない
        Select Case ComboBox1.SelectedItem?.ToString()
            Case "出荷待ち"
                statusCondition = " AND order_status = 0"
            Case "出荷済み"
                statusCondition = " AND order_status = 1"
        End Select

        ' TextBox1の条件を取得　SQL文の追加部分の生成
        'ブランクの場合は何もしない
        If Not String.IsNullOrEmpty(TextBox1.Text) Then
            idCondition = " AND order_id = '" & TextBox1.Text & "'"
        End If

        ' クエリを生成
        Dim query As String = "SELECT order_status, order_id, order_create_date, total_amount, member_name, send_name FROM orders"

        ' 状態またはIDの条件が指定されている場合にWHERE句を追加
        If Not String.IsNullOrEmpty(statusCondition & idCondition) Then
            query &= " WHERE 1=1" & statusCondition & idCondition
        End If

        ' データを再読み込みして表示、ロード時と異なり今度は引数あり
        LoadOrderLists(query)

        ' 結果が0件だったらメッセージを表示
        If DataGridView1.Rows.Count = 0 Then
            MessageBox.Show("検索結果なし", "検索結果", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    ' button2がクリックされたときの処理
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' 選択された行があるか確認
        If DataGridView1.SelectedRows.Count > 0 Then
            ' 選択された行のorder_idを取得
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
            Dim orderId As String = selectedRow.Cells("order_id").Value.ToString()

            ' Form11を開き、orderIdを渡す
            Dim form11 As New Form11(orderId)
            form11.Show()
        Else
            MessageBox.Show("注文を選択してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    '閉じるボタン
    Private Sub Button3_Click(sender As Object, e As EventArgs)

    End Sub
End Class
