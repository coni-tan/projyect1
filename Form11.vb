Imports MySql.Data.MySqlClient

Public Class Form11
    Private orderId As String

    ' 前画面からorderIdを受け取るコンストラクタ
    Public Sub New(orderId As String)
        InitializeComponent()
        Me.orderId = orderId
    End Sub

    ' フォームロード時の処理
    Private Sub Form11_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 注文情報をロード
        LoadOrderDetails()
        ' 明細情報をロード
        LoadMeisaiDetails()

        ' 行ヘッダー(デフォルトの選択カラム)を非表示にする
        DataGridView1.RowHeadersVisible = False

        ' 新しい行(デフォルトの追加用のブランクレコード)の追加を無効にする
        DataGridView1.AllowUserToAddRows = False
    End Sub

    ' 注文情報をロードするメソッド
    Private Sub LoadOrderDetails()
        Dim query As String = "SELECT order_id, order_status, send_name, send_address, send_tell, member_id, member_name, member_address, member_tell, member_mail, payment, total_amount  
                               FROM orders 
                               WHERE order_id = @order_id"
        Dim orderDetails As New DataTable()

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            connection.Open()
            Using command As New MySqlCommand(query, connection)
                command.Parameters.AddWithValue("@order_id", orderId)
                Using adapter As New MySqlDataAdapter(command)
                    adapter.Fill(orderDetails)
                End Using
            End Using
        End Using

        ' 注文情報をラベルに表示
        If orderDetails.Rows.Count > 0 Then
            Dim row As DataRow = orderDetails.Rows(0)
            Label3.Text = row("order_id").ToString()
            Label7.Text = If(Convert.ToInt32(row("order_status")) = 0, "出荷待ち", "出荷済み")
            Label4.Text = row("send_name").ToString()
            Label9.Text = row("send_address").ToString()
            Label10.Text = row("send_tell").ToString()
            Label20.Text = row("member_id").ToString()
            Label16.Text = row("member_name").ToString()
            Label15.Text = row("member_address").ToString()
            Label13.Text = row("member_tell").ToString()
            Label22.Text = row("member_mail").ToString()
            Label11.Text = row("payment").ToString()
            Label26.Text = String.Format("{0:N0}", Convert.ToDecimal(row("total_amount")))
        End If
    End Sub

    ' 明細情報をロードするメソッド
    Private Sub LoadMeisaiDetails()
        Dim query As String = "SELECT meisai_suffix, item_name, item_price, item_quantity 
                               FROM meisai 
                               WHERE meisai_id = @order_id"
        Dim meisaiDetails As New DataTable()

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            connection.Open()
            Using command As New MySqlCommand(query, connection)
                command.Parameters.AddWithValue("@order_id", orderId)
                Using adapter As New MySqlDataAdapter(command)
                    adapter.Fill(meisaiDetails)
                End Using
            End Using
        End Using

        ' DataGridViewにデータをバインド
        DataGridView1.DataSource = meisaiDetails
        DataGridView1.Columns("meisai_suffix").HeaderText = "注文明細"
        DataGridView1.Columns("item_name").HeaderText = "商品名"
        DataGridView1.Columns("item_price").HeaderText = "販売価格"
        DataGridView1.Columns("item_price").DefaultCellStyle.Format = "N0"
        DataGridView1.Columns("item_quantity").HeaderText = "数量"
        DataGridView1.Columns("item_quantity").DefaultCellStyle.Format = "N0"

        ' 全カラムを編集不可に設定
        For Each column As DataGridViewColumn In DataGridView1.Columns
            column.ReadOnly = True
        Next
    End Sub

    ' 閉じる(button1)がクリックされたときの処理
    Private Sub button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form10.Show()
        ' 現在のフォームを閉じる
        Me.Close()
    End Sub

    ' 更新(button2)がクリックされたときの処理
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim selectedStatus As String = ComboBox1.SelectedItem?.ToString()

        If String.IsNullOrEmpty(selectedStatus) Then
            MessageBox.Show("注文ステータスが選択されていません")
            Return
        End If

        Dim orderStatus As Integer
        If selectedStatus = "出荷待ち" Then
            orderStatus = 0
        ElseIf selectedStatus = "出荷済み" Then
            orderStatus = 1
        Else
            MessageBox.Show("注文ステータスが無効です")
            Return
        End If

        Dim query As String = "UPDATE orders SET order_status = @order_status, order_update_date = @order_update_date WHERE order_id = @order_id"

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            connection.Open()
            Using command As New MySqlCommand(query, connection)
                command.Parameters.AddWithValue("@order_status", orderStatus)
                command.Parameters.AddWithValue("@order_update_date", DateTime.Now)
                command.Parameters.AddWithValue("@order_id", orderId)
                command.ExecuteNonQuery()
            End Using
        End Using

        MessageBox.Show("注文情報が更新されました")
        LoadOrderDetails()  ' 更新後のデータを再読み込み
    End Sub

End Class
