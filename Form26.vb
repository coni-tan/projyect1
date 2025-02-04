Imports MySql.Data.MySqlClient

Public Class Form26
    Private memberId As String
    Private orderId As String

    ' memberIdとorderIdを受け取るコンストラクタ
    Public Sub New(memberId As String, orderId As String)
        InitializeComponent()
        Me.memberId = memberId
        Me.orderId = orderId
    End Sub

    ' フォームロード時の処理
    Private Sub Form26_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

        ' 注文情報をラベルに設定
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

    ' button1がクリックされたときの処理
    Private Sub button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Form25を開き、memberIdを渡す
        Dim form25 As New Form25(memberId)
        form25.Show()
        ' 現在のフォームを閉じる
        Me.Close()
    End Sub
End Class
