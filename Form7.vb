Imports MySql.Data.MySqlClient

Public Class Form7
    Private memberId As String
    Private initialLoad As Boolean = True ' 初回ロードフラグ
    Private initialItemNameWidth As Integer ' 初期の商品名カラムの横幅

    ' 初期値としてmember_idを受け取るコンストラクタ
    Public Sub New(memberId As String)
        InitializeComponent()
        Me.memberId = memberId
    End Sub

    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadCartItems(memberId)
        CalculateTotal()
        initialLoad = False ' 初回ロード完了
    End Sub

    Private Sub LoadCartItems(memberId As String)
        Dim query As String = "SELECT cart_suffix, item_id, cart_item_quantity FROM cart WHERE cart_id = @member_id"
        Dim cartItems As New DataTable()

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            connection.Open()
            Using command As New MySqlCommand(query, connection)
                command.Parameters.AddWithValue("@member_id", memberId)
                Using adapter As New MySqlDataAdapter(command)
                    adapter.Fill(cartItems)
                End Using
            End Using
        End Using

        ' item_nameとitem_priceを取得してDataGridViewに表示
        cartItems.Columns.Add("item_name", GetType(String))
        cartItems.Columns.Add("item_price", GetType(Decimal))

        For Each row As DataRow In cartItems.Rows
            Dim itemId As String = row("item_id").ToString()
            Dim itemName As String = GetItemName(itemId)
            Dim itemPrice As Decimal = GetItemPrice(itemId)
            row("item_name") = itemName
            row("item_price") = itemPrice
        Next

        DataGridView1.DataSource = cartItems
        DataGridView1.Columns("cart_suffix").Visible = False
        DataGridView1.Columns("item_id").Visible = False

        ' 列の表示順と見出しを設定
        DataGridView1.Columns("item_name").DisplayIndex = 0
        DataGridView1.Columns("item_name").HeaderText = "商品名"
        DataGridView1.Columns("item_price").DisplayIndex = 1
        DataGridView1.Columns("item_price").HeaderText = "単価"
        DataGridView1.Columns("cart_item_quantity").DisplayIndex = 2
        DataGridView1.Columns("cart_item_quantity").HeaderText = "数量"

        ' item_nameとitem_priceを編集不可に設定
        DataGridView1.Columns("item_name").ReadOnly = True
        DataGridView1.Columns("item_price").ReadOnly = True

        ' 初回ロード時に商品名カラムの横幅を2倍に設定し、横幅を記憶
        If initialLoad Then
            initialItemNameWidth = DataGridView1.Columns("item_name").Width
            DataGridView1.Columns("item_name").Width = initialItemNameWidth * 2
        Else
            ' 初回ロード以降は記憶した横幅を設定
            DataGridView1.Columns("item_name").Width = initialItemNameWidth * 2
        End If

        ' 単価と数量を右寄せし、カンマ区切りで表示
        DataGridView1.Columns("item_price").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns("item_price").DefaultCellStyle.Format = "N0"
        DataGridView1.Columns("cart_item_quantity").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns("cart_item_quantity").DefaultCellStyle.Format = "N0"

        ' 削除ボタンの列が既に存在するか確認
        If Not DataGridView1.Columns.Contains("deleteButton") Then
            ' 削除ボタンの列を追加
            Dim deleteButtonColumn As New DataGridViewButtonColumn()
            deleteButtonColumn.Name = "deleteButton"
            deleteButtonColumn.HeaderText = "削除"
            deleteButtonColumn.Text = "削除"
            deleteButtonColumn.UseColumnTextForButtonValue = True
            deleteButtonColumn.DisplayIndex = 3
            deleteButtonColumn.Width = 40 ' 横幅を設定（ここで調整してください）
            DataGridView1.Columns.Add(deleteButtonColumn)
        End If
    End Sub

    Private Function GetItemName(itemId As String) As String
        Dim query As String = "SELECT item_name FROM item WHERE item_id = @item_id"
        Dim itemName As String = ""

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            connection.Open()
            Using command As New MySqlCommand(query, connection)
                command.Parameters.AddWithValue("@item_id", itemId)
                itemName = command.ExecuteScalar().ToString()
            End Using
        End Using

        Return itemName
    End Function

    Private Function GetItemPrice(itemId As String) As Decimal
        Dim query As String = "SELECT item_price FROM item WHERE item_id = @item_id"
        Dim itemPrice As Decimal = 0

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            connection.Open()
            Using command As New MySqlCommand(query, connection)
                command.Parameters.AddWithValue("@item_id", itemId)
                itemPrice = Convert.ToDecimal(command.ExecuteScalar())
            End Using
        End Using

        Return itemPrice
    End Function

    ' 合計金額を計算して表示するメソッド
    Private Sub CalculateTotal()
        Dim total As Decimal = 0

        For Each row As DataGridViewRow In DataGridView1.Rows
            Dim itemPrice As Decimal = Convert.ToDecimal(row.Cells("item_price").Value)
            Dim quantity As Integer = Convert.ToInt32(row.Cells("cart_item_quantity").Value)
            total += itemPrice * quantity
        Next

        Label1.Text = "合計金額: " & total.ToString("N0") & " 円"
    End Sub

    ' 数量の変更を反映するメソッド
    Private Sub UpdateQuantity(cartSuffix As Integer, newQuantity As Integer)
        Dim query As String = "UPDATE cart SET cart_item_quantity = @new_quantity WHERE cart_id = @member_id AND cart_suffix = @cart_suffix"

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            connection.Open()
            Using command As New MySqlCommand(query, connection)
                command.Parameters.AddWithValue("@new_quantity", newQuantity)
                command.Parameters.AddWithValue("@member_id", memberId)
                command.Parameters.AddWithValue("@cart_suffix", cartSuffix)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ' 商品をカートから削除するメソッド
    Private Sub RemoveItem(cartSuffix As Integer)
        Dim query As String = "DELETE FROM cart WHERE cart_id = @member_id AND cart_suffix = @cart_suffix"

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            connection.Open()
            Using command As New MySqlCommand(query, connection)
                command.Parameters.AddWithValue("@member_id", memberId)
                command.Parameters.AddWithValue("@cart_suffix", cartSuffix)
                command.ExecuteNonQuery()
            End Using
        End Using

        ' 最新のカート情報を再読み込み
        LoadCartItems(memberId)
        CalculateTotal()

        ' カートの連番を更新
        ReindexCart()
    End Sub

    ' カートの連番を更新するメソッド
    Private Sub ReindexCart()
        Dim query As String = "SELECT cart_suffix FROM cart WHERE cart_id = @member_id ORDER BY cart_suffix"
        Dim cartItems As New DataTable()

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            connection.Open()
            Using command As New MySqlCommand(query, connection)
                command.Parameters.AddWithValue("@member_id", memberId)
                Using adapter As New MySqlDataAdapter(command)
                    adapter.Fill(cartItems)
                End Using
            End Using
        End Using

        Dim count As Integer = 1
        For Each row As DataRow In cartItems.Rows
            Dim updateQuery As String = "UPDATE cart SET cart_suffix = @count WHERE cart_id = @member_id AND cart_suffix = @old_suffix"
            Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
                connection.Open()
                Using command As New MySqlCommand(updateQuery, connection)
                    command.Parameters.AddWithValue("@count", count)
                    command.Parameters.AddWithValue("@member_id", memberId)
                    command.Parameters.AddWithValue("@old_suffix", row("cart_suffix"))
                    command.ExecuteNonQuery()
                End Using
            End Using
            count += 1
        Next

        ' 最新のカート情報を再読み込み
        LoadCartItems(memberId)
        CalculateTotal()
    End Sub

    ' DataGridViewのセル内容クリックイベント
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = DataGridView1.Columns("deleteButton").Index AndAlso e.RowIndex >= 0 Then
            Dim cartSuffix As Integer = DataGridView1.Rows(e.RowIndex).Cells("cart_suffix").Value
            RemoveItem(cartSuffix)
        End If
    End Sub

    ' DataGridViewのセル編集完了イベント（数量変更）
    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        If e.ColumnIndex = DataGridView1.Columns("cart_item_quantity").Index Then
            Dim cartSuffix As Integer = DataGridView1.Rows(e.RowIndex).Cells("cart_suffix").Value
            Dim newQuantity As Integer = DataGridView1.Rows(e.RowIndex).Cells("cart_item_quantity").Value
            UpdateQuantity(cartSuffix, newQuantity)
            CalculateTotal()
        End If
    End Sub


    ' 再計算ボタンのクリックイベント
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        DataGridView1.DataSource = Nothing
        LoadCartItems(memberId)
        CalculateTotal()
    End Sub

    ' 買い物を続けるボタンのクリックイベント
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' フォームを閉じて前の画面に戻る
        Me.Close()
        Dim form5 As New Form5(memberId)
        form5.Show()
    End Sub

    ' カート内容を確定するボタンのクリックイベント
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If CheckStockAvailability() Then
            MessageBox.Show("在庫確認に問題はありません。次の画面へ進みます。")
            Dim form8 As New Form8(memberId)
            form8.Show()
            Me.Close()
        End If
    End Sub

    ' 在庫確認を行うメソッド
    Private Function CheckStockAvailability() As Boolean
        Dim query As String = "SELECT cart.item_id, cart.cart_item_quantity, item.item_name, 
                           COALESCE(SUM(COALESCE(stock.stock_in_val, 0)) - SUM(COALESCE(stock.stock_out_val, 0)), 0) AS stock_quantity
                           FROM cart
                           JOIN item ON cart.item_id = item.item_id
                           LEFT JOIN stock ON cart.item_id = stock.item_id
                           WHERE cart.cart_id = @member_id
                           GROUP BY cart.item_id, item.item_name, cart.cart_item_quantity"
        Dim cartItems As New DataTable()

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            connection.Open()
            Using command As New MySqlCommand(query, connection)
                command.Parameters.AddWithValue("@member_id", memberId)
                Using adapter As New MySqlDataAdapter(command)
                    adapter.Fill(cartItems)
                End Using
            End Using
        End Using

        If cartItems.Rows.Count = 0 Then
            MessageBox.Show("カートに商品がありません。")
            Return False
        End If

        For Each row As DataRow In cartItems.Rows
            Dim itemName As String = row("item_name").ToString()
            Dim cartQuantity As Integer = Convert.ToInt32(row("cart_item_quantity"))
            Dim stockQuantity As Integer = Convert.ToInt32(row("stock_quantity"))

            ' デバッグ用メッセージボックスを表示
            MessageBox.Show("商品名: " & itemName & vbCrLf & "カート数量: " & cartQuantity & vbCrLf & "在庫数量: " & stockQuantity)

            If cartQuantity > stockQuantity Then
                MessageBox.Show(itemName & "の在庫が不足しています。")
                Return False
            End If
        Next

        Return True
    End Function



End Class
