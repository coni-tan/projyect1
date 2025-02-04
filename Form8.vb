Imports MySql.Data.MySqlClient

Public Class Form8
    ' 初期値としてmember_idを受け取るコンストラクタ
    Private memberId As String
    Public Sub New(memberId As String)
        InitializeComponent()
        Me.memberId = memberId
    End Sub

    Private Sub Form8_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' お届け先情報のロード
        LoadDeliveryAddress()
        ' ご注文商品のロード
        LoadOrderedProducts()
        ' 合計金額の計算と表示
        CalculateTotalAmount()
    End Sub

    Private Sub LoadDeliveryAddress()
        ' お届け先情報をデータベースから取得
        Dim query As String = "SELECT member_name, member_address, member_tell FROM member WHERE member_id = @member_id"
        Dim deliveryInfo As New DataTable()

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            connection.Open()
            Using command As New MySqlCommand(query, connection)
                command.Parameters.AddWithValue("@member_id", memberId)
                Using adapter As New MySqlDataAdapter(command)
                    adapter.Fill(deliveryInfo)
                End Using
            End Using
        End Using

        If deliveryInfo.Rows.Count > 0 Then
            Dim row As DataRow = deliveryInfo.Rows(0)
            TextBox1.Text = row("member_name").ToString()
            TextBox2.Text = row("member_address").ToString()
            TextBox3.Text = row("member_tell").ToString()
        End If
    End Sub

    Private Sub LoadOrderedProducts()
        ' ご注文商品の情報をデータベースから取得
        Dim query As String = "SELECT item.item_id, item.item_name, item.item_price, cart.cart_item_quantity 
                               FROM cart 
                               JOIN item ON cart.item_id = item.item_id 
                               WHERE cart.cart_id = @member_id"
        Dim orderedProducts As New DataTable()

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            connection.Open()
            Using command As New MySqlCommand(query, connection)
                command.Parameters.AddWithValue("@member_id", memberId)
                Using adapter As New MySqlDataAdapter(command)
                    adapter.Fill(orderedProducts)
                End Using
            End Using
        End Using

        DataGridView1.DataSource = orderedProducts
        DataGridView1.Columns("item_name").HeaderText = "商品名"
        DataGridView1.Columns("item_price").HeaderText = "単価"
        DataGridView1.Columns("cart_item_quantity").HeaderText = "数量"

        ' item_nameとitem_price、cart_item_quantityを編集不可に設定
        DataGridView1.Columns("item_name").ReadOnly = True
        DataGridView1.Columns("item_price").ReadOnly = True
        DataGridView1.Columns("cart_item_quantity").ReadOnly = True

        ' 単価と数量を右寄せし、カンマ区切りで表示
        DataGridView1.Columns("item_price").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns("item_price").DefaultCellStyle.Format = "N0"
        DataGridView1.Columns("cart_item_quantity").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns("cart_item_quantity").DefaultCellStyle.Format = "N0"
    End Sub

    ' 合計金額の計算と表示
    Private Sub CalculateTotalAmount()
        Dim total As Decimal = 0

        For Each row As DataGridViewRow In DataGridView1.Rows
            Dim itemPrice As Decimal = Convert.ToDecimal(row.Cells("item_price").Value)
            Dim quantity As Integer = Convert.ToInt32(row.Cells("cart_item_quantity").Value)
            total += itemPrice * quantity
        Next

        Label12.Text = total.ToString("N0") & " 円"
    End Sub

    ' 在庫確認を行うメソッド(カート表、在庫表を比較)
    ' stock_quantity(現在庫)=累計入庫(stock_in_valのサマリ)-累計出庫(stock_out_valのサマリ)在庫
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

        For Each row As DataRow In cartItems.Rows
            Dim itemName As String = row("item_name").ToString()
            Dim cartQuantity As Integer = Convert.ToInt32(row("cart_item_quantity"))
            Dim stockQuantity As Integer = Convert.ToInt32(row("stock_quantity"))

            If cartQuantity > stockQuantity Then
                MessageBox.Show(itemName & "の在庫が不足しています。")
                Return False
            End If
        Next

        Return True
    End Function

    ' 注文番号採番メソッド（最新のorder_idを取得し、次のorder_idを生成）
    Private Function GetNextOrderId() As String
        Dim query As String = "SELECT MAX(order_id) FROM orders"
        Dim nextOrderId As String = "00000001"

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            connection.Open()
            Using command As New MySqlCommand(query, connection)
                Dim result As Object = command.ExecuteScalar()
                If result IsNot DBNull.Value Then
                    Dim maxOrderId As Integer = Convert.ToInt32(result)
                    nextOrderId = (maxOrderId + 1).ToString("D8")
                End If
            End Using
        End Using

        Return nextOrderId
    End Function

    ' 在庫をFIFO方式で更新するメソッド(引数あり：注文番号、商品番号、注文数)
    Private Sub UpdateStock(orderId As String, itemId As String, quantity As Integer)
        ' SQL文を生成（注文商品の在庫情報を、在庫表から入庫が古い順にソートし取得)
        Dim stockQuery As String = "SELECT partner_id, stock_in_day, stock_in_suf, stock_in_val, stock_out_val, reserve_flg
                                FROM stock 
                                WHERE item_id = @item_id AND reserve_flg = 0 
                                ORDER BY stock_in_day ASC, stock_in_suf ASC"
        Dim stockData As New DataTable()

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            connection.Open()
            Using command As New MySqlCommand(stockQuery, connection)
                command.Parameters.AddWithValue("@item_id", itemId)
                Using adapter As New MySqlDataAdapter(command)
                    adapter.Fill(stockData)
                End Using
            End Using

            Dim remainingQuantity As Integer = quantity
            Dim orderSuf As Integer = 1

            ' 取得した在庫情報を1行づつ引き当て判定
            For Each row As DataRow In stockData.Rows
                Dim partnerId As String = row("partner_id").ToString()
                Dim stockInDay As Date = Convert.ToDateTime(row("stock_in_day"))
                Dim stockInSuf As Integer = Convert.ToInt32(row("stock_in_suf"))
                Dim stockInVal As Integer = If(row("stock_in_val") Is DBNull.Value, 0, Convert.ToInt32(row("stock_in_val")))
                Dim stockOutVal As Integer = If(row("stock_out_val") Is DBNull.Value, 0, Convert.ToInt32(row("stock_out_val")))
                Dim availableQuantity As Integer = stockInVal - stockOutVal

                If availableQuantity > 0 Then
                    If availableQuantity >= remainingQuantity Then
                        ' 完全に引き当てる
                        Dim updateQuery As String = "UPDATE stock 
                                                 SET stock_out_val = stock_out_val + @quantity, 
                                                     stock_out_day = NOW(), 
                                                     reserve_flg = IF(stock_out_val + @quantity = stock_in_val, 1, 0)
                                                 WHERE item_id = @item_id AND partner_id = @partner_id AND stock_in_day = @stock_in_day AND stock_in_suf = @stock_in_suf"
                        Using updateCommand As New MySqlCommand(updateQuery, connection)
                            updateCommand.Parameters.AddWithValue("@quantity", remainingQuantity)
                            updateCommand.Parameters.AddWithValue("@partner_id", partnerId)
                            updateCommand.Parameters.AddWithValue("@stock_in_day", stockInDay)
                            updateCommand.Parameters.AddWithValue("@stock_in_suf", stockInSuf)
                            updateCommand.Parameters.AddWithValue("@item_id", itemId)
                            updateCommand.ExecuteNonQuery()
                        End Using

                        ' stock_out_valレコードを生成する
                        Dim insertQuery As String = "INSERT INTO stock (item_id, partner_id, stock_in_day, stock_out_day, stock_in_suf, stock_out_val, order_id, order_suf, reserve_flg, create_date) 
                                                 VALUES (@item_id, @partner_id, @stock_in_day, NOW(), @stock_in_suf, @quantity, @order_id, @order_suf, 0, NOW())"
                        Using insertCommand As New MySqlCommand(insertQuery, connection)
                            insertCommand.Parameters.AddWithValue("@item_id", itemId)
                            insertCommand.Parameters.AddWithValue("@partner_id", partnerId)
                            insertCommand.Parameters.AddWithValue("@stock_in_day", stockInDay)
                            insertCommand.Parameters.AddWithValue("@stock_in_suf", stockInSuf)
                            insertCommand.Parameters.AddWithValue("@quantity", remainingQuantity)
                            insertCommand.Parameters.AddWithValue("@order_id", orderId)
                            insertCommand.Parameters.AddWithValue("@order_suf", orderSuf)
                            insertCommand.ExecuteNonQuery()
                        End Using

                        remainingQuantity = 0
                        Exit For
                    Else
                        ' 一部を引き当てて次へ進む
                        Dim updateQuery As String = "UPDATE stock 
                                                 SET stock_out_val = stock_out_val + @quantity, 
                                                     stock_out_day = NOW(), 
                                                     reserve_flg = 1
                                                 WHERE item_id = @item_id AND partner_id = @partner_id AND stock_in_day = @stock_in_day AND stock_in_suf = @stock_in_suf"
                        Using updateCommand As New MySqlCommand(updateQuery, connection)
                            updateCommand.Parameters.AddWithValue("@quantity", availableQuantity)
                            updateCommand.Parameters.AddWithValue("@partner_id", partnerId)
                            updateCommand.Parameters.AddWithValue("@stock_in_day", stockInDay)
                            updateCommand.Parameters.AddWithValue("@stock_in_suf", stockInSuf)
                            updateCommand.Parameters.AddWithValue("@item_id", itemId)
                            updateCommand.ExecuteNonQuery()
                        End Using

                        ' stock_out_valレコードを生成する
                        Dim insertQuery As String = "INSERT INTO stock (item_id, partner_id, stock_in_day, stock_out_day, stock_in_suf, stock_out_val, order_id, order_suf, reserve_flg, create_date) 
                                                 VALUES (@item_id, @partner_id, @stock_in_day, NOW(), @stock_in_suf, @quantity, @order_id, @order_suf, 0, NOW())"
                        Using insertCommand As New MySqlCommand(insertQuery, connection)
                            insertCommand.Parameters.AddWithValue("@item_id", itemId)
                            insertCommand.Parameters.AddWithValue("@partner_id", partnerId)
                            insertCommand.Parameters.AddWithValue("@stock_in_day", stockInDay)
                            insertCommand.Parameters.AddWithValue("@stock_in_suf", stockInSuf)
                            insertCommand.Parameters.AddWithValue("@quantity", availableQuantity)
                            insertCommand.Parameters.AddWithValue("@order_id", orderId)
                            insertCommand.Parameters.AddWithValue("@order_suf", orderSuf)
                            insertCommand.ExecuteNonQuery()
                        End Using

                        remainingQuantity -= availableQuantity
                        orderSuf += 1
                    End If
                End If
            Next

            If remainingQuantity > 0 Then
                MessageBox.Show("在庫が不足しています。注文数量を満たすために追加の在庫が必要です。")
            End If
        End Using
    End Sub



    ' 注文を確定するボタンのクリックイベント
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If CheckStockAvailability() Then '在庫確認メソッドを呼び出しtrueなら進む
            Dim orderId As String = GetNextOrderId() '注文番号採番メソッドを呼び出す

            ' 商品ごとに在庫を引き当て(FIFOメソッドにカート表の注文番号、商品番号、注文数を引数で1行づつ渡し引き当て処理させる)
            For Each row As DataGridViewRow In DataGridView1.Rows
                Dim itemId As String = row.Cells("item_id").Value.ToString()
                Dim quantity As Integer = Convert.ToInt32(row.Cells("cart_item_quantity").Value)
                UpdateStock(orderId, itemId, quantity)
            Next

            ' member表から情報を取得
            Dim memberName As String = ""
            Dim memberAddress As String = ""
            Dim memberTell As String = ""
            Dim memberMail As String = ""

            Dim memberQuery As String = "SELECT member_name, member_address, member_tell, member_mail FROM member WHERE member_id = @member_id"
            Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
                connection.Open()
                Using memberCommand As New MySqlCommand(memberQuery, connection)
                    memberCommand.Parameters.AddWithValue("@member_id", memberId)
                    Using reader As MySqlDataReader = memberCommand.ExecuteReader()
                        If reader.Read() Then
                            memberName = reader("member_name").ToString()
                            memberAddress = reader("member_address").ToString()
                            memberTell = reader("member_tell").ToString()
                            memberMail = reader("member_mail").ToString()
                        End If
                    End Using
                End Using

                ' orders表とmeisai表に注文情報を挿入
                Dim totalAmount As Decimal = Convert.ToDecimal(Label12.Text.Replace(" 円", ""))
                Dim paymentMethod As String = ""

                ' 支払い方法を取得
                For Each control As Control In GroupBox2.Controls
                    If TypeOf control Is RadioButton AndAlso DirectCast(control, RadioButton).Checked Then
                        paymentMethod = control.Text
                        Exit For
                    End If
                Next

                Using transaction = connection.BeginTransaction()
                    Try
                        ' orders表に注文を挿入
                        Dim orderQuery As String = "INSERT INTO orders (order_id, order_status, member_id, member_name, member_address, member_tell, member_mail, send_name, send_address, send_tell, payment, total_amount, order_create_date, order_update_date)
                                                VALUES (@order_id, 0, @member_id, @member_name, @member_address, @member_tell, @member_mail, @send_name, @send_address, @send_tell, @payment, @total_amount, NOW(), NOW())"
                        Using command As New MySqlCommand(orderQuery, connection, transaction)
                            command.Parameters.AddWithValue("@order_id", orderId)
                            command.Parameters.AddWithValue("@member_id", memberId)
                            command.Parameters.AddWithValue("@member_name", memberName)
                            command.Parameters.AddWithValue("@member_address", memberAddress)
                            command.Parameters.AddWithValue("@member_tell", memberTell)
                            command.Parameters.AddWithValue("@member_mail", memberMail)
                            command.Parameters.AddWithValue("@send_name", TextBox1.Text)
                            command.Parameters.AddWithValue("@send_address", TextBox2.Text)
                            command.Parameters.AddWithValue("@send_tell", TextBox3.Text)
                            command.Parameters.AddWithValue("@payment", paymentMethod)
                            command.Parameters.AddWithValue("@total_amount", totalAmount)
                            command.ExecuteNonQuery()
                        End Using

                        ' 各商品をmeisai表に挿入
                        For Each row As DataGridViewRow In DataGridView1.Rows
                            Dim meisaiQuery As String = "INSERT INTO meisai (meisai_id, meisai_suffix, item_id, item_name, item_price, item_quantity)
                                                     VALUES (@meisai_id, @meisai_suffix, @item_id, @item_name, @item_price, @item_quantity)"
                            Using meisaiCommand As New MySqlCommand(meisaiQuery, connection, transaction)
                                meisaiCommand.Parameters.AddWithValue("@meisai_id", orderId)
                                meisaiCommand.Parameters.AddWithValue("@meisai_suffix", row.Index + 1)
                                meisaiCommand.Parameters.AddWithValue("@item_id", row.Cells("item_id").Value)
                                meisaiCommand.Parameters.AddWithValue("@item_name", row.Cells("item_name").Value)
                                meisaiCommand.Parameters.AddWithValue("@item_price", Convert.ToDecimal(row.Cells("item_price").Value))
                                meisaiCommand.Parameters.AddWithValue("@item_quantity", Convert.ToInt32(row.Cells("cart_item_quantity").Value))
                                meisaiCommand.ExecuteNonQuery()
                            End Using
                        Next

                        transaction.Commit()
                        MessageBox.Show("注文を確定しました。注文番号：" & orderId)
                        Me.Close()
                    Catch ex As Exception
                        transaction.Rollback()
                        MessageBox.Show("注文の確定中にエラーが発生しました: " & ex.Message)
                    End Try
                End Using
            End Using
        End If
    End Sub
End Class
