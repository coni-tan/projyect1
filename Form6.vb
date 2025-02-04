Imports MySql.Data.MySqlClient

Public Class Form6
    Private itemId As String
    Private memberId As String

    ' 初期値としてitem_idとmember_idを受け取るコンストラクタ
    Public Sub New(itemId As String, memberId As String)
        InitializeComponent()
        Me.itemId = itemId
        Me.memberId = memberId
    End Sub

    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadItemDetails(itemId)
    End Sub

    Private Sub LoadItemDetails(itemId As String)
        Dim itemQuery As String = "SELECT item_name, item_image, item_price, item_description FROM item WHERE item_id = @item_id"
        Dim stockQuery As String = "SELECT COALESCE(SUM(stock_in_val), 0) - COALESCE(SUM(stock_out_val), 0) AS stock_quantity FROM stock WHERE item_id = @item_id"
        ' COALESCE はNULLを指定の値に変換する関数、ここでは0

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            connection.Open()

            ' item情報を取得
            Using itemCommand As New MySqlCommand(itemQuery, connection)
                itemCommand.Parameters.AddWithValue("@item_id", itemId)
                Using itemReader As MySqlDataReader = itemCommand.ExecuteReader()
                    If itemReader.Read() Then
                        Label2.Text = itemReader("item_name").ToString()

                        ' item_priceをカンマ区切りにフォーマットして表示
                        Label4.Text = "販売価格： " & Convert.ToDecimal(itemReader("item_price")).ToString("N0") & " 円" ' 直接フォーマット適用

                        TextBox1.Text = itemReader("item_description").ToString()

                        Dim imagePath As String = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "img", itemReader("item_image").ToString())
                        If System.IO.File.Exists(imagePath) Then
                            PictureBox1.Image = Image.FromFile(imagePath)
                        Else
                            ' 商品画像が未登録の場合はimg99を表示
                            Dim fallbackImagePath As String = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "img", "img99.png")
                            If System.IO.File.Exists(fallbackImagePath) Then
                                PictureBox1.Image = Image.FromFile(fallbackImagePath)
                            End If
                        End If
                    End If
                End Using
            End Using

            ' stock表から在庫数計(in - out)
            Using stockCommand As New MySqlCommand(stockQuery, connection)
                stockCommand.Parameters.AddWithValue("@item_id", itemId)
                Dim stockQuantity As Object = stockCommand.ExecuteScalar()
                If stockQuantity IsNot Nothing Then
                    Label7.Text = Convert.ToDecimal(stockQuantity).ToString("N0") ' 直接フォーマット適用
                Else
                    Label7.Text = "0"
                End If
            End Using
        End Using
    End Sub

    ' 「カートへ入れる」ボタンが押されたときの処理
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '在庫ある場合のみカート画面へ遷移、在庫0の場合はそのまま
        If Label7.Text > 0 Then
            ' カートに追加する処理（ここでmember_idを利用）
            AddToCart(itemId, memberId)
            ' Form7を開く処理を追加
            Dim form7 As New Form7(memberId)
            form7.Show()
            Me.Hide()
        Else
            MessageBox.Show("在庫不足のためカートに追加できません。")
        End If
    End Sub

    Private Sub AddToCart(itemId As String, memberId As String)
        ' カートに商品を追加または更新する処理を実装
        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            connection.Open()

            ' 同じitem_idが存在するか確認
            Dim checkItemQuery As String = "SELECT cart_suffix, item_id FROM cart WHERE cart_id = @cart_id AND item_id = @item_id"
            Using checkItemCommand As New MySqlCommand(checkItemQuery, connection)
                checkItemCommand.Parameters.AddWithValue("@cart_id", memberId)
                checkItemCommand.Parameters.AddWithValue("@item_id", itemId)
                Dim reader As MySqlDataReader = checkItemCommand.ExecuteReader()

                If reader.Read() Then
                    ' item_idが存在する場合は数量を更新
                    Dim existingCartSuffix As Object = reader("cart_suffix")
                    reader.Close()
                    Dim updateQuantityQuery As String = "UPDATE cart SET cart_item_quantity = cart_item_quantity + 1 WHERE cart_id = @cart_id AND cart_suffix = @cart_suffix AND item_id = @item_id"
                    Using updateQuantityCommand As New MySqlCommand(updateQuantityQuery, connection)
                        updateQuantityCommand.Parameters.AddWithValue("@cart_id", memberId)
                        updateQuantityCommand.Parameters.AddWithValue("@cart_suffix", existingCartSuffix)
                        updateQuantityCommand.Parameters.AddWithValue("@item_id", itemId)
                        updateQuantityCommand.ExecuteNonQuery()
                    End Using
                    MessageBox.Show("商品がカートに追加されました。")
                Else
                    reader.Close()
                    ' 同一会員のカートに存在する商品の数を確認
                    Dim checkCartCountQuery As String = "SELECT COUNT(*) FROM cart WHERE cart_id = @cart_id"
                    Using checkCartCountCommand As New MySqlCommand(checkCartCountQuery, connection)
                        checkCartCountCommand.Parameters.AddWithValue("@cart_id", memberId)
                        Dim cartCount As Integer = Convert.ToInt32(checkCartCountCommand.ExecuteScalar())

                        If cartCount >= 9 Then
                            ' 商品の種類が9種類以上の場合、エラーメッセージを表示
                            MessageBox.Show("カートに入る商品は9商品までです。")
                        Else
                            ' item_idが存在しない場合は新しいレコードを挿入
                            Dim newCartSuffixQuery As String = "SELECT IFNULL(MAX(cart_suffix), 0) + 1 FROM cart WHERE cart_id = @cart_id"
                            Using newCartSuffixCommand As New MySqlCommand(newCartSuffixQuery, connection)
                                newCartSuffixCommand.Parameters.AddWithValue("@cart_id", memberId)
                                Dim newCartSuffix As Integer = Convert.ToInt32(newCartSuffixCommand.ExecuteScalar())

                                Dim insertQuery As String = "INSERT INTO cart (cart_id, cart_suffix, item_id, cart_item_quantity) VALUES (@cart_id, @cart_suffix, @item_id, 1)"
                                Using insertCommand As New MySqlCommand(insertQuery, connection)
                                    insertCommand.Parameters.AddWithValue("@cart_id", memberId)
                                    insertCommand.Parameters.AddWithValue("@cart_suffix", newCartSuffix)
                                    insertCommand.Parameters.AddWithValue("@item_id", itemId)
                                    insertCommand.ExecuteNonQuery()
                                End Using
                            End Using
                            MessageBox.Show("商品がカートに追加されました。")
                        End If
                    End Using
                End If
            End Using
        End Using
    End Sub
End Class
