Imports System.IO
Imports MySql.Data.MySqlClient

Public Class Form13
    ' FormUtilsをインスタンス化
    Private formUtils As New FormUtils()



    ' 商品新規登録ボタンの処理
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' 入力内容のチェック
        If TextBox1.Text = "" OrElse TextBox1.Text.Length > 100 Then
            MessageBox.Show("商品名は100文字以内で入力してください。")
            Return
        End If

        If ComboBox2.SelectedIndex = -1 Then
            MessageBox.Show("掲載を選択してください。")
            Return
        End If

        If Not IsNumeric(TextBox3.Text) OrElse TextBox3.Text.Length > 7 Then
            MessageBox.Show("価格は7桁以内の半角数字で入力してください。")
            Return
        End If

        If ComboBox1.SelectedIndex = -1 Then
            MessageBox.Show("カテゴリーを選択してください。")
            Return
        End If

        If Not IsNumeric(TextBox7.Text) AndAlso TextBox7.Text <> "" Then
            MessageBox.Show("在庫数は最大3桁の半角数字で入力してください。")
            Return
        End If

        If ComboBox3.SelectedIndex = -1 Then
            MessageBox.Show("主要取引先を選択してください。")
            Return
        End If

        ' 新しいitem_idを生成
        Dim newItemId As String = GenerateNewItemId()

        ' データベース接続
        Dim connectionString As String = "Database=jyutyuu;Data Source=localhost;User Id=root"
        Using conn As New MySqlConnection(connectionString)
            conn.Open()

            ' item表に新規レコードを追加
            Dim itemQuery As String = "INSERT INTO item (item_id, item_name, item_keisai, item_price, item_image, item_description, category_id, item_constant_stock, item_create_date, item_update_date) " &
                                  "VALUES (@item_id, @item_name, @item_keisai, @item_price, @item_image, @item_description, @category_id, @item_constant_stock, @item_create_date, @item_update_date)"
            Using cmd As New MySqlCommand(itemQuery, conn)
                cmd.Parameters.AddWithValue("@item_id", newItemId)
                cmd.Parameters.AddWithValue("@item_name", TextBox1.Text)
                cmd.Parameters.AddWithValue("@item_keisai", ComboBox2.SelectedItem)
                cmd.Parameters.AddWithValue("@item_price", TextBox3.Text)
                cmd.Parameters.AddWithValue("@item_image", Label14.Text)
                cmd.Parameters.AddWithValue("@item_description", TextBox2.Text)
                cmd.Parameters.AddWithValue("@category_id", DirectCast(ComboBox1.SelectedItem, FormUtils.ComboBoxItem).Value)
                cmd.Parameters.AddWithValue("@item_constant_stock", If(TextBox7.Text = "", DBNull.Value, TextBox7.Text))
                cmd.Parameters.AddWithValue("@item_create_date", DateTime.Now)
                cmd.Parameters.AddWithValue("@item_update_date", DateTime.Now)
                cmd.ExecuteNonQuery()
            End Using

            ' multi_partner表に新規レコードを追加
            Dim partnerQuery As String = "INSERT INTO multi_partner (item_id, partner_id, ratio) " &
                                     "VALUES (@item_id, @partner_id, @ratio)"
            Using cmd As New MySqlCommand(partnerQuery, conn)
                cmd.Parameters.AddWithValue("@item_id", newItemId)
                cmd.Parameters.AddWithValue("@partner_id", DirectCast(ComboBox3.SelectedItem, FormUtils.ComboBoxItem).Value)
                cmd.Parameters.AddWithValue("@ratio", 1)
                cmd.ExecuteNonQuery()
            End Using
        End Using

        MessageBox.Show("新規登録が完了しました。")
    End Sub

    ' 商品番号自動採番
    Private Function GenerateNewItemId() As String
        Dim newItemId As String = ""
        Dim connectionString As String = "Database=jyutyuu;Data Source=localhost;User Id=root"
        Using conn As New MySqlConnection(connectionString)
            conn.Open()

            Dim query As String = "SELECT MAX(item_id) FROM item"
            Using cmd As New MySqlCommand(query, conn)
                Dim result As Object = cmd.ExecuteScalar()
                If result IsNot DBNull.Value Then
                    Dim maxId As Integer = Integer.Parse(result.ToString())
                    newItemId = (maxId + 1).ToString().PadLeft(8, "0"c)
                Else
                    newItemId = "00000001"
                End If
            End Using
        End Using

        Return newItemId
    End Function

    ' 画像選択ボタンの処理
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim targetDirectory As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "img")
        formUtils.CopyAndDisplayImage(OpenFileDialog1, PictureBox1, Label16, Label14, targetDirectory)
    End Sub



    'ロード時の処理
    Private Sub Form13_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' フルスクリーン設定
        formUtils.SetFullScreen(Me)

        ' combobox1にcategory表のcategory_idとcategory_nameを設定
        formUtils.PopulateComboBox(ComboBox1, "SELECT category_id, category_name FROM category", "category_id", "category_name")

        ' combobox3にpartner表のpartner_idとpartner_nameを設定
        formUtils.PopulateComboBox(ComboBox3, "SELECT partner_id, partner_name FROM partner", "partner_id", "partner_name")

    End Sub

    ' 画面がリサイズされたとき
    Private Sub Form13_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        ' GroupBoxのリサイズおよび中央配置
        formUtils.ResizeAndCenterGroupBox(GroupBox1, Me)
    End Sub

    '閉じるボタンの処理
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' 現在のフォームを閉じる
        Me.Close()
    End Sub
End Class
