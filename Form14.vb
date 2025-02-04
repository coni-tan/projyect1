Imports System.IO
Imports MySql.Data.MySqlClient
Imports 商品受発注システム.FormUtils
Public Class Form14
    ' FormUtilsをインスタンス化
    Private formUtils As New FormUtils()

    ' 商品一覧をデータベースから取得
    Private Sub LoadItemLists()
        Try
            ' クエリの生成
            Dim query As String = "SELECT item_id, item_name, item_keisai, item_price, item_image, item_description, category_id, item_constant_stock, item_create_date, item_update_date FROM item"

            ' データテーブルの作成
            Dim itemLists As New DataTable()

            ' データテーブル接続実行
            Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
                connection.Open()
                Using command As New MySqlCommand(query, connection)
                    Using adapter As New MySqlDataAdapter(command)
                        adapter.Fill(itemLists)
                    End Using
                End Using
            End Using

            ' 見出しを設定
            DataGridView1.DataSource = itemLists
            DataGridView1.Columns("item_id").HeaderText = "商品番号"
            DataGridView1.Columns("item_name").HeaderText = "商品名"
            DataGridView1.Columns("item_keisai").HeaderText = "掲載フラグ"
            DataGridView1.Columns("item_price").HeaderText = "販売価格"
            DataGridView1.Columns("item_price").DefaultCellStyle.Format = "N0" ' カンマ区切りで表示
            DataGridView1.Columns("item_price").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight ' 数字を右寄せに設定
            DataGridView1.Columns("item_image").HeaderText = "商品画像"
            DataGridView1.Columns("item_description").HeaderText = "商品説明"
            DataGridView1.Columns("category_id").HeaderText = "カテゴリコード"
            DataGridView1.Columns("item_constant_stock").HeaderText = "定数在庫"
            DataGridView1.Columns("item_create_date").HeaderText = "登録日"
            DataGridView1.Columns("item_update_date").HeaderText = "更新日"

            ' 全カラムを編集不可に設定
            For Each column As DataGridViewColumn In DataGridView1.Columns
                column.ReadOnly = True
            Next

            ' 不要レコード(追加用のブランクのレコード)の非表示
            DataGridView1.AllowUserToAddRows = False

            ' 行ヘッダーを表示し、見出しを「選択」に設定
            DataGridView1.RowHeadersVisible = True
            DataGridView1.TopLeftHeaderCell.Value = "選択"

        Catch ex As Exception
            MessageBox.Show($"商品リストの読み込み中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 画像選択ボタンの処理
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim targetDirectory As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "img")
        formUtils.CopyAndDisplayImage(OpenFileDialog1, PictureBox1, Label16, Label14, targetDirectory)
    End Sub

    'ロード時の処理
    Private Sub Form14_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' フルスクリーン設定
        formUtils.SetFullScreen(Me)

        ' combobox1にcategory表のcategory_idとcategory_nameを設定
        formUtils.PopulateComboBox(ComboBox1, "SELECT category_id, category_name FROM category", "category_id", "category_name")

        ' combobox3にpartner表のpartner_idとpartner_nameを設定
        formUtils.PopulateComboBox(ComboBox3, "SELECT partner_id, partner_name FROM partner", "partner_id", "partner_name")

        'DGVに商品一覧を表示
        LoadItemLists()
    End Sub

    ' 画面がリサイズされたとき
    Private Sub Form14_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        ' GroupBoxのリサイズおよび中央配置
        formUtils.ResizeAndCenterGroupBox(GroupBox1, Me)
    End Sub

    '閉じるボタンの処理
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' 現在のフォームを閉じる
        Me.Close()
    End Sub

    ' 商品削除ボタンの処理_削除は重大なのでトランザクションを入れる
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ' DataGridViewで選択された行を取得
        If DataGridView1.SelectedRows.Count > 0 Then
            ' 選択された行のitem_idを取得
            Dim itemId As Integer = Convert.ToInt32(DataGridView1.SelectedRows(0).Cells("item_id").Value)

            ' データベース接続文字列
            Dim connectionString As String = "Database=jyutyuu;Data Source=localhost;User Id=root"

            Using connection As New MySqlConnection(connectionString)
                connection.Open()

                ' トランザクションを宣言
                Dim transaction As MySqlTransaction = Nothing

                Try
                    ' トランザクションを開始
                    transaction = connection.BeginTransaction()

                    ' item表から該当するレコードを削除
                    Dim deleteItemQuery As String = "DELETE FROM item WHERE item_id = @itemId"
                    Using itemCommand As New MySqlCommand(deleteItemQuery, connection, transaction)
                        itemCommand.Parameters.AddWithValue("@itemId", itemId)
                        itemCommand.ExecuteNonQuery()
                    End Using

                    ' multi_partner表から該当するレコードを削除
                    Dim deleteMultiPartnerQuery As String = "DELETE FROM multi_partner WHERE item_id = @itemId"
                    Using multiPartnerCommand As New MySqlCommand(deleteMultiPartnerQuery, connection, transaction)
                        multiPartnerCommand.Parameters.AddWithValue("@itemId", itemId)
                        multiPartnerCommand.ExecuteNonQuery()
                    End Using

                    ' トランザクションをコミット
                    transaction.Commit()

                    MessageBox.Show("商品が正常に削除されました。", "削除完了", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    ' DataGridViewを再読み込み
                    LoadItemLists()
                Catch ex As Exception
                    ' エラーが発生した場合、トランザクションをロールバック
                    If Not IsNothing(transaction) Then
                        transaction.Rollback()
                    End If
                    MessageBox.Show($"商品削除中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End Using
        Else
            MessageBox.Show("削除する商品を選択してください。", "選択エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    ' 表示ボタンの処理
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' 選択された行の数をチェック
        If DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("表示する商品を選択してください", "選択エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        ElseIf DataGridView1.SelectedRows.Count > 1 Then
            MessageBox.Show("1商品のみ選択してください", "選択エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' 選択された商品レコードを取得
        Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

        ' 各コントロールに値を表示
        Label3.Text = selectedRow.Cells("item_id").Value.ToString()
        TextBox1.Text = selectedRow.Cells("item_name").Value.ToString()
        TextBox2.Text = selectedRow.Cells("item_description").Value.ToString()
        TextBox3.Text = selectedRow.Cells("item_price").Value.ToString()
        ComboBox2.SelectedItem = selectedRow.Cells("item_keisai").Value.ToString()
        ComboBox1.SelectedValue = selectedRow.Cells("category_id").Value
        TextBox7.Text = selectedRow.Cells("item_constant_stock").Value.ToString()

        ' item_idに基づいてmulti_partnerからpartner_idを取得
        Dim itemId As Integer = Convert.ToInt32(selectedRow.Cells("item_id").Value)
        Dim connectionString As String = "Database=jyutyuu;Data Source=localhost;User Id=root"
        Dim partnerId As String = ""

        Using connection As New MySqlConnection(connectionString)
            connection.Open()
            Dim query As String = "SELECT partner_id FROM multi_partner WHERE item_id = @itemId LIMIT 1"
            Using command As New MySqlCommand(query, connection)
                command.Parameters.AddWithValue("@itemId", itemId)
                Using reader As MySqlDataReader = command.ExecuteReader()
                    If reader.Read() Then
                        partnerId = reader("partner_id").ToString()
                    End If
                End Using
            End Using
        End Using

        ' ComboBox3の選択値を設定
        If Not String.IsNullOrEmpty(partnerId) Then
            ComboBox3.SelectedValue = partnerId
        End If

        ' 画像の表示
        Dim imageDirectory As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "img")
        Dim imagePath As String = Path.Combine(imageDirectory, selectedRow.Cells("item_image").Value.ToString())
        If File.Exists(imagePath) Then
            Using fileStream As New FileStream(imagePath, FileMode.Open, FileAccess.Read)
                Dim tempImage As Image = Image.FromStream(fileStream)
                PictureBox1.Image = New Bitmap(tempImage)
            End Using
            Label14.Text = selectedRow.Cells("item_image").Value.ToString()
            Label16.Text = imagePath
        Else
            PictureBox1.Image = Nothing
            Label14.Text = ""
            Label16.Text = ""
            MessageBox.Show("画像が見つかりませんでした。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

        ' ComboBox1とComboBox3の選択状態を確認
        For Each item As ComboBoxItem In ComboBox1.Items
            If item.Value = selectedRow.Cells("category_id").Value.ToString() Then
                ComboBox1.SelectedItem = item
                Exit For
            End If
        Next

        For Each item As ComboBoxItem In ComboBox3.Items
            If item.Value = partnerId Then
                ComboBox3.SelectedItem = item
                Exit For
            End If
        Next
    End Sub

    ' 更新ボタンの処理
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' 商品情報が表示されているかチェック
        If String.IsNullOrEmpty(Label3.Text) Then
            MessageBox.Show("更新内容がありません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

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

        ' データベース接続
        Dim connectionString As String = "Database=jyutyuu;Data Source=localhost;User Id=root"
        Using conn As New MySqlConnection(connectionString)
            conn.Open()

            ' item表の該当レコードを更新
            Dim updateItemQuery As String = "UPDATE item SET item_name = @item_name, item_keisai = @item_keisai, item_price = @item_price, item_image = @item_image, item_description = @item_description, category_id = @category_id, item_constant_stock = @item_constant_stock, item_update_date = @item_update_date WHERE item_id = @item_id"
            Using cmd As New MySqlCommand(updateItemQuery, conn)
                cmd.Parameters.AddWithValue("@item_id", Label3.Text)
                cmd.Parameters.AddWithValue("@item_name", TextBox1.Text)
                cmd.Parameters.AddWithValue("@item_keisai", ComboBox2.SelectedItem)
                cmd.Parameters.AddWithValue("@item_price", TextBox3.Text)
                cmd.Parameters.AddWithValue("@item_image", Label14.Text)
                cmd.Parameters.AddWithValue("@item_description", TextBox2.Text)
                cmd.Parameters.AddWithValue("@category_id", DirectCast(ComboBox1.SelectedItem, FormUtils.ComboBoxItem).Value)
                cmd.Parameters.AddWithValue("@item_constant_stock", If(TextBox7.Text = "", DBNull.Value, TextBox7.Text))
                cmd.Parameters.AddWithValue("@item_update_date", DateTime.Now)
                cmd.ExecuteNonQuery()
            End Using
        End Using

        MessageBox.Show("更新が完了しました。", "更新完了", MessageBoxButtons.OK, MessageBoxIcon.Information)

        ' DataGridViewを再読み込み
        LoadItemLists()
    End Sub




End Class