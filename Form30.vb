Imports MySql.Data.MySqlClient
Imports System.Text.RegularExpressions

Public Class Form30

    ' コンストラクタ
    Public Sub New()
        InitializeComponent()
    End Sub

    ' フォームロード時の処理
    Private Sub Form30_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 初期化処理等があればここに記載
    End Sub

    ' 検索ボタンButton1がクリックされたときの処理
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' textbox1の値を取得
        Dim itemId As String = TextBox1.Text
        ' データベースクエリの準備
        Dim query As String = "SELECT * FROM multi_partner WHERE item_id = @item_id"
        ' データテーブルの作成
        Dim dataTable As New DataTable()

        ' データベース接続
        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root;Convert Zero Datetime=True")
            Try
                connection.Open()
                ' クエリの実行準備
                Using command As New MySqlCommand(query, connection)
                    ' パラメータの設定
                    command.Parameters.AddWithValue("@item_id", itemId)
                    ' データアダプタの作成
                    Using adapter As New MySqlDataAdapter(command)
                        ' データテーブルにデータを埋め込む
                        adapter.Fill(dataTable)
                    End Using
                End Using
            Catch ex As Exception
                ' エラーメッセージの表示
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End Using

        ' DataGridView1にデータを表示
        DataGridView1.DataSource = dataTable

        ' 全カラムを編集不可に設定
        For Each column As DataGridViewColumn In DataGridView1.Columns
            column.ReadOnly = True
        Next

        ' 追加用レコードを非表示にする
        DataGridView1.AllowUserToAddRows = False

        ' 行ヘッダーを表示し、見出しを「選択」に設定
        DataGridView1.RowHeadersVisible = True
        DataGridView1.TopLeftHeaderCell.Value = "選択"
    End Sub

    ' 削除ボタンButton2がクリックされたときの処理
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' 行が選択されているか確認
        If DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("レコードを選択してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' 選択された行のitem_idを取得
        Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
        Dim itemId As String = selectedRow.Cells("item_id").Value.ToString()

        ' データベースクエリの準備
        Dim countQuery As String = "SELECT COUNT(*) FROM multi_partner WHERE item_id = @item_id"
        Dim deleteQuery As String = "DELETE FROM multi_partner WHERE item_id = @item_id AND partner_id = @partner_id"

        ' データベース接続
        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root;Convert Zero Datetime=True")
            Try
                connection.Open()

                ' item_idのレコード数を確認
                Using countCommand As New MySqlCommand(countQuery, connection)
                    countCommand.Parameters.AddWithValue("@item_id", itemId)
                    Dim recordCount As Integer = Convert.ToInt32(countCommand.ExecuteScalar())

                    ' レコードが1件のみの場合、削除を禁止
                    If recordCount <= 1 Then
                        MessageBox.Show("取引先を0件にすることはできません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End If
                End Using

                ' レコードの削除
                Using deleteCommand As New MySqlCommand(deleteQuery, connection)
                    deleteCommand.Parameters.AddWithValue("@item_id", itemId)
                    deleteCommand.Parameters.AddWithValue("@partner_id", selectedRow.Cells("partner_id").Value.ToString())
                    deleteCommand.ExecuteNonQuery()
                End Using

                ' 完了メッセージの表示
                MessageBox.Show("商品と取引先の紐づけを解除しました", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information)

                ' DataGridViewの再読み込み
                Button1_Click(sender, e)

            Catch ex As Exception
                ' エラーメッセージの表示
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End Using
    End Sub

    ' 追加・更新ボタンButton4がクリックされたときの処理
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' textbox1の値が空か確認
        If String.IsNullOrEmpty(TextBox1.Text) Then
            MessageBox.Show("まず商品番号を検索してください", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' 入力規則のチェック
        If Not Regex.IsMatch(TextBox2.Text, "^\d{2}$") Then
            MessageBox.Show("取引先コードは2桁の半角数字で入力してください", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If Not Regex.IsMatch(TextBox3.Text, "^\d{1,11}$") Then
            MessageBox.Show("発注割合は最大11桁の半角数字で入力してください", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' partner_idがpartner表に存在するか確認
        Dim partnerId As String = TextBox2.Text
        Dim checkPartnerQuery As String = "SELECT COUNT(*) FROM partner WHERE partner_id = @partner_id"
        Dim exists As Boolean

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root;Convert Zero Datetime=True")
            Try
                connection.Open()

                Using checkCommand As New MySqlCommand(checkPartnerQuery, connection)
                    checkCommand.Parameters.AddWithValue("@partner_id", partnerId)
                    exists = Convert.ToInt32(checkCommand.ExecuteScalar()) > 0
                End Using

                If Not exists Then
                    MessageBox.Show("取引先コードが存在しません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                ' multi_partnerにitem_idとpartner_idの組みが存在するか確認
                Dim itemId As String = TextBox1.Text
                Dim checkMultiPartnerQuery As String = "SELECT COUNT(*) FROM multi_partner WHERE item_id = @item_id AND partner_id = @partner_id"
                Dim isUpdating As Boolean

                Using checkMultiPartnerCommand As New MySqlCommand(checkMultiPartnerQuery, connection)
                    checkMultiPartnerCommand.Parameters.AddWithValue("@item_id", itemId)
                    checkMultiPartnerCommand.Parameters.AddWithValue("@partner_id", partnerId)
                    isUpdating = Convert.ToInt32(checkMultiPartnerCommand.ExecuteScalar()) > 0
                End Using

                ' 更新または新規作成の処理
                Dim saveQuery As String

                If isUpdating Then
                    saveQuery = "UPDATE multi_partner SET ratio = @ratio WHERE item_id = @item_id AND partner_id = @partner_id"
                Else
                    saveQuery = "INSERT INTO multi_partner (item_id, partner_id, ratio) VALUES (@item_id, @partner_id, @ratio)"
                End If

                Using saveCommand As New MySqlCommand(saveQuery, connection)
                    saveCommand.Parameters.AddWithValue("@item_id", itemId)
                    saveCommand.Parameters.AddWithValue("@partner_id", partnerId)
                    saveCommand.Parameters.AddWithValue("@ratio", TextBox3.Text)
                    saveCommand.ExecuteNonQuery()
                End Using

                ' 完了メッセージの表示
                MessageBox.Show("商品と取引先の紐づけを更新しました", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information)

                ' DataGridViewの再読み込み
                Button1_Click(sender, e)

            Catch ex As Exception
                ' エラーメッセージの表示
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End Using
    End Sub

End Class
