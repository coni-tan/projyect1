Imports MySql.Data.MySqlClient
Imports System.Text.RegularExpressions

Public Class Form16
    Private memberId As String

    ' フォームロード時の処理
    Private Sub Form16_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 初期状態では何も表示しない
    End Sub

    ' 検索ボタンがクリックされた時の処理 (Button2)
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        memberId = TextBox9.Text.Trim()

        If String.IsNullOrEmpty(memberId) Then
            MessageBox.Show("会員番号を入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' 会員情報のロード
        LoadMemberInfo()
    End Sub

    ' 会員情報をデータベースから取得して表示するメソッド
    Private Sub LoadMemberInfo()
        Dim query As String = "SELECT member_id, member_name, member_address, member_tell, member_mail, member_status
                               FROM member 
                               WHERE member_id = @member_id"

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            Using command As New MySqlCommand(query, connection)
                command.Parameters.AddWithValue("@member_id", memberId)
                connection.Open()
                Using reader As MySqlDataReader = command.ExecuteReader()
                    If reader.Read() Then
                        ' 各コントロールに会員情報を表示
                        TextBox9.Text = reader("member_id").ToString()
                        TextBox1.Text = reader("member_name").ToString()
                        TextBox2.Text = reader("member_address").ToString()
                        TextBox3.Text = reader("member_tell").ToString()
                        TextBox4.Text = reader("member_mail").ToString()
                        ComboBox2.SelectedItem = If(reader("member_status").ToString() = "0", "通常", "退会")
                    Else
                        MessageBox.Show("会員情報が見つかりません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End Using
            End Using
        End Using
    End Sub

    ' 更新ボタンがクリックされた時の処理 (Button1)
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' 各テキストボックスから入力データを取得しトリムする
        Dim name As String = TextBox1.Text.Trim()
        Dim address As String = TextBox2.Text.Trim()
        Dim phone As String = TextBox3.Text.Trim()
        Dim email As String = TextBox4.Text.Trim()
        Dim status As Integer = If(ComboBox2.SelectedItem.ToString() = "通常", 0, 1)

        ' 入力チェック
        If String.IsNullOrEmpty(name) OrElse String.IsNullOrEmpty(address) OrElse
           String.IsNullOrEmpty(phone) OrElse String.IsNullOrEmpty(email) OrElse
           String.IsNullOrEmpty(status.ToString()) Then
            MessageBox.Show("全ての項目を入力してください。")
            Return
        End If

        If name.Length > 50 Then
            MessageBox.Show("氏名は50文字以内で入力してください。")
            Return
        End If

        If address.Length > 100 Then
            MessageBox.Show("住所は100文字以内で入力してください。")
            Return
        End If

        If Not Regex.IsMatch(phone, "^\d{1,11}$") Then
            MessageBox.Show("電話番号は半角数字11文字以内で入力してください。")
            Return
        End If

        If Not Regex.IsMatch(email, "^[\w\.\-]+@([\w\-]+\.)+[\w\-]{2,4}$") OrElse email.Length > 50 Then
            MessageBox.Show("メールは正しい形式で50文字以内で入力してください。")
            Return
        End If

        ' メール重複チェック
        Dim checkEmailQuery As String = "SELECT COUNT(*) FROM member WHERE member_mail = @Email AND member_id <> @member_id"
        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            Using checkEmailCmd As New MySqlCommand(checkEmailQuery, connection)
                checkEmailCmd.Parameters.AddWithValue("@Email", email)
                checkEmailCmd.Parameters.AddWithValue("@member_id", memberId)
                connection.Open()
                Dim count As Integer = Convert.ToInt32(checkEmailCmd.ExecuteScalar())
                If count > 0 Then
                    MessageBox.Show("このメールアドレスは既に登録されています。")
                    Return
                End If
            End Using
        End Using

        ' 更新処理
        Dim query As String = "UPDATE member 
                               SET member_name = @Name, member_address = @Address, member_tell = @Phone, 
                                   member_mail = @Email, member_status = @Status
                               WHERE member_id = @member_id"

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            Using command As New MySqlCommand(query, connection)
                command.Parameters.AddWithValue("@member_id", memberId)
                command.Parameters.AddWithValue("@Name", name)
                command.Parameters.AddWithValue("@Address", address)
                command.Parameters.AddWithValue("@Phone", phone)
                command.Parameters.AddWithValue("@Email", email)
                command.Parameters.AddWithValue("@Status", status)
                connection.Open()
                Dim rowsAffected As Integer = command.ExecuteNonQuery()

                If rowsAffected > 0 Then
                    MessageBox.Show("会員情報を更新しました。", "更新成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("更新に失敗しました。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End Using
        End Using
    End Sub

    ' 削除ボタンがクリックされた時の処理 (Button3)
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' 未出荷の注文がないか確認
        Dim checkOrderQuery As String = "SELECT COUNT(*) FROM orders WHERE member_id = @member_id AND order_status = 0"
        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            Using checkOrderCmd As New MySqlCommand(checkOrderQuery, connection)
                checkOrderCmd.Parameters.AddWithValue("@member_id", memberId)
                connection.Open()
                Dim count As Integer = Convert.ToInt32(checkOrderCmd.ExecuteScalar())
                If count > 0 Then
                    MessageBox.Show("未処理の注文が存在する会員情報は削除できません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            End Using
        End Using

        Dim confirmResult As DialogResult = MessageBox.Show("本当に会員情報を削除しますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If confirmResult = DialogResult.No Then
            Return
        End If

        ' 削除処理
        Dim query As String = "DELETE FROM member WHERE member_id = @member_id"

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            Using command As New MySqlCommand(query, connection)
                command.Parameters.AddWithValue("@member_id", memberId)
                connection.Open()
                Dim rowsAffected As Integer = command.ExecuteNonQuery()

                If rowsAffected > 0 Then
                    MessageBox.Show("会員情報を削除しました。", "削除成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    ' 削除後にフィールドをクリア
                    TextBox9.Text = ""
                    TextBox1.Text = ""
                    TextBox2.Text = ""
                    TextBox3.Text = ""
                    TextBox4.Text = ""
                    ComboBox2.SelectedIndex = -1
                Else
                    MessageBox.Show("削除に失敗しました。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End Using
        End Using
    End Sub

End Class
