Imports MySql.Data.MySqlClient
Imports System.Text.RegularExpressions　' Regexメソッドの正規表現チェックに使用する

Public Class Form3
    Private Sub button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' 各テキストボックスから入力データを取得しトリムする
        Dim name As String = TextBox1.Text.Trim()
        Dim address As String = TextBox2.Text.Trim()
        Dim phone As String = TextBox3.Text.Trim()
        Dim email As String = TextBox4.Text.Trim()
        Dim password As String = TextBox5.Text.Trim()

        ' 入力チェック
        If String.IsNullOrEmpty(name) OrElse String.IsNullOrEmpty(address) OrElse
           String.IsNullOrEmpty(phone) OrElse String.IsNullOrEmpty(email) OrElse
           String.IsNullOrEmpty(password) Then
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

        If Not Regex.IsMatch(password, "^[a-zA-Z0-9!@#\$%\^&\*\(\)\-_+={}\[\]:;""'<>?,.\/\\]{1,10}$") OrElse password.Length > 10 Then
            MessageBox.Show("パスワードは半角英数記号10文字以内で入力してください。")
            Return
        End If

        ' MySQL接続情報
        Dim connectionString As String = "Database=jyutyuu;Data Source=localhost;User Id=root;"
        Dim newId As Integer ' newId を外で宣言
        Dim memberId As String ' memberId をUsingの外で宣言
        Using connection As New MySqlConnection(connectionString)
            connection.Open()

            ' メール重複チェック
            Dim checkEmailQuery As String = "SELECT COUNT(*) FROM member WHERE member_mail = @Email"
            Using checkEmailCmd As New MySqlCommand(checkEmailQuery, connection)
                checkEmailCmd.Parameters.AddWithValue("@Email", email)
                Dim count As Integer = Convert.ToInt32(checkEmailCmd.ExecuteScalar())
                If count > 0 Then
                    MessageBox.Show("このメールアドレスは既に登録があります。")
                    Return
                End If
            End Using

            ' 新しいmember_idの取得
            Dim getIdQuery As String = "SELECT COALESCE(MAX(member_id), 0) + 1 FROM member"
            Using getIdCmd As New MySqlCommand(getIdQuery, connection)
                newId = Convert.ToInt32(getIdCmd.ExecuteScalar())
                ' 8桁のゼロ埋め形式に変換
                memberId = newId.ToString("D8")
                ' デバッグメッセージを追加して memberId の値を確認
                MessageBox.Show("取得した member_id: " & memberId)
            End Using

            ' 新規レコードの挿入
            Dim insertQuery As String = "INSERT INTO member (member_id, member_name, member_address, member_tell, member_mail, member_pass, member_status, member_date) " &
                                        "VALUES (@Id, @Name, @Address, @Phone, @Email, @Password, 0, @Date)"
            Using insertCmd As New MySqlCommand(insertQuery, connection)
                insertCmd.Parameters.AddWithValue("@Id", memberId)
                insertCmd.Parameters.AddWithValue("@Name", name)
                insertCmd.Parameters.AddWithValue("@Address", address)
                insertCmd.Parameters.AddWithValue("@Phone", phone)
                insertCmd.Parameters.AddWithValue("@Email", email)
                insertCmd.Parameters.AddWithValue("@Password", password)
                insertCmd.Parameters.AddWithValue("@Date", DateTime.Now)
                insertCmd.ExecuteNonQuery()
            End Using
        End Using

        ' 登録完了メッセージ
        MessageBox.Show($"会員登録ができました。あなたの会員番号は「{memberId}」です。OKボタンを押しログイン画面へお進みください。", "登録完了", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Form2.Show()

        ' 現在のフォームを閉じる
        'Me.Close()　’自身を閉じるとデバックモードが爆速終了してform2の展開を確認できないので開発中はコメントアウトしておく

    End Sub
End Class
