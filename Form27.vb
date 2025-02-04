Imports MySql.Data.MySqlClient
Imports System.Text.RegularExpressions　'正規表現(特定パターンの文字列に一致しているか)

Public Class Form27
    Private memberId As String

    ' 開発用コンストラクタ(会員ログイン省略)
    Public Sub New()
        InitializeComponent()
        ' 仮のmember_idを設定
        Me.memberId = "00000001"
    End Sub

    ' 本番用コンストラクタ(前画面から会員IDを引き継ぐ)
    Public Sub New(memberId As String)
        InitializeComponent()
        Me.memberId = memberId
    End Sub

    ' フォームロード時の処理
    Private Sub Form27_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 会員情報のロード
        LoadMemberInfo()
    End Sub

    ' 会員情報をデータベースから取得して表示するメソッド
    Private Sub LoadMemberInfo()
        Dim query As String = "SELECT member_id, member_name, member_address, member_tell, member_mail, member_pass 
                               FROM member 
                               WHERE member_id = @member_id"

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            Using command As New MySqlCommand(query, connection)
                command.Parameters.AddWithValue("@member_id", memberId)
                connection.Open()
                Using reader As MySqlDataReader = command.ExecuteReader()
                    If reader.Read() Then
                        ' 各コントロールに会員情報を表示
                        Label20.Text = reader("member_id").ToString()
                        TextBox1.Text = reader("member_name").ToString()
                        TextBox2.Text = reader("member_address").ToString()
                        TextBox3.Text = reader("member_tell").ToString()
                        TextBox4.Text = reader("member_mail").ToString()
                        TextBox5.Text = reader("member_pass").ToString()
                    Else
                        MessageBox.Show("会員情報が見つかりません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End Using
            End Using
        End Using
    End Sub

    ' 更新ボタンがクリックされた時の処理
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
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
                                   member_mail = @Email, member_pass = @Password 
                               WHERE member_id = @member_id"

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            Using command As New MySqlCommand(query, connection)
                command.Parameters.AddWithValue("@member_id", memberId)
                command.Parameters.AddWithValue("@Name", name)
                command.Parameters.AddWithValue("@Address", address)
                command.Parameters.AddWithValue("@Phone", phone)
                command.Parameters.AddWithValue("@Email", email)
                command.Parameters.AddWithValue("@Password", password)
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

    ' 閉じるボタンがクリックされた時の処理
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' member_id を引き継いで Form24 を表示
        Dim form24 As New Form24(memberId)
        form24.Show()

        ' 現在のフォームを閉じる
        Me.Close()
    End Sub
End Class
