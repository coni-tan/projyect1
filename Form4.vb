Imports MySql.Data.MySqlClient

Public Class Form4
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' テキストボックスの宣言
        Dim empId As String = TextBox1.Text
        Dim empPw As String = TextBox2.Text

        ' 未入力なのにボタン押されたらエラー
        If String.IsNullOrEmpty(empId) OrElse String.IsNullOrEmpty(empPw) Then
            MessageBox.Show("IDまたはパスワードを入力してください。")
            Return
        End If

        ' データベース接続と認証処理
        Dim empPermissions As String = AuthenticateUser(empId, empPw)
        If empPermissions IsNot Nothing Then
            MessageBox.Show("ログインしました。")

            ' Form9のインスタンスを作成し、表示
            Dim form9 As New Form9(empId, empPermissions)
            form9.Show()

            ' 現在のフォームを隠す
            Me.Hide()
        Else
            MessageBox.Show("IDまたはパスワードが一致しません。")
        End If
    End Sub

    ' ログインを form4 以外で使う場合に備え関数にしておく
    Private Function AuthenticateUser(empId As String, empPw As String) As String
        Dim empPermissions As String = Nothing

        ' データベース接続情報
        Dim connectionString As String = "Database=jyutyuu;Data Source=localhost;User Id=root"

        Try
            Using connection As New MySqlConnection(connectionString) 'usingで自動close
                Using command As New MySqlCommand() 'usingで自動close
                    command.Connection = connection

                    ' SQLクエリをセットする + SQLインジェクション対策
                    command.CommandText = "SELECT emp_id, emp_pass, emp_permissions FROM emp WHERE emp_id = @empId AND emp_pass = @empPw"

                    ' パラメータを紐づけ
                    command.Parameters.AddWithValue("@empId", empId)
                    command.Parameters.AddWithValue("@empPw", empPw)

                    connection.Open()

                    ' データリーダーにデータ取得
                    Dim DataReader As MySqlDataReader = command.ExecuteReader()

                    ' 入力したID/PWの組がデータが存在するか確認し、権限を取得
                    If DataReader.Read() Then
                        empPermissions = DataReader("emp_permissions").ToString()
                    End If

                    ' データリーダーを閉じる
                    DataReader.Close()
                End Using
            End Using
        Catch ex As MySqlException
            MessageBox.Show("エラーが発生しました: " & ex.Message)
        End Try

        Return empPermissions '権限（文字列型）を返す
    End Function
End Class
