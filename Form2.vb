Imports MySql.Data.MySqlClient

Public Class Form2
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' テキストボックスの宣言
        Dim memberId As String = TextBox1.Text
        Dim memberPw As String = TextBox2.Text

        ' 未入力なのにボタン押されたらエラー
        If String.IsNullOrEmpty(memberId) OrElse String.IsNullOrEmpty(memberPw) Then
            MessageBox.Show("IDまたはパスワードを入力してください。")
            Return
        End If

        ' データベース接続と認証処理
        If AuthenticateUser(memberId, memberPw) Then
            MessageBox.Show("ログインしました。")

            ' Form5のインスタンスを作成し、表示
            Dim form5 As New Form5(memberId)
            form5.Show()

            ' 現在のフォームを隠す
            Me.Hide()
        Else
            MessageBox.Show("IDまたはパスワードが一致しません。")
        End If
    End Sub

    ' ログインを form2 以外で使う場合に備え関数にしておく
    Private Function AuthenticateUser(memberId As String, memberPw As String) As Boolean
        Dim isAuthenticated As Boolean = False

        ' データベース接続情報
        Dim connectionString As String = "Database=jyutyuu;Data Source=localhost;User Id=root"

        Try
            Using connection As New MySqlConnection(connectionString) 'usingで自動close
                Using command As New MySqlCommand() 'usingで自動close
                    command.Connection = connection

                    ' SQLクエリをセットする + SQLインジェクション対策
                    command.CommandText = "SELECT member_id, member_pass FROM member WHERE member_id = @memberId AND member_pass = @memberPw"

                    ' パラメータを紐づけ
                    command.Parameters.AddWithValue("@memberId", memberId)
                    command.Parameters.AddWithValue("@memberPw", memberPw)

                    connection.Open()

                    ' データリーダーにデータ取得
                    Dim DataReader As MySqlDataReader = command.ExecuteReader()

                    ' 入力したID/PWの組がデータが存在するか確認
                    If DataReader.HasRows Then
                        isAuthenticated = True
                    Else
                        isAuthenticated = False
                    End If

                    ' データリーダーを閉じる
                    DataReader.Close()
                End Using
            End Using
        Catch ex As MySqlException
            MessageBox.Show("エラーが発生しました: " & ex.Message)
        End Try

        Return isAuthenticated 'Boolean型
    End Function

    ' 会員情報を取得する関数
    Private Function GetMemberInfo(memberId As String) As Dictionary(Of String, String)
        Dim memberInfo As New Dictionary(Of String, String)()

        ' データベース接続情報
        Dim connectionString As String = "Database=jyutyuu;Data Source=localhost;User Id=root"

        Try
            Using connection As New MySqlConnection(connectionString) 'usingで自動close
                Using command As New MySqlCommand() 'usingで自動close
                    command.Connection = connection

                    ' SQLクエリをセットする + SQLインジェクション対策
                    command.CommandText = "SELECT member_id, member_name, member_email FROM member WHERE member_id = @memberId"

                    ' パラメータを紐づけ
                    command.Parameters.AddWithValue("@memberId", memberId)

                    connection.Open()

                    ' データリーダーにデータ取得
                    Dim DataReader As MySqlDataReader = command.ExecuteReader()

                    ' 会員情報を取得
                    If DataReader.Read() Then
                        memberInfo("member_id") = DataReader("member_id").ToString()
                        memberInfo("member_name") = DataReader("member_name").ToString()
                        memberInfo("member_email") = DataReader("member_email").ToString()
                    End If

                    ' データリーダーを閉じる
                    DataReader.Close()
                End Using
            End Using
        Catch ex As MySqlException
            MessageBox.Show("エラーが発生しました: " & ex.Message)
        End Try

        Return memberInfo 'Dictionary型
    End Function
End Class
