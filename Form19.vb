Imports MySql.Data.MySqlClient
Imports System.Text.RegularExpressions

Public Class Form19
    ' 新規登録ボタンがクリックされた時の処理 (Button1)
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' テキストボックスから入力値を取得
        Dim empId As String = TextBox1.Text.Trim()
        Dim empName As String = TextBox9.Text.Trim()
        Dim empPass As String = TextBox2.Text.Trim()

        ' 入力値の検証
        If String.IsNullOrEmpty(empId) OrElse empId.Length <> 8 OrElse Not IsNumeric(empId) Then
            MessageBox.Show("社員番号は8桁の数字で入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If String.IsNullOrEmpty(empName) OrElse empName.Length > 50 Then
            MessageBox.Show("氏名は50文字以内で入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' TextBox2の入力チェック（半角英数記号、10文字以内）
        If String.IsNullOrEmpty(empPass) OrElse empPass.Length > 10 OrElse Not Regex.IsMatch(empPass, "^[a-zA-Z0-9!@#\$%\^&\*\(\)\-_+={}\[\]:;""'<>?,.\/\\]{1,10}$") Then
            MessageBox.Show("パスワードは半角英数記号10文字以内で入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' コンボボックスの選択値を結合してemp_permissionsを作成
        Dim empPermissions As String = String.Concat(
            If(ComboBox1.SelectedItem.ToString() = "利用可", "0", "1"),
            If(ComboBox2.SelectedItem.ToString() = "利用可", "0", "1"),
            If(ComboBox3.SelectedItem.ToString() = "利用可", "0", "1"),
            If(ComboBox4.SelectedItem.ToString() = "利用可", "0", "1"),
            If(ComboBox5.SelectedItem.ToString() = "利用可", "0", "1")
        )

        ' 重複チェック
        Try
            Using conn As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
                conn.Open()
                Dim checkSql As String = "SELECT COUNT(*) FROM emp WHERE emp_id = @emp_id"
                Using checkCmd As New MySqlCommand(checkSql, conn)
                    checkCmd.Parameters.AddWithValue("@emp_id", empId)
                    Dim count As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
                    If count > 0 Then
                        MessageBox.Show("その社員番号は既に登録されています。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End If
                End Using

                ' 新規レコードの追加
                Dim insertSql As String = "INSERT INTO emp (emp_id, emp_name, emp_pass, emp_permissions, emp_date) VALUES (@emp_id, @emp_name, @emp_pass, @emp_permissions, NOW())"
                Using insertCmd As New MySqlCommand(insertSql, conn)
                    insertCmd.Parameters.AddWithValue("@emp_id", empId)
                    insertCmd.Parameters.AddWithValue("@emp_name", empName)
                    insertCmd.Parameters.AddWithValue("@emp_pass", empPass)
                    insertCmd.Parameters.AddWithValue("@emp_permissions", empPermissions)
                    insertCmd.ExecuteNonQuery()
                End Using

                MessageBox.Show("新規社員が登録されました。", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Using
        Catch ex As Exception
            MessageBox.Show("エラーが発生しました: " & ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
