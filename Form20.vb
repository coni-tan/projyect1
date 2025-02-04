Imports MySql.Data.MySqlClient
Imports System.Text.RegularExpressions

Public Class Form20
    Private empId As String

    ' 検索ボタンがクリックされた時の処理 (Button2)
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        empId = TextBox1.Text.Trim()

        If String.IsNullOrEmpty(empId) OrElse empId.Length <> 8 OrElse Not IsNumeric(empId) Then
            MessageBox.Show("社員番号は8桁の数字で入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' 社員情報のロード
        LoadEmployeeInfo()
    End Sub

    ' 社員情報をデータベースから取得して表示するメソッド
    Private Sub LoadEmployeeInfo()
        Dim query As String = "SELECT emp_id, emp_name, emp_pass, emp_permissions FROM emp WHERE emp_id = @emp_id"

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            Using command As New MySqlCommand(query, connection)
                command.Parameters.AddWithValue("@emp_id", empId)
                connection.Open()
                Using reader As MySqlDataReader = command.ExecuteReader()
                    If reader.Read() Then
                        TextBox1.Text = reader("emp_id").ToString()
                        TextBox9.Text = reader("emp_name").ToString()
                        TextBox2.Text = reader("emp_pass").ToString()
                        Dim permissions As String = reader("emp_permissions").ToString()

                        ' コンボボックスの選択状態を設定
                        ComboBox1.SelectedIndex = If(permissions(0) = "0", 0, 1)
                        ComboBox2.SelectedIndex = If(permissions(1) = "0", 0, 1)
                        ComboBox3.SelectedIndex = If(permissions(2) = "0", 0, 1)
                        ComboBox4.SelectedIndex = If(permissions(3) = "0", 0, 1)
                        ComboBox5.SelectedIndex = If(permissions(4) = "0", 0, 1)
                    Else
                        MessageBox.Show("社員情報が見つかりません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End Using
            End Using
        End Using
    End Sub

    ' 更新ボタンがクリックされた時の処理 (Button1)
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' 社員情報が表示されていない場合のエラーメッセージ
        If String.IsNullOrEmpty(empId) Then
            MessageBox.Show("社員情報を表示してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' 各テキストボックスから入力値を取得
        Dim empName As String = TextBox9.Text.Trim()
        Dim empPass As String = TextBox2.Text.Trim()

        ' 入力値の検証
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

        ' 更新処理
        Try
            Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
                connection.Open()
                Dim updateSql As String = "UPDATE emp SET emp_name = @emp_name, emp_pass = @emp_pass, emp_permissions = @emp_permissions WHERE emp_id = @emp_id"
                Using command As New MySqlCommand(updateSql, connection)
                    command.Parameters.AddWithValue("@emp_id", empId)
                    command.Parameters.AddWithValue("@emp_name", empName)
                    command.Parameters.AddWithValue("@emp_pass", empPass)
                    command.Parameters.AddWithValue("@emp_permissions", empPermissions)
                    command.ExecuteNonQuery()
                End Using
                MessageBox.Show("社員情報が更新されました。", "更新成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Using
        Catch ex As Exception
            MessageBox.Show("エラーが発生しました: " & ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 削除ボタンがクリックされた時の処理 (Button3)
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' 社員情報が表示されていない場合のエラーメッセージ
        If String.IsNullOrEmpty(empId) Then
            MessageBox.Show("社員情報を表示してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim confirmResult As DialogResult = MessageBox.Show("本当に削除しますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If confirmResult = DialogResult.No Then
            Return
        End If

        ' 削除処理
        Try
            Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
                connection.Open()
                Dim deleteSql As String = "DELETE FROM emp WHERE emp_id = @emp_id"
                Using command As New MySqlCommand(deleteSql, connection)
                    command.Parameters.AddWithValue("@emp_id", empId)
                    command.ExecuteNonQuery()
                End Using
                MessageBox.Show("社員情報が削除されました。", "削除成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
                ' フィールドをクリア
                TextBox1.Text = ""
                TextBox9.Text = ""
                TextBox2.Text = ""
                ComboBox1.SelectedIndex = -1
                ComboBox2.SelectedIndex = -1
                ComboBox3.SelectedIndex = -1
                ComboBox4.SelectedIndex = -1
                ComboBox5.SelectedIndex = -1
                empId = Nothing
            End Using
        Catch ex As Exception
            MessageBox.Show("エラーが発生しました: " & ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
