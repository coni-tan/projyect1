Imports MySql.Data.MySqlClient
Imports System.Text.RegularExpressions

Public Class Form22
    ' 新規登録ボタンがクリックされた時の処理 (Button1)
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' テキストボックスから入力値を取得
        Dim partnerId As String = TextBox1.Text.Trim()
        Dim partnerName As String = TextBox2.Text.Trim()

        ' 入力値の検証
        If String.IsNullOrEmpty(partnerId) OrElse partnerId.Length <> 2 OrElse Not IsNumeric(partnerId) Then
            MessageBox.Show("取引先コードは2桁の数字で入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If String.IsNullOrEmpty(partnerName) OrElse partnerName.Length > 50 Then
            MessageBox.Show("取引先名は50文字以内で入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' 重複チェック
        Try
            Using conn As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
                conn.Open()
                Dim checkSql As String = "SELECT COUNT(*) FROM partner WHERE partner_id = @partner_id"
                Using checkCmd As New MySqlCommand(checkSql, conn)
                    checkCmd.Parameters.AddWithValue("@partner_id", partnerId)
                    Dim count As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
                    If count > 0 Then
                        MessageBox.Show("その取引先コードは既に登録されています。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End If
                End Using

                ' 新規レコードの追加
                Dim insertSql As String = "INSERT INTO partner (partner_id, partner_name, partner_date) VALUES (@partner_id, @partner_name, NOW())"
                Using insertCmd As New MySqlCommand(insertSql, conn)
                    insertCmd.Parameters.AddWithValue("@partner_id", partnerId)
                    insertCmd.Parameters.AddWithValue("@partner_name", partnerName)
                    insertCmd.ExecuteNonQuery()
                End Using

                MessageBox.Show("新規取引先が登録されました。", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
                TextBox1.Clear()
                TextBox2.Clear()
            End Using
        Catch ex As Exception
            MessageBox.Show("エラーが発生しました: " & ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
