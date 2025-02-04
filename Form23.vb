Imports MySql.Data.MySqlClient
Imports System.Text.RegularExpressions

Public Class Form23
    Private partnerId As String

    ' 検索ボタンがクリックされた時の処理 (Button2)
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        partnerId = TextBox1.Text.Trim()

        If String.IsNullOrEmpty(partnerId) OrElse partnerId.Length <> 2 OrElse Not IsNumeric(partnerId) Then
            MessageBox.Show("取引先コードは2桁の数字で入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' 取引先情報のロード
        LoadPartnerInfo()
    End Sub

    ' 取引先情報をデータベースから取得して表示するメソッド
    Private Sub LoadPartnerInfo()
        Dim query As String = "SELECT partner_id, partner_name FROM partner WHERE partner_id = @partner_id"

        Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
            Using command As New MySqlCommand(query, connection)
                command.Parameters.AddWithValue("@partner_id", partnerId)
                connection.Open()
                Using reader As MySqlDataReader = command.ExecuteReader()
                    If reader.Read() Then
                        TextBox1.Text = reader("partner_id").ToString()
                        TextBox2.Text = reader("partner_name").ToString()
                    Else
                        MessageBox.Show("取引先情報が見つかりません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End Using
            End Using
        End Using
    End Sub

    ' 更新ボタンがクリックされた時の処理 (Button1)
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' 取引先情報が表示されていない場合のエラーメッセージ
        If String.IsNullOrEmpty(partnerId) Then
            MessageBox.Show("取引先情報を表示してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' 各テキストボックスから入力値を取得
        Dim partnerName As String = TextBox2.Text.Trim()

        ' 入力値の検証
        If String.IsNullOrEmpty(partnerName) OrElse partnerName.Length > 50 Then
            MessageBox.Show("取引先名は50文字以内で入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' 更新処理
        Try
            Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
                connection.Open()
                Dim updateSql As String = "UPDATE partner SET partner_name = @partner_name WHERE partner_id = @partner_id"
                Using command As New MySqlCommand(updateSql, connection)
                    command.Parameters.AddWithValue("@partner_id", partnerId)
                    command.Parameters.AddWithValue("@partner_name", partnerName)
                    command.ExecuteNonQuery()
                End Using
                MessageBox.Show("取引先情報が更新されました。", "更新成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Using
        Catch ex As Exception
            MessageBox.Show("エラーが発生しました: " & ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 削除ボタンがクリックされた時の処理 (Button3)
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' 取引先情報が表示されていない場合のエラーメッセージ
        If String.IsNullOrEmpty(partnerId) Then
            MessageBox.Show("取引先情報を表示してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim confirmResult As DialogResult = MessageBox.Show("本当に削除しますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If confirmResult = DialogResult.No Then
            Return
        End If

        ' multi_partner表に関連するレコードが存在するか確認
        Dim checkMultiPartnerQuery As String = "SELECT COUNT(*) FROM multi_partner WHERE partner_id = @partner_id"
        Try
            Using connection As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
                connection.Open()
                Using checkCommand As New MySqlCommand(checkMultiPartnerQuery, connection)
                    checkCommand.Parameters.AddWithValue("@partner_id", partnerId)
                    Dim count As Integer = Convert.ToInt32(checkCommand.ExecuteScalar())
                    If count > 0 Then
                        MessageBox.Show("商品に紐づいている取引先は削除できません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End If
                End Using

                ' 削除処理
                Dim deleteSql As String = "DELETE FROM partner WHERE partner_id = @partner_id"
                Using deleteCommand As New MySqlCommand(deleteSql, connection)
                    deleteCommand.Parameters.AddWithValue("@partner_id", partnerId)
                    deleteCommand.ExecuteNonQuery()
                End Using
                MessageBox.Show("取引先情報が削除されました。", "削除成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
                ' フィールドをクリア
                TextBox1.Text = ""
                TextBox2.Text = ""
                partnerId = Nothing
            End Using
        Catch ex As Exception
            MessageBox.Show("エラーが発生しました: " & ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
