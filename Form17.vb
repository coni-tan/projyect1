Imports MySql.Data.MySqlClient

Public Class Form17
    ' フォームロード時のイベントハンドラ
    Private Sub Form17_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadTalkRecords("すべて")
    End Sub

    ' DataGridView1に表示するレコードをロードする関数
    Private Sub LoadTalkRecords(filter As String)
        Try
            Using conn As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
                conn.Open()

                Dim sql As String = "SELECT talk_id, member_id, take_create_date, talk_status, talk_que, talk_ans FROM talk WHERE 1=1"
                If filter = "未回答" Then
                    sql &= " AND talk_status = 0"
                ElseIf filter = "回答済" Then
                    sql &= " AND talk_status = 1"
                End If

                Using cmd As New MySqlCommand(sql, conn)
                    Using adapter As New MySqlDataAdapter(cmd)
                        Dim dt As New DataTable()
                        adapter.Fill(dt)

                        ' ステータス列の値を変換するための新しい列を追加
                        dt.Columns.Add("talk_status_text", GetType(String))
                        For Each row As DataRow In dt.Rows
                            Dim status As Integer = Convert.ToInt32(row("talk_status"))
                            row("talk_status_text") = If(status = 0, "未回答", "回答済")
                        Next

                        ' DataGridViewにデータをバインドする前に列を再設定
                        DataGridView1.DataSource = dt

                        ' 列ヘッダーを設定
                        DataGridView1.Columns("talk_id").HeaderText = "問合せ番号"
                        DataGridView1.Columns("member_id").HeaderText = "会員番号"
                        DataGridView1.Columns("take_create_date").HeaderText = "問合せ日"
                        DataGridView1.Columns("talk_status_text").HeaderText = "ステータス"
                        DataGridView1.Columns("talk_que").HeaderText = "問合せ内容"
                        DataGridView1.Columns("talk_ans").HeaderText = "回答内容"

                        ' 元のtalk_status列を隠す
                        DataGridView1.Columns("talk_status").Visible = False

                        ' セルの編集を無効化
                        For Each column As DataGridViewColumn In DataGridView1.Columns
                            column.ReadOnly = True
                        Next

                        ' 不要レコード(追加用のブランクのレコード)の非表示
                        DataGridView1.AllowUserToAddRows = False

                        ' 行ヘッダーを表示し、見出しを「選択」に設定
                        DataGridView1.RowHeadersVisible = True
                        DataGridView1.TopLeftHeaderCell.Value = "選択"
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("エラーが発生しました: " & ex.Message)
        End Try
    End Sub

    ' 検索ボタンのクリックイベントハンドラ (Button3)
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' フィルター条件の取得
        Dim statusFilter As String = ComboBox1.SelectedItem.ToString()
        Dim talkIdFilter As String = TextBox3.Text.Trim()
        Dim memberIdFilter As String = TextBox4.Text.Trim()

        ' クエリの生成
        Dim sql As String = "SELECT talk_id, member_id, take_create_date, talk_status, talk_que, talk_ans FROM talk WHERE 1=1"

        ' ステータスフィルタを追加
        If statusFilter = "未回答" Then
            sql &= " AND talk_status = 0"
        ElseIf statusFilter = "回答済" Then
            sql &= " AND talk_status = 1"
        End If

        ' 問い合わせIDフィルタを追加
        If Not String.IsNullOrEmpty(talkIdFilter) Then
            sql &= " AND talk_id = @talk_id"
        End If

        ' 会員番号フィルタを追加
        If Not String.IsNullOrEmpty(memberIdFilter) Then
            sql &= " AND member_id = @member_id"
        End If

        ' データベース接続とクエリ実行
        Try
            Using conn As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
                conn.Open()
                Using cmd As New MySqlCommand(sql, conn)
                    ' パラメータを設定
                    If Not String.IsNullOrEmpty(talkIdFilter) Then
                        cmd.Parameters.AddWithValue("@talk_id", talkIdFilter)
                    End If
                    If Not String.IsNullOrEmpty(memberIdFilter) Then
                        cmd.Parameters.AddWithValue("@member_id", memberIdFilter)
                    End If

                    ' データを取得
                    Using adapter As New MySqlDataAdapter(cmd)
                        Dim dt As New DataTable()
                        adapter.Fill(dt)

                        ' ステータス列の値を変換するための新しい列を追加
                        dt.Columns.Add("talk_status_text", GetType(String))
                        For Each row As DataRow In dt.Rows
                            Dim status As Integer = Convert.ToInt32(row("talk_status"))
                            row("talk_status_text") = If(status = 0, "未回答", "回答済")
                        Next

                        ' DataGridViewにデータをバインドする前に列を再設定
                        DataGridView1.DataSource = dt

                        ' 列ヘッダーを設定
                        DataGridView1.Columns("talk_id").HeaderText = "問合せ番号"
                        DataGridView1.Columns("member_id").HeaderText = "会員番号"
                        DataGridView1.Columns("take_create_date").HeaderText = "問合せ日"
                        DataGridView1.Columns("talk_status_text").HeaderText = "ステータス"
                        DataGridView1.Columns("talk_que").HeaderText = "問合せ内容"
                        DataGridView1.Columns("talk_ans").HeaderText = "回答内容"

                        ' 元のtalk_status列を隠す
                        DataGridView1.Columns("talk_status").Visible = False

                        ' セルの編集を無効化
                        For Each column As DataGridViewColumn In DataGridView1.Columns
                            column.ReadOnly = True
                        Next

                        ' 不要レコード(追加用のブランクのレコード)の非表示
                        DataGridView1.AllowUserToAddRows = False

                        ' 行ヘッダーを表示し、見出しを「選択」に設定
                        DataGridView1.RowHeadersVisible = True
                        DataGridView1.TopLeftHeaderCell.Value = "選択"
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("エラーが発生しました: " & ex.Message)
        End Try
    End Sub

    ' DataGridView1の選択された行が変更された際のイベントハンドラ
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
            TextBox2.Text = selectedRow.Cells("talk_que").Value.ToString()
            TextBox2.ReadOnly = True ' TextBox2を閲覧のみで編集不可に設定
            TextBox1.Text = selectedRow.Cells("talk_ans").Value.ToString()
            Dim talkStatus As Integer = Convert.ToInt32(selectedRow.Cells("talk_status").Value)

            ' ComboBox2の選択状態を設定
            If talkStatus = 0 Then
                ComboBox2.SelectedIndex = 0 ' "未回答"が選択された状態
            ElseIf talkStatus = 1 Then
                ComboBox2.SelectedIndex = 1 ' "回答済"が選択された状態
            End If
        End If
    End Sub

    ' 更新ボタンがクリックされた時の処理 (Button1)
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' DataGridViewで選択されているか確認
        If DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("問い合わせを選択してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' TextBox1の入力内容を取得し、トリムする
        Dim responseContent As String = TextBox1.Text.Trim()

        ' TextBox1の入力チェック
        If responseContent.Length > 1000 Then
            MessageBox.Show("回答内容は1000文字以内で入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' DataGridViewで選択された行のtalk_idを取得
        Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
        Dim talkId As String = selectedRow.Cells("talk_id").Value.ToString()

        ' ComboBox2の選択内容でtalk_statusを設定
        Dim talkStatus As Integer
        If ComboBox2.SelectedItem.ToString() = "未回答" Then
            talkStatus = 0
        ElseIf ComboBox2.SelectedItem.ToString() = "回答済" Then
            talkStatus = 1
        Else
            MessageBox.Show("有効なステータスを選択してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' 回答の更新処理
        Try
            Using conn As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
                conn.Open()

                Dim sql As String = "UPDATE talk SET talk_ans = @response, talk_status = @status, take_update_date = NOW() WHERE talk_id = @talk_id"

                Using cmd As New MySqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@response", responseContent)
                    cmd.Parameters.AddWithValue("@status", talkStatus)
                    cmd.Parameters.AddWithValue("@talk_id", talkId)

                    cmd.ExecuteNonQuery()
                End Using

                MessageBox.Show("回答が送信されました。", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
                TextBox1.Clear()
                TextBox2.Clear()
                LoadTalkRecords("すべて") ' DataGridView1を更新
            End Using
        Catch ex As Exception
            MessageBox.Show("エラーが発生しました: " & ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
