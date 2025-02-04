Imports MySql.Data.MySqlClient
Imports System.Windows.Forms

Public Class Form28
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

    ' フォームロード時のイベントハンドラ
    Private Sub Form28_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadTalkRecords("すべて")
        CenterGroupBox()
    End Sub

    ' フォームサイズ変更時のイベントハンドラ
    Private Sub Form28_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        CenterGroupBox()
    End Sub

    ' GroupBox3を画面の中央に配置するメソッド
    Private Sub CenterGroupBox()
        GroupBox3.Left = (Me.ClientSize.Width - GroupBox3.Width) / 2
        GroupBox3.Top = (Me.ClientSize.Height - GroupBox3.Height) / 2
    End Sub

    ' DataGridView1に表示するレコードをロードする関数
    Private Sub LoadTalkRecords(filter As String)
        Try
            Using conn As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
                conn.Open()

                Dim sql As String = "SELECT talk_id, take_create_date, talk_status, talk_que, talk_ans FROM talk WHERE member_id = @member_id"
                If filter = "回答済み" Then
                    sql &= " AND talk_status = 1"
                ElseIf filter = "回答待ち" Then
                    sql &= " AND talk_status = 0"
                End If

                Using cmd As New MySqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@member_id", Me.memberId)

                    Using adapter As New MySqlDataAdapter(cmd)
                        Dim dt As New DataTable()
                        adapter.Fill(dt)

                        ' ステータス列の値を変換するための新しい列を追加
                        dt.Columns.Add("talk_status_text", GetType(String))
                        For Each row As DataRow In dt.Rows
                            Dim status As Integer = Convert.ToInt32(row("talk_status"))
                            row("talk_status_text") = If(status = 0, "回答待ち", "回答済み")
                        Next

                        ' DataGridViewにデータをバインドする前に列を再設定
                        DataGridView1.DataSource = dt

                        ' 列ヘッダーを設定
                        DataGridView1.Columns("talk_id").HeaderText = "問合せ番号"
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

    ' 検索ボタンのクリックイベントハンドラ
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim filter As String = ComboBox1.SelectedItem.ToString()
        LoadTalkRecords(filter)
    End Sub

    ' DataGridView1の選択された行が変更された際のイベントハンドラ
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
            TextBox3.Text = selectedRow.Cells("talk_que").Value.ToString()
            TextBox4.Text = selectedRow.Cells("talk_ans").Value.ToString()
        End If
    End Sub

    ' 新規問い合わせを送信するボタンクリックイベントハンドラ
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim talkContent As String = TextBox2.Text.Trim()

        ' 入力内容のバリデーション
        If talkContent = "" Then
            MessageBox.Show("問い合わせ内容が入力されていません。")
            Exit Sub
        End If

        If talkContent.Length > 1000 Then
            MessageBox.Show("問い合わせ内容は1000文字以内で入力してください。")
            Exit Sub
        End If

        ' 新規問い合わせの作成
        Try
            Using conn As New MySqlConnection("Database=jyutyuu;Data Source=localhost;User Id=root")
                conn.Open()

                ' 新しいtalk_idの生成
                Dim newTalkId As String = GenerateNewTalkId(conn)

                Dim sql As String = "INSERT INTO talk (talk_id, member_id, talk_que, talk_ans, talk_status, take_create_date, take_update_date) " &
                                    "VALUES (@talk_id, @member_id, @talk_que, '', 0, NOW(), NOW())"

                Using cmd As New MySqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@talk_id", newTalkId)
                    cmd.Parameters.AddWithValue("@member_id", Me.memberId)
                    cmd.Parameters.AddWithValue("@talk_que", talkContent)

                    cmd.ExecuteNonQuery()
                End Using

                MessageBox.Show("問い合わせが送信されました。問い合わせ番号：" & newTalkId)
                TextBox2.Clear()
                LoadTalkRecords("すべて") ' DataGridView1を更新
            End Using
        Catch ex As Exception
            MessageBox.Show("エラーが発生しました: " & ex.Message)
        End Try
    End Sub

    ' 新しいtalk_idを生成する関数
    Private Function GenerateNewTalkId(conn As MySqlConnection) As String
        Dim newTalkId As String = "00000001"
        Dim sql As String = "SELECT MAX(talk_id) FROM talk"

        Using cmd As New MySqlCommand(sql, conn)
            Dim result As Object = cmd.ExecuteScalar()

            If result IsNot DBNull.Value Then
                newTalkId = (CInt(result) + 1).ToString("D8")
            End If
        End Using

        Return newTalkId
    End Function

    ' 閉じるボタンがクリックされた時の処理
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' member_id を引き継いで Form24 を表示
        Dim form24 As New Form24(memberId)
        form24.Show()

        ' 現在のフォームを閉じる
        Me.Close()
    End Sub
End Class
