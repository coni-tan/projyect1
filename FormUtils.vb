Imports System.IO
Imports MySql.Data.MySqlClient

Public Class FormUtils
    ' 画像処理機能
    Public Sub CopyAndDisplayImage(openFileDialog As OpenFileDialog, pictureBox As PictureBox, labelFilePath As Label, labelFileName As Label, targetDirectory As String)
        ' OpenFileDialogの設定
        With openFileDialog
            .Title = "画像ファイルの選択"
            .CheckFileExists = True
            .RestoreDirectory = True
            .Filter = "イメージファイル|*.bmp;*.jpg;*.gif;*.png"
        End With

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim sourceFile As String = openFileDialog.FileName
            Dim fileName As String = Path.GetFileName(sourceFile)
            Dim targetFile As String = Path.Combine(targetDirectory, fileName)

            Dim fileInfo As New FileInfo(sourceFile)
            If fileInfo.Length > 1048576 Then
                MessageBox.Show("ファイルサイズが1MBを超えています。1MB以下の画像を選択してください。")
                Return
            End If

            Try
                ' ディレクトリが存在しない場合は作成
                If Not Directory.Exists(targetDirectory) Then
                    Directory.CreateDirectory(targetDirectory)
                End If

                ' ファイルをコピー
                File.Copy(sourceFile, targetFile, True)

                ' FileStreamを使用して画像を読み込み、ファイルをロックしないようにする
                Using fileStream As New FileStream(targetFile, FileMode.Open, FileAccess.Read)
                    Dim tempImage As Image = Image.FromStream(fileStream)
                    pictureBox.Image = New Bitmap(tempImage)
                End Using

                ' ラベルに情報を表示
                labelFilePath.Text = targetFile
                labelFileName.Text = fileName
            Catch ex As Exception
                MessageBox.Show("画像のコピー中にエラーが発生しました: " & ex.Message)
                labelFilePath.Text = ""
                labelFileName.Text = ""
                pictureBox.Image = Nothing
            End Try
        Else
            labelFilePath.Text = ""
            labelFileName.Text = ""
            pictureBox.Image = Nothing
        End If
    End Sub

    ' フルスクリーン設定機能
    Public Sub SetFullScreen(form As Form)
        form.WindowState = FormWindowState.Maximized
        form.FormBorderStyle = FormBorderStyle.None
    End Sub

    ' GroupBoxの中央配置機能
    Public Sub CenterGroupBox(groupBox As GroupBox, parentForm As Form)
        groupBox.Left = (parentForm.ClientSize.Width - groupBox.Width) / 2
        groupBox.Top = (parentForm.ClientSize.Height - groupBox.Height) / 2
    End Sub

    ' GroupBoxとその中のコントロールのリサイズ機能
    Public Sub ResizeAndCenterGroupBox(groupBox As GroupBox, parentForm As Form)
        ' GroupBoxを親フォームのサイズに合わせてリサイズ
        Dim groupBoxOriginalSize As Size = groupBox.Size
        groupBox.Width = parentForm.ClientSize.Width * 0.8
        groupBox.Height = parentForm.ClientSize.Height * 0.8

        ' GroupBoxを中央に配置
        CenterGroupBox(groupBox, parentForm)

        ' GroupBox内のコントロールをリサイズおよび再配置
        For Each control As Control In groupBox.Controls
            '拡大されたグループボックスのサイズと元のサイズの比率を出す
            Dim widthRatio As Double = groupBox.ClientSize.Width / groupBoxOriginalSize.Width
            Dim heightRatio As Double = groupBox.ClientSize.Height / groupBoxOriginalSize.Height
            'CIntは計算結果を指定した方(整数型)に変換する
            control.Width = CInt(control.Width * widthRatio)
            control.Height = CInt(control.Height * heightRatio)
            control.Left = CInt(control.Left * widthRatio)
            control.Top = CInt(control.Top * heightRatio)
        Next
    End Sub

    ' ComboBoxにデータベースの値を設定する機能
    Public Sub PopulateComboBox(comboBox As ComboBox, query As String, valueMember As String, displayMember As String)
        Dim connectionString As String = "Database=jyutyuu;Data Source=localhost;User Id=root"
        Using conn As New MySqlConnection(connectionString)
            conn.Open()

            Using cmd As New MySqlCommand(query, conn)
                Using reader As MySqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        'リーダーから取り出した値をコンボボックスに追加していく
                        comboBox.Items.Add(New ComboBoxItem(reader.GetString(displayMember), reader.GetString(valueMember)))
                    End While
                End Using
            End Using
        End Using

        comboBox.DisplayMember = "DisplayText"
        comboBox.ValueMember = "Value"
    End Sub

    ' ComboBoxに表示する項目のためのクラス
    Public Class ComboBoxItem
        Public Property DisplayText As String
        Public Property Value As String

        Public Sub New(displayText As String, value As String)
            Me.DisplayText = displayText
            Me.Value = value
        End Sub

        ' オーバーライドせずともコンボボックスの機能によってdisplaytextが表示はされるが明示的にするため記述しておいた
        Public Overrides Function ToString() As String
            Return DisplayText
        End Function
    End Class
End Class
