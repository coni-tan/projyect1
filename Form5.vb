Imports MySql.Data.MySqlClient
Imports System.Text
Imports System.Windows.Forms

Public Class Form5
    ' データベース接続情報
    Dim connectionString As String = "Database=jyutyuu;Data Source=localhost;User Id=root"
    Dim items As New List(Of Dictionary(Of String, Object))()　'リスト型で商品情報を格納
    Dim currentIndex As Integer = 0
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

    ' 初期状態で商品を表示
    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadProductData()
        DisplayProductData(0)
    End Sub

    Private Sub LoadProductData(Optional query As String = Nothing)
        ' デフォルトのクエリ　検索条件を絞ったクエリが特に場合はこちらを使用
        If query Is Nothing Then
            query = "SELECT item_id, item_name, item_image, item_price FROM item WHERE item_keisai = 0 ORDER BY item_id"
        End If

        items.Clear() 'リストをクリア
        ' コネクト、コマンド、リーダーは一連の処理ではあるが、個々にusingでリソースを解放しないとパフォーマンス低下の恐れがある
        Using connection As New MySqlConnection(connectionString)
            Using command As New MySqlCommand(query, connection)
                connection.Open()
                Using reader As MySqlDataReader = command.ExecuteReader()
                    ' 取得した商品情報をリストに追加
                    While reader.Read()
                        Dim item As New Dictionary(Of String, Object) From {
                            {"item_id", reader("item_id").ToString()},
                            {"item_name", reader("item_name").ToString()},
                            {"item_image", reader("item_image").ToString()},
                            {"item_price", reader("item_price")}
                        }
                        items.Add(item)
                    End While
                End Using
            End Using
        End Using
    End Sub

    ' 画像クリックで商品詳細画面(form6)を表示
    Private Sub PictureBox_Click(sender As Object, e As EventArgs)
        Dim pictureBox As PictureBox = CType(sender, PictureBox)
        Dim itemId As String = pictureBox.Tag.ToString()
        Dim form6 As New Form6(itemId, memberId)
        form6.Show()
    End Sub

    ' 商品一覧を表示
    Private Sub DisplayProductData(startIndex As Integer)
        ' 現在表示している商品情報をクリア
        ClearProductData()

        Dim count As Integer = 1
        For i As Integer = startIndex To Math.Min(startIndex + 9, items.Count - 1)
            Dim item As Dictionary(Of String, Object) = items(i)

            ' 商品情報を表示するコントロールを取得
            Dim pictureBox As PictureBox = CType(Me.Controls.Find("PictureBox" & count, True).FirstOrDefault(), PictureBox)
            Dim nameLabel As Label = CType(Me.Controls.Find("Label" & count, True).FirstOrDefault(), Label)
            Dim priceLabel As Label = CType(Me.Controls.Find("Label" & (count + 10), True).FirstOrDefault(), Label)

            ' item_nameを関数処理しバイト数で制限
            Dim itemName As String = item("item_name").ToString()
            Dim truncatedName As String = TruncateStringByBytes(itemName, 50) ' バイト数を60に制限

            ' 文字数制限したitem_nameと単位付きのitem_priceをコントロールに表示
            nameLabel.Text = truncatedName

            ' item_priceをカンマ区切りにフォーマットして表示
            Dim itemPrice As Decimal = Convert.ToDecimal(item("item_price"))
            priceLabel.Text = itemPrice.ToString("N0") & "円" ' "N0"はカンマ区切りを追加するフォーマット

            ' item_imageの画像ファイル名から画像パスを生成しピクチャーボックスに表示
            Dim imagePath As String = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "img", item("item_image").ToString())
            If Not String.IsNullOrEmpty(item("item_image").ToString()) AndAlso System.IO.File.Exists(imagePath) Then
                pictureBox.Image = Image.FromFile(imagePath)
            Else
                ' 商品画像が未登録の場合はimg99を表示
                Dim fallbackImagePath As String = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "img", "img99.png")
                If System.IO.File.Exists(fallbackImagePath) Then
                    pictureBox.Image = Image.FromFile(fallbackImagePath)
                End If
            End If

            ' ToolTip(ツールヒント)を設定
            ToolTip1.SetToolTip(pictureBox, itemName)

            ' 既存のクリックイベントハンドラを削除
            RemoveHandler pictureBox.Click, AddressOf PictureBox_Click

            ' PictureBoxにクリックイベントハンドラを追加
            pictureBox.Tag = item("item_id").ToString()
            AddHandler pictureBox.Click, AddressOf PictureBox_Click

            count += 1
        Next
    End Sub

    Private Sub ClearProductData()
        For count As Integer = 1 To 10
            ' 商品情報を表示するコントロールを取得
            Dim pictureBox As PictureBox = CType(Me.Controls.Find("PictureBox" & count, True).FirstOrDefault(), PictureBox)
            Dim nameLabel As Label = CType(Me.Controls.Find("Label" & count, True).FirstOrDefault(), Label)
            Dim priceLabel As Label = CType(Me.Controls.Find("Label" & (count + 10), True).FirstOrDefault(), Label)

            ' 各種コントロールをクリア
            If pictureBox IsNot Nothing Then
                pictureBox.Image = Nothing
                ToolTip1.SetToolTip(pictureBox, String.Empty) ' ToolTipをクリア

                ' クリックイベントハンドラを削除
                RemoveHandler pictureBox.Click, AddressOf PictureBox_Click
            End If
            If nameLabel IsNot Nothing Then nameLabel.Text = String.Empty
            If priceLabel IsNot Nothing Then priceLabel.Text = String.Empty
        Next
    End Sub


    ' 「次へ」ボタンが押されたときの処理
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        currentIndex += 10
        If currentIndex < items.Count Then
            DisplayProductData(currentIndex)
        Else
            MessageBox.Show("これ以上の商品はありません。")
            currentIndex -= 10 ' インデックスを元に戻す
        End If
    End Sub

    ' 「前へ」ボタンが押されたときの処理
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If currentIndex >= 10 Then
            currentIndex -= 10
            DisplayProductData(currentIndex)
        Else
            MessageBox.Show("これ以上前の商品はありません。")
        End If
    End Sub

    ' 検索ボタンが押されたときの処理
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim categoryQueryPart As String = ""
        Dim keywordQueryPart As String = ""

        ' ジャンルの条件を取得
        If ComboBox1.SelectedItem IsNot Nothing Then
            Dim selectedCategory As String = ComboBox1.SelectedItem.ToString()
            If selectedCategory <> "すべて" Then
                Dim categoryID As String = GetCategoryID(selectedCategory)
                If Not String.IsNullOrEmpty(categoryID) Then
                    categoryQueryPart = " AND category_id = '" & categoryID & "'"
                End If
            End If
        End If

        ' 検索キーワードの条件を取得
        Dim keyword As String = TextBox1.Text
        If Not String.IsNullOrEmpty(keyword) Then
            keywordQueryPart = " AND item_name LIKE '%" & keyword & "%'"
        End If

        ' クエリを生成
        Dim query As String = "SELECT item_id, item_name, item_image, item_price FROM item WHERE item_keisai = 0" & categoryQueryPart & keywordQueryPart & " ORDER BY item_id"

        ' データを再読み込みして表示
        LoadProductData(query)
        currentIndex = 0
        DisplayProductData(currentIndex)
    End Sub

    ' category_nameに対応するcategory_idを取得する関数
    Private Function GetCategoryID(categoryName As String) As String
        Dim categoryID As String = ""
        Dim query As String = "SELECT category_id FROM category WHERE category_name = @category_name"

        Using connection As New MySqlConnection(connectionString)
            Using command As New MySqlCommand(query, connection)
                command.Parameters.AddWithValue("@category_name", categoryName)
                connection.Open()
                Dim result = command.ExecuteScalar()
                If result IsNot Nothing Then
                    categoryID = result.ToString()
                End If
            End Using
        End Using

        Return categoryID
    End Function

    ' 商品名の文字数制限
    Private Function TruncateStringByBytes(input As String, byteLimit As Integer) As String
        Dim utf8 As Encoding = Encoding.UTF8
        Dim bytes As Byte() = utf8.GetBytes(input)
        Dim outputBytes As New List(Of Byte)
        Dim currentByteCount As Integer = 0

        For Each c As Char In input
            Dim charBytes As Byte() = utf8.GetBytes(c.ToString())
            If currentByteCount + charBytes.Length > byteLimit Then
                Exit For
            End If
            outputBytes.AddRange(charBytes)
            currentByteCount += charBytes.Length
        Next

        Return utf8.GetString(outputBytes.ToArray())
    End Function

    ' 「マイページ」が押されたときの処理
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim form24 As New Form24(memberId)
        form24.Show()
    End Sub
End Class
