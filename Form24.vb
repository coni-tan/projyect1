Public Class Form24
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

    ' 「買い物かご」が押された時の処理
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim form7 As New Form7(memberId)
        form7.Show()
        Me.Hide()
    End Sub


    ' 「注文履歴」が押された時の処理
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim form25 As New Form25(memberId)
        form25.Show()
        Me.Hide()
    End Sub

    ' 「会員情報」が押された時の処理
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim form27 As New Form27(memberId)
        form27.Show()
        Me.Hide()
    End Sub

    ' 「問い合わせ」が押された時の処理
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim form28 As New Form28(memberId)
        form28.Show()
        Me.Hide()
    End Sub
End Class