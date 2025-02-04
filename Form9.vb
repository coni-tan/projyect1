Public Class Form9
    ' 前画面から社員IDと権限情報を引き継ぐコンストラクタ
    Private empId As String
    Private empPermissions As String

    Public Sub New(empId As String, empPermissions As String)
        InitializeComponent()
        Me.empId = empId
        Me.empPermissions = empPermissions
    End Sub

    Private Sub Form9_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' フォームがロードされた時にラベルにテキストを設定
        Label1.Text = empId
        Label2.Text = empPermissions

        ' 各ボタンの利用権限を判定
        SetButtonState()
    End Sub


    ' subはサブプロシージャ(手続き)、戻り値がない場合
    ' functionはファンクション(関数)、戻り値がある場合
    ' 両方のことを区別せず言いたいときはメソッド(method)
    Private Sub SetButtonState()
        ' 各ボタンのリストを作成
        Dim buttons As Button() = {Button1, Button2, Button3, Button4, Button5}

        ' empPermissionsの各桁に基づいてボタンの状態を設定
        For i As Integer = 1 To buttons.Length
            If Mid(empPermissions, i, 1) = "1" Then
                buttons(i - 1).Enabled = False
                buttons(i - 1).BackColor = Color.Gray
            End If
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '　注文管理画面を表示
        Form10.Show()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '　商品管理画面を表示
        Form12.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        '　会員管理画面を表示
        Form15.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' 社員管理
        Form18.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ' 取引先管理
        Form21.Show()
    End Sub
End Class
