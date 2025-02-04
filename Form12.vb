Public Class Form12
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        '　商品編集を表示
        Form14.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '　商品新規登録を表示
        Form13.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        '　在庫管理を表示
        Form29.Show()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '　複数取引先を表示
        Form30.Show()
    End Sub
End Class