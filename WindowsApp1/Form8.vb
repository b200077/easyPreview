Public Class Form8
    Private Sub Mybase_Load(sender As Object, e As EventArgs) Handles MyBase.Shown
        TextBox1.Text = Clipboard.GetText

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form1.目錄檔案集合.Items(Form1.目錄檔案集合.SelectedIndex) = TextBox1.Text
        Hide()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Hide()
    End Sub
End Class