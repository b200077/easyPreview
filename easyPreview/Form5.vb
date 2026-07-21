Public Class Form5
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form1.Remove_Button(sender, New EventArgs)
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form2.checkWebSite()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form2.Button3_Click(sender, e)
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Form2.Button4_Click(sender, e)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Form1.file_Preview(False)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim fakeMenuItem As New ToolStripMenuItem With {
            .Text = "搜尋"
        }
        Form2.Form1List_operate(fakeMenuItem, e)
    End Sub

    Private Sub CheckBox_CheckedChanged(sender As CheckBox, e As EventArgs) Handles CheckBox1.CheckedChanged, CheckBox2.CheckedChanged
        Form1.CheckedListBox1_setCheck(sender.Text, sender.Checked)
    End Sub
End Class