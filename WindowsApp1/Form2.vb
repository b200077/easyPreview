Public Class Form2
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        TextBox1.Text = (Convert.ToInt32(TextBox1.Text) + 5).ToString
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        TextBox1.Text = (Convert.ToInt32(TextBox1.Text) - 5).ToString
    End Sub
    Public Function CopyRectangle(ByVal sourceImage As Image, ByVal area As Rectangle) As Image
        Dim outPut As New Bitmap(area.Width, area.Height)
        Dim DescREctangle As Rectangle = New Rectangle(0, 0, outPut.Width, outPut.Height)
        Dim g As Graphics = Drawing.Graphics.FromImage(outPut)
        g.DrawImage(sourceImage, DescREctangle, area, GraphicsUnit.Pixel)
        g.Save()
        Return outPut
    End Function

    Public Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim ori_image = Form1.PictureBox1.Image
        If ori_image IsNot Nothing Then
            Dim point = picture_set(ori_image)
            Dim rect As Rectangle = New Rectangle(point(0), point(1), point(2), point(3))
            Dim boximage As Bitmap = CopyRectangle(ori_image, rect)
            PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
            PictureBox1.Image = CType(boximage, Image)
        End If
    End Sub



    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Form1.壓縮檔案集合.Items.Count = 0 And Form1.壓縮檔案集合.SelectedItem = Nothing Then
            If Form1.目錄檔案集合.SelectedIndices.Cast(Of Integer)().ToArray()(0) >= 1 Then
                Form1.目錄檔案集合.SetSelected(Form1.目錄檔案集合.SelectedIndex - 1, True)
                Form1.目錄檔案集合.SetSelected(Form1.目錄檔案集合.SelectedIndex, False)
                ' obj.SelectedIndices.Remove(obj.SelectedIndex + 1)
            End If
        Else
            If Form1.壓縮檔案集合.SelectedIndices.Cast(Of Integer)().ToArray()(0) >= 1 Then
                Form1.壓縮檔案集合.SetSelected(Form1.壓縮檔案集合.SelectedIndex - 1, True)
                Form1.壓縮檔案集合.SetSelected(Form1.壓縮檔案集合.SelectedIndex, False)
                ' obj.SelectedIndices.Remove(obj.SelectedIndex + 1)
            End If
        End If
        If Form1.壓縮檔案集合.Items.Count = 0 And Form1.壓縮檔案集合.SelectedItem = Nothing Then
            Form1.file_Preview(sender, e)
        Else
            Form1.Archive_Preview(sender, e)
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Form1.壓縮檔案集合.Items.Count = 0 And Form1.壓縮檔案集合.SelectedItem = Nothing Then
            If Form1.目錄檔案集合.SelectedIndices.Cast(Of Integer)().ToArray()(Form1.目錄檔案集合.SelectedIndices.Count - 1) < Form1.目錄檔案集合.Items.Count - 1 Then
                Form1.目錄檔案集合.SetSelected(Form1.目錄檔案集合.SelectedIndex + 1, True)
                Form1.目錄檔案集合.SetSelected(Form1.目錄檔案集合.SelectedIndex, False)
                'obj.SelectedIndices.Remove(obj.SelectedIndex - 1)
            End If
        Else
            If Form1.壓縮檔案集合.SelectedIndices.Cast(Of Integer)().ToArray()(Form1.壓縮檔案集合.SelectedIndices.Count - 1) < Form1.目錄檔案集合.Items.Count - 1 Then
                Form1.壓縮檔案集合.SetSelected(Form1.壓縮檔案集合.SelectedIndex + 1, True)
                Form1.壓縮檔案集合.SetSelected(Form1.壓縮檔案集合.SelectedIndex, False)
                'obj.SelectedIndices.Remove(obj.SelectedIndex - 1)
            End If
        End If

        If Form1.壓縮檔案集合.Items.Count = 0 And Form1.壓縮檔案集合.SelectedItem = Nothing Then
            Form1.file_Preview(sender, e)
        Else
            Form1.Archive_Preview(sender, e)
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Form1.Next_Page(sender, e)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Form1.Privious_Page(sender, e)
    End Sub
    Private Sub PictureBox1_PictureChange(sender As Object, e As EventArgs) Handles PictureBox1.BackgroundImageChanged
        TextBox1_TextChanged(sender, e)
    End Sub
    Public Function picture_set(ori_image)
        Dim width As Integer = ori_image.Width
        Dim height As Integer = ori_image.Height
        Dim range = Convert.ToInt32(TextBox1.Text) - 100
        Dim Xzoom_range As Integer = width * range / 100
        Dim Yzoom_range As Integer = height * range / 100
        width -= Xzoom_range * 2
        height -= Yzoom_range * 2
        Dim out = New Integer() {Xzoom_range, Yzoom_range, width, height}
        Return out
    End Function
    Private Sub PictureBox1_DragDrug(sender As Object, e As DragEventArgs) Handles PictureBox1.DragDrop
        Dim ori_image = Form1.PictureBox1.Image
        Dim point = picture_set(ori_image)
        Dim moveX = e.X
        Dim moveY = e.Y
        Dim rect As Rectangle = New Rectangle(point(0) + moveX, point(1) + moveY, point(2) + moveX, point(3) + moveY)
        Dim boximage As Bitmap = CopyRectangle(ori_image, rect)
        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
        PictureBox1.Image = CType(boximage, Image)
    End Sub

    Private Sub PictureBox1_MouseWheel(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseWheel
        If e.Delta > 0 Then
            TextBox1.Text = (Convert.ToInt32(TextBox1.Text) + 5).ToString
        Else
            TextBox1.Text = (Convert.ToInt32(TextBox1.Text) - 5).ToString
        End If
    End Sub
End Class