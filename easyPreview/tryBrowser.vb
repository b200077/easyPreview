

Imports System.IO
Imports System.Net
Imports System.Reflection.Emit
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Menu
Imports DiscUtils
Imports ImageMagick
Imports Microsoft.WindowsAPICodePack.Shell

Public Class tryBrowser
    Dim beginIndex As Integer = 0
    Dim endIndex As Integer = Form1.fileCollection.Items.Count - 1
    Dim scrollPosition As Integer

    ' 在表單的Activated事件中，設定滾軸的值為之前保存的值

    Private Async Sub TryBrowser_LoadWebImage(sender As Object, e As EventArgs) Handles MyBase.Load
        If Not Form1.PathComboBox.Text.Contains("twitter") Then Return
        Dim range As String() = Form1.orderComboBox.Text.Split("~")
        If range.Count = 2 Then
            beginIndex = Convert.ToInt32(range(0))
            endIndex = Convert.ToInt32(range(1))
        End If
        For i = beginIndex To endIndex
            If Visible = False Then Exit For
            If i > Form1.FileCollection.Items.Count - 1 Then Exit For
            Dim add_panel As New UserControl1
            add_panel.CheckBox1.Text = i.ToString 'Form1.目錄檔案集合.Items(i)
            Call Form2.WebMode(Form1.FileCollection.Items(i))
            Dim source3 As New List(Of String)
            Dim tar_index As Integer = 2
            Dim stopwatch As New Stopwatch()
            stopwatch.Start()

            While True

                source3 = Await Form2.GetWebImageSrc("css-9pa8cd", 4)
                If source3.Count > tar_index Then
                    If source3(tar_index) <> "null" Then
                        Exit While
                    End If
                End If


                Await Form2.WaitForLoad()

                ' do something in the loop
                Dim elapsed As TimeSpan = stopwatch.Elapsed
                Console.WriteLine(elapsed)
                Dim minutes As Double = elapsed.TotalMilliseconds
                Console.WriteLine(minutes) ' 90
                If minutes > 5000 And source3.Count < tar_index Then Exit While
                If minutes > 2000 Then
                    Dim pageAvailable As String = Await Form2.WebView21.ExecuteScriptAsync("document.documentElement.outerHTML;")
                    If pageAvailable.Contains("此頁面不存在") Then Exit While
                    If pageAvailable.Contains("此推文來自遭到停用的帳戶。") Then Exit While
                End If

            End While
            stopwatch.Stop()
            If source3.Count < tar_index + 1 Then Continue For
            ' 創建一個WebClient類別的實例
            Dim tClient As New WebClient()
            source3(tar_index) = source3(tar_index).Replace("""", "")
            ' 使用DownloadData方法下載網路上的圖片，並將其轉換為Byte陣列
            Dim tImage As Byte() = tClient.DownloadData(source3(tar_index))
            ' 使用MemoryStream類別將Byte陣列轉換為Bitmap類別的實例
            Dim bmp As Bitmap = Image.FromStream(New MemoryStream(tImage))
            ' 設定PictureBox控制項的Image屬性
            add_panel.PictureBox1.Image = bmp
            FlowLayoutPanel2.Controls.Add(add_panel)
        Next
    End Sub
    Private Async Sub TryBrowser_LoadPixivImage(sender As Object, e As EventArgs) Handles MyBase.Load
        If Not Form1.PathComboBox.Text.Contains("pixiv") Then Return
        Dim range As String() = Form1.orderComboBox.Text.Split("~")

        If range.Count = 2 Then
            beginIndex = Convert.ToInt32(range(0))
            endIndex = Convert.ToInt32(range(1))
        End If
        For i = beginIndex To endIndex
            If Visible = False Then Exit For
            If i > Form1.FileCollection.Items.Count - 1 Then Exit For
            Dim target As String = Form1.FileCollection.Items(i)
            Dim add_panel As New UserControl1
            add_panel.CheckBox1.Text = i.ToString 'Form1.目錄檔案集合.Items(i)
            Call Form2.WebMode(target)
            Dim source3 As New List(Of String)
            Dim tar_index As Integer = 0
            Dim stopwatch As New Stopwatch()
            stopwatch.Start()

            While True
                Dim pageAvailable As String = Await Form2.WebView21.ExecuteScriptAsync("document.documentElement.outerHTML;")
                If pageAvailable.Contains("此頁面不存在") Then Exit While
                source3 = Await Form2.GetWebImageSrc("sc-1qpw8k9-1 jOmqKq", 4)
                If source3.Count > tar_index + 1 Then
                    If source3(tar_index) <> "null" Then
                        Exit While
                    End If
                End If

                Await Form2.WaitForLoad()
                ' do something in the loop
                Dim elapsed As TimeSpan = stopwatch.Elapsed
                Console.WriteLine(elapsed)
                Dim minutes As Double = elapsed.TotalMilliseconds
                Console.WriteLine(minutes) ' 90
                If minutes > 5000 Then Exit While
            End While

            stopwatch.Stop()
            If source3.Count < tar_index + 1 Then Continue For
            ' 創建一個WebClient類別的實例
            Console.WriteLine(source3(tar_index))
            Dim tClient As New WebClient()
            source3(tar_index) = source3(tar_index).Replace("""", "")
            tClient.Headers.Add("Referer", target)
            ' 使用DownloadData方法下載網路上的圖片，並將其轉換為Byte陣列
            Dim tImage As Byte() = tClient.DownloadData(source3(tar_index))

            ' 使用MemoryStream類別將Byte陣列轉換為Bitmap類別的實例
            Dim bmp As Bitmap = Bitmap.FromStream(New MemoryStream(tImage))

            ' 設定PictureBox控制項的Image屬性
            add_panel.PictureBox1.Image = bmp
            FlowLayoutPanel2.Controls.Add(add_panel)
        Next
    End Sub

    Private Async Sub TryBrowser_LoadNhentaiImage(sender As Object, e As EventArgs) Handles MyBase.Load
        If Not Form1.PathComboBox.Text.Contains("nhentai") Then Return
        Dim range As String() = Form1.orderComboBox.Text.Split("~")

        If range.Count = 2 Then
            beginIndex = Convert.ToInt32(range(0))
            endIndex = Convert.ToInt32(range(1))
        End If
        For i = beginIndex To endIndex
            If Visible = False Then Exit For
            If i > Form1.FileCollection.Items.Count - 1 Then Exit For
            Dim target As String = Form1.FileCollection.Items(i)
            Dim add_panel As New UserControl1
            add_panel.CheckBox1.Text = i.ToString 'Form1.目錄檔案集合.Items(i)
            Call Form2.WebMode(target)
            Dim source3 As New List(Of String)
            Dim tar_index As Integer = 0
            Dim stopwatch As New Stopwatch()
            stopwatch.Start()

            While True
                Dim pageAvailable As String = Await Form2.WebView21.ExecuteScriptAsync("document.documentElement.outerHTML;")
                If pageAvailable.Contains("此頁面不存在") Then Exit While
                source3 = Await Form2.GetWebImageSrc("https://t5.nhentai.net", 2)
                If source3.Count > tar_index + 1 Then
                    If source3(tar_index) <> "null" Then
                        Exit While
                    End If
                End If

                Await Form2.WaitForLoad()
                ' do something in the loop
                Dim elapsed As TimeSpan = stopwatch.Elapsed
                Console.WriteLine(elapsed)
                Dim minutes As Double = elapsed.TotalMilliseconds
                Console.WriteLine(minutes) ' 90
                If minutes > 10000 Then Exit While
            End While

            stopwatch.Stop()
            If source3.Count < tar_index + 1 Then Continue For
            ' 創建一個WebClient類別的實例
            Console.WriteLine(source3(tar_index))
            Dim tClient As New WebClient()
            source3(tar_index) = source3(tar_index).Replace("""", "")
            tClient.Headers.Add("Referer", target)
            ' 使用DownloadData方法下載網路上的圖片，並將其轉換為Byte陣列
            Dim tImage As Byte() = tClient.DownloadData(source3(tar_index))

            ' 使用MemoryStream類別將Byte陣列轉換為Bitmap類別的實例
            Dim bmp As Bitmap = Bitmap.FromStream(New MemoryStream(tImage))

            ' 設定PictureBox控制項的Image屬性
            add_panel.PictureBox1.Image = bmp
            FlowLayoutPanel2.Controls.Add(add_panel)
        Next
    End Sub
    Private Sub TryBrowser_LoadLocationImage(sender As Object, e As EventArgs) Handles MyBase.Load
        If Form1.PathComboBox.Text.Contains("twitter") Or Form1.PathComboBox.Text.Contains("pixiv") Then Return
        If Not File.Exists(Form1.PathComboBox.Text) AndAlso Not Directory.Exists(Form1.PathComboBox.Text) Then Return
        MagickNET.Initialize()
        Dim range As String() = Form1.orderComboBox.Text.Split("~")
        If range.Count = 2 Then
            Integer.TryParse(range(0), beginIndex)
            Integer.TryParse(range(1), endIndex)
        End If
        For i = beginIndex To endIndex
            If Visible = False Then Exit For
            If i > Form1.FileCollection.Items.Count - 1 Then Exit For
            Dim add_panel As New UserControl1
            add_panel.CheckBox1.Text = i.ToString 'Form1.目錄檔案集合.Items(i)
            Dim filename As String = Form1.FileCollection.Items(i)
            Dim ext = Path.GetExtension(filename)
            If Not File.Exists(filename) Then Continue For
            If Form1.Ext_judge(ext, "其他圖片") Then
                Try

                    Using img As New MagickImage(filename)
                        Using memStream As New MemoryStream()
                            img.Format = MagickFormat.Bmp
                            img.Write(memStream)
                            add_panel.PictureBox1.Image = New Bitmap(memStream)
                        End Using
                    End Using
                Catch ex As Exception
                End Try
            ElseIf Form1.Ext_judge(ext, "Images") Then
                Try
                    Using bmp As New Bitmap(filename)
                        add_panel.PictureBox1.Image = CType(bmp.Clone, Image)
                    End Using
                Catch ex As Exception
                End Try
            End If
            FlowLayoutPanel2.Controls.Add(add_panel)
        Next
    End Sub
    ' 在表單的Activated事件中，設定滾軸的值為之前保存的值
    Private Sub Form_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        FlowLayoutPanel2.HorizontalScroll.Value = scrollPosition
    End Sub
    ' 定義FlowLayoutPanel的Scroll事件處理器，用於記錄滾軸的位置
    Private Sub HandleScroll(sender As Object, e As ScrollEventArgs) Handles FlowLayoutPanel2.Scroll
        scrollPosition = e.NewValue
    End Sub

End Class