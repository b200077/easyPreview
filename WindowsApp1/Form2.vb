Imports System.Drawing.Imaging
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Menu
Imports AngleSharp.Html.Dom
Imports Aspose.Slides
Imports CefSharp
Imports CefSharp.DevTools
Imports CefSharp.DevTools.DOM
Imports CefSharp.Handler
Imports CefSharp.WinForms
Imports Microsoft.UI.Xaml.Controls
Imports Microsoft.VisualBasic.FileIO
Imports Microsoft.Web.WebView2.Core
Imports Pango
Imports VisioForge.Libs.Accord.Math
Imports VisioForge.Libs.MediaFoundation.OPM
Imports VisioForge.Libs.NDI
Imports VisioForge.Libs.ZXing


'Public Class Browser
'    Public Property Page As ChromiumWebBrowser
'    Public Property RequestContext As RequestContext
'    Private manualResetEvent As ManualResetEvent = New ManualResetEvent(False)


'    Public Sub New()

'        Dim settings = New CefSettings() With {
'            .CachePath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\cef"
'            }    ' .CachePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CefSharp\Cache")


'        CefSharpSettings.ShutdownOnExit = True
'        Cef.Initialize(settings, performDependencyCheck:=True, browserProcessHandler:=Nothing)
'        RequestContext = New RequestContext()
'        Page = New ChromiumWebBrowser("", Nothing)
'        Page.RequestContext = RequestContext
'        PageInitialize()
'    End Sub

'    Public Sub OpenUrl(ByVal url As String)
'        Try
'            AddHandler Page.LoadingStateChanged, AddressOf PageLoadingStateChanged

'            If Page.IsBrowserInitialized Then
'                Page.Load(url)
'                Dim isSignalled As Boolean = manualResetEvent.WaitOne(TimeSpan.FromSeconds(60))
'                manualResetEvent.Reset()

'                If Not isSignalled Then
'                    Page.[Stop]()
'                End If
'            End If

'        Catch __unusedObjectDisposedException1__ As ObjectDisposedException
'        End Try

'        RemoveHandler Page.LoadingStateChanged, AddressOf PageLoadingStateChanged
'    End Sub

'    Private Sub PageLoadingStateChanged(ByVal sender As Object, ByVal e As LoadingStateChangedEventArgs)
'        If Not e.IsLoading Then
'            manualResetEvent.[Set]()
'        End If
'    End Sub

'    Private Sub PageInitialize()
'        SpinWait.SpinUntil(Function() Page.IsBrowserInitialized)
'    End Sub
'End Class
Public Class Form2
    Public Index As Integer = 0

    Public webCompeleteLoading As Boolean = True
    Public webFrameLoadingEnd As Boolean = True
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        TextBox1.Text = (Convert.ToInt32(TextBox1.Text) + 5).ToString
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        TextBox1.Text = (Convert.ToInt32(TextBox1.Text) - 5).ToString
    End Sub
    Public Function CopyRectangle(ByVal sourceImage As Image, ByVal area As Drawing.Rectangle) As Image
        Dim outPut As New Bitmap(area.Width, area.Height)
        Dim DescREctangle As Drawing.Rectangle = New Drawing.Rectangle(0, 0, outPut.Width, outPut.Height)
        Dim g As Graphics = Drawing.Graphics.FromImage(outPut)
        g.DrawImage(sourceImage, DescREctangle, area, GraphicsUnit.Pixel)
        g.Save()
        Return outPut
    End Function
    Private Sub Form2_Base_load() Handles MyBase.Load


        Adjustment_Form()
        setbase_contextmenustrip()


    End Sub

    Private Sub setbase_contextmenustrip()
        Dim rightHotKey = New ToolStripMenuItem("複製圖像")
        AddHandler rightHotKey.Click, AddressOf Clipboard_SetIImage
        ContextMenuStrip1.Items.Add(rightHotKey)
        rightHotKey = New ToolStripMenuItem("複製地址")
        AddHandler rightHotKey.Click, AddressOf Clipboard_SetText
        ContextMenuStrip1.Items.Add(rightHotKey)
        rightHotKey = New ToolStripMenuItem("剪下")
        AddHandler rightHotKey.Click, AddressOf text_operate
        ContextMenuStrip2.Items.Add(rightHotKey)
        rightHotKey = New ToolStripMenuItem("複製")
        AddHandler rightHotKey.Click, AddressOf text_operate
        ContextMenuStrip2.Items.Add(rightHotKey)
        rightHotKey = New ToolStripMenuItem("貼上")
        AddHandler rightHotKey.Click, AddressOf text_operate
        ContextMenuStrip2.Items.Add(rightHotKey)
        rightHotKey = New ToolStripMenuItem("刪除")
        AddHandler rightHotKey.Click, AddressOf text_operate
        ContextMenuStrip2.Items.Add(rightHotKey)
        ContextMenuStrip2.Items.Add(New ToolStripSeparator())
        rightHotKey = New ToolStripMenuItem("全選")
        AddHandler rightHotKey.Click, AddressOf text_operate
        ContextMenuStrip2.Items.Add(rightHotKey)
        ContextMenuStrip2.Items.Add(New ToolStripSeparator())
        rightHotKey = New ToolStripMenuItem("加入連結")
        AddHandler rightHotKey.Click, AddressOf form1List_operate
        ContextMenuStrip2.Items.Add(rightHotKey)
        rightHotKey = New ToolStripMenuItem("取代連結")
        AddHandler rightHotKey.Click, AddressOf form1List_operate
        ContextMenuStrip2.Items.Add(rightHotKey)
    End Sub
    Private Sub text_operate(sender As Object, e As EventArgs)
        Dim btn_name As String = CType(sender, ToolStripMenuItem).Text
        Dim operate_text As String = TextBox3.SelectedText
        Select Case btn_name
            Case "剪下"
                TextBox3.Cut()
            Case "複製"
                TextBox3.Copy()
            Case "貼上"
                TextBox3.Paste()
            Case "刪除"
                If String.IsNullOrEmpty(TextBox3.SelectedText.ToString) Then
                    TextBox3.Text = ""
                    Return
                End If
                TextBox3.Text.Replace(TextBox3.SelectedText, "")
            Case "全選"
                TextBox3.SelectAll()
        End Select
    End Sub
    Private Sub form1List_operate(sender As Object, e As EventArgs)
        Dim btn_name As String = CType(sender, ToolStripMenuItem).Text
        Select Case btn_name
            Case "加入連結"
                Form1.目錄檔案集合.Items.Add(TextBox3.Text)
            Case "取代連結"
                Dim index As Short = Form1.目錄檔案集合.SelectedIndex
                If index = -1 Then Return
                Form1.目錄檔案集合.Items(index) = TextBox3.Text
        End Select
        Form1.delete_repeat_item()
        Form1.refresh_backup()
    End Sub
    Public Sub Clipboard_SetText()

        '加1個把數組轉成換行文字
        Clipboard.SetText(Form1.目錄檔案集合.SelectedItem.ToString)
    End Sub
    Private Sub Clipboard_SetIImage()
        '加1個把數組轉成換行文字
        Clipboard.SetData(DataFormats.Bitmap, PictureBox1.Image)
    End Sub
    Public Delegate Sub UpdateTextBoxDelegate(controls As System.Windows.Forms.TextBox, value As String)
    Private Function setTextboxText(controls As System.Windows.Forms.TextBox, value As String)
        If (controls.InvokeRequired) Then
            Dim del As UpdateTextBoxDelegate = New UpdateTextBoxDelegate(AddressOf setTextboxText)
            controls.Invoke(del, controls, value)
        Else
            controls.Text = value
        End If
        Return Nothing
    End Function

    Private Sub Adjustment_Form()
        If Not PictureBox1.Visible Then Return
        Dim imageProportion As Single = PictureBox1.Image.Width / PictureBox1.Image.Height
        Dim standard As Drawing.Rectangle = Screen.PrimaryScreen.WorkingArea
        WebView21.Width = standard.Width * 2 / 3
        WebView21.Height = standard.Height / 3
        Dim form2Width As Single = standard.Height * imageProportion
        Size = New Size(form2Width, standard.Height)
        Dim locationX As Integer = standard.Width - form2Width
        If locationX < 0 Then locationX = 0
        Location = New Drawing.Point(locationX, 0)
    End Sub
    Public Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If WebView21.Visible Then Return
        Dim ori_image = Form1.PictureBox1.Image
        If ori_image Is Nothing Then Return
        Dim point = picture_set(ori_image)
        Dim rect As Drawing.Rectangle = New Drawing.Rectangle(point(0), point(1), point(2), point(3))
        Dim boximage As Bitmap = CopyRectangle(ori_image, rect)
        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
        PictureBox1.Image = CType(boximage, Image)
    End Sub
    Private Sub startPreview(sender, e)
        If Form1.壓縮檔案集合.Items.Count = 0 And Form1.壓縮檔案集合.SelectedItem = Nothing Then
            Form1.file_Preview(sender, e)
        Else
            Form1.Archive_Preview(sender, e)
        End If
    End Sub
    'Private Sub textBoxTest_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles TextBox2.Enter
    '    If e.KeyCode = Keys.Enter Then
    '        Button3_Click(sender, New EventArgs())
    '    End If
    'End Sub

    Public Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, Button4.Click

        Dim oriIndex As Integer
        Dim targetList As ListBox
        If Form1.壓縮檔案集合.Items.Count = 0 And Form1.壓縮檔案集合.SelectedItem = Nothing Then

            targetList = Form1.目錄檔案集合
        Else
            targetList = Form1.壓縮檔案集合
        End If
        'legitimate check

        Select Case sender.Name
            Case "Button3"
                If Index = targetList.Items.Count - 1 Then Return
            Case "Button4"
                If Index = 0 Then Return
        End Select
        oriIndex = CInt(targetList.SelectedIndex)
        'change index
        Select Case sender.Name
            Case "Button3"
                Index += 1
            Case "Button4"
                Index -= 1
        End Select
        'If oriIndex = Index Then
        '    startPreview(sender, e)
        '    PictureBox1.Image = New Bitmap(targetList.SelectedItem.ToString)
        '    Return
        'End If

        targetList.SetSelected(Index, True)
        targetList.SetSelected(oriIndex, False)
        TextBox2.Text = Index.ToString
        'startPreview(sender, e)
        'Console.WriteLine(targetList.SelectedItem.ToString)
        'PictureBox1.Image = CType(Form1.PictureBox1.Image.Clone, Image)
        'targetList.SetSelected(oriIndex, True)
        'targetList.SetSelected(Index, False)
        'startPreview(sender, e)
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged, MyBase.KeyPress
        Dim textWidth As Integer = TextRenderer.MeasureText(TextBox2.Text, TextBox2.Font).Width
        TextBox1.Width = textWidth + 10 ' Add some padding to the width
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
        Dim out = New Integer() {Xzoom_range - newPoint.X, Yzoom_range - newPoint.Y, width, height}
        Return out
    End Function
    Public IsDragging As Boolean = False
    Public oriPoint As Drawing.Point = New Drawing.Point(0, 0)
    Public newPoint As Drawing.Point = New Drawing.Point(0, 0)
    Private Sub PictureBox1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseDown
        IsDragging = True
        oriPoint = e.Location
    End Sub
    Private Sub PictureBox1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseUp
        IsDragging = False
        newPoint += e.Location - oriPoint
    End Sub
    Public Sub PictureMode()
        'ChromiumWebBrowser1.Hide()
        WebView21.Hide()
        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
        PictureBox1.Image = Form1.PictureBox1.Image
        PictureBox1.Show()
        TextBox3.Hide()
    End Sub
    Public Async Sub WebMode(url As String)
        PictureBox1.Hide()

        If Not String.IsNullOrEmpty(url) Then
            ' ChromiumWebBrowser1.LoadUrlAsync(url)
            'ChromiumWebBrowser1.Show()

            '確保webview2已經初始化
            Await WebView21.EnsureCoreWebView2Async(Nothing)
            ' 訪問網址
            WebView21.CoreWebView2.Navigate(url)
            WebView21.CoreWebView2.Settings.IsZoomControlEnabled = True
            WebView21.Show()
        End If

        'ChromiumWebBrowser1.Dock = DockStyle.Top
        ' PictureBox1.Dock = DockStyle.None
        TextBox3.Show()


    End Sub
    Private Sub PictureBox1_DragEnter(sender As System.Object, e As System.Windows.Forms.DragEventArgs) Handles Panel1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub
    Private Sub PictureBox1_DragDrop(sender As System.Object, e As System.Windows.Forms.DragEventArgs) Handles Panel1.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        For Each sourcrPath In files
            If Form1.目錄檔案集合.Items.Contains(sourcrPath) Then Form1.目錄檔案集合.Items.Remove(sourcrPath)
            Form1.目錄檔案集合.Items.Add(sourcrPath)
            If Not String.IsNullOrEmpty(Form1.TextBox5.Text) Then Form1.backup.Add(sourcrPath)
        Next
        Form1.refresh_backup()
    End Sub

    Private Sub PictureBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove
        If IsDragging Then
            Dim ori_image = Form1.PictureBox1.Image
            Dim point = picture_set(ori_image)
            Dim moveX = e.X - oriPoint.X
            Dim moveY = e.Y - oriPoint.Y
            Dim rect As Drawing.Rectangle = New Drawing.Rectangle(point(0) - moveX, point(1) - moveY, point(2), point(3))
            Dim boximage As Bitmap = CopyRectangle(ori_image, rect)
            PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
            PictureBox1.Image = CType(boximage, Image)
        End If
    End Sub

    Private Sub refresh_web() Handles WebView21.NavigationCompleted
        TextBox3.Text = WebView21.Source.AbsoluteUri
    End Sub

    Public Sub checkWebSite()
        Dim weburl = WebView21.Source.AbsoluteUri
        If webUrl.Contains("twitter") Then
            crawler("twitter")
        ElseIf webUrl.Contains("facebook") Then
            crawler("facebook")
        ElseIf webUrl.Contains("pixiv") Then
            crawler("pixiv")
        ElseIf webUrl.Contains("youtube") Then
            crawler("youtube")
        End If

    End Sub
    Public Sub crawler(webSite As String)
        Dim srcs As New ArrayList
        Dim positions As New ArrayList
        Select Case webSite
            Case "twitter"
                tweetDownload()
                'srcs = Await GetWebImageSrc("pbs.twimg.com/media", 1)
                'positions = Await getPosition(srcs)
                'ClickWebPicures(positions)
            Case "facebook"
                facebookDownload()
                'Await ClickPositionByClass("_39pi")
                'Await ClickPositionByClass("_blank")
                'Dim value As ArrayList = Await GetWebImageSrc("", 2)
                'DownloadOriginalImage(value(0))
            Case "pixiv"
                pixivDownload()
                'srcs = await getwebimagesrc("i.pximg.net", 1)
                'positions = await getposition(srcs)
                'clickwebpicures(positions)
        End Select
    End Sub
    Public Async Sub tweetDownload()


        '--------------Method1------------------
        'Dim scripts = "document.documentElement.outerHTML"
        'Dim result = Await ChromiumWebBrowser1.EvaluateScriptAsync(scripts)
        'Dim source1 As String = result.Result.ToString

        '--------------Method2------------------
        'Dim devTool As DevToolsClient = ChromiumWebBrowser1.GetDevToolsClient
        'Dim response As DevToolsMethodResponse = Await devTool.ExecuteDevToolsMethodAsync("DOM.getOuterHTML")
        'Dim source2 = response.ResponseAsJsonString

        '--------------Method3------------------
        'Dim scripts = "document.documentElement.outerHTML"
        'Dim result = Await ChromiumWebBrowser1.EvaluateScriptAsync(scripts)
        'Dim source3 As String = result.Result.ToString

        If Not WebView21.Source.ToString().Contains("photo/") Then
            Await WebView21.ExecuteScriptAsync("document.getElementsByClassName('css-4rbku5 css-18t94o4 css-1dbjc4n r-1loqt21 r-1pi2tsx r-1ny4l3l')[1].click()")

        End If
        'check if need click

        Dim source3 As ArrayList = Await GetWebImageSrc("css-9pa8cd", 4)
        Dim exceptForm As String() = {"large", "4800x4800", "medium", "900x900", "4096x4096"}
        For Each matched As String In source3
            Console.WriteLine(matched)
            For Each formName In exceptForm
                If matched.Contains(formName) Then DownloadOriginalImage(matched)
            Next
        Next
        '------------Original Method-------------
        'Dim source As String = Await ChromiumWebBrowser1.GetSourceAsync()

        'Dim rgx As Regex = New Regex("""(https:\/\/pbs.twimg.com\/media\/[\w\.\:\?\&\=\;\-]*)""")
        'Dim fn As MatchCollection = rgx.Matches(source)
        'Dim MyHandler As MyCustomMenuHandler = New MyCustomMenuHandler()
        'Dim downloaded As New ArrayList
        'For Each match As Match In fn
        '    Dim value As String = match.Groups(1).Value
        '    If value.Contains("format") Then Continue For
        '    If value.Contains("large") Or value.Contains("4800x4800") Or value.Contains("medium") Then
        '        Console.WriteLine(value)
        '        ' Dim saveFileName As String = MyHandler.getFileName(value)
        '        ' If downloaded.Contains(saveFileName.Split(".")(0)) Then Continue For
        '        DownloadOriginalImage(value)
        '        '  downloaded.Add(saveFileName.Split(".")(0))
        '    End If
        'Next
    End Sub
    Public Async Sub pixivDownload()
        Dim source As String = Await WebView21.CoreWebView2.ExecuteScriptAsync("document.documentElement.outerHTML")
        Dim rgx As Regex = New Regex("href=""(https:\/\/i.pximg.net\/img-original\/img[\d\w\/\.]*)""")
        Dim fn As MatchCollection = rgx.Matches(source)
        For Each match As Match In fn
            Dim value As String = match.Groups(1).Value
            If value.Contains("master1200") Then
                Continue For
            End If
            Console.WriteLine(value)
            DownloadOriginalImage(value)
        Next
    End Sub
    Public Async Sub fanboxDownload()
        Dim source As ArrayList = Await GetWebImageSrc("https://downloads.fanbox.cc/images/post", 3)

        For Each matched As String In source

            Console.WriteLine(matched)
            DownloadOriginalImage(matched)
        Next
    End Sub
    Public Async Sub facebookDownload()
        Dim source As ArrayList = Await GetWebImageSrc("https://scontent.ftpe4", 3)

        For Each matched As String In source

            Console.WriteLine(matched)
            DownloadOriginalImage(matched)
        Next
    End Sub
    Public Async Sub instagramDownload()
        Dim source As ArrayList = Await GetWebImageSrc("https://instagram.ftpe4", 2)
        For Each matched As String In source

            Console.WriteLine(matched)
            DownloadOriginalImage(matched)
        Next
    End Sub
    Public Async Sub youtubeDownload()
        Dim source As ArrayList = Await GetWebImageSrc("https://yt3.ggpht.com", 3)
        'Dim source As String = Await ChromiumWebBrowser1.GetSourceAsync()
        'Dim rgx As Regex = New Regex("src=\'(https:\/\/yt3.ggpht.com\/[\w\.\:\?\&\=\;\-]*)\'")
        'Dim fn As MatchCollection = rgx.Matches(source)
        For Each match As Match In source
            Dim value As String = match.Groups(1).Value
            Console.WriteLine(value)
            DownloadOriginalImage(value)
        Next
    End Sub
    Public Async Sub wmcSearch(keyword As String)
        If ChromiumWebBrowser1.IsBrowserInitialized = False Then Return
        'Await waitForInitial()

        Dim source As String = Await ChromiumWebBrowser1.GetSourceAsync()
        Dim rgx As Regex = New Regex(keyword)
        Dim fn As MatchCollection = rgx.Matches(source)
        Dim matches As New ArrayList
        For Each match As Match In fn
            matches.Add(match.Groups(1).Value)
            Form1.目錄檔案集合.Items.Add(match.Groups(1).Value)
        Next
    End Sub
    Public Async Function GetElement() As Task(Of IHtmlElement)
        Dim browser As IWebBrowser = ChromiumWebBrowser1 ' Replace this with a reference to your CefSharp browser control

        ' Wait for the page to finish loading
        Await browser.EvaluateScriptAsync("void(0);")

        ' Get the element using querySelector
        Dim element As IHtmlElement = Await browser.EvaluateScriptAsync("document.querySelector('input[name=submit]');")

        ' If the element was not found, try using getElementsByName
        If element Is Nothing Then
            element = Await browser.EvaluateScriptAsync("document.getElementsByName('submit')[0];")
        End If

        ' If the element was still not found, try using getElementsByClassName
        If element Is Nothing Then
            element = Await browser.EvaluateScriptAsync("document.getElementsByClassName('searchbutton')[0];")
        End If

        Return element
    End Function
    Public Sub DownloadOriginalImage(src As String)
        If Form1.ComboBox1.Text.Contains("noDownload") Then Return

        Dim MyHandler As MyCustomMenuHandler = New MyCustomMenuHandler()

        Dim saveFolder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString & "\" & System.DateTime.Now.ToString("yyyy_MM") & "\"
        If Form1.ComboBox5.Text = "下載在combobox4中" Then
            saveFolder = Form1.ComboBox4.Text
        End If
        If Not Directory.Exists(saveFolder) Then Directory.CreateDirectory(saveFolder)
        Dim saveFileName = MyHandler.getFileName(src)
        MyHandler.SaveImage(src, saveFolder & saveFileName, ImageFormat.Jpeg)
        Form1.Label11.Text = saveFileName & "下載成功"
        Form1.壓縮檔案集合.Items.Add(saveFolder & saveFileName)


    End Sub



    Private Async Function getPositionsFromClass(className As String) As Task(Of String)
        Dim Script = " 
                      var rect = document.getElementsByClassName(""" & className & """);
                      var curRect = rect.getBoundingClientRect();
                      var center = String(curRect.left + curRect.width / 2) + "","" + String(curRect.top + curRect.height / 2);
                      center;     "
        Dim result As String = Await WebView21.ExecuteScriptAsync(Script)
        Return result
    End Function
    Private Sub Form2_Close(sender As System.Object, e As System.EventArgs) Handles MyBase.FormClosing
        If Not Form1.Visible Then Form1.Close()
    End Sub
    Public Async Function GetWebImageSrc(keyword As String, type As Integer) As Task(Of ArrayList)
        'type1 回傳目錄
        'type2 回傳src
        'type3 如果長*寬 > 250000 回傳src
        'type4 keyword改搜尋class名稱 回傳src
        Dim script As String = "document.querySelectorAll(""img"").length"
        If type = 4 Then script = "document.getElementsByClassName('" & keyword & "').length;"
        Console.WriteLine(script)
        Dim result As String = Await WebView21.ExecuteScriptAsync(script)
        Dim canUseLength As Integer
        Console.WriteLine(canUseLength)
        Dim canUseIndex As New ArrayList

        For i = 0 To canUseLength - 1
            script = "document.querySelectorAll(""img"")[" & i.ToString & "].getAttribute(""src"")"
            If type = 4 Then script = "document.getElementsByClassName('" & keyword & "')[" & i.ToString & "].getAttribute(""src"")"
            Dim script2 = "var curRect = document.querySelectorAll(""img"")[" & i.ToString & "];
                          var area = curRect.naturalWidth * curRect.naturalHeight
                          area"
            Dim script3 = "document.getElementsByClassName('" & keyword & "');"
            result = Await WebView21.ExecuteScriptAsync(script)
            If result = Nothing Then Continue For
            If Not result.Contains(keyword) And type <> 4 Then Continue For
            Console.WriteLine(result)
            If type = 1 Then
                canUseIndex.Add(i)
            ElseIf type = 2 Or type = 4 Then
                canUseIndex.Add(result)
            ElseIf type = 3 Then
                result = Await WebView21.ExecuteScriptAsync(script2)
                If Convert.ToInt32(result) < 250000 Then Continue For
                result = Await WebView21.ExecuteScriptAsync(script)
                canUseIndex.Add(result)
            End If
        Next
        Return canUseIndex
    End Function

    Public Async Function clickByClass(className As String) As Task(Of ArrayList)
        Dim positions As New ArrayList
        Dim result As String
        Dim Script = " 
                      var rect = Document.getElementsByClassName(""" & className & """)
                      var curRect = rect[" & CInt(Index) & "].getBoundingClientRect();
                      var center = String(curRect.left + curRect.width / 2) + "","" + String(curRect.top + curRect.height / 2);
                      center;     "
        result = Await WebView21.ExecuteScriptAsync(Script)
        Console.WriteLine(result)
        positions.Add(result)
        Return positions
    End Function
    Private Sub checkWebLoading() Handles WebView21.NavigationCompleted
        webCompeleteLoading = True
    End Sub

    Public Async Function waitForLoad() As Task(Of Task)
        While True
            Await Task.Delay(10)
            If webCompeleteLoading = True Then
                webCompeleteLoading = False
                Exit While
            End If
        End While
        Return Task.CompletedTask
    End Function


    Private Sub PictureBox1_MouseWheel(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseWheel
        If e.Delta > 0 Then
            TextBox1.Text = (Convert.ToInt32(TextBox1.Text) + 5).ToString
        Else
            TextBox1.Text = (Convert.ToInt32(TextBox1.Text) - 5).ToString
        End If
    End Sub



    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If Form1.Visible Then
            Form1.Hide()
        Else
            Form1.Show()
        End If

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar <> Convert.ToChar(Keys.Enter) Then Return
        If PictureBox1.Visible Then
            Dim oriIndex As Integer
            Dim targetList As ListBox
            If Form1.壓縮檔案集合.Items.Count = 0 And Form1.壓縮檔案集合.SelectedItem = Nothing Then

                targetList = Form1.目錄檔案集合
            Else
                targetList = Form1.壓縮檔案集合
            End If
            'legitimate check

            If CInt(TextBox2.Text) > targetList.Items.Count - 1 Then Return
            If CInt(TextBox2.Text) < 0 Then Return
            oriIndex = CInt(targetList.SelectedIndex)

            Index = CInt(TextBox2.Text)


            targetList.SetSelected(Index, True)
            targetList.SetSelected(oriIndex, False)
            TextBox2.Text = Index.ToString
        ElseIf webview21.Visible Then
            WebView21.CoreWebView2.Navigate(TextBox3.Text)
        End If

    End Sub

    Private Sub TextBox3_TextSelected(sender As Object, e As EventArgs) Handles TextBox3.Click
        If TextBox3.SelectedText.Length <> 0 Then Return
        TextBox3.SelectAll()
    End Sub


End Class