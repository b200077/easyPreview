
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports ImageMagick
Imports LibVLCSharp.Shared
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports Mono.Nat
Imports Newtonsoft.Json

Module EverythingSearcher

    ' Import Everything64.dll functions (use Everything64.dll for 64-bit systems)
    <DllImport("Everything64.dll", CharSet:=CharSet.Unicode)>
    Public Function Everything_SetSearchW(ByVal lpSearchString As String) As Integer
    End Function

    <DllImport("Everything64.dll")>
    Public Function Everything_QueryW(ByVal bWait As Boolean) As Integer
    End Function

    <DllImport("Everything64.dll")>
    Public Function Everything_GetNumResults() As Integer
    End Function

    <DllImport("Everything64.dll", CharSet:=CharSet.Unicode)>
    Public Function Everything_GetResultFullPathNameW(ByVal nIndex As Integer, ByVal lpString As System.Text.StringBuilder, ByVal nMaxCount As Integer) As Integer
    End Function

    ' Public method that can be called externally
    Public Function SearchFiles(ByVal searchFile As String) As List(Of String)
        Dim results As New List(Of String)

        ' 設置查詢字串
        Everything_SetSearchW(searchFile)

        ' 執行查詢
        Everything_QueryW(True)

        ' 獲取結果數量
        Dim numResults As Integer = Everything_GetNumResults()

        ' 檢查是否有結果
        If numResults > 0 Then
            ' 遍歷結果
            For i As Integer = 0 To numResults - 1
                ' 初始化字串來接收文件全路徑
                Dim resultPath As New System.Text.StringBuilder(260) ' 最大長度
                Everything_GetResultFullPathNameW(i, resultPath, resultPath.Capacity)
                Dim pathname As String = resultPath.ToString
                If pathname.Contains("AppData\Roaming\Microsoft\Windows\Recent") And pathname.Contains(".lnk") Then Continue For                 ' 添加結果到列表
                results.Add(pathname)
            Next
        End If

        Return results ' 返回所有找到的文件路徑
    End Function
    Public Function SearchFiles(ByVal folderPath As String, ByVal keyword As String) As List(Of String)
        Dim results As New List(Of String)

        ' 組合查詢字串，指定資料夾並使用 parent: 限定範圍
        Dim query As String = folderPath + keyword
        Console.WriteLine($"query = {query}")
        ' 設置查詢字串
        Everything_SetSearchW(query)

        ' 執行查詢
        Everything_QueryW(True)

        ' 獲取結果數量
        Dim numResults As Integer = Everything_GetNumResults()

        ' 檢查是否有結果
        If numResults > 0 Then
            ' 遍歷結果
            For i As Integer = 0 To numResults - 1
                ' 初始化字串來接收文件全路徑
                Dim resultPath As New System.Text.StringBuilder(260) ' 最大長度
                Everything_GetResultFullPathNameW(i, resultPath, resultPath.Capacity)

                ' 添加結果到列表
                results.Add(resultPath.ToString())
            Next
        End If

        Return results ' 返回所有找到的文件路徑
    End Function
End Module



Public Class Form2
    Public Index As Integer = 0

    Public webCompeleteLoading As Boolean = False
    Public webFrameLoadingEnd As Boolean = True

    ' WebView2 初始化控制
    Private _webView2InitializationTask As Task
    Private _webView2InitializationLock As New Object()

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        TextBox1.Text = (Convert.ToInt32(TextBox1.Text) + 5).ToString
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        TextBox1.Text = (Convert.ToInt32(TextBox1.Text) - 5).ToString
    End Sub
    Public Function CopyRectangle(ByVal sourceImage As Image, ByVal area As Rectangle) As Image
        Dim outPut As New Bitmap(area.Width, area.Height)
        Dim DescREctangle As New Rectangle(0, 0, outPut.Width, outPut.Height)
        Dim g As Graphics = Graphics.FromImage(outPut)
        g.DrawImage(sourceImage, DescREctangle, area, GraphicsUnit.Pixel)
        g.Save()
        Return outPut
    End Function
    Private Sub Form2_Base_load() Handles MyBase.Load

        'AddHandler _mediaPlayer.EndReached, AddressOf Form1.MediaPlayer_EndReached
        'AddHandler Form1._mediaPlayer.TimeChanged, AddressOf OnTimeChange

        Setbase_contextmenustrip()
        InitializeWebView2OnceAsync()

    End Sub
    ''' <summary>
    ''' 確保 WebView2 只初始化一次，避免多重初始化
    ''' </summary>
    Private Async Sub InitializeWebView2OnceAsync()
        ' 使用鎖定確保多執行緒安全
        SyncLock _webView2InitializationLock
            If _webView2InitializationTask Is Nothing Then
                _webView2InitializationTask = InitializeWebView2CoreAsync()
            End If
        End SyncLock

        ' 等待初始化完成
        Await _webView2InitializationTask
    End Sub
    Private Async Function InitializeWebView2CoreAsync() As Task
        Try
            If WebView21 IsNot Nothing AndAlso WebView21.CoreWebView2 Is Nothing Then
                Await WebView21.EnsureCoreWebView2Async(Nothing)
            End If
        Catch ex As Exception
            ' 記錄錯誤但不中斷程式
            Console.WriteLine("WebView2 初始化失敗: " & ex.Message)
        End Try
    End Function
    Private Sub HandleAddLink(linkUri As String, sender As Object, e As EventArgs, insert As Boolean)
        ' 使用 Invoke 確保在 UI 執行緒
        If Me.InvokeRequired Then
            Me.Invoke(Sub() HandleAddLink(linkUri, sender, e, insert))
            Return
        End If

        ' 原本邏輯
        store_trf_records(linkUri, "addLinkRecord", "加入歷史")
        Form1.CheckedListBox1_setCheck("插入選取的下一項", insert)
        Dim nowsort As String = Path.GetFileNameWithoutExtension(Form1.PathComboBox.Text)
        Dim stdLink As String = Form1.standardized_denomination(linkUri)

        If Form1.PathComboBox.Text.Contains("個別網站") AndAlso Not stdLink.Contains(nowsort) Then
            Form3.AutoSort(stdLink, Form1.fileCollection)
        Else

            Form4.form4_load()
            Form4.setdata(stdLink)
            Form4.Button1_Click(sender, e)
        End If
    End Sub
    Private Sub NavigateTarget(targetUrl As String, useWebView2 As Boolean)
        ' 使用 Invoke 確保在 UI 執行緒
        If Me.InvokeRequired Then
            Me.Invoke(Sub() NavigateTarget(targetUrl, useWebView2))
            Return
        End If

        Dim finalUrl As String = targetUrl

        ' 處理 pixiv jump.php 的特殊格式
        If finalUrl.Contains("www.pixiv.net/jump.php") Then
            finalUrl = finalUrl.Replace("https://www.pixiv.net/jump.php?", "")
            If finalUrl.StartsWith("url=") Then
                finalUrl = finalUrl.Substring(finalUrl.IndexOf("url=") + 4)
            End If
            finalUrl = Uri.UnescapeDataString(finalUrl)
        End If

        ' 去除空格與括號
        finalUrl = finalUrl.Replace(" ", "").Replace("(", "").Replace(")", "")

        ' 導向
        If useWebView2 Then
            WebView21.CoreWebView2.Navigate(finalUrl)
        Else
            Process.Start(finalUrl)
        End If
    End Sub
    Private Sub WebView2_ContextMenuRequested(sender As Object, e As CoreWebView2ContextMenuRequestedEventArgs)
        ' Dim deferral = e.GetDeferral()
        Dim browser As WebView2 = WebView21
        Dim tarlist As ListBox = Form1.targetList
        Dim menutarget As CoreWebView2ContextMenuTarget = e.ContextMenuTarget
        Console.WriteLine("menutarget.HasLinkUri:" & menutarget.HasLinkUri)
        Console.WriteLine("menutarget.HasSelection:" & menutarget.HasSelection)
        Dim LinkUri As String = If(menutarget.HasLinkUri, menutarget.LinkUri, String.Empty)
        Dim selectionword As String = If(menutarget.HasSelection, menutarget.SelectionText, String.Empty)
        Dim target As String = If(menutarget.HasLinkUri, menutarget.LinkUri, menutarget.SelectionText)
        Console.WriteLine("LinkUri:" & LinkUri)
        Console.WriteLine("selectionword:" & selectionword)
        Console.WriteLine("target:" & target)
        ' add new item to end of collection

        Dim menuList As IList(Of CoreWebView2ContextMenuItem) = e.MenuItems


        Dim addlink As CoreWebView2ContextMenuItem = browser.CoreWebView2.Environment.CreateContextMenuItem("加入連結", Nothing, CoreWebView2ContextMenuItemKind.Command)

        AddHandler addlink.CustomItemSelected, Sub(s, ex)
                                                   e.Handled = True
                                                   HandleAddLink(LinkUri, s, ex, False)
                                               End Sub
        menuList.Insert(0, addlink)
        Dim insertlink As CoreWebView2ContextMenuItem = browser.CoreWebView2.Environment.CreateContextMenuItem("加入下一項", Nothing, CoreWebView2ContextMenuItemKind.Command)

        AddHandler insertlink.CustomItemSelected, Sub(s, ex)
                                                      e.Handled = True
                                                      HandleAddLink(LinkUri, s, ex, True)
                                                  End Sub
        menuList.Insert(1, insertlink)
        'Dim dlImg As CoreWebView2ContextMenuItem = browser.CoreWebView2.Environment.CreateContextMenuItem("下載圖片", Nothing, CoreWebView2ContextMenuItemKind.Command)

        'AddHandler dlImg.CustomItemSelected, Sub(send, ex)
        '                                         Me.Invoke(Sub()
        '                                                       WebView21.CoreWebView2.Navigate(LinkUri)
        '                                                       checkWebSite()
        '                                                   End Sub)
        '                                     End Sub
        'menuList.Insert(2, dlImg)
        Dim openLink As CoreWebView2ContextMenuItem = browser.CoreWebView2.Environment.CreateContextMenuItem("開啟連結", Nothing, CoreWebView2ContextMenuItemKind.Command)
        AddHandler openLink.CustomItemSelected, Sub(s, ex)
                                                    e.Handled = True
                                                    NavigateTarget(target, False)
                                                End Sub
        menuList.Insert(2, openLink)
        Dim goLink As CoreWebView2ContextMenuItem = browser.CoreWebView2.Environment.CreateContextMenuItem("前往", Nothing, CoreWebView2ContextMenuItemKind.Command)
        AddHandler goLink.CustomItemSelected, Sub(s, ex)
                                                  e.Handled = True
                                                  NavigateTarget(target, True)
                                              End Sub
        menuList.Insert(3, goLink)
        Dim custom As CoreWebView2ContextMenuItem = browser.CoreWebView2.Environment.CreateContextMenuItem("自訂", Nothing, CoreWebView2ContextMenuItemKind.Submenu)
        menuList.Insert(4, custom)

        'Dim copyLink As CoreWebView2ContextMenuItem = browser.CoreWebView2.Environment.CreateContextMenuItem("複製連結", Nothing, CoreWebView2ContextMenuItemKind.Command)
        'AddHandler copyLink.CustomItemSelected, Sub(send, ex)
        '                                            target = target.Replace(" ", "").Replace("(", "").Replace(")", "")
        '                                            Clipboard.SetText(target)
        '                                            e.Handled = True
        '                                        End Sub
        'custom.Children.Add(copyLink)

        Dim replaceLink As CoreWebView2ContextMenuItem = browser.CoreWebView2.Environment.CreateContextMenuItem("取代連結", Nothing, CoreWebView2ContextMenuItemKind.Command)
        AddHandler replaceLink.CustomItemSelected, Sub(send, ex)
                                                       'Dim menutarget As CoreWebView2ContextMenuTarget = e.ContextMenuTarget
                                                       Me.Invoke(Sub()
                                                                     If tarlist.SelectedIndex = -1 Then Return
                                                                     tarlist.Items(tarlist.SelectedIndex) = target
                                                                     Form1.refresh_backup()
                                                                 End Sub)
                                                   End Sub
        custom.Children.Add(replaceLink)
        'Dim clearCache As CoreWebView2ContextMenuItem = browser.CoreWebView2.Environment.CreateContextMenuItem("清除快取", Nothing, CoreWebView2ContextMenuItemKind.Command)
        'AddHandler newItem.CustomItemSelected, Sub(send, ex)
        '                                           ClearBrowsingDataAsync()
        '                                       End Sub
        'custom.Children.Add(clearCache)
        Dim search As CoreWebView2ContextMenuItem = browser.CoreWebView2.Environment.CreateContextMenuItem("搜尋", Nothing, CoreWebView2ContextMenuItemKind.Submenu)
        menuList.Insert(5, search)
        Dim searchGoogle_word As CoreWebView2ContextMenuItem = browser.CoreWebView2.Environment.CreateContextMenuItem("搜尋文字google", Nothing, CoreWebView2ContextMenuItemKind.Command)
        AddHandler searchGoogle_word.CustomItemSelected, Sub(send, ex)
                                                             Console.WriteLine("selectionword:" & selectionword)
                                                             If Not String.IsNullOrEmpty(selectionword) Then
                                                                 ' 拼接 Bing 搜索的網址
                                                                 Dim url As String = "https://www.google.com/search?q=" & selectionword
                                                                 ' 使用 Process 類打開默認瀏覽器並訪問該網址
                                                                 browser.CoreWebView2.Navigate(url)
                                                             End If
                                                         End Sub
        search.Children.Add(searchGoogle_word)

        Dim searchLocal_word As CoreWebView2ContextMenuItem = browser.CoreWebView2.Environment.CreateContextMenuItem("搜尋文字本機", Nothing, CoreWebView2ContextMenuItemKind.Command)
        AddHandler searchLocal_word.CustomItemSelected, Sub(send, ex)
                                                            Console.WriteLine("selectionword:" & selectionword)
                                                            If Not String.IsNullOrEmpty(selectionword) Then
                                                                Me.Invoke(Sub()
                                                                              Dim files As List(Of String) = SearchFiles(selectionword)
                                                                              For Each file In files
                                                                                  Form1.subFileCollection.Items.Add(file)
                                                                              Next
                                                                          End Sub)

                                                            End If
                                                        End Sub
        search.Children.Add(searchLocal_word)

        ' 創建自定義項目
        Dim searchAscii2d As CoreWebView2ContextMenuItem = browser.CoreWebView2.Environment.CreateContextMenuItem("搜尋圖片ascii2d", Nothing, CoreWebView2ContextMenuItemKind.Command)
        ' 添加點擊事件
        AddHandler searchAscii2d.CustomItemSelected, Sub(send, args)
                                                         If e.ContextMenuTarget Is Nothing Then Exit Sub
                                                         ' 取得選中的項目類型
                                                         Dim targetKind As CoreWebView2ContextMenuTargetKind = e.ContextMenuTarget.Kind

                                                         ' 檢查是否點擊的是圖片
                                                         If targetKind = CoreWebView2ContextMenuTargetKind.Image Then
                                                             Dim imageUrl As String = e.ContextMenuTarget.SourceUri
                                                             browser.CoreWebView2.Navigate($"https://ascii2d.net/search/url/{imageUrl}")
                                                         End If

                                                     End Sub
        ' 將自定義項目添加到上下文選單
        search.Children.Add(searchAscii2d) ' 插入到選單的最上方
        Dim searchGoogle_pic As CoreWebView2ContextMenuItem = browser.CoreWebView2.Environment.CreateContextMenuItem("搜尋圖片google", Nothing, CoreWebView2ContextMenuItemKind.Command)
        ' 添加點擊事件
        AddHandler searchGoogle_pic.CustomItemSelected, Sub(send, args)
                                                            If e.ContextMenuTarget Is Nothing Then Exit Sub
                                                            ' 取得選中的項目類型
                                                            Dim targetKind As CoreWebView2ContextMenuTargetKind = e.ContextMenuTarget.Kind

                                                            ' 檢查是否點擊的是圖片
                                                            If targetKind = CoreWebView2ContextMenuTargetKind.Image Then
                                                                Dim imageUrl As String = e.ContextMenuTarget.SourceUri
                                                                browser.CoreWebView2.Navigate($"https://lens.google.com/uploadbyurl?url={imageUrl}")
                                                            End If

                                                        End Sub
        ' 將自定義項目添加到上下文選單
        search.Children.Add(searchGoogle_pic) ' 插入到選單的最上方
        'deferral.Complete()
    End Sub

    Private Async Sub Form2_shown() Handles MyBase.Shown
        Adjustment_Form()
        ' WebView2 在 Form_Load 時已初始化，此處只需確保已完成
        If _webView2InitializationTask IsNot Nothing Then
            Await _webView2InitializationTask
        End If

    End Sub
    Private contextMenuSubscribed As Boolean = False
    Private Sub WebView21CoreWebView2Initialization() Handles WebView21.CoreWebView2InitializationCompleted
        If WebView21.CoreWebView2 IsNot Nothing AndAlso Not contextMenuSubscribed Then
            AddHandler WebView21.CoreWebView2.ContextMenuRequested, AddressOf WebView2_ContextMenuRequested
            contextMenuSubscribed = True
        Else
            MessageBox.Show("CoreWebView2 is not initialized.")
        End If
    End Sub
    Sub ClearBrowsingDataAsync()
        ' 確保WebView已初始化並且有一個有效的CoreWebView2
        If WebView21.CoreWebView2 IsNot Nothing Then
            ' 清除所有類型的瀏覽數據
            Call WebView21.CoreWebView2.Profile.ClearBrowsingDataAsync()
        End If
    End Sub

    Public Sub Add_to_record(link_url As String, btn_name As Object)
        Form1.trf_backup(btn_name)
        Using sr As New StreamWriter(btn_name, True)
            sr.WriteLine(link_url)
        End Using
    End Sub
    Private Sub AddItem_from_clipboard(sender, e)
        Form1.AddItem_from_clipboard(Form1.ContextMenuStrip1.Controls(0), e)
    End Sub

    Private Sub Setbase_contextmenustrip()
        Form1.add_ContextMenuStrip(ContextMenuStrip1, "複製圖像", AddressOf Clipboard_SetIImage)
        Form1.add_ContextMenuStrip(ContextMenuStrip1, "複製地址", AddressOf Clipboard_SetText)
        Dim rightHotKey As New ToolStripMenuItem("尋找資源")
        ContextMenuStrip1.Items.Add(rightHotKey)
        For Each opt In Constants.GetValue("search_option")
            Form1.add_subStripMenuItem(rightHotKey, opt, AddressOf FindSource)
        Next
        Form1.add_ContextMenuStrip(ContextMenuStrip2, "剪下", AddressOf Text_operate)
        Form1.add_ContextMenuStrip(ContextMenuStrip2, "複製", AddressOf Text_operate)
        Form1.add_ContextMenuStrip(ContextMenuStrip2, "貼上", AddressOf Text_operate)
        Form1.add_ContextMenuStrip(ContextMenuStrip2, "刪除", AddressOf Text_operate)
        ContextMenuStrip2.Items.Add(New ToolStripSeparator())
        Form1.add_ContextMenuStrip(ContextMenuStrip2, "全選", AddressOf Text_operate)
        ContextMenuStrip2.Items.Add(New ToolStripSeparator())
        Form1.add_ContextMenuStrip(ContextMenuStrip2, "加入連結", AddressOf Form1List_operate)
        Form1.add_ContextMenuStrip(ContextMenuStrip2, "開啟連結", AddressOf Form1List_operate)
        Form1.add_ContextMenuStrip(ContextMenuStrip2, "搜尋", AddressOf Form1List_operate)
        Form1.add_ContextMenuStrip(ContextMenuStrip2, "取代連結", AddressOf Form1List_operate)
    End Sub
    Private Sub Text_operate(sender As Object, e As EventArgs)
        Dim btn_name As String = CType(sender, ToolStripMenuItem).Text
        Dim operate_text As String = TextBox3.SelectedText
        Select Case btn_name
            Case "剪下"
                TextBox3.Cut()
            Case "複製"
                If TextBox3.SelectionLength > 0 Then
                    TextBox3.Copy()
                Else
                    Clipboard.SetText(TextBox3.Text)
                End If
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

    Public Sub Form1List_operate(sender As ToolStripMenuItem, e As EventArgs)
        If sender Is Nothing Then Return
        Dim btn_name As String = sender.Text
        Dim target As String = TextBox3.Text
        Select Case btn_name
            Case "加入連結"
                Dim LinkUri = Form1.standardized_denomination(target)
                Form1.addLink(LinkUri)
            Case "取代連結"
                Dim index As Short = Form1.FileCollection.SelectedIndex
                If index = -1 Then Return
                Form1.FileCollection.Items(index) = target
            Case "開啟連結"
                Process.Start(target)
            Case "搜尋"
                Dim selectedText As String = target
                ' 如果有選取文字，則將其作為查詢參數傳遞給 Bing 搜索引擎
                If TextBox3.Text.Contains("x.com") Then
                    '自動偵測用戶明並搜尋
                    Dim xUserPattern As New String("https?://x\.com/([A-Za-z0-9_]+)")
                    Dim match As Match = Regex.Match(target, xUserPattern)
                    If match.Success Then
                        Dim username As String = match.Groups(1).Value
                        If target = $"https://x.com/{username}" Then
                            WebView21.CoreWebView2.Navigate($"https://www.google.com/search?q={username}")
                        Else
                            WebView21.CoreWebView2.Navigate($"https://x.com/{username}")
                            Console.WriteLine("Username: " & username)  ' 會輸出 myai04
                        End If

                    Else
                        Console.WriteLine("No match found.")
                    End If

                ElseIf Not String.IsNullOrEmpty(selectedText) Then
                    ' 拼接 Bing 搜索的網址
                    Dim url As String = $"https://www.google.com/search?q={selectedText}"
                    If TextBox3.Text.Contains("x.com") Then
                        'x.com直接搜尋
                        WebView21.CoreWebView2.Navigate(url)
                    Else
                        ' 使用 Process 類打開默認瀏覽器並訪問該網址
                        Process.Start(url)
                    End If
                End If
        End Select
        Form1.refresh_backup()
    End Sub
    Public Sub Clipboard_SetText()
        '加1個把數組轉成換行文字
        Clipboard.SetText(Form1.fileCollection.SelectedItem.ToString)
    End Sub
    Private Sub Clipboard_SetIImage()
        '加1個把數組轉成換行文字
        Clipboard.SetData(DataFormats.Bitmap, PictureBox1.Image)
    End Sub
    Public Delegate Sub UpdateTextBoxDelegate(controls As System.Windows.Forms.TextBox, value As String)
    Private Function SetTextboxText(controls As System.Windows.Forms.TextBox, value As String)
        If (controls.InvokeRequired) Then
            Dim del As New UpdateTextBoxDelegate(AddressOf SetTextboxText)
            controls.Invoke(del, controls, value)
        Else
            controls.Text = value
        End If
        Return Nothing
    End Function

    Public Sub Adjustment_Form()
        Dim standard As Rectangle = Screen.PrimaryScreen.WorkingArea
        Dim locationX As Integer
        If PictureBox1.Visible Then
            Dim imageProportion As Single = PictureBox1.Image.Width / PictureBox1.Image.Height
            Dim form2Width As Single = standard.Height * imageProportion
            Size = New Size(form2Width, standard.Height)
            locationX = standard.Width - form2Width
        Else
            locationX = standard.Width - Width
        End If
        If locationX < 0 Then locationX = 0
        Location = New Point(locationX, 0)
    End Sub
    Public Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If WebView21.Visible Then Return
        KeyPreview = Not String.IsNullOrEmpty(TextBox1.Text)
        Dim ori_image As Image = Form1.PictureBox1.Image
        If ori_image Is Nothing Then Return
        Dim point = picture_set(ori_image)
        Dim rect As New Rectangle(point(0), point(1), point(2), point(3))
        Dim boximage As Bitmap = CopyRectangle(ori_image, rect)
        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
        PictureBox1.Image = boximage
    End Sub
    Private Sub Webview2_DragDrop(sender As ListBox, e As DragEventArgs) Handles WebView21.DragEnter
        ' 顯示有哪些資料格式
        Dim formats = e.Data.GetFormats()
        MessageBox.Show("Available formats:" & vbCrLf & String.Join(vbCrLf, formats))

        ' 以下是你原本的邏輯...
    End Sub
    Public Sub Button3_Click(sender As Object, e As MouseEventArgs) Handles Button3.Click
        If Not select_lastFolder(Index = Form1.targetList.Items.Count - 1) Then Return
        Index += 1
        IndexChange()
    End Sub
    Public Sub Button4_Click(sender As Object, e As MouseEventArgs) Handles Button4.Click
        If Not select_lastFolder(Index = 0) Then Return
        Index -= 1
        IndexChange()
    End Sub
    Private Function Select_lastFolder(islimit As Boolean)
        If islimit = False Then Return True
        If Form1.targetList.Name Is "壓縮檔案集合" Then
            Form1.targetList = Form1.fileCollection
            Index = Form1.targetList.SelectedIndex
            Return True
        End If
        Return False
    End Function
    Private Sub IndexChange()
        If Form1.targetList.Items.Count < Index Then
            Return
            If Form1.targetList Is Form1.fileCollection Then
                Form1.targetList = Form1.subFileCollection
            Else
                Form1.targetList = Form1.fileCollection
            End If
        End If
        Form1.targetList.ClearSelected()
        Form1.targetList.SelectedIndex = Index
        TextBox2.Text = Index + 1
        Adjustment_Form()
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged, MyBase.KeyPress
        Dim textWidth As Integer = TextRenderer.MeasureText(TextBox2.Text, TextBox2.Font).Width
        TextBox1.Width = textWidth + 10 ' Add some padding to the width
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Form1.Next_Page(sender, e)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Form1.Next_Page(sender, e)
    End Sub
    Private Sub PictureBox1_PictureChange(sender As Object, e As EventArgs) Handles PictureBox1.BackgroundImageChanged
        TextBox1_TextChanged(sender, e)
    End Sub
    Public Function Picture_set(ori_image)
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
    Public oriPoint As New Point(0, 0)
    Public newPoint As New Point(0, 0)
    Private Sub PictureBox1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseDown
        If e.Button <> MouseButtons.Left Then Return
        IsDragging = True
        oriPoint.X = e.X
        oriPoint.Y = e.Y
    End Sub
    Private Sub PictureBox1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseUp
        If e.Button <> MouseButtons.Left Then Return
        IsDragging = False
        newPoint.Offset(e.X - oriPoint.X, e.Y - oriPoint.Y)
    End Sub


    Public Sub PictureMode()
        WebView21.Hide()
        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
        PictureBox1.Image = Form1.PictureBox1.Image
        PictureBox1.Show()
        TextBox3.Hide()
    End Sub
    Public Sub MediaMode()
        WebView21.Hide()
        PictureBox1.Hide()
        TextBox3.Hide()
        videoPanel.Show()
        VideoView1.MediaPlayer = Form1._mediaPlayer
    End Sub
    Public Async Function WebMode(url As String) As Task
        Show()
        PictureBox1.Hide()
        Size = New Size(Width, Screen.PrimaryScreen.WorkingArea.Height)

        Try
            If _webView2InitializationTask IsNot Nothing Then
                Await _webView2InitializationTask
            ElseIf WebView21 IsNot Nothing Then
                Await WebView21.EnsureCoreWebView2Async(Nothing)
            Else
                MessageBox.Show("WebView2 控件尚未初始化", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
        Catch ex As Exception
            MessageBox.Show("webview2初始化失敗", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End Try

        ' WebView2の設定を取得します。
        If WebView21.CoreWebView2 IsNot Nothing Then
            If Not String.IsNullOrEmpty(url) Then WebView21.CoreWebView2.Navigate(url)
            WebView21.CoreWebView2.Settings.IsZoomControlEnabled = True
        End If

        WebView21.Show()
        ' PictureBox1.Dock = DockStyle.None
        TextBox3.Show()


    End Function
    Private Sub PictureBox1_DragEnter(sender As System.Object, e As System.Windows.Forms.DragEventArgs) Handles Panel1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub
    Private Sub PictureBox1_DragDrop(sender As System.Object, e As System.Windows.Forms.DragEventArgs) Handles Panel1.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        For Each sourcrPath In files
            If Form1.fileCollection.Items.Contains(sourcrPath) Then Form1.fileCollection.Items.Remove(sourcrPath)
            Form1.fileCollection.Items.Add(sourcrPath)
            If Not String.IsNullOrEmpty(Form1.SearchTextBox.Text) Then Form1.backup.Add(sourcrPath)
        Next
        Form1.refresh_backup()
    End Sub

    Private Sub PictureBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove
        If IsDragging Then
            Dim ori_image = Form1.PictureBox1.Image
            Dim point = picture_set(ori_image)
            Dim moveX = e.X - oriPoint.X
            Dim moveY = e.Y - oriPoint.Y
            Dim rect As New Rectangle(point(0) - moveX, point(1) - moveY, point(2), point(3))
            Dim boximage As Bitmap = CopyRectangle(ori_image, rect)
            PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
            PictureBox1.Image = CType(boximage, Image)
        End If
    End Sub

    'Private Sub WebNavigationStart(sender As Object, e As CoreWebView2NavigationStartingEventArgs) Handles WebView21.NavigationStarting
    '    TextBox3.Text = WebView21.Source.AbsoluteUri
    'End Sub
    Private Async Sub WebNavigationCompleted(sender As Object, e As CoreWebView2NavigationCompletedEventArgs) Handles WebView21.NavigationCompleted
        Dim uri = WebView21.Source.AbsoluteUri
        TextBox3.Text = uri
        store_trf_records(uri, "browserRecord", "網頁歷史")
        If Not e.IsSuccess Then
            webCompeleteLoading = False
            Return
        End If
        webCompeleteLoading = True

        ' If Form1.CheckedListBox1_isCheck("當網頁跳轉時下載圖片") And uri.Contains("/photo/") Then checkWebSite()
        If Form1.CheckedListBox1_isCheck("獲取標題") Then
            Dim browser As WebView2 = WebView21
            Dim title As String = Await get_web_title()
            Form1.Text = title
        End If
        If Form1.fileCollection.SelectedItem IsNot Nothing Then
            If Form1.fileCollection.SelectedItem.Contains("click.mg.dlsite.com") And Not Form1.CheckedListBox1_isCheck("關閉自動修正") Then
                Form1.fileCollection.Items(Form1.fileCollection.SelectedIndex) = uri
                Form1.refresh_backup()
            End If
        End If
    End Sub
    Public Sub Store_trf_records(address As String, directoryName As String, fileScript As String)
        Dim openPath As String = $"{Form1.commonUsed}\Records\{directoryName}\" & Now.ToString("yyyy_MM") & $"{fileScript}.trf"
        Directory.CreateDirectory($"{Form1.commonUsed}\Records\{directoryName}")
        Using fs As New FileStream(openPath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite)
            Using sw As New StreamWriter(fs)
                sw.WriteLine(address)
            End Using
        End Using
    End Sub

    Private Sub WebNavigationStarting() Handles WebView21.NavigationStarting
        webCompeleteLoading = False
    End Sub

    Public Async Sub CheckWebSite()
        'type1 回傳目錄
        'type2 回傳src instagram
        'type3 如果長*寬 > 250000 且非600*480 回傳src  facebook youtube
        'type4 keyword改搜尋class名稱 回傳src  x
        Dim weburl = WebView21.Source.AbsoluteUri
        Dim source As New List(Of String)
        Dim downloadCount As Integer = 1
        Dim downloadSuccess As Integer = 0
        If Not Form1.CheckedListBox1_isCheck("下載時不按讚") Then
            Dim Script2 As String = ""
            If weburl.Contains("twitter") Or weburl.Contains("x.com") Then
                Script2 = "document.querySelector('div [data-testid=""like""]').click();"
            ElseIf weburl.Contains("pixiv") Then
                Script2 = "document.querySelector("".gtm-main-bookmark"").click();"
            End If
            Await WebView21.ExecuteScriptAsync(Script2)
        End If
        If weburl.Contains("twitter") Or weburl.Contains("x.com") Then
            ' 要執行的 JavaScript 程式碼
            Dim script As String = "
                (function() {
                    return document.querySelector('[data-testid=""videoComponent""]') !== null;
                })();
            "

            ' 執行 JavaScript 並取得結果
            Dim result As String = Await WebView21.ExecuteScriptAsync(script)
            If result = "true" Then

                Dim saveFolder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString & "\" & Now.ToString("yyyy_MM") & "\"
                Dim psi As New ProcessStartInfo With {
                    .FileName = "cmd.exe",
                    .Arguments = $"/c yt-dlp --cookies ""x.com_cookies.txt"" -P {saveFolder} {weburl}",
                    .WorkingDirectory = "C:\Users\b2000",
                    .WindowStyle = ProcessWindowStyle.Hidden,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True,
                    .UseShellExecute = False,
                    .CreateNoWindow = True
                }
                Dim proc As New Process With {
                    .StartInfo = psi
                }

                AddHandler proc.OutputDataReceived,
                Sub(sender, e)
                    If String.IsNullOrEmpty(e.Data) Then Return
                    Form1.subFileCollection.Items.Add(e.Data)
                    ' 偵測下載開始
                    If e.Data.Contains("[download]") Then
                        Form1.subFileCollection.Items.Add("下載中...")
                    End If

                    ' 偵測完成訊息
                    If e.Data.Contains("[download] 100% of") Then
                        Form1.subFileCollection.Items.Add("下載完成！")
                    End If

                    ' 偵測檔名
                    If e.Data.Contains("[FixupM3u8]") Then
                        Dim rgx As New Regex("""(\.+)""")
                        Dim match = rgx.Match(e.Data)
                        If match.Success Then
                            Form1.subFileCollection.Items.Add("檔案名稱：" & match.Value)
                        End If
                    End If
                End Sub

                AddHandler proc.ErrorDataReceived,
                Sub(sender, e)
                    If Not String.IsNullOrEmpty(e.Data) Then
                        Form1.subFileCollection.Items.Add("錯誤：" & e.Data)
                    End If
                End Sub

                proc.Start()
                proc.BeginOutputReadLine()
                proc.BeginErrorReadLine()

            End If
            If Not WebView21.Source.ToString().Contains("photo/") Then
                script = " 
                    var links = document.getElementsByClassName(""css-175oi2r r-1pi2tsx r-1ny4l3l r-1loqt21"");
                    var max = 0;
                    var target = null;
                    for (var i = 0; i < links.length; i++) {
                    var match = links[i].href.match(/photo\/(\d+)/);
                    if (match) {
                        var num = parseInt(match[1]);
                        if (num <= max) {
                        break;}
                        max = num;
                        target = links[i]; }
                                }
                    if (target) {
                        target.click();}"
                'script = $"var links = document.getElementsByClassName(""css-9pa8cd"");
                '              if (links.length != 2 ) {{
                '                links[{downloadCount + 1}].click();
                '            }}"
                result = Await WebView21.ExecuteScriptAsync(script)
                Console.WriteLine(result)
                Await waitForLoad()
            End If
            source = Await GetWebImageSrc("css-9pa8cd", "class-src")
            Dim exceptForm As String() = {"large", "4800x4800", "medium", "4096x4096", "900x900"}
            Dim targetsource = source.Where(Function(s) exceptForm.Any(Function(form) s.Contains(form))).ToList()

            If targetsource.Count = 0 Then
                targetsource = source.Where(Function(s) s.Contains("small")).ToList()
            End If

            source = targetsource
        ElseIf weburl.Contains("facebook") Then
            source = Await GetWebImageSrc("https://scontent.ftpe", "src-special")
            source.RemoveAt(0) 'facebook更新預載入
            '如果為group跳過第一個圖片(社團封面)
            If TextBox3.Text.Contains("groups") And source.Count <> 0 Then source.RemoveAt(0)
        ElseIf weburl.Contains("pixiv") Then
            Dim script = "const allButtons = document.querySelectorAll('[type=""button""]');
                          const target = Array.from(allButtons).find(btn => btn.innerText.includes('查看全部'));
                          target.click();"
            Await WebView21.ExecuteScriptAsync(script)
            Await waitForLoad()
            '".sc-dba767bd-3.cnkiZk.gtm-expand-full-size-illust"
            source = Await GetWebImageSrc(".gtm-expand-full-size-illust", "css-href")
        ElseIf weburl.Contains("youtube") Then
            source = Await GetWebImageSrc("https://yt3.ggpht.com", "src-special")
        ElseIf weburl.Contains("jmcomic") Then
            source = Await GetWebImageSrc("https://cdn-msp.jmcomic1.me/media/photos", "src-special")
        ElseIf weburl.Contains("instagram") Or weburl.Contains("threads.com") Then
            source = Await GetWebImageSrc("instagram.ftpe14-1", "src-special")
        ElseIf weburl.Contains("fanbox") Then
            source = Await GetWebImageSrc("https://downloads.fanbox.cc/images/post", "src-special")
        ElseIf weburl.Contains("gamer.com.tw") Then
            source = Await GetWebImageSrc(".photoswipe-image", "css-href")
            If source.Count = 0 Then
                source = Await GetWebImageSrc("https://truth.bahamut.com.tw/artwork", "src-special")
            End If
        End If
        If source.Count = 0 Then
            MessageBox.Show("找不到資源!")
            Return
        End If
        Form6.Form6_setup(source)
        Form6.Show()
    End Sub
    Public Async Sub FindSource(sender As Object, e As EventArgs)
        Await WebMode("")
        Dim browser As WebView2 = WebView21
        Dim menuItem As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        If menuItem Is Nothing Then Return
        Dim owner As ToolStrip = menuItem.Owner
        Form1.ToolStripMenuItemsetTargetlistbox(owner)
        Dim rgx As Regex = Form1.dlsiteRgx
        Dim dlrgxResult As Match = rgx.Match(Form1.targetList.SelectedItem.ToString())
        Dim fn As String = If(dlrgxResult.Success, dlrgxResult.Value, Nothing)
        Dim pattern As String = "\]\s*(.+)$"
        Dim itemTitleMatch As Match = Regex.Match(Form1.targetList.SelectedItem, pattern)
        ' Get the value of the element
        Dim title As String = ""
        Dim btn_name As String = CType(sender, ToolStripMenuItem).Text
        Dim search_site As String()
        Console.WriteLine(btn_name)
        If {"nhentai", "nyaa", "google"}.Contains(btn_name) Or (Form1.CheckedListBox1_isCheck("獲取標題") And btn_name = "dlsite") Then
            search_site = {"get_title", btn_name}
        ElseIf {"google鏡頭", "二次元画像詳細検索", "sauceNAO"}.Contains(btn_name) Then
            search_site = {"get_address", btn_name}
        Else
            search_site = {btn_name}
        End If
        If Not search_site.Contains("本機") Then Show()
        For Each i As String In search_site
            Select Case i
                Case "get_title"
                    If Form1.CheckedListBox1_isCheck("獲取標題") Then
                        title = Form1.Text
                    Else
                        Dim newName = Await Get_web_title()
                        title = If(newName = "null", title, newName)
                    End If

                Case "get_address"
                    fn = Form1.targetList.SelectedItem
                    ' 對檔案名稱進行 URL 編碼
                Case "dlsite"
                    If dlrgxResult.Success Then
                        Form1.SubFileCollection.Items.Add($"https://www.dlsite.com/maniax/work/=/product_id/{fn}.html")
                        browser.CoreWebView2.Navigate("https://www.dlsite.com/maniax/work/=/product_id/" & fn & ".html")
                    Else
                        browser.CoreWebView2.Navigate("https://www.dlsite.com/maniax/")
                        If itemTitleMatch.Success Then
                            title = itemTitleMatch.Groups(1).Value
                            Await WaitForLoad()
                            Await browser.ExecuteScriptAsync($"document.getElementById('search_text').value = '{title}';")
                            Await browser.ExecuteScriptAsync("document.getElementById('search_button').click();")
                        End If
                    End If
                Case "kikoeru"
                    Form1.SubFileCollection.Items.Add($"https://www.asmr.one/work/{fn}")
                    browser.CoreWebView2.Navigate($"https://www.asmr.one/work/{fn}")
                Case "mikocon"
                    browser.CoreWebView2.Navigate("https://www.mikocon.com/search.php?Mod=forum")
                    Await WaitForLoad()
                    Await browser.ExecuteScriptAsync($"document.getElementById('scform_srchtxt').value = '{fn}';")
                    Await browser.ExecuteScriptAsync("document.getElementById('scform_submit').click();")
                Case "google鏡頭"
                    browser.CoreWebView2.Navigate($"https://www.google.com/searchbyimage?image_url=https://easypreview.unified-storage.com/read-file?file={fn}")
                Case "二次元画像詳細検索"
                    Dim encodedTarget As String = fn.ToString() _
                                             .Replace(" ", "%20") _
                                             .Replace(":", "%3A") _
                                             .Replace("\", "%5C")
                    browser.CoreWebView2.Navigate($"https://ascii2d.net/search/url/https://easypreview.unified-storage.com/read-file?file={encodedTarget}")
                Case "sauceNAO"
                    browser.CoreWebView2.Navigate($"https://saucenao.com/search.php?db=999&url=https://easypreview.unified-storage.com/read-file?file={fn}")
                Case "anime-sharing"
                    browser.CoreWebView2.Navigate("https://www.anime-sharing.com/search/")
                    Await WaitForLoad()
                    Await browser.ExecuteScriptAsync($"document.querySelector('[type=search]').value = '{fn}';")
                    'Await waitForLoad()
                    '無法等待輸入
                    'Await browser.ExecuteScriptAsync("document.querySelectorAll('button')[1].click();")
                Case "nhentai"
                    browser.CoreWebView2.Navigate("https://nhentai.net/")
                    Await WaitForLoad()
                    Await browser.ExecuteScriptAsync($"document.querySelector('.search input').value = '{title}';")
                    Await browser.ExecuteScriptAsync("document.querySelector('.search .btn.btn-square').click();")
                Case "nyaa"
                    browser.CoreWebView2.Navigate("https://sukebei.nyaa.si/")
                    Await WaitForLoad()
                    Await browser.ExecuteScriptAsync($"document.querySelector('.form-control').value = '{title}';")
                    Await browser.ExecuteScriptAsync("document.querySelector('.btn-primary').click();")
                Case "Tokyo図書館"
                    browser.CoreWebView2.Navigate("https://www.tokyotosho.info/index.php")
                    Await WaitForLoad()
                    Await browser.ExecuteScriptAsync("document.querySelector('input[type=""text"" i]').value = '" & title & "';")
                    Await browser.ExecuteScriptAsync("document.querySelector('input[type=""submit"" i]').click();")
                Case "eyny"
                    browser.CoreWebView2.Navigate("https://www.eyny.com/search.php?mod=forum")
                    Await WaitForLoad()
                    Await browser.ExecuteScriptAsync("document.querySelector('#scform_srchtxt').value = '" & title & "';")
                    Await browser.ExecuteScriptAsync("document.querySelector('#scform_submit]').click();")
                Case "hanime1"
                    browser.CoreWebView2.Navigate("https://hanime1.me/search")
                    Await WaitForLoad()
                    Await browser.ExecuteScriptAsync("document.getElementById('nav-query').value = '" & title & "';")
                    Await browser.ExecuteScriptAsync("document.querySelector('search-btn').click();")
                Case "google"
                    Dim url As String = "https://www.google.com/search?q=" & title
                    browser.CoreWebView2.Navigate(url)
            End Select
            If i.Contains(":\") Or i.Contains("本機") Then
                Dim files As List(Of String) = SearchFiles(fn)
                For Each file In files
                    Form1.SubFileCollection.Items.Add(file)
                Next
            Else
                Form1.SubFileCollection.Items.Add(browser.Source.ToString())
            End If
        Next



        'search item name

    End Sub
    Public Async Function Get_web_title() As Task(Of String)
        Dim title = ""
        Dim browser As WebView2 = WebView21
        Dim url As String = Form1.fileCollection.SelectedItem
        If url.Contains("dlsite") Then
            title = Await browser.ExecuteScriptAsync("document.querySelector('#work_name').innerHTML;")
            'title = Await WebView21.ExecuteScriptAsync("document.documentElement.outerHTML;")
            'Dim rgx2 As Regex = New Regex("<h1 itemprop='name' id='work_name'>(.*)<\/h1>")
            'Dim fn2 As String = rgx2.Match(title).Groups(1).Value
            'title = fn2
            Console.WriteLine("name:" & title)
        ElseIf url.Contains("melonbooks") Then
            title = Await browser.ExecuteScriptAsync("document.querySelector('.page-header').innerHTML;")
        ElseIf url.Contains("dl.getchu") Then
            title = Await browser.ExecuteScriptAsync("document.querySelector('.bold').innerHTML;")
        ElseIf url.Contains("getchu") Then
            'title = Await browser.ExecuteScriptAsync("document.querySelector('#soft-title').innerHTML;")
            'title = title.Replace("\n", "").Trim
            title = Await browser.ExecuteScriptAsync("document.querySelector('#text').innerHTML;")
        End If
        Return title.Replace("""", "")
    End Function

    Private Sub Form_click(sender As Object, e As EventArgs) Handles Me.Activated
        If Form1.Visible And Form1.CheckedListBox1_isCheck("form1與form2連動") Then
            Form1.TopMost = True
            Form1.TopMost = False
            Me.TopMost = True
            Me.TopMost = False
        End If
    End Sub



    Public Async Sub WmcSearch(keyword As String)
        ' 等待 WebView2 初始化完成（已在 Form_Load 時初始化）
        If _webView2InitializationTask IsNot Nothing Then
            Await _webView2InitializationTask
        End If

        ' JavaScript 產生 JSON 陣列
        Dim jsCode As String = "
        JSON.stringify(
            Array.from(document.querySelectorAll('.text'))
                 .map(el => el.innerText)
        );
    "

        ' 執行 JavaScript
        Dim resultJson As String = Await WebView21.CoreWebView2.ExecuteScriptAsync(jsCode)

        ' WebView2 傳回的結果會多一層引號包住整個 JSON 字串，要去掉開頭和結尾
        If resultJson.StartsWith("""") AndAlso resultJson.EndsWith("""") Then
            resultJson = resultJson.Substring(1, resultJson.Length - 2)
        End If

        ' 還原被轉義的雙引號
        resultJson = resultJson.Replace("\""", """") ' 把 \" 變成 "

        ' 反序列化為陣列
        Dim matches As String() = JsonConvert.DeserializeObject(Of String())(resultJson)

        ' 關鍵字比對並顯示
        For Each match As String In matches
            If Not match.Contains(keyword) Then Continue For
            Form1.subFileCollection.Items.Add(match)
        Next
    End Sub

    Public Function SearchAllFiles(savename As String) As List(Of String)
        ' 調用 EverythingSearcher 模組中的 SearchFiles 方法
        If savename Is Nothing Then Return New List(Of String)
        Dim noExt As String = savename.Split(".")(0)
        Dim foundFiles As List(Of String) = EverythingSearcher.SearchFiles(noExt)
        Return foundFiles
    End Function
    Private Function Add_download_record(savename As String, savepath As String) As String
        If String.IsNullOrEmpty(savename) Then Return "fileNameIsEmpty"
        'If Not Form1.CheckedListBox1_isCheck("下載時紀錄") Then Return "noSave"

        Dim recordPath As String = Form1.recordPath & "downloadRecord.txt"
        Dim downloadRecords As New Dictionary(Of String, String)

        ' 讀取現有記錄到字典
        If File.Exists(recordPath) Then
            For Each line As String In File.ReadAllLines(recordPath, Encoding.UTF8)
                Dim parts() As String = line.Split(New String() {" download in "}, StringSplitOptions.None)
                If parts.Length = 2 Then
                    downloadRecords(parts(0)) = parts(1)
                End If
            Next
        End If
        'SQLRead()
        ' 檢查記錄是否存在
        If downloadRecords.ContainsKey(savename) Then
            Return $"{savename} download in {downloadRecords(savename)}"
        Else
            ' 添加新記錄
            Using sw As New StreamWriter(recordPath, True, Encoding.UTF8)
                sw.WriteLine($"{savename} download in {savepath}")
            End Using
            Return "canDownload"
        End If
    End Function



    Public Async Function DownloadOriginalImage(src As String) As Task(Of Boolean)
        'If Form1.orderComboBox.Text.Contains("noDownload") Then Return False

        Dim saveFolder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString & "\" & Now.ToString("yyyy_MM") & "\"
        If Form1.CheckedListBox1_isCheck("下載在TargetBox") Then
            saveFolder = Form1.TargetComboBox.Text & "\"
        End If
        Directory.CreateDirectory(saveFolder)
        Dim saveFileName As String = getFileName(src)
        Console.WriteLine(src)
        saveFileName = saveFileName.Replace("""", "")
        If String.IsNullOrEmpty(saveFileName) Then
            Form1.MessageLabel.Text = "fileNameIsEmpty"
            Return True
        End If
        Dim downloadMessage As List(Of String) = SearchAllFiles(saveFileName)
        Dim labelMessage As String
        If downloadMessage.Count = 0 Then
            Dim result As String = Await SaveImage(src, saveFolder & saveFileName)
            If result.Contains("下載失敗") Then
                Form1.SubFileCollection.Items.Add(result)
            End If
            labelMessage = saveFileName & result
            Set_text_unity(labelMessage)
        Else
            labelMessage = saveFileName & "已下載"
            Set_text_unity(labelMessage)
            downloadMessage.RemoveAll(Function(s) s.Contains("CrossDevice\V2041"))
            Form1.PictureBox1.Image = New Bitmap(downloadMessage(0))
            If Form5.Visible Then Form5.PictureBox1.Image = Form1.PictureBox1.Image
            For Each downloadfile In downloadMessage
                Form1.subFileCollection.Items.Add(downloadfile)
            Next
        End If

        Return True
    End Function

    Private Sub Set_text_unity(labelMessage As String)
        Form1.MessageLabel.Text = labelMessage
        Label3.Text = labelMessage
        If Form5.Visible Then Form5.Label1.Text = labelMessage
        Form1.SubFileCollection.Items.Add(labelMessage)
    End Sub



    ' 在WebView2控制項的CoreWebView2Ready事件中註冊Network.responseReceived事件
    Private Sub WebView21_CoreWebView2Ready(sender As Object, e As EventArgs) Handles WebView21.CoreWebView2InitializationCompleted
        Return
        ' 獲取Network.responseReceived事件的接收器
        Dim receiver As CoreWebView2DevToolsProtocolEventReceiver = WebView21.CoreWebView2.GetDevToolsProtocolEventReceiver("Network.responseReceived")
        ' 添加事件處理程序

        ' 啟用Network域
        WebView21.CoreWebView2.CallDevToolsProtocolMethodAsync("Network.enable", "{}")
    End Sub


    Public Function GetFileName(oriName As String)
        Dim rgx As New Regex("([^/]*)$")
        If oriName Is Nothing Then Return ""
        Dim fn As String = rgx.Match(oriName).Groups(1).Value

        'pixiv處理
        If oriName.Contains("i.pximg.net") Then
            Dim prgx As New Regex("(\d+)-\w+(_.+)")
            Dim pmatch = prgx.Match(fn)
            If pmatch.Success Then
                fn = pmatch.Groups(1).Value + pmatch.Groups(2).Value
            End If
        ElseIf fn.Contains("?") Then
            Dim subs As String() = fn.Split("?")
            If subs(0).Contains(".") Then
                fn = subs(0)
            Else
                fn = subs(0) & ".jpg"
            End If
        ElseIf fn.Contains(".") Then
            Dim subs As String() = fn.Split(".")
            fn = subs(0) & "." & subs(1).Split(":")(0)
        Else
            ' 替換檔案名稱中的特殊字元，例如inputjpg
            fn = Form1.Make_Character_Valid(fn, "jpg")
        End If
        Return fn
    End Function

    Public Async Function SaveImage(ByVal imageUrl As String, ByVal filename As String) As Task(Of String)
        Dim result = "下載失敗"
        If imageUrl = "null" OrElse String.IsNullOrWhiteSpace(imageUrl) Then Return result & " 原因:空網址"
        imageUrl = imageUrl.Replace("""", "")

        Form1.MessageLabel.Text = $"開始下載{filename}"
        Form5.Label1.Text = $"開始下載{filename}"
        Form1.subFileCollection.Items.Add($"開始下載{filename}")
        Using client As New HttpClient()
            Dim browser As WebView2 = WebView21
            Dim tarUrl As String = browser.Source.AbsoluteUri
            tarUrl = tarUrl.Replace("""", "")
            ' 建立單次請求
            Dim request As New HttpRequestMessage(HttpMethod.Get, imageUrl)

            ' 加上 Referer
            request.Headers.Referrer = New Uri(tarUrl)
            ' 發送請求
            Dim response As HttpResponseMessage = Await client.SendAsync(request)
            response.EnsureSuccessStatusCode()

            Dim bytes As Byte() = Await response.Content.ReadAsByteArrayAsync()
            ' 在背景工作處理圖片
            Dim bmp As Bitmap = Await Task.Run(Function()
                                                   Using stream As New MemoryStream(Bytes)
                                                       Using image As New MagickImage(stream)
                                                           ' 自動判斷格式，不強制轉 BMP
                                                           image.Write(filename)

                                                           ' 轉成 PNG 的 byte array
                                                           Dim pngBytes As Byte() = image.ToByteArray(MagickFormat.Png)

                                                           Using ms As New MemoryStream(pngBytes)
                                                               ' Clone，避免依賴 ms
                                                               Return New Bitmap(ms)
                                                           End Using
                                                       End Using
                                                   End Using
                                               End Function)

            ' 回 UI 更新 PictureBox
            Form1.Invoke(Sub()
                             Form1.PictureBox1.Image = bmp
                             If Form5.Visible Then Form5.PictureBox1.Image = bmp
                             Form1.Label15.Text = $"{bmp.Width}px {bmp.Height}px"
                             Form1.SubFileCollection.Items.Add(filename)

                         End Sub)

            result = "下載成功"
        End Using
        Return result
    End Function


    Private Sub Form2_Close(sender As System.Object, e As System.EventArgs) Handles MyBase.FormClosing
        If Not Form1.Visible Then Form1.Show()
    End Sub
    Private Sub Picturebox1_DoubleClick(sender As System.Object, e As System.EventArgs) Handles PictureBox1.DoubleClick
        Form1.Visible = True
        Close()
    End Sub

    Public Async Function GetWebImageSrc(keyword As String, type As String) As Task(Of List(Of String))
        'type img-src 回傳src  
        'type src-special 如果長*寬 > 250000 且非600*480 回傳src  facebook youtube instagram
        'type class-src keyword改搜尋class名稱 回傳src  x 
        'type css-href keyword改搜尋class名稱 回傳href  pixiv

        Dim canUseIndex As New List(Of String)
        Select Case type
            Case "img-src"
                Dim script As String = "(() => {
                                return JSON.stringify(
                                    Array.from(document.querySelectorAll('img'))
                                    .map(img => img.getAttribute('src'))
                                    .filter(src => src !== null && src !== '')
                                );
                            })();"
                Dim result As String = Await WebView21.ExecuteScriptAsync(script)
                ' 修正 JSON 格式
                Dim cleanResult As String = result.Trim()

                ' 判斷是否為雙重轉義的 JSON
                If cleanResult.StartsWith("""") AndAlso cleanResult.EndsWith("""") Then
                    cleanResult = cleanResult.Substring(1, cleanResult.Length - 2) ' 去掉開頭和結尾的引號
                    cleanResult = cleanResult.Replace("\""", """") ' 取消 JavaScript 轉義的雙引號
                End If
                ' 解析 JSON 結果
                Dim srcList As List(Of String) = JsonConvert.DeserializeObject(Of List(Of String))(cleanResult)

                ' 過濾符合關鍵字的圖片
                'Dim canUseIndex As New List(Of String)
                For Each src In srcList
                    If src.Contains(keyword) Then
                        Console.WriteLine(src)
                        canUseIndex.Add(src)
                    End If
                Next
            Case "src-special"
                Dim script As String = "document.querySelectorAll(""img"").length"
                Dim result As String = Await WebView21.ExecuteScriptAsync(script)
                Dim canUseLength As Integer = CInt(result)
                Console.WriteLine("CanUseLength:" & canUseLength)
                '若沒有結果回傳空list
                If canUseLength = 0 Then Return canUseIndex
                For i = 0 To canUseLength - 1
                    script = "document.querySelectorAll(""img"")[" & i.ToString & "].getAttribute(""src"")"
                    Dim script2 = "var curRect = document.querySelectorAll(""img"")[" & i.ToString & "];
                          var area = curRect.naturalWidth.toString() + '_' +  curRect.naturalHeight.toString();
                          area"
                    result = Await WebView21.ExecuteScriptAsync(script)
                    If result = Nothing Then Continue For
                    If Not result.Contains(keyword) Then Continue For
                    Console.WriteLine(result)
                    result = Await WebView21.ExecuteScriptAsync(script2)
                    result = result.Replace("""", "")
                    Dim size As String() = result.Split("_")
                    If CInt(size(0)) * CInt(size(1)) < 250000 Or result = "600_480" Or result = "600_600" Then Continue For
                    result = Await WebView21.ExecuteScriptAsync(script)
                    canUseIndex.Add(result)
                Next

            Case "class-src"
                Dim script As String = "
                            (() => {
                                return JSON.stringify(
                                    Array.from(document.getElementsByClassName('" & keyword & "'))
                                    .map(el => el.getAttribute('src'))
                                    .filter(src => src !== null && src !== '')
                                );
                            })();"

                Dim result As String = Await WebView21.ExecuteScriptAsync(script)

                ' 修正 JSON 格式
                Dim cleanResult As String = result.Trim()

                ' 判斷是否為雙重轉義的 JSON
                If cleanResult.StartsWith("""") AndAlso cleanResult.EndsWith("""") Then
                    cleanResult = cleanResult.Substring(1, cleanResult.Length - 2) ' 去掉開頭和結尾的引號
                    cleanResult = cleanResult.Replace("\""", """") ' 取消 JavaScript 轉義的雙引號
                End If
                ' 解析 JSON 結果
                'Dim canUseIndex As New List(Of String)
                Try
                    canUseIndex = JsonConvert.DeserializeObject(Of List(Of String))(cleanResult)
                Catch ex As Exception
                    Console.WriteLine("JSON 解析失敗: " & ex.Message)
                    Console.WriteLine("錯誤的 JSON: " & cleanResult)
                End Try

                ' 輸出結果
                For Each href In canUseIndex
                    Console.WriteLine(href)
                Next

            'Case "outputHtml"
            '    Dim script As String = "document.documentElement.outerHTML; // get the entire HTML of the document"
            '    Dim result As String = Await WebView21.ExecuteScriptAsync(script)
            '    Console.WriteLine(result)
            '    Dim rgx As Regex = New Regex(keyword)
            '    Dim fn As MatchCollection = rgx.Matches(result)
            '    For Each col In fn
            '        canUseIndex.Add(col.Groups(1).Value)
            '        Console.WriteLine(col.Groups(1).Value)
            '    Next
            Case "css-href"
                Dim script As String = "
                                        (() => {
                                            return JSON.stringify(
                                                Array.from(document.querySelectorAll('" & keyword & "'))
                                                .map(el => el.getAttribute('href'))
                                                .filter(href => href !== null && href !== '')
                                            );
                                        })();
                                    "
                Dim result As String = Await WebView21.ExecuteScriptAsync(script)
                ' 修正 JSON 格式
                Dim cleanResult As String = result.Trim()

                ' 判斷是否為雙重轉義的 JSON
                If cleanResult.StartsWith("""") AndAlso cleanResult.EndsWith("""") Then
                    cleanResult = cleanResult.Substring(1, cleanResult.Length - 2) ' 去掉開頭和結尾的引號
                    cleanResult = cleanResult.Replace("\""", """") ' 取消 JavaScript 轉義的雙引號
                End If
                ' 解析 JSON 結果
                Try
                    canUseIndex = JsonConvert.DeserializeObject(Of List(Of String))(cleanResult)
                Catch ex As Exception
                    Console.WriteLine("JSON 解析失敗: " & ex.Message)
                    Console.WriteLine("錯誤的 JSON: " & cleanResult)
                End Try

                ' 輸出結果
                For Each href In canUseIndex
                    Console.WriteLine(href)
                Next
        End Select
        Return canUseIndex
    End Function
    Public Async Function WaitForLoad() As Task
        While True
            Await Task.Delay(120)
            If webCompeleteLoading = True Then Exit While
        End While
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

    Private Sub TextBox2_TextEdit(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode <> Keys.Enter Then Return
        Dim oriIndex As Integer
        Dim targetList As ListBox
        If Form1.subFileCollection.Items.Count = 0 Or Form1.subFileCollection.SelectedItem = Nothing Then

            targetList = Form1.fileCollection
        Else
            targetList = Form1.subFileCollection
        End If
        'legitimate check

        If CInt(TextBox2.Text) > targetList.Items.Count - 1 Then Return
        If CInt(TextBox2.Text) < 0 Then Return
        oriIndex = targetList.SelectedIndex

        Index = CInt(TextBox2.Text) - 1


        targetList.SetSelected(Index, True)
        If oriIndex <> -1 Then targetList.SetSelected(oriIndex, False)
        TextBox2.Text = Index + 1
        Adjustment_Form()
        Form1.file_Preview(True)
    End Sub
    Private Sub TextBox3_TextEdit(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode <> Keys.Enter Then Return
        'If e.KeyChar <> Convert.ToChar(Keys.Enter) Then Return
        WebView21.CoreWebView2.Navigate(TextBox3.Text)
    End Sub

    Private Sub TextBox3_Ctrl(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        '判斷是否按下alt+方向上下
        If e.Control Then
            If e.KeyCode = Keys.C Then
                Clipboard.SetText(TextBox3.SelectedText)
            End If
            If e.KeyCode = Keys.X Then
                Clipboard.SetText(TextBox3.SelectedText)
                TextBox3.Text.Replace(TextBox3.SelectedText, "")
            End If
            'If e.KeyCode = Keys.V Then
            '    TextBox3.Text.Replace(TextBox3.SelectedText, Clipboard.GetText)
            'End If
            If e.KeyCode = Keys.A Then
                TextBox3.SelectAll()
            End If
        End If
    End Sub
    Private Sub Form2_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F5 And WebView21.Visible Then
            WebView21.Reload()
        End If
        If Not e.Control Then Return
        Select Case e.KeyCode
            Case Keys.D
                checkWebSite()
            Case Keys.R
                Form1.Remove_Button(sender, New EventArgs)
            Case Keys.Down
                Form1.targetList.SelectedIndex = Form1.targetList.SelectedIndex + 1
                Form1.targetList.SetSelected(Form1.targetList.SelectedIndex, False)
            Case Keys.Up
                Form1.targetList.SelectedIndex = Form1.targetList.SelectedIndex - 1
                Form1.targetList.SetSelected(Form1.targetList.SelectedIndex + 1, False)
            Case Keys.Enter
                Form1.OnButton4Click(sender, New EventArgs)
            Case Keys.Z
                Form1.listUndo()
            Case Keys.C
                Clipboard_SetIImage()
            Case Keys.V
                Form4.form4_load()
                Form4.setdata_from_clipboard()
                Form4.Button1_Click(sender, e)
        End Select
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.Click
        TopMost = CheckBox1.Checked
    End Sub


    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Form5.Show()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.Click
        Dim rgx As New Regex("x.com/(\.+)/status")
        Dim result As Match = rgx.Match(TextBox3.Text)
        If result.Success Then
            TextBox3.SelectedText = result.Groups(1).Value
        End If
    End Sub

    Private Sub VolumeTrackBar_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.ValueChanged
        Form1._mediaPlayer.Volume = TrackBar1.Value
    End Sub
    ' 當前播放時間改變時更新 ProgressBar
    Private Sub OnTimeChanged(sender As Object, e As MediaPlayerTimeChangedEventArgs)
        If ProgressBar2.InvokeRequired Then
            ProgressBar2.Invoke(New Action(Sub() UpdateProgressBar()))
        Else
            UpdateProgressBar()
        End If
    End Sub
    ' 將更新進度條的邏輯封裝到單獨的方法中
    ' 更新進度條，根據當前播放進度
    Private Sub UpdateProgressBar()
        If Form1._mediaPlayer.Media IsNot Nothing AndAlso Form1._mediaPlayer.Media.Duration > 0 Then
            Dim newValue As Integer = CInt((Form1._mediaPlayer.Time / Form1._mediaPlayer.Media.Duration) * ProgressBar2.Maximum)
            ProgressBar2.Value = newValue ' 直接設置值，而不是逐漸增加
        End If
    End Sub

    ' 點擊 ProgressBar 時跳轉進度
    Private Sub ProgressBar2_Click(sender As Object, e As EventArgs) Handles ProgressBar2.Click
        ' 將游標位置轉換為 ProgressBar2 內的相對位置
        Dim pos As Point = ProgressBar2.PointToClient(Cursor.Position)

        ' 計算點擊位置的比例
        Dim clickRatio As Double = pos.X / ProgressBar2.Width

        ' 計算新的進度值，並根據比例設定播放時間
        Dim newValue As Integer = CInt(clickRatio * ProgressBar2.Maximum)
        Dim newTime As Long = CLng(clickRatio * Form1._mediaPlayer.Media.Duration)


        ' 設置 MediaPlayer 的播放時間
        Form1._mediaPlayer.Time = newTime

        ' 更新進度條以反映新進度
        ProgressBar2.Value = Math.Min(newValue, ProgressBar2.Maximum)

        ' 手動更新進度條以反映新的進度
        'UpdateProgressBar()
    End Sub
    ' 撥放/暫停按鈕
    'Private Sub PlayPauseButton_Click(sender As Object, e As EventArgs) Handles Button24.Click
    '    PlayOrPause()
    'End Sub

    ' 停止按鈕
    Private Sub StopButton_Click(sender As Object, e As EventArgs) Handles Button25.Click
        Form1._mediaPlayer.Stop()
        Button24.Text = "播放"
        ProgressBar2.Value = 0
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Form1.Remove_Button(sender, e)
    End Sub
End Class

