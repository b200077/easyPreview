Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Drawing.Design
Imports System.Drawing.Imaging
Imports System.IO
Imports System.IO.Compression
Imports System.IO.Pipes
Imports System.Net
Imports System.Net.Http
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports DiscUtils.Iso9660
Imports ImageMagick
Imports LibVLCSharp.Shared
Imports Microsoft.VisualBasic.FileIO
Imports Microsoft.Win32
Imports Microsoft.WindowsAPICodePack.Shell
Imports Microsoft.WindowsAPICodePack.Shell.PropertySystem
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Imports SharpCompress.Archives
Imports SharpCompress.Common
Imports Clipboard = System.Windows.Forms.Clipboard

Partial Public Class Form1


    <DllImport("shlwapi.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function StrCmpLogicalW(x As String, y As String) As Integer
    End Function

    Private Function CompareNatural(x As String, y As String) As Integer
        Return StrCmpLogicalW(x, y)
    End Function

    Private Class MyControl
        Inherits Form
        Implements IMessageFilter

        Private InterfaceClassGuid As New Guid(&H4D1E55B2, &HF16F, &H11CF, &H88, &HCB, &H0, &H11, &H11, &H0, &H0, &H30)
        Private Const WM_DEVICECHANGE As UInteger = &H219
        Private Const DBT_DEVICEARRIVAL As UInteger = &H8000
        Private Const DBT_DEVICEREMOVEPENDING As UInteger = &H8003
        Private Const DBT_DEVICEREMOVECOMPLETE As UInteger = &H8004
        Private Const DBT_CONFIGCHANGED As UInteger = &H18
        Private Const DBT_DEVTYP_DEVICEINTERFACE As UInteger = &H5
        Private Const DEVICE_NOTIFY_WINDOW_HANDLE As UInteger = &H0

        Private Structure DEV_BROADCAST_DEVICEINTERFACE
            Friend dbcc_size As UInteger
            Friend dbcc_devicetype As UInteger
            Friend dbcc_reserved As UInteger
            Friend dbcc_classguid As Guid
            Friend dbcc_name As Char()
        End Structure

        <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Private Shared Function RegisterDeviceNotification(ByVal hRecipient As IntPtr, ByVal NotificationFilter As IntPtr, ByVal Flags As UInteger) As IntPtr

        End Function

        Public Sub New()
            Dim DeviceBroadcastHeader As New DEV_BROADCAST_DEVICEINTERFACE With {
                .dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE
            }
            DeviceBroadcastHeader.dbcc_size = CUInt(Marshal.SizeOf(DeviceBroadcastHeader))
            DeviceBroadcastHeader.dbcc_reserved = 0
            DeviceBroadcastHeader.dbcc_classguid = InterfaceClassGuid
            Dim pDeviceBroadcastHeader As IntPtr
            pDeviceBroadcastHeader = Marshal.AllocHGlobal(Marshal.SizeOf(DeviceBroadcastHeader))
            Marshal.StructureToPtr(DeviceBroadcastHeader, pDeviceBroadcastHeader, False)
            RegisterDeviceNotification(Me.Handle, pDeviceBroadcastHeader, DEVICE_NOTIFY_WINDOW_HANDLE)
        End Sub

        Public Event DeviceConnected As EventHandler

        Protected Overrides Sub WndProc(ByRef m As Message)
            If m.Msg = WM_DEVICECHANGE Then

                If (CInt(m.WParam) = DBT_DEVICEARRIVAL) OrElse (CInt(m.WParam) = DBT_DEVICEREMOVEPENDING) OrElse (CInt(m.WParam) = DBT_DEVICEREMOVECOMPLETE) OrElse (CInt(m.WParam) = DBT_CONFIGCHANGED) Then
                    RaiseEvent DeviceConnected(Me, Nothing)
                End If
            End If

            MyBase.WndProc(m)
        End Sub

        Public Function PreFilterMessage(ByRef m As Message) As Boolean Implements IMessageFilter.PreFilterMessage
            Throw New NotImplementedException()
        End Function
    End Class

    <DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function ShowWindow(hWnd As IntPtr, nCmdShow As Integer) As Boolean
    End Function

    Private Const SW_RESTORE As Integer = 9


    Public doc As New Spire.Pdf.PdfDocument()
    Public instObj As Stream = Stream.Null
    Public backup As New List(Of String)
    Public backupUndo As New List(Of String)
    Public archive As IArchive = Nothing
    Public allProcesses As Process() = Nothing
    Dim lineCount As Integer = 0
    Dim output As New StringBuilder()
    Public targetList As ListBox
    Dim isclosing As Boolean = False
    Private recentAddedMenu As ToolStripMenuItem  ' 類別層級變數，存住這個子選單的參考

    ' ==========================================
    'Dim isopening As Boolean = False
    Public recordPath As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\Documents\easyPreview\"
    Public commonUsed As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\常用\"

    Public dataHasRead As Boolean = False
    '--------------------------------
    ' 在類級別預編譯正則表達式
    Public Shared twitterRgx As New Regex("mobile\.|/photo/\d|\?t.*|\?s.*|\?fb.*|\?cn.*", RegexOptions.Compiled)
    Public Shared pixivRgx As New Regex("\#.*|\?.*|(?<=pixiv\.net)\/en(?=\/)", RegexOptions.Compiled)
    Public Shared normalRgx As New Regex("\?.*", RegexOptions.Compiled)
    Public Shared DMMRgx As New Regex("\&.*", RegexOptions.Compiled)
    Public Shared dlsiteRgx As New Regex("(?:RJ|VJ|BJ)\d+", RegexOptions.Compiled)
    Public Shared nhentaiRgx As New Regex("/(\d+)/", RegexOptions.Compiled)

    Public Shared youtubeRgx As New Regex("https?://(?:www\.|m\.)?youtube\.com", RegexOptions.Compiled)
    Public Shared bilibiliRgx As New Regex("https?://(?:www\.|m\.)?bilibili\.com", RegexOptions.Compiled)
    Public libvlcOptions As String() = {"--aout=directsound"}
    Public _libvlc As New LibVLC(libvlcOptions)
    Public _mediaPlayer As New MediaPlayer(_libvlc)




    Private Sub PathOption(sender As RadioButton, e As EventArgs) Handles RadioButton14.Click, RadioButton1.Click
        If sender Is RadioButton14 Then
            'TextBox1.Text = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile).ToString & "\Downloads"
            Dim downloadPath As String = Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "{374DE290-123F-4565-9164-39C4925E467B}", Nothing)

            ' 如果沒有設定，則使用預設的下載目錄

            ' 取得使用者設定的下載目錄
            If Not downloadPath Then
                downloadPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\Downloads"
            End If
            PathComboBox.Text = downloadPath
        ElseIf sender Is RadioButton1 Then
            PathComboBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        End If
    End Sub
    Public Sub Form1_Close(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles Me.FormClosing
        'bug

        '取消關閉表單
        isclosing = True
        If Not PathComboBox.Text.EndsWith("開啟.trf") Then
            Form2.Store_trf_records(PathComboBox.Text, "openRecord", "開啟")
        End If
        'If Not CheckedListBox1_isCheck("Auto Save") Then Return
        If String.IsNullOrEmpty(PathComboBox.Text) Then Return
        '重選撥放音訊

        Refresh_backup()
        SaveASMRPlayRecord(FileCollection.SelectedItem)
        SendPipeMessage($"add@{PathComboBox.Text}", "easyPreview CloseRecord")
        'SendPipeMessage($"remove {Me.Text}")
        isclosing = False

    End Sub
    Private Sub Form1_beforeShown(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        '設定表單的主題和樣式
        Check_document()
        'load_passward()
        Add_additional_operation()
        Load_past_path()
        Load_stripMenuItem()
        Add_handlers()
        getOutputDevices()
        PictureBox1.Image = New Bitmap(1, 1)
        Form2.PictureBox1.Image = New Bitmap(1, 1)
        Core.Initialize()
        MagickNET.Initialize()
        targetList = FileCollection
        VideoView1.MediaPlayer = _mediaPlayer
    End Sub

    Private Sub Add_handlers()
        AddHandler PictureBox1.Click, AddressOf Form1_UnActiveControl
        AddHandler FileCollection.Click, AddressOf Form1_UnActiveControl
        AddHandler SubFileCollection.Click, AddressOf Form1_UnActiveControl
        AddHandler Panel35.Click, AddressOf Form1_UnActiveControl
        AddHandler Panel5.Click, AddressOf Form1_UnActiveControl
        AddHandler Panel34.Click, AddressOf Form1_UnActiveControl
        AddHandler Panel2.Click, AddressOf Form1_UnActiveControl
        AddHandler PathComboBox.SelectedIndexChanged, AddressOf Form1_UnActiveControl
        AddHandler TargetComboBox.SelectedIndexChanged, AddressOf Form1_UnActiveControl
        AddHandler ComboBox5.SelectedIndexChanged, AddressOf Form1_UnActiveControl
        'AddHandler OptionCheckedListBox.Click, AddressOf Form1_UnActiveControl
        'AddHandler OptionCheckedListBox.SelectedIndexChanged, AddressOf Form1_UnActiveControl
        'AddHandler Button10.Click, AddressOf Form1_UnActiveControl
        'AddHandler Button3.Click, AddressOf Form1_UnActiveControl
        AddHandler _mediaPlayer.EndReached, AddressOf MediaPlayer_EndReached
        AddHandler _mediaPlayer.TimeChanged, AddressOf OnTimeChanged
    End Sub

    Private Sub Check_document()
        Directory.CreateDirectory(recordPath)
        If Not File.Exists(recordPath & "downloadRecord.txt") Then File.Create(recordPath & "downloadRecord.txt").Close()
        If Not File.Exists(recordPath & "des_pathrecord.txt") Then File.Create(recordPath & "des_pathrecord.txt").Close()
        If Not File.Exists(recordPath & "ori_pathrecord.txt") Then File.Create(recordPath & "ori_pathrecord.txt").Close()
        If Not File.Exists(recordPath & "passward.txt") Then File.Create(recordPath & "passward.txt").Close()
        If Not File.Exists(recordPath & "add_record.txt") Then File.Create(recordPath & "add_record.txt").Close()
    End Sub
    ' 全局變量保存參數
    Dim commandargs As String = String.Empty
    Private Sub ImageMode() Handles MyBase.Shown
        If Environment.GetCommandLineArgs().Length < 2 Then Return
        Dim ext = Path.GetExtension(commandargs)
        If Ext_judge(ext, "Images") Then
            PictureBox1_DoubleClick(New Object, New EventArgs)
        End If

    End Sub
    Private Sub ProcessCommandArgs() Handles MyBase.Load

        If Environment.GetCommandLineArgs().Length < 2 Then Return
        ' System.Threading.Thread.Sleep(500)
        commandargs = Environment.GetCommandLineArgs()(1)



        Dim ext = Path.GetExtension(commandargs)

        If Ext_judge(ext, "easyRecord") Or Ext_judge(ext, ".txt") Then
            PathComboBox.Text = commandargs
            Return
        End If
        If Ext_judge(ext, "Images") Then
            PathComboBox.Text = Path.GetDirectoryName(commandargs)
            ExtList.SetItemChecked(0, True)
            ListBox_Click(New Object, New EventArgs)
            FileCollection.SelectedItem = commandargs
        ElseIf Ext_judge(ext, "Media") Then
            PathComboBox.Text = Path.GetDirectoryName(commandargs)
            ExtList.SetItemChecked(1, True)
            ListBox_Click(New Object, New EventArgs)
            FileCollection.SelectedItem = commandargs
        ElseIf Ext_judge(commandargs, "目錄") Then
            PathComboBox.Text = commandargs
            ListBox_Click(New Object, New EventArgs)
        End If

    End Sub
    Private Sub StartRecordClosePipe()
        While CheckedListBox1_isCheck("紀錄關閉視窗")

            Try
                Using pipeServer As New NamedPipeServerStream("easyPreview CloseRecord", PipeDirection.InOut)

                    pipeServer.WaitForConnection()
                    'Dim connectTask = pipeServer.WaitForConnectionAsync()
                    'If Not connectTask.Wait(500) Then
                    '    Continue While ' 每 0.5 秒檢查一次取消狀態
                    'End If


                    ' 讀取客戶端發送的消息
                    Using reader As New StreamReader(pipeServer, Encoding.UTF8)
                        Dim message As String = reader.ReadToEnd()
                        'Invoke(Sub() 壓縮檔案集合.Items.Add("收到的消息: " & message))
                        'If Not message.Contains(Me.Text) Then Return
                        If message.Contains("add") Then
                            Dim newitem As String = message.Split({"@"}, StringSplitOptions.RemoveEmptyEntries)(1)
                            Invoke(Sub()
                                       If Not CheckedListBox1_isCheck("紀錄關閉視窗") Then Return
                                       If Not newitem.EndsWith("開啟.trf") Then
                                           targetList = FileCollection
                                           AddLink(newitem)
                                       End If
                                   End Sub)
                        End If
                    End Using
                End Using
            Catch ex As Exception
                ' 處理可能的錯誤
                Invoke(Sub() PasswordBox.Items.Add("錯誤: " & ex.Message))
                'Return
            End Try
        End While
        Invoke(Sub() PasswordBox.Items.Add("停止記錄關閉視窗..."))
    End Sub
    Private Sub StartPipeServer(pipename As String)
        'Dim pipeIndex As Integer = 0 ' 初始索引
        Try
            Using checkClient As New NamedPipeClientStream(".", pipename, PipeDirection.InOut)
                checkClient.Connect(100) ' 嘗試連線 100ms
                ' 如果能連上，代表已有伺服器在運作，直接退出
                Using writer As New StreamWriter(checkClient)
                    writer.AutoFlush = True
                    writer.Write("ShowWindow")
                End Using
                Console.WriteLine("已有其他伺服器啟動，取消啟動本伺服器。")
                BeginInvoke(Sub() Me.Close())

            End Using
        Catch ex As Exception
            Console.WriteLine(ex.Message)

        End Try
        While True

            Try
                ' 動態生成唯一的管道名稱

                'pipeIndex += 1 ' 遞增索引，保證名稱唯一
                Using pipeServer As New NamedPipeServerStream(pipename, PipeDirection.InOut)


                    pipeServer.WaitForConnection()



                    ' 讀取客戶端發送的消息
                    Using reader As New StreamReader(pipeServer, Encoding.UTF8)
                        Dim message As String = reader.ReadToEnd()
                        Invoke(Sub() SubFileCollection.Items.Add("收到的消息: " & message))
                        'If Not message.Contains(Me.Text) Then Return
                        If message.Contains("add") Then
                            Dim newitem As String = message.Split({"@"}, StringSplitOptions.RemoveEmptyEntries)(1)
                            Invoke(Sub()
                                       If Not newitem.EndsWith("開啟.trf") Then AddLink(newitem)
                                   End Sub)
                        ElseIf message.Contains("Download") Then
                            Invoke(Sub() Button9_Click(New Object, New EventArgs))
                        ElseIf message.Contains("Next") Then
                            Dim tebutton As New Button With {
                                .Text = "next"
                            }
                            Invoke(Sub() Next_Page(tebutton, New EventArgs))
                        ElseIf message.Contains("Previous") Then
                            Dim tebutton As New Button With {
                                .Text = "previous"
                            }
                            Invoke(Sub() Next_Page(tebutton, New EventArgs))
                        ElseIf message.Contains("PlayPause") Then
                            Invoke(Sub() PlayOrPause())
                        ElseIf message = "ShowWindow" Then
                            Invoke(Sub() BringToFront(Me))
                        End If
                    End Using
                End Using
            Catch ex As Exception
                ' 處理可能的錯誤
                Invoke(Sub() PasswordBox.Items.Add("錯誤: " & ex.Message))
                'Return
            End Try
        End While
    End Sub
    Public Overloads Shared Sub BringToFront(targetForm As Form)
        If targetForm.WindowState = FormWindowState.Minimized Then
            ShowWindow(targetForm.Handle, SW_RESTORE)
        End If
        SetForegroundWindow(targetForm.Handle)
    End Sub

    Private Function SendPipeMessage(message As String, Optional file As String = "open_sample")
        Using pipeClient As New NamedPipeClientStream(".", file, PipeDirection.InOut)
            Try
                ' 嘗試在 100ms 內連接到伺服器
                pipeClient.Connect(100)
            Catch ex As TimeoutException
                ' 無法連接到伺服器，直接跳出，不發訊息
                Return False
            End Try

            ' 成功連接才執行以下內容
            Invoke(Sub() PasswordBox.Items.Add("連接到伺服器..."))

            ' 向伺服器發送消息
            Using writer As New StreamWriter(pipeClient)
                writer.AutoFlush = True
                writer.Write(message)
            End Using

            Invoke(Sub() PasswordBox.Items.Add("消息已發送."))
            Return True
        End Using
    End Function
    Private Sub Open_Directory(sender As Object, e As EventArgs)
        Dim listBox As ListBox = GetSourceControl(Of ListBox)(sender)
        If listBox Is Nothing Then Return

        For Each item In listBox.SelectedItems
            Dim argument As String = "/select, """ & Remove_label(item) & """"
            System.Diagnostics.Process.Start("explorer.exe", argument)
        Next
    End Sub

    Private Sub Top_dir()
        Dim target = FileCollection.SelectedIndices.Cast(Of Integer).ToList
        For i = 0 To FileCollection.SelectedIndices.Count - 1
            Dim index = target(i)
            If String.IsNullOrEmpty(FileCollection.Items(index)) Then Continue For
            FileCollection.Items(index) = Directory.GetParent(FileCollection.Items(index)).ToString
        Next
        delete_repeat_item()
        Refresh_backup()
    End Sub
    Public Function Add_subStripMenuItem(rightHotKey As ToolStripMenuItem, subName As String, Optional handler As EventHandler = Nothing) As ToolStripMenuItem
        ' 創建子選單項目
        Dim subHotKey As New ToolStripMenuItem(subName)
        If handler IsNot Nothing Then
            AddHandler subHotKey.Click, handler
        End If
        ' 將子選單項目添加到主選單項目
        rightHotKey.DropDownItems.Add(subHotKey)
        Return subHotKey
    End Function
    Public Function Add_ContextMenuStrip(targetMenu As ContextMenuStrip, mainName As String, Optional handler As EventHandler = Nothing) As ToolStripMenuItem
        ' 創建子選單項目
        Dim rightHotKey As New ToolStripMenuItem(mainName)
        If handler IsNot Nothing Then
            ' 為選單項目添加事件處理
            AddHandler rightHotKey.Click, handler
        End If
        ' 將子選單項目添加到指定的選單中
        targetMenu.Items.Add(rightHotKey)
        Return rightHotKey
    End Function

    Private Sub Load_stripMenuItem()
        Dim rightHotKey As ToolStripMenuItem
        Dim subHotKey As ToolStripMenuItem
        Dim subHotKey2 As ToolStripMenuItem
        '---------------------- ContextMenuStrip1---------------------------------
        rightHotKey = New ToolStripMenuItem("操控列表")
        ContextMenuStrip1.Items.Add(rightHotKey)
        Add_subStripMenuItem(rightHotKey, "移到最底", AddressOf rearrangeButton)
        Add_subStripMenuItem(rightHotKey, "移到頂端", AddressOf rearrangeButton)
        Add_subStripMenuItem(rightHotKey, "上移", AddressOf rearrangeButton)
        Add_subStripMenuItem(rightHotKey, "下移", AddressOf rearrangeButton)
        Add_subStripMenuItem(rightHotKey, "Delete_Repeat_Item", AddressOf delete_repeat_item)
        Add_subStripMenuItem(rightHotKey, "反轉列表", AddressOf Reverse_item_list)
        '------------------------------------------------------
        rightHotKey = New ToolStripMenuItem("尋找資源")
        ContextMenuStrip1.Items.Add(rightHotKey)
        For Each opt In Constants.GetValue("search_option")
            Add_subStripMenuItem(rightHotKey, opt, AddressOf Form2_operate)
        Next
        Add_ContextMenuStrip(ContextMenuStrip1, "Open File Directory", AddressOf Open_Directory)
        Add_ContextMenuStrip(ContextMenuStrip1, "Copy Path", AddressOf Clipboard_SetText)
        If Directory.Exists(commonUsed) Then
            rightHotKey = New ToolStripMenuItem("加入")
            ContextMenuStrip1.Items.Add(rightHotKey)
            subHotKey = New ToolStripMenuItem("最近加入")
            rightHotKey.DropDownItems.Add(subHotKey)
            recentAddedMenu = subHotKey        ' 存住參考
            Refresh_recentAddedMenu()          ' 首次填充
            subHotKey = New ToolStripMenuItem("ASMR")
            rightHotKey.DropDownItems.Add(subHotKey)
            For Each recordDir As String In GetAllDirectories($"{commonUsed}ASMR")
                subHotKey2 = New ToolStripMenuItem(recordDir)
                subHotKey.DropDownItems.Add(subHotKey2)
                For Each subrecord As String In Getallfiles(recordDir, "")
                    Add_subStripMenuItem(subHotKey2, subrecord, AddressOf Add_to_record)
                Next
            Next
            For Each record As String In Getallfiles($"{commonUsed}ASMR", "")
                Add_subStripMenuItem(subHotKey, record, AddressOf Add_to_record)
            Next
            'subHotKey = New ToolStripMenuItem("畫")
            'rightHotKey.DropDownItems.Add(subHotKey)
            'For Each record As String In getallfiles($"{commonUsed}畫", "")
            '    add_subStripMenuItem(subHotKey, record, AddressOf add_to_record)
            'Next
            subHotKey = New ToolStripMenuItem("網路整合")
            rightHotKey.DropDownItems.Add(subHotKey)
            For Each record As String In Getallfiles($"{commonUsed}", "*.trf")
                Add_subStripMenuItem(subHotKey, record, AddressOf Add_to_record)
            Next
            'subHotKey = New ToolStripMenuItem("個別網站")
            'rightHotKey.DropDownItems.Add(subHotKey)
            'For Each record As String In getallfiles($"{commonUsed}個別網站", "")
            '    subHotKey2 = New ToolStripMenuItem(record)
            '    AddHandler subHotKey2.Click, AddressOf add_to_record
            '    subHotKey.DropDownItems.Add(subHotKey2)
            '    Dim dir As String = Path.GetDirectoryName(record) & "\" & Path.GetFileNameWithoutExtension(record)
            '    If Not Directory.Exists(dir) Then Continue For
            '    For Each subrecord As String In getallfiles(dir, "")
            '        add_subStripMenuItem(subHotKey2, subrecord, AddressOf add_to_record)
            '    Next
            'Next

        End If
        Add_ContextMenuStrip(ContextMenuStrip1, "從剪貼簿匯入", AddressOf AddItem_from_clipboard)
        Add_ContextMenuStrip(ContextMenuStrip1, "從文字框輸入", AddressOf AddItem_from_textbox)
        Add_ContextMenuStrip(ContextMenuStrip1, "產生網址", AddressOf Generate_webAddress)
        Add_ContextMenuStrip(ContextMenuStrip1, "標記", AddressOf Handle_label_sender)
        Add_ContextMenuStrip(ContextMenuStrip1, "上一步", AddressOf listUndo)
        rightHotKey = New ToolStripMenuItem("編輯")
        ContextMenuStrip1.Items.Add(rightHotKey)
        Add_subStripMenuItem(rightHotKey, "重新命名", AddressOf EditItem)
        Add_subStripMenuItem(rightHotKey, "上一層", AddressOf Top_dir)
        Add_subStripMenuItem(rightHotKey, "上一集", AddressOf Find_Episode)
        Add_subStripMenuItem(rightHotKey, "下一集", AddressOf Find_Episode)
        '---------------------- ContextMenuStrip2---------------------------------
        Add_ContextMenuStrip(ContextMenuStrip2, "複製圖像", AddressOf Clipboard_SetIImage)
        Add_ContextMenuStrip(ContextMenuStrip2, "刪除縮圖", AddressOf DeleteSelectedThumbnail)
        '---------------------- ContextMenuStrip3---------------------------------
        Add_ContextMenuStrip(ContextMenuStrip3, "Copy Path", AddressOf Clipboard_SetText)
        rightHotKey = New ToolStripMenuItem("尋找資源")
        ContextMenuStrip3.Items.Add(rightHotKey)
        For Each opt In Constants.GetValue("search_option")
            Add_subStripMenuItem(rightHotKey, opt, AddressOf Form2_operate)
        Next
        Add_ContextMenuStrip(ContextMenuStrip3, "Open File Directory", AddressOf Open_Directory)
    End Sub
    Private Sub Refresh_recentAddedMenu()
        If recentAddedMenu Is Nothing Then Return
        recentAddedMenu.DropDownItems.Clear()
        Dim add_record_path = recordPath & "add_record.txt"
        If Not File.Exists(add_record_path) Then Return
        For Each record As String In File.ReadAllLines(add_record_path, Encoding.UTF8)
            If String.IsNullOrWhiteSpace(record) Then Continue For
            Add_subStripMenuItem(recentAddedMenu, record, AddressOf Add_to_record)
        Next
    End Sub

    Private Sub Handle_label_sender(sender As Object, e As EventArgs)
        Dim control As ListBox = GetSourceControl(Of ListBox)(sender)
        If control Is Nothing Then Return

        Dim index As Integer() = control.SelectedIndices.Cast(Of Integer).ToArray
        For Each i In index
            Handle_label(i)
        Next
    End Sub
    Private Sub Handle_label(checkIndex As Integer)
        Dim sitem As String = FileCollection.Items(checkIndex)
        If sitem.StartsWith("★") Then
            FileCollection.Items(checkIndex) = FileCollection.Items(checkIndex).ToString.Substring(1)
        Else
            FileCollection.Items(checkIndex) = "★" + FileCollection.Items(checkIndex)
        End If
    End Sub
    Public Sub Generate_webAddress(sender As Object, e As EventArgs)
        Dim control As ListBox = GetSourceControl(Of ListBox)(sender)
        If control Is Nothing Then Return

        Console.WriteLine($"add item from {control}")
        For Each target In control.SelectedItems
            ' 同時處理空格、冒號與反斜線，確保手機版 LINE 絕對不發瘋
            Dim encodedTarget As String = target.ToString() _
                                             .Replace(" ", "%20") _
                                             .Replace(":", "%3A") _
                                             .Replace("\", "%5C")

            SubFileCollection.Items.Add($"https://easypreview.unified-storage.com/read-file?file={encodedTarget}")
        Next
    End Sub


    Private Sub InitializeContextMenu(menu As ContextMenuStrip, items As List(Of Tuple(Of String, EventHandler)))
        For Each item In items
            Dim menuItem = New ToolStripMenuItem(item.Item1)
            If item.Item2 IsNot Nothing Then AddHandler menuItem.Click, item.Item2
            menu.Items.Add(menuItem)
        Next
    End Sub

    ' 取得 ContextMenuStrip1 的選單項目和事件處理函數


    Private Sub DeleteSelectedThumbnail()
        Dim curitem = FileCollection.SelectedItem.ToString
        Delete_thumbnail(curitem)
    End Sub
    Private Sub Delete_thumbnail(curitem As String)
        release_all_process()
        Dim fileName = Make_Character_Valid(curitem, "png")
        ' 組合檔案名稱和副檔名，例如inputjpg.png
        Dim preImagePath = recordPath & "clip圖像\" & fileName
        File.Delete(preImagePath)
        Dim psi As New ProcessStartInfo With {
            .FileName = "cmd.exe",
            .Arguments = "/c nconvert -out png -q 50 -o """ & preImagePath & """ """ & curitem & """",
            .WindowStyle = ProcessWindowStyle.Hidden
        }
        Process.Start(psi)
    End Sub

    Public Sub Add_to_record(sender As Object, e As EventArgs)
        Dim btn_name As String = CType(sender, ToolStripMenuItem).Text
        Console.WriteLine(btn_name)
        Add_Record(btn_name)
        Refresh_recentAddedMenu()
        Dim files As List(Of String) = FileCollection.SelectedItems.Cast(Of String).ToList
        If CheckBox5.Checked Then
            files = FileCollection.Items.Cast(Of String).ToList
        End If
        If SendPipeMessage($"connect", $"easyPreview {btn_name}") Then
            For Each foundfile In files
                SendPipeMessage($"add@{foundfile}", $"easyPreview {btn_name}")
            Next
        Else
            Trf_backup(btn_name)
            Using sr As New StreamWriter(btn_name, True)
                For Each foundfile In files
                    sr.WriteLine(foundfile)
                Next
            End Using
        End If
        If CheckedListBox1_isCheck("執行後移出") Then
            System.Threading.SynchronizationContext.Current.Post(Sub()
                                                                     Remove_Button(sender, e)
                                                                 End Sub, Nothing)

        End If
    End Sub
    Private Sub Form2_operate(sender As Object, e As EventArgs)
        Form2.FindSource(sender, e)
    End Sub
    Private Sub Reverse_item_list(sender, e)
        Dim newlist As Array = FileCollection.Items.Cast(Of String).ToArray
        Array.Reverse(newlist)
        FileCollection.Items.Clear()
        For Each i In newlist
            FileCollection.Items.Add(i)
        Next
        Refresh_backup()
    End Sub
    Public Function Remove_label(name As String) As String
        If name Is Nothing Then Return name
        If name.StartsWith("★") Then name = name.Substring(1)
        Return name
    End Function
    Public Function Standardized_denomination(name As String) As String
        '移除開頭標記
        name = Remove_label(name)
        '判斷是否為單獨網站分析，並定向
        If name.Contains("x.com") Or name.IndexOf("twitter.com") > 0 Then Return twitter_repeat(name)
        If name.Contains("%") Or name.Contains("google.com/search") Then Return DecodeUri(name)
        If name.Contains("pixiv") Then Return pixivRgx.Replace(name, "")
        If name.Contains("dlsite") Then Return dlsite_repeat(name)
        If name.Contains("nhentai") Then Return nhentai_repeat(name)
        If name.Contains("youtube") Then Return Youtube_repeat(name)
        If name.Contains("bilibili") Then Return Bilibili_repeat(name)
        If name.Contains("facebook") Then Return Facebook_repeat(name)
        If name.Contains("fanbox") Or name.Contains("dl.getchu") Then Return normalRgx.Replace(name, "")
        If name.Contains("dmm.co") Then Return DMMRgx.Replace(name, "")
        Return name
    End Function
    Public Function Youtube_repeat(name As String) As String
        name = youtubeRgx.Replace(name, "https://youtube.com")
        Return DecodeUri(name)
    End Function
    Public Function Bilibili_repeat(name As String) As String
        name = bilibiliRgx.Replace(name, "https://www.bilibili.com")
        Return DecodeUri(name)
    End Function

    Public Function DecodeUri(encodedUrl As String) As String
        Try
            ' 解析 URL
            Dim uri As New Uri(encodedUrl)

            ' 解碼 Path
            Dim decodedPath As String = Uri.UnescapeDataString(uri.AbsolutePath)

            ' 解碼 Fragment（無論是否 Google 搜尋都處理）
            Dim decodedFragment As String = Uri.UnescapeDataString(uri.Fragment)
            If decodedFragment.Equals("#google_vignette", StringComparison.OrdinalIgnoreCase) Then decodedFragment = ""
            ' 處理 Query
            Dim decodedQuery As String = ""
            If uri.Host.Contains("google.") AndAlso uri.AbsolutePath.StartsWith("/search") Then
                Dim queryParams = System.Web.HttpUtility.ParseQueryString(uri.Query)
                Dim qValue As String = queryParams("q")
                If Not String.IsNullOrEmpty(qValue) Then
                    decodedQuery = "?q=" & System.Web.HttpUtility.UrlDecode(qValue)
                End If
                ' 🎨 Pixiv jump.php 特例：query 直接是被 encode 的 URL
            ElseIf uri.Host.Contains("pixiv.net") AndAlso uri.AbsolutePath.Equals("/jump.php", StringComparison.OrdinalIgnoreCase) Then
                If uri.Query.StartsWith("?http") Then
                    ' 去掉 ? 再解碼整個 URL
                    decodedQuery = Uri.UnescapeDataString(uri.Query.Substring(1))
                    Return decodedQuery ' 直接返回跳轉的真實網址
                End If
                ' ⭐ 新增 YouTube 處理
            ElseIf uri.Host.Contains("youtube.com") AndAlso uri.AbsolutePath.StartsWith("/watch") Then
                Dim queryParams = System.Web.HttpUtility.ParseQueryString(uri.Query)
                Dim vValue As String = queryParams("v")

                If Not String.IsNullOrEmpty(vValue) Then
                    Return uri.Scheme & "://" & uri.Host & "/watch?v=" & vValue
                End If
            ElseIf uri.Host.Contains("youtube.com") AndAlso uri.AbsolutePath.StartsWith("/shorts") Then
                Return uri.Scheme & "://" & uri.Host & decodedPath
            ElseIf uri.Host.Contains("bilibili.com") AndAlso uri.AbsolutePath.StartsWith("/video") Then
                ' bilibili規則：只保留主機與路徑，完全剔除追蹤參數（spm_id_from, vd_source 等）
                Dim cleanUri As New UriBuilder(uri) With {
                    .Query = String.Empty, ' 直接清空所有 Query 參數
                    .Scheme = "https",     ' 強制使用 https
                    .Host = "www.bilibili.com" ' 建議統一規格為 www.bilibili.com 瀏覽體驗較佳
                }

                ' 如果你有後續的解碼需求（例如處理中間的 %XX 編碼），可以在這裡呼叫。
                ' 若無，直接返回去除了 Query 的字串
                Return cleanUri.Uri.ToString()
            Else
                decodedQuery = Uri.UnescapeDataString(uri.Query)
            End If

            ' 重建 URL
            Dim decodedUrl As String = uri.Scheme & "://" & uri.Host & decodedPath & decodedQuery & decodedFragment
            Return decodedUrl

        Catch ex As Exception
            Console.WriteLine(ex)
            Return encodedUrl ' 解析失敗則返回原始 URL
        End Try
    End Function
    Public Function Twitter_repeat(name As String)
        If name.Contains("www.youtube.com/redirect") Then
            ' 檢查網址是否包含 "www.youtube.com/redirect"
            name = ExtractRedirectUrl(name)
        End If
        name = twitterRgx.Replace(name, "")
        'Console.WriteLine(name)
        name = name.Replace("twitter.com", "x.com")
        Return name
    End Function
    Public Function ExtractRedirectUrl(youtubeUrl As String) As String

        ' 解析 URL 中的 "q=" 參數值
        Dim queryStart As Integer = youtubeUrl.IndexOf("q=")

        If queryStart > -1 Then
            ' 提取 "q=" 之後的字串並解碼
            Dim encodedUrl As String = youtubeUrl.Substring(queryStart + 2)
            Dim endIndex As Integer = encodedUrl.IndexOf("&")

            If endIndex > -1 Then
                encodedUrl = encodedUrl.Substring(0, endIndex)
            End If

            ' 將 URL 解碼
            Dim redirectUrl As String = WebUtility.UrlDecode(encodedUrl)
            Return redirectUrl
        End If
        ' 如果條件不符合或沒有找到 "q" 參數，返回原始網址
        Return youtubeUrl
    End Function
    Public Function Dlsite_repeat(name As String)
        Dim dlmatch As Match = dlsiteRgx.Match(name)
        If dlmatch.Success Then
            Dim fn As String = dlmatch.Value
            Dim sb As String = ""
            If fn.Contains("RJ") Then
                Dim ifAnocement As Boolean = name.Contains("announce")
                If Not ifAnocement Then
                    sb = "maniax/work"
                Else
                    sb = "maniax/announce"
                End If
            ElseIf fn.Contains("VJ") Then
                sb = "pro/work"
            ElseIf fn.Contains("BJ") Then
                sb = "books/work"
            End If
            name = $"https://www.dlsite.com/{sb}/=/product_id/{fn}.html"
        End If
        Return name
    End Function
    Public Function Facebook_repeat(name As String)
        If name.Contains("m.facebook") Then
            name = name.Replace("m.facebook", "www.facebook")
        End If
        Return name
    End Function
    Public Function Nhentai_repeat(name As String)
        Dim nhentaiMatch As Match = nhentaiRgx.Match(name)
        If nhentaiMatch.Success Then
            Dim fn As String = nhentaiMatch.Groups(1).Value
            name = "https://nhentai.net/g/" & fn
        End If
        Return name
    End Function
    Public Async Sub Delete_repeat_item()
        If CheckedListBox1_isCheck("取消自動刪除重複") Then Return
        Dim removeFromStart As Boolean = CheckedListBox1_isCheck("重複移到最底")
        Dim ori_item As String = If(FileCollection.SelectedItem?.ToString(), String.Empty)
        Dim ori_index As Integer = If(FileCollection.SelectedIndex >= 0, FileCollection.SelectedIndex, -1)

        If FileCollection.Items.Count = 0 Then Return

        Dim items As List(Of String) = FileCollection.Items.Cast(Of String).ToList()

        ' 處理重複項目
        Dim processedItems As List(Of String) = Await ProcessItemsAsync(items, removeFromStart)

        ' 更新列表
        FileCollection.BeginUpdate()
        FileCollection.Items.Clear()
        FileCollection.Items.AddRange(processedItems.ToArray())
        FileCollection.EndUpdate()
        ' 恢復選擇
        If Not String.IsNullOrEmpty(ori_item) AndAlso FileCollection.Items.Contains(ori_item) Then
            FileCollection.SelectedItem = ori_item
        ElseIf FileCollection.Items.Count > ori_index Then
            FileCollection.SelectedIndex = ori_index
        End If
        Refresh_listbox_numbers()
    End Sub
    Async Function ProcessItemsAsync(items As List(Of String), removeFromStart As Boolean) As Task(Of List(Of String))
        Return Await Task.Run(Function()
                                  Dim dict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                                  Dim indices As IEnumerable(Of Integer) =
                                  If(removeFromStart,
                                     Enumerable.Range(0, items.Count),
                                     Enumerable.Range(0, items.Count).Reverse())


                                  For Each i In indices
                                      Dim item = items(i)
                                      Dim isLabel As Boolean = item.StartsWith("★")
                                      Dim standardized As String = Standardized_denomination(item)
                                      Dim finalName As String = If(isLabel, "★", "") + standardized

                                      If Not dict.ContainsKey(standardized) Then
                                          dict.Add(standardized, finalName)
                                      ElseIf isLabel Then
                                          dict(standardized) = finalName
                                      End If
                                  Next

                                  Dim result As List(Of String) = dict.Values.ToList()

                                  ' 如果你要維持原順序（可選）
                                  If Not removeFromStart Then
                                      result.Reverse()
                                  End If

                                  Return result
                              End Function)
    End Function

    Public Function GetDedupKey(standardized As String) As String
        If standardized.Contains("x.com") Then
            Dim m As Match = Regex.Match(standardized, "/status/(\d+)")
            If m.Success Then Return "x_status_" & m.Groups(1).Value
        End If
        Return standardized
    End Function

    Public Async Function ResolveIToUsername(url As String) As Task(Of String)
        Try
            Using handler As New HttpClientHandler() With {.AllowAutoRedirect = True}
                Using client As New HttpClient(handler) With {.Timeout = TimeSpan.FromSeconds(5)}
                    Using request As New HttpRequestMessage(HttpMethod.Head, url)
                        Dim response As HttpResponseMessage = Await client.SendAsync(request)
                        Return response.RequestMessage.RequestUri.ToString()
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Return url ' 失敗就保留原本的
        End Try
    End Function

    ''' <summary>
    ''' 從 ContextMenuStrip 事件取得特定類型的來源控制項 (通用泛型版本)
    ''' 用法: Dim listBox As ListBox = GetSourceControl(Of ListBox)(sender)
    ''' </summary>
    Private Function GetSourceControl(Of T As Control)(sender As Object) As T
        Try
            Dim menuItem As ToolStripMenuItem = TryCast(sender, ToolStripMenuItem)
            If menuItem Is Nothing Then Return Nothing

            Dim strip As ContextMenuStrip = TryCast(menuItem.Owner, ContextMenuStrip)
            If strip Is Nothing Then Return Nothing

            Return TryCast(strip.SourceControl, T)
        Catch ex As Exception
            Console.WriteLine($"錯誤: 無法轉換控制項為 {GetType(T).Name} - {ex.Message}")
            Return Nothing
        End Try
    End Function

    Public Sub ToolStripMenuItemsetTargetlistbox(owner As Object)


        ' 往上找到 ContextMenuStrip，再透過 SourceControl 取得觸發它的 ListBox
        Do While TypeOf owner IsNot ContextMenuStrip
            If TypeOf owner Is ToolStripDropDownMenu Then
                owner = CType(owner, ToolStripDropDownMenu).OwnerItem.Owner
            Else
                Exit Do
            End If
        Loop
        Dim cms As ContextMenuStrip = TryCast(owner, ContextMenuStrip)
        Console.WriteLine(cms.Name)
        If cms IsNot Nothing Then
            Dim findlistbox As ListBox = TryCast(cms.SourceControl, ListBox)
            Console.WriteLine(findlistbox.Name)
            If findlistbox IsNot Nothing Then
                targetList = findlistbox
            End If
        End If
    End Sub

    Private Sub RearrangeButton(sender As Object, e As EventArgs)
        Dim menuItem As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        Dim btn_name As String = menuItem.Text
        Dim owner As ToolStrip = menuItem.Owner
        ToolStripMenuItemsetTargetlistbox(owner)

        ' 根據按鈕名稱執行相應操作
        Select Case btn_name
            Case "下移" : ListRearrange(1)
            Case "上移" : ListRearrange(-1)
            Case "移到最底" : ListRearrange(targetList.Items.Count)
            Case "移到頂端" : ListRearrange(0 - targetList.Items.Count)
        End Select
    End Sub

    'listRearrange只能移動選取的第一項
    Private Sub ListRearrange(offset As Integer)
        Dim selectedItem As String = targetList.SelectedItem
        'Dim itemCount As Integer = targetList.Items.Count
        Dim targets As String() = targetList.SelectedItems.Cast(Of String).ToArray
        Dim targetsindex As Integer() = targetList.SelectedIndices.Cast(Of Integer).ToArray
        ' 复制索引，以便排序和修改
        Dim oriForm As List(Of String) = targetList.Items.Cast(Of String).ToList()

        ' 开始移动
        Dim targetSet As New List(Of String)
        For Each target In targets                          ' 如果插入的項目在index後面移除
            oriForm.Remove(target)
        Next
        Dim newIndex = targetList.Items.IndexOf(selectedItem) + offset
        If newIndex < 0 Then newIndex = 0
        If newIndex > targetList.Items.Count Then newIndex = targetList.Items.Count
        ' 插入新項目
        targetSet.AddRange(oriForm.Take(newIndex))       ' 插入前面的部分
        targetSet.AddRange(targets)                          ' 插入目標項目
        targetSet.AddRange(oriForm.Skip(newIndex))      ' 插入後面的部分
        ' 更新顯示
        targetList.Items.Clear()
        targetList.Items.AddRange(targetSet.ToArray())
        ' 重新选中新的索引
        If CheckedListBox1_isCheck("移動時選取原項") Then
            ' 重新选中新的索引
            For Each selectItem In targets
                Dim index = targetList.Items.IndexOf(selectItem)
                targetList.SetSelected(index, True)
            Next
        Else
            For Each selectIndex In targetsindex
                targetList.SetSelected(selectIndex, True)
            Next
        End If
        Refresh_listbox_numbers()
    End Sub
    Public Sub ListUndo()
        Dim select_index = FileCollection.SelectedIndex
        Dim c = backup
        backup = backupUndo
        backupUndo = c
        FileCollection.BeginUpdate()
        FileCollection.Items.Clear()
        For Each i In backup
            FileCollection.Items.Add(i)
        Next
        FileCollection.EndUpdate()
        If select_index >= 0 AndAlso select_index < FileCollection.Items.Count Then
            FileCollection.SetSelected(select_index, True)
        End If
        Refresh_listbox_numbers()
    End Sub
    Private Function Select_keyword(files As String(), keywords As String()) As String()
        For Each keyword In keywords
            Console.WriteLine($"現在排除{keyword}前，剩餘{files.Count}個檔案")
            files = Array.FindAll(Of String)(files, Function(x) x.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) = -1)
        Next
        Return files
    End Function
    Private Function Check_file_ext(files As String())
        Dim exts As String() = Constants.GetValue("asmr_ext_in_turn")
        Dim extfiles As New List(Of Array)()
        For Each fileExt In exts
            'Console.WriteLine(fileExt)
            Dim matchingFiles = Array.FindAll(files, Function(s) Path.GetExtension(s) = fileExt)
            extfiles.Add(matchingFiles)
        Next
        ' 找到數組的最大值
        Dim maxValue As Integer = extfiles.Max(Function(x) x.Length)

        ' 找到最大值的索引位置
        Dim maxIndex As Integer = extfiles.FindIndex(Function(x) x.Length = maxValue)
        'Console.WriteLine(exts(maxIndex))
        'Console.WriteLine(maxValue)
        If maxValue = 0 Then Return New String() {}
        If extfiles(0).Length > maxValue / 2 Then
            Return extfiles(0)
        Else
            Return extfiles(maxIndex)
        End If
    End Function
    Private Function Check_file_keyword(files As String())
        Dim allRadiosUnchecked As Boolean = Not GroupBox3.Controls.OfType(Of RadioButton)().Any(Function(r) r.Checked)
        If allRadiosUnchecked Then
            ' 如果所有 RadioButton 都沒有被選中，這裡將執行相應的代碼
            Return files
        End If
        Dim controls As RadioButton = GroupBox3.Controls.OfType(Of RadioButton)().FirstOrDefault(Function(n) n.Checked)
        Dim content As String() = Array.FindAll(files, Function(s) s.IndexOf($"-{controls.Text}", StringComparison.Ordinal) > 0)
        If content.Length <> 0 Then
            Return content
        Else
            Return "no_compare"
        End If
    End Function
    Private Function Find_all_asmr_sound(target As Array) As Array
        ' 取得目錄下所有 .mp3 檔案
        Dim files As String() = Select_keyword(target, Constants.GetValue("asmrExtraKeywords"))
        Console.WriteLine($"找到{files.Count}個檔案")
        files = Check_file_ext(files)

        If FileCollection.SelectedItem.Contains("テグラユウキ") Then
            files = Check_file_keyword(files)
        End If
        ' 篩選關鍵字
        files = Select_keyword(files, Constants.GetValue("asmrSoundKeywords"))
        If files.Length = 0 Then
            Console.WriteLine("沒有找到任何音訊檔")
            Return New String() {}
        End If
        Array.Sort(files, AddressOf CompareNatural)
        Return files
    End Function

    Private Sub MediaPlayer_EndReached(sender As Object, e As EventArgs)
        If SubFileCollection.InvokeRequired Then
            SubFileCollection.Invoke(Sub() HandleMediaEnd())
        Else
            HandleMediaEnd()
        End If
    End Sub
    Public newMediaList As Boolean = False
    Private Sub HandleMediaEnd()
        Try
            lastPosition = 0
            Dim playList As ListBox = Nothing
            Dim asmrSounds As String() = Nothing
            Dim mediaPath As String = New Uri(_mediaPlayer.Media.Mrl).LocalPath
            asmrSounds = find_all_asmr_sound(SubFileCollection.Items.Cast(Of String).ToArray)
            If Not SubFileCollection.SelectedIndex > -1 Or SubFileCollection.Items.Count = 0 Or asmrSounds Is Nothing Then
                playList = FileCollection
            Else
                playList = SubFileCollection

                Dim checkitem As String = asmrSounds(asmrSounds.Count - 1)
                Dim checkitemUri As String = New Uri(checkitem).LocalPath

                If mediaPath = checkitemUri Then
                    playList = FileCollection
                    playover = True
                End If
            End If
            Console.WriteLine(playList.Text)
            If ComboBox5.SelectedItem IsNot Nothing AndAlso ComboBox5.SelectedItem.ToString() = "自動重播(單首)" Then
                If playList Is FileCollection Then
                    File_Preview(False)
                Else
                    Archive_Preview(New Object, New EventArgs)
                End If

                Return
            End If
            If playList.SelectedIndex < playList.Items.Count - 1 Then
                If FileCollection.SelectedItem.ToString.Contains("テグラユウキ") Or playList Is SubFileCollection Then
                    Dim nextIndex = Array.IndexOf(asmrSounds, mediaPath) + 1
                    Console.WriteLine(mediaPath)
                    playList.ClearSelected()
                    playList.SelectedItem = asmrSounds(nextIndex)
                Else
                    Dim selectindex As Integer = playList.SelectedIndex + 1
                    playList.ClearSelected()
                    newMediaList = True

                    Console.WriteLine($"s:{selectindex}")
                    playList.SelectedIndex = selectindex
                End If
            Else
                If ComboBox5.SelectedItem = "自動重播(全部)" Then
                    If playList.SelectedIndex = playList.Items.Count - 1 Then
                        playList.SetSelected(playList.SelectedIndex, False)
                        playList.SetSelected(0, True)
                    End If
                Else
                    playList.SetSelected(playList.SelectedIndex, False)
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("HandleMediaEnd 發生錯誤：" & ex.ToString())
        End Try


    End Sub
    Private Sub VolumeTrackBar_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.ValueChanged
        _mediaPlayer.Volume = TrackBar1.Value
    End Sub
    ' 當前播放時間改變時更新 ProgressBar
    Private lastUpdate As Long = 0
    Private Sub OnTimeChanged(sender As Object, e As MediaPlayerTimeChangedEventArgs)
        Dim now = Environment.TickCount

        ' 限制最短更新間隔：例如 50ms（每秒最多 20 次）
        If now - lastUpdate < 50 Then Return
        lastUpdate = now
        If ProgressBar2.InvokeRequired Then
            ProgressBar2.BeginInvoke(New Action(Sub() UpdateProgressBar()))
        Else
            UpdateProgressBar()
        End If
    End Sub
    ' 將更新進度條的邏輯封裝到單獨的方法中
    ' 更新進度條，根據當前播放進度
    Private Sub UpdateProgressBar()
        ' Media 尚未準備好時忽略
        If _mediaPlayer.Media Is Nothing Then Exit Sub
        If _mediaPlayer.Media.Duration <= 0 Then Exit Sub
        Dim newValue As Integer = CInt((_mediaPlayer.Time / _mediaPlayer.Media.Duration) * ProgressBar2.Maximum)
        If Double.IsNaN(newValue) OrElse Double.IsInfinity(newValue) Then Exit Sub
        If newValue > 1000 Then newValue = 1000
        ProgressBar2.Value = newValue ' 直接設置值，而不是逐漸增加

        ' 顯示時間（假設有 Label11）
        Dim currentTime As TimeSpan = TimeSpan.FromMilliseconds(_mediaPlayer.Time)
        Dim totalTime As TimeSpan = TimeSpan.FromMilliseconds(_mediaPlayer.Media.Duration)
        Label11.Text = $"{currentTime:mm\:ss} / {totalTime:mm\:ss}"
    End Sub

    ' 點擊 ProgressBar 時跳轉進度
    Private Sub ProgressBar2_Click(sender As Object, e As EventArgs) Handles ProgressBar2.Click
        ' 將游標位置轉換為 ProgressBar2 內的相對位置
        Dim pos As System.Drawing.Point = ProgressBar2.PointToClient(Cursor.Position)

        ' 計算點擊位置的比例
        Dim clickRatio As Double = pos.X / ProgressBar2.Width

        ' 計算新的進度值，並根據比例設定播放時間
        Dim newValue As Integer = CInt(clickRatio * ProgressBar2.Maximum)
        Dim newTime As Long = CLng(clickRatio * _mediaPlayer.Media.Duration)


        ' 設置 MediaPlayer 的播放時間
        _mediaPlayer.Time = newTime

        ' 更新進度條以反映新進度
        ProgressBar2.Value = Math.Min(newValue, ProgressBar2.Maximum)

        ' 手動更新進度條以反映新的進度
        'UpdateProgressBar()
    End Sub


    Private Sub StopButton_Click(sender As Object, e As EventArgs) Handles Button25.Click
        _mediaPlayer.Stop()
        Button24.Text = "播放"
        ProgressBar2.Value = 0
    End Sub

    Private Sub Button2_Dialog(sender As Button, e As EventArgs) Handles Button1.Click, Button2.Click
        Dim target As ComboBox = If(sender Is Button1, PathComboBox, TargetComboBox)
        Using fbd As New FolderBrowserDialog(Me)
            fbd.DirectoryPath = target.Text
            If fbd.ShowDialog = DialogResult.OK Then
                target.Text = fbd.DirectoryPath
            End If
        End Using
        If CheckBox2.Checked And sender Is Button2 Then
            TargetComboBox.Text += Path.DirectorySeparatorChar & System.DateTime.Now.ToString("yyyy_MM_dd")
        End If
    End Sub

    Private Sub Checkbox7_Dialog(sender As Object, e As EventArgs) Handles Button6.Click
        OpenFileDialog1.Filter = "trf files (*.trf)|*.trf|wmc files (*.wmc)|*.trf|All files (*.*)|*.*"
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            PathComboBox.Text = OpenFileDialog1.FileName
        End If
    End Sub
    Public Sub Trf_backup(filename As String)
        Try
            Dim saveFolder As String = $"{commonUsed}\backup\" & DateTime.Now.ToString("yyyy_MM_dd") & Path.DirectorySeparatorChar
            Directory.CreateDirectory(saveFolder)
            Dim backupFile As String = saveFolder & Path.GetFileName(filename)
            If File.Exists(backupFile) Then Return
            My.Computer.FileSystem.CopyFile(filename, backupFile)
            MessageLabel.Text = Path.GetFileName(backupFile) & "備份成功"
        Catch ex As Exception
            MessageLabel.Text = "備份失敗：" & ex.Message
        End Try
    End Sub
    Private Sub CleanOldBackups(rootFolder As String, keepDays As Integer)
        Try
            If Not Directory.Exists(rootFolder) Then Return
            For Each ddDir In Directory.GetDirectories(rootFolder)
                Dim folderDate As Date
                If Date.TryParseExact(Path.GetFileName(ddDir), "yyyy_MM_dd",
                    Nothing, Globalization.DateTimeStyles.None, folderDate) Then
                    If (Date.Now - folderDate).TotalDays > keepDays Then
                        Directory.Delete(ddDir, True)
                    End If
                End If
            Next
        Catch
            ' 清理失敗不影響主流程
        End Try
    End Sub

    Dim isReadingData As Boolean = False
    Public Async Function Read_text_data(oripath As String, targetList As ListBox, Optional ignoreReading As Boolean = False) As Task
        If isReadingData And Not ignoreReading Then
            Return
        End If
        isReadingData = True
        Dim items As New HashSet(Of String)
        Dim operates As New HashSet(Of String)
        Dim rgx As New Regex("(http.*)")
        If oripath.Contains(".txt") Then
            rgx = New Regex("(htp.*)")
        End If
        Dim operate_rgx As New Regex("^(\[.*\])")
        Dim firencording As Encoding = Encoding.UTF8
        ' 進度回報器
        Dim progress = New Progress(Of Integer)(Sub(p)
                                                    ProgressBar1.Value = p
                                                End Sub)

        Await Task.Run(Sub()
                           Using sr As New StreamReader(oripath, firencording)
                               Dim totalLines = File.ReadLines(oripath).Count() ' 預估總行數
                               Dim i As Integer = 0

                               While Not sr.EndOfStream
                                   Dim item = sr.ReadLine()
                                   If rgx IsNot Nothing AndAlso rgx.IsMatch(item) Then
                                       items.Add(rgx.Match(item).Groups(1).Value)
                                   ElseIf operate_rgx.IsMatch(item) And Not item.Contains("[crawler]") Then
                                       operates.Add(item)
                                   Else
                                       items.Add(item)
                                   End If

                                   i += 1
                                   If i Mod 100 = 0 Then
                                       CType(progress, IProgress(Of Integer)).Report(CInt(i / totalLines * 100))
                                   End If
                               End While
                           End Using
                       End Sub)

        For Each operate In operates
            Console.WriteLine(operate)
            Dim prefix As String() = operate.Split(“]”)
            ' Get the prefix of the item 
            Select Case prefix(0)
                Case “[selected”
                    'If targetList.Items.Count = 0 Then Return
                    Dim index As Integer = CInt(prefix(1))
                    If index <> -1 Then lastSelectedIndex = index
                Case “[subselected”

                    Dim index = CInt(prefix(1))

                    If targetList.Name IsNot "subFileCollection" Then sublastSelectedIndex = index
                Case "[playTime"
                    lastPosition = CLng(prefix(1))
                Case "[volume"
                    Console.WriteLine(CLng(prefix(1)))
                    _mediaPlayer.Volume = CLng(prefix(1))
                Case "[Form2_open"
                    If targetList Is FileCollection Then PictureBox1_DoubleClick(New Object, New EventArgs)
            End Select
        Next

        FileCollection.BeginUpdate()
        targetList.Items.AddRange(items.ToArray)
        FileCollection.EndUpdate()
        If targetList IsNot FileCollection Then Return
        Additional_operate(oripath)
        If FileCollection.SelectedIndex = -1 And lastSelectedIndex < targetList.Items.Count And lastSelectedIndex <> -1 Then
            targetList.Invoke(Sub() FileCollection.SelectedIndex = lastSelectedIndex)
            last_selectedItem = FileCollection.SelectedItem
        End If
        If sublastSelectedIndex < SubFileCollection.Items.Count And sublastSelectedIndex <> -1 Then
            SubFileCollection.SelectedIndex = sublastSelectedIndex
        End If

        isReadingData = False
        ' 傳回 HashSet 物件
    End Function
    Private Sub Additional_operate(oripath As String)
        'isCodeChange = True
        Dim ext = Path.GetExtension(oripath)
        Console.WriteLine(oripath)
        If ext = ".apr" Then Return
        If ext <> ".txt" Then CheckedListBox1_setCheck("Auto Save", True)
        If oripath.Contains("clip") Then
            CheckedListBox1_setCheck("clip啟動", True)
            CheckedListBox1_setCheck("使用nconvert", True)
            RadioButton2.Checked = True
            RadioButton2_Watch(New Object, New EventArgs)
        ElseIf oripath.Contains("網頁") Or oripath.Contains("[web]") Or oripath.Contains("dlsite") Or oripath.Contains("hanime") Then
            CheckedListBox1_setCheck("重複移到最底", True)
        ElseIf oripath.Contains("開啟") Then
            CheckedListBox1_setCheck("執行後移到最後一項", True)
            CheckedListBox1_setCheck("重複移到最底", True)
        ElseIf oripath.IndexOf("ASMR", StringComparison.OrdinalIgnoreCase) > 0 Then
            CheckedListBox1_setCheck("ASMR播放模式", True)
            CheckedListBox1_setCheck("移動時選取原項", True)
            CheckedListBox1_setCheck("隱藏壓縮檔案集合", False)
            MeasureDevice("耳機")
            ' If ComboBox6.MaxLength = 0 Then ComboBox6_click(New Object, New EventArgs)
        ElseIf oripath.Contains("遊戲公告") Then
            CheckedListBox1_setCheck("取消自動刪除重複", True)

        ElseIf oripath.Contains("遊戲") Then
            ' If ComboBox6.MaxLength = 0 Then ComboBox6_click(New Object, New EventArgs)

            game_Watch(New Object, New EventArgs)
            CheckedListBox1_setCheck("locale emulator", True)
            CheckedListBox1_setCheck("隱藏預覽", True)
            CheckedListBox1_setCheck("隱藏壓縮檔案集合", True)
            CheckedListBox1_setCheck("重複移到最底", True)
            'CheckedListBox1_setCheck("關閉自動修正", True)
        End If
        Dim filename As String = Path.GetFileName(oripath)
        Form3.FileNameTextBox.Text = filename
    End Sub
    Public Sub Output_file_label(sender As Object, e As EventArgs) Handles Button8.Click
        Form3.Show()
    End Sub
    Private Sub RadioButton2_Text(sender As Object, e As EventArgs) Handles RadioButton13.Click, RadioButton11.Click, CheckBox2.Click
        If RadioButton13.Checked = True Then
            TargetComboBox.Text = PathComboBox.Text

        ElseIf RadioButton11.Checked = True Then

            If PathComboBox.Text.Length = 0 Then
                Using fbd As New FolderBrowserDialog(Me)
                    If fbd.ShowDialog = DialogResult.OK Then
                        MsgBox(fbd.DirectoryPath)
                    End If
                End Using
            End If
            TargetComboBox.Text = PathComboBox.Text
        End If
        If CheckBox2.Checked = True Then
            TargetComboBox.Text += Path.DirectorySeparatorChar & System.DateTime.Now.ToString("yyyy_MM_dd")
        End If
    End Sub
    Private Sub Find_Episode(sender As Object, e As EventArgs)
        If FileCollection.SelectedItem Is Nothing Then
            Return
        End If
        Dim rgx As New Regex("(.*\W*)\W(\d{1,2})(\W.*)")
        Dim ext As String = Path.GetExtension(FileCollection.SelectedItem)
        Dim match_result As Match = rgx.Match(FileCollection.SelectedItem)
        Dim episode As Integer = Integer.Parse(match_result.Groups(2).Value)

        Dim btn_name As String = CType(sender, ToolStripMenuItem).Text
        Console.WriteLine(match_result.Groups(1).Value & episode.ToString & match_result.Groups(3).Value)
        If btn_name = "下一集" Then
            episode += 1
        ElseIf btn_name = "上一集" Then
            episode -= 1
        End If
        For Each files In Directory.GetFiles(Path.GetDirectoryName(FileCollection.SelectedItem))
            If Path.GetExtension(files) <> ext Then Continue For
            If files.Contains(match_result.Groups(1).Value) And rgx.Match(files).Groups(2).Value = episode And files.Contains(match_result.Groups(3).Value) Then
                FileCollection.Items(FileCollection.SelectedIndex) = files
                Return
            End If
        Next
        For Each files In Directory.GetFiles(Path.GetDirectoryName(FileCollection.SelectedItem))
            If Path.GetExtension(files) <> ext Then Continue For
            If rgx.Match(files).Groups(2).Value = episode Then
                FileCollection.Items(FileCollection.SelectedIndex) = files
                Return
            End If
        Next
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles SearchTextBox.TextChanged
        Dim cal As String = "intersection"
        Dim keywords As String() = If(SearchTextBox.Text.Contains(" "), SearchTextBox.Text.Split(" "), {SearchTextBox.Text})
        Dim selectitem As String = If(String.IsNullOrEmpty(last_selectedItem), FileCollection.SelectedItem, last_selectedItem)
        'KeyPreview = If(String.IsNullOrEmpty(TextBox5.Text), False, True)
        If backup.Count = 0 Then backup = FileCollection.Items.Cast(Of String).ToList
        Dim new_array As List(Of String) = backup.ToList
        FileCollection.BeginUpdate()
        For j = 0 To keywords.Count - 1
            Dim keyword = keywords(j)
            If keyword.StartsWith("-") Then
                cal = "subsection"
                keyword = keyword.Remove(0, 1)
            End If
            If New Regex("(\d+)\*(\d+)").IsMatch(SearchTextBox.Text) Then cal = "pixel"
            If SearchTextBox.Text.Contains("\desktop") Then cal = "desktop"
            If SearchTextBox.Text.IndexOf("\same", StringComparison.OrdinalIgnoreCase) > 0 Then cal = "same"
            If SearchTextBox.Text = "()" Then cal = "repeat"
            Select Case cal
                Case "intersection"
                    For Each i In backup
                        If i.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) = -1 Then new_array.Remove(i)
                    Next
                Case "subsection"
                    For Each i In backup
                        If i.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) > 0 Then new_array.Remove(i)
                    Next
                Case "pixel"
                    For Each i In backup
                        If Not Ext_judge(Path.GetExtension(i).ToLower(), "Images") Then Continue For
                        Try
                            'Using image As Image = Image.FromFile(i)
                            '    If image.Width.ToString & "*" & image.Height.ToString <> keyword Then new_array.Remove(i)
                            'End Using


                            '------------ImageDimensions2------------
                            Dim format As String = GetImageFormat(i)
                            Dim size As New Size
                            Select Case format
                                Case "JPEG"
                                    size = GetJpegDimensions(i)
                                Case "PNG"
                                    size = GetPngDimensions(i)
                            End Select
                            If size.Width.ToString & "*" & size.Height.ToString <> keyword Then new_array.Remove(i)
                        Catch ex As Exception
                            Console.WriteLine(ex.Message)
                        End Try
                    Next
                Case "desktop"
                    For Each i In backup
                        If Not Ext_judge(Path.GetExtension(i).ToLower(), "Images") Then Continue For
                        'horizontally 橫向
                        If Not Image_horizontally(i) Then
                            new_array.Remove(i)
                        End If
                    Next
                Case "same"
                    ' 第一步：計算所有檔案的 hash，依大小先分組加速
                    Dim sizeGroups As New Dictionary(Of Long, List(Of String))
                    For Each i In backup
                        If Not File.Exists(i) Then Continue For
                        Dim size As Long = New FileInfo(i).Length
                        If Not sizeGroups.ContainsKey(size) Then
                            sizeGroups(size) = New List(Of String)
                        End If
                        sizeGroups(size).Add(i)
                    Next

                    ' 第二步：只對大小相同的檔案算 hash，收集有重複的路徑
                    Dim dupeSet As New HashSet(Of String)
                    For Each kvp In sizeGroups
                        If kvp.Value.Count < 2 Then Continue For
                        Dim hashGroups As New Dictionary(Of String, List(Of String))
                        For Each path In kvp.Value
                            Dim h As String = GetFileHash(path)
                            If Not hashGroups.ContainsKey(h) Then
                                hashGroups(h) = New List(Of String)
                            End If
                            hashGroups(h).Add(path)
                        Next
                        For Each hkvp In hashGroups
                            If hkvp.Value.Count > 1 Then
                                For Each path In hkvp.Value
                                    dupeSet.Add(path)
                                Next
                            End If
                        Next
                    Next

                    ' 第三步：不在重複集合裡的就從 new_array 移除
                    For Each i In backup
                        If Not dupeSet.Contains(i) Then new_array.Remove(i)
                    Next
                Case "repeat"
                    Dim rgx As New Regex("\(\d+\)")
                    For Each i In backup
                        If Not rgx.IsMatch(i) Then new_array.Remove(i)
                    Next
            End Select
        Next
        FileCollection.Items.Clear()
        For Each i In new_array
            FileCollection.Items.Add(i)
        Next
        FileCollection.EndUpdate()
        If Not String.IsNullOrEmpty(last_selectedItem) Then
            If FileCollection.Items.Contains(selectitem) Then
                Dim index = FileCollection.Items.IndexOf(selectitem)
                FileCollection.SetSelected(index, True)
            End If
        End If

        Refresh_listbox_numbers()
    End Sub
    ' 計算單一檔案的 MD5
    Function GetFileHash(filePath As String) As String
        Using md5 As MD5 = MD5.Create()
            Using stream As FileStream = File.OpenRead(filePath)
                Dim hashBytes() As Byte = md5.ComputeHash(stream)
                Return BitConverter.ToString(hashBytes).Replace("-", "").ToLower()
            End Using
        End Using
    End Function
    Function GetImageFormat(filePath As String) As String
        Dim pngSignature As Byte() = {&H89, &H50, &H4E, &H47, &HD, &HA, &H1A, &HA}
        Dim jpgSignature As Byte() = {&HFF, &HD8}

        Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read)
            Using br As New BinaryReader(fs)
                Dim header As Byte() = br.ReadBytes(8) ' 讀取前 8 個字節

                ' 檢查 PNG 簽名
                If header.Take(8).SequenceEqual(pngSignature) Then
                    Return "PNG"
                End If

                ' 檢查 JPEG 簽名 (只需要前 2 個字節)
                If header.Take(2).SequenceEqual(jpgSignature) Then
                    Return "JPEG"
                End If
            End Using
        End Using

        Return "未知格式"
    End Function
    Function GetPngDimensions(filePath As String) As Size
        ' PNG 文件以 8 個字節的簽名開始: 89 50 4E 47 0D 0A 1A 0A
        Dim pngSignature As Byte() = {&H89, &H50, &H4E, &H47, &HD, &HA, &H1A, &HA}

        Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read)
            Using br As New BinaryReader(fs)
                ' 檢查簽名
                Dim signature As Byte() = br.ReadBytes(8)
                If Not signature.SequenceEqual(pngSignature) Then
                    Return System.Drawing.Size.Empty ' 文件不是有效的 PNG
                End If

                ' 跳過接下來的 4 個字節 (IHDR 塊長度) 和 "IHDR" 字節
                br.ReadBytes(4 + 4)

                ' 讀取寬度和高度
                Dim width As Integer = ReadBigEndianInt32(br)
                Dim height As Integer = ReadBigEndianInt32(br)

                Return New Size(width, height)
            End Using
        End Using
    End Function

    Function ReadBigEndianInt32(br As BinaryReader) As Integer
        Dim bytes As Byte() = br.ReadBytes(4)
        Array.Reverse(bytes) ' 將字節順序轉換為大端
        Return BitConverter.ToInt32(bytes, 0)
    End Function
    Function GetJpegDimensions(filePath As String) As Size
        Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read)
            Using br As New BinaryReader(fs)
                br.ReadBytes(2) ' 跳過文件起始字節
                Do
                    Dim marker As UShort = ReadBigEndianUInt16(br)
                    Dim blockLength As UShort = ReadBigEndianUInt16(br)

                    ' Start of Frame marker in JPEG 是 0xFFC0 或 0xFFC2
                    If marker = &HFFC0 Or marker = &HFFC2 Then
                        br.ReadByte() ' 跳過精度字節
                        Dim height As UShort = ReadBigEndianUInt16(br)
                        Dim width As UShort = ReadBigEndianUInt16(br)
                        Return New Size(width, height)
                    End If
                    br.BaseStream.Seek(blockLength - 2, SeekOrigin.Current) ' 跳過當前區塊
                Loop
            End Using
        End Using
        Return Size.Empty
    End Function

    Function ReadBigEndianUInt16(br As BinaryReader) As UShort
        Dim bytes = br.ReadBytes(2)
        Array.Reverse(bytes)
        Return BitConverter.ToUInt16(bytes, 0)
    End Function
    Public Function Image_horizontally(imageFile As String)
        Try
            'Using image As Image = Image.FromFile(imageFile)
            '    If image.Width < image.Height Then Return False
            '    Return True
            'End Using
            '------------ImageDimensions2------------
            Dim format As String = GetImageFormat(imageFile)
            Dim size As New Size
            Select Case format
                Case "JPEG"
                    size = GetJpegDimensions(imageFile)
                Case "PNG"
                    size = GetPngDimensions(imageFile)
            End Select
            If size.Width < size.Height Then Return False
            Return True

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Return False
        End Try
    End Function

    Private Sub CheckedListBox1_ItemCheck(ByVal sender As Object, ByVal e As ItemCheckEventArgs) Handles OptionCheckedListBox.ItemCheck
        Me.BeginInvoke(CType(Function()
                                 OptionBox_CheckedChanged(sender, e)
                                 Return True
                             End Function, MethodInvoker))
    End Sub

    Private Sub OptionBox_CheckedChanged(sender As Object, e As EventArgs)
        Dim obj As ListBox = Nothing
        Dim new_array As New List(Of String)
        Dim shouldTopMost As Boolean = CheckedListBox1_isCheck("固定在最上層")
        If Me.TopMost <> shouldTopMost Then
            Me.TopMost = shouldTopMost
        End If
        Panel9.Visible = If(CheckedListBox1_isCheck("隱藏壓縮檔案集合"), False, True)
        If CheckedListBox1_isCheck("隱藏預覽") Then
            Button3.Visible = True
            Panel13.Visible = False
            Width = 800
        Else
            Button3.Visible = False
            Panel13.Visible = True
            Width = 1412
        End If
        If CheckedListBox1_isCheck("紀錄關閉視窗") Then
            Call Task.Run(Sub() StartRecordClosePipe())
        End If
        If CheckedListBox1_isCheck("獲取標題") Then
            Call Form2.Get_web_title()
        Else
            Try
                FileNameComboBox.Text = Path.GetFileNameWithoutExtension(FileCollection.SelectedItems(FileCollection.SelectedItems.Count - 1).ToString)
            Catch
            End Try
        End If
        Form1_scale()
        'KeyPreview = If(CheckedListBox1_isCheck("KeyPreview"), True, False)
        ' Form2.KeyPreview = If(CheckedListBox1_isCheck("KeyPreview"), True, False)
        For Each checkeditem In OptionCheckedListBox.CheckedItems
            Select Case checkeditem.ToString
                Case "取樣"
                    If CheckBox11.Checked = True Then
                        obj = SubFileCollection
                    Else
                        obj = FileCollection
                    End If

                    For i = 0 To obj.Items.Count - 2
                        Dim item = Path.GetFileNameWithoutExtension(obj.Items(i).ToString)
                        Dim rgx As New Regex(item.Substring(0, item.Length - 1) & ".")
                        Dim item2 = Path.GetFileNameWithoutExtension(obj.Items(i + 1).ToString)
                        If rgx.IsMatch(item2) Then
                            new_array.Add(obj.Items(i + 1))
                        End If
                    Next
                    For Each i As String In new_array
                        obj.Items.Remove(i)
                    Next

                '檔名去ext
                '如果檔名去掉最後一位字符後，匹配成功
                '去掉數字較大的那一份檔案
                Case "Combine trf files"
                    CheckedListBox1_setCheck("Auto Save", False)

                    'Panel14.Size = New System.Drawing.Size(Panel14.Size.Width, Panel1.Size.Height * 2)

            End Select
        Next

    End Sub
    'Dim previous_combobox3_text As String = String.Empty
    Private Async Sub PathComboBox_TextChanged(sender As Object, e As EventArgs) Handles PathComboBox.TextChanged
        ExtList.Items.Clear()
        If dataHasRead Or savingTrf Or isReadingData Then Return
        If Directory.Exists(PathComboBox.Text) Then ReadDirectory()

        Dim ext As String = Path.GetExtension(PathComboBox.Text)

        If Ext_judge(ext, "文字") Then
            If File.Exists(PathComboBox.Text) Then Trf_backup(PathComboBox.Text)
            'OpenFileDialog1.FileName = PathComboBox.Text
            Await Read_text_data(PathComboBox.Text, FileCollection)
            If PathComboBox.Text.EndsWith(".trf") Then
                delete_repeat_item()
                Dim directory = Path.GetDirectoryName(PathComboBox.Text)
                '每次儲存都會重新讀取
                ' get_file_watch(directory)
            End If
            Refresh_backup()
            dataHasRead = True
        End If
        Me.Text = Path.GetFileName(PathComboBox.Text)
        If Not isCodeChange Then
            Add_new_path(PathComboBox, "ori_pathrecord.txt")
            Add_new_path(TargetComboBox, "des_pathrecord.txt")
        End If
        Dim pipeName As String = "easyPreview " & PathComboBox.Text
        Console.WriteLine(pipeName)
        If File.Exists(PathComboBox.Text) Then
            Call Task.Run(Sub() StartPipeServer(pipeName))
        End If

    End Sub
    Private Sub ReadDirectory()
        ExtList.Items.Clear()
        Dim files As List(Of String) = Getallfiles(PathComboBox.Text, "")

        If CheckBox10.Checked = True Then
            Dim extensionSet As New HashSet(Of String)
            For Each foundFile As String In files
                Dim ext As String = Path.GetExtension(foundFile)
                If String.IsNullOrWhiteSpace(ext) Then Continue For
                extensionSet.Add(ext)
            Next
            Dim extSetArray As String() = extensionSet.ToArray()
            Array.Sort(extSetArray, AddressOf CompareNatural)
            ExtList.Items.AddRange(extSetArray)
        Else
            ExtList.Items.Add("Images")
            ExtList.Items.Add("Media")
            ExtList.Items.Add("壓縮檔")
            ExtList.Items.Add("光盤")
            ExtList.Items.Add("遊戲")
            ExtList.Items.Add("文件")
            ExtList.Items.Add(".torrent")
            ExtList.Items.Add(".ovpn")
            If PathComboBox.Text.Contains("音樂") Then
                ExtList.SetItemChecked(1, True)
                'ListBox_Click(sender, e)
            ElseIf PathComboBox.Text.Contains("H漫") Then
                If Not IsBottomDirectory(PathComboBox.Text) Then CheckBox11.Checked = True
                SubDirCheckBox.Checked = True
                'extList.SetItemChecked(0, True)
                'ListBox_Click(sender, e)

            End If
        End If
    End Sub
    ' Using the Dir function


    ' Using the Directory.EnumerateDirectories and Directory.EnumerateFiles methods
    Public Function IsBottomDirectory(path As String) As Boolean
        For Each subDir In Directory.EnumerateDirectories(path)
            ' Check if there are any files in the subdirectory
            If Directory.EnumerateFiles(subDir).Any() Then
                ' The directory is not the bottom
                Return False
            End If
        Next
        ' The directory is the bottom
        Return True
    End Function

    'Public Function load_passward()
    '    Dim path1 = recordPath & "passward.txt"
    '    If File.Exists(path1) AndAlso PasswordBox.Items.Count = 0 Then
    '        Using sr As New StreamReader(File.OpenRead(path1))

    '            While (sr.Peek() >= 0)
    '                Try
    '                    PasswordBox.Items.Add(sr.ReadLine())
    '                Catch ex As Exception
    '                    Console.WriteLine(ex.Message)
    '                End Try
    '            End While
    '        End Using
    '    End If
    '    Return True
    'End Function
    Public Sub Load_past_path()
        Dim ori_path = recordPath & "ori_pathrecord.txt"
        Dim des_path = recordPath & "des_pathrecord.txt"
        Try
            ' 一次性讀取所有行，減少 I/O 操作
            Dim allLines As String() = File.ReadAllLines(ori_path)

            ' 將所有行添加到 ComboBox 中
            PathComboBox.Items.AddRange(allLines)

        Catch ex As Exception
            ' 捕獲並處理異常
            Console.WriteLine(ex.Message)
        End Try
        Try
            ' 一次性讀取所有行，減少 I/O 操作
            Dim allLines As String() = File.ReadAllLines(des_path)

            ' 將所有行添加到 ComboBox 中
            TargetComboBox.Items.AddRange(allLines)

        Catch ex As Exception
            ' 捕獲並處理異常
            Console.WriteLine(ex.Message)
        End Try
    End Sub
    Public Function SearchAllSubDirectories() As Boolean
        Return SubDirCheckBox.Checked
    End Function
    Private Function GetAllDirectories(path As String) As IEnumerable(Of String)

        Return Directory.EnumerateDirectories(path).Union(Directory.EnumerateDirectories(path).SelectMany(Function(d)

                                                                                                              Try
                                                                                                                  Return FileSystem.GetDirectories(path)
                                                                                                              Catch e As UnauthorizedAccessException
                                                                                                                  Return Enumerable.Empty(Of String)()
                                                                                                              End Try
                                                                                                          End Function))
    End Function
    Dim filereadtimes As Integer = 0
    Public Function Getallfiles(path As String, keyword As String) As List(Of String)
        Try
            If SearchAllSubDirectories() Then
                Return EverythingSearcher.SearchFiles($"""{path}""", $" file:{keyword}")
            Else
                Return EverythingSearcher.SearchFiles($"parent:""{path}"" ", $"file:{keyword}")
            End If
        Catch e As UnauthorizedAccessException
            Return Enumerable.Empty(Of String)()
        End Try
    End Function
    Private Function GetallfilesExt(path As String, exts As String()) As List(Of String)
        ' 使用 String.Join 來組合字串，避免手動迴圈拼接
        Dim keyword As String = String.Join("|*", exts)
        Return EverythingSearcher.SearchFiles($"parent:""{path}"" ", $" file:{keyword}")
    End Function
    Private Sub 副檔名列表_ItemCheck(ByVal sender As Object, ByVal e As ItemCheckEventArgs) Handles ExtList.ItemCheck
        Me.BeginInvoke(CType(Function()
                                 'If isCodeChange = True Then Return True
                                 ListBox_Click(sender, e)
                                 Return True
                             End Function, MethodInvoker))
    End Sub

    Private Sub ListBox_Click(sender As Object, e As EventArgs)
        If String.IsNullOrEmpty(PathComboBox.Text) Then Return
        Dim downloadPath As String = Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "{374DE290-123F-4565-9164-39C4925E467B}", Nothing)

        ' 如果沒有設定，則使用預設的下載目錄

        ' 取得使用者設定的下載目錄
        If Not downloadPath Then
            downloadPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\Downloads"
        End If
        Dim record As New List(Of Integer)(2)
        If record.Count < 2 Then
            record.Insert(0, 1)
            record.Insert(0, 1)
        End If
        record(0) = FileCollection.SelectedIndex
        record(1) = SubFileCollection.SelectedIndex
        FileCollection.BeginUpdate()
        FileCollection.Items.Clear()
        SubFileCollection.Items.Clear()
        If CheckBox11.Checked Then
            Dim newTable As List(Of String) = ExtList.CheckedItems.Cast(Of String).ToList
            'Console.WriteLine(newTable)
            If newTable.Count = 0 Then
                Dim Directories As IEnumerable(Of String) = GetAllDirectories(PathComboBox.Text)
                FileCollection.Items.AddRange(Directories.ToArray)
            End If
            If CheckBox10.Checked = False Then
                For Each cont As Object In newTable
                    Dim Directories As IEnumerable = GetAllDirectories(PathComboBox.Text)

                    For Each foundDirectory As String In Directories
                        Dim all_files_count As Integer = Getallfiles(foundDirectory, "").Count
                        Dim target_count As Integer = GetallfilesExt(foundDirectory, Ext_table(cont).ToArray).Count
                        Console.WriteLine(target_count & "/" & all_files_count)
                        If target_count > all_files_count / 3 Then
                            FileCollection.Items.Add(foundDirectory)
                            'If Path.GetDirectoryName(foundDirectory) <> downloadPath Then 目錄檔案集合.Items.Add(foundDirectory)
                        ElseIf cont = "Media" Or cont = "games" Then
                            FileCollection.Items.Add(foundDirectory)
                            'If Path.GetDirectoryName(foundDirectory) <> downloadPath Then 目錄檔案集合.Items.Add(foundDirectory)
                        End If
                    Next
                Next
            Else
                For Each cont As Object In newTable
                    Dim Directories As IEnumerable = GetAllDirectories(PathComboBox.Text)
                    For Each foundDirectory As String In Directories
                        Dim all_files_count As Short = Getallfiles(foundDirectory, "").Count
                        Dim target_count As Short = Getallfiles(foundDirectory, cont).Count
                        If target_count > all_files_count - target_count Then
                            FileCollection.Items.Add(foundDirectory)
                        End If
                    Next
                Next
            End If

        Else
            For Each cont As Object In ExtList.CheckedItems
                If CheckBox10.Checked Then
                    '搜尋ext
                    'Dim searchCondition = My.Computer.FileSystem.GetFiles(
                    ' ComboBox3.Text, searchOptionCheck(), "*" & cont)
                    Dim searchCondition As String() = Getallfiles(PathComboBox.Text, "*" & cont).ToArray

                    FileCollection.Items.AddRange(searchCondition)
                Else
                    Dim files As String() = Getallfiles(PathComboBox.Text, "").ToArray
                    Array.Sort(files, AddressOf CompareNatural)
                    Dim passFiles As String() = Array.FindAll(Of String)(files, Function(s) Ext_judge(Path.GetExtension(s).ToLower(), cont))
                    FileCollection.Items.AddRange(passFiles)
                End If
            Next
        End If
        FileCollection.EndUpdate()
        If record.Count > 1 Then
            If record(0) <> -1 And record(0) < FileCollection.Items.Count Then FileCollection.SelectedIndex = record(0)
            If record(1) <> -1 And record(1) < SubFileCollection.Items.Count Then SubFileCollection.SelectedIndex = record(1)
        End If
        If Not String.IsNullOrEmpty(SearchTextBox.Text) Then
            backup = FileCollection.Items.Cast(Of String).ToList  ' 強制更新
        End If
        Refresh_backup()
        If Not String.IsNullOrEmpty(SearchTextBox.Text) Then TextBox5_TextChanged(sender, e)
    End Sub
    Dim lastSelectedIndex As Integer = -1
    Dim sublastSelectedIndex As Integer = -1
    Dim lastPosition As Long = 0
    Dim volume As Integer = 0
    Public Sub Refresh_backup()
        If savingTrf Then Return
        Refresh_listbox_numbers()

        'If PathComboBox.Text.EndsWith(".wmc") Then Return
        If PathComboBox.Text.Contains("backup") Then Return
        If Not String.IsNullOrEmpty(SearchTextBox.Text) Then Return
        If Directory.Exists(PathComboBox.Text) Then Return
        If Not FileCollection.Items.Cast(Of String).SequenceEqual(backup) Then
            backupUndo = backup
            backup = FileCollection.Items.Cast(Of String).ToList
        End If
        If FileCollection.Items.Count = 0 Then Return
        If File.Exists(PathComboBox.Text) Then Trf_backup(PathComboBox.Text)
        If Not CheckedListBox1_isCheck("Auto Save") Then Return
        Form3.Form3_Load(New Object, New EventArgs)
        Form3.Button1_Click(New Object, New EventArgs)

    End Sub
    Public Sub SaveASMRPlayRecord(taritem As String)
        'If isopening Then Return
        If CheckedListBox1_isCheck("鎖定") Then Return
        If Not CheckedListBox1_isCheck("ASMR播放模式") Then Return
        taritem = Remove_label(taritem)
        'If taritem.StartsWith("★") Then taritem.Substring(1)
        Dim record As String = taritem & "/playRecord.apr"
        Dim lines As New List(Of String)

        If videoPanel.Visible And _mediaPlayer.Media IsNot Nothing Then
            If isclosing Then
                Dim mediaPath As String = New Uri(_mediaPlayer.Media.Mrl).LocalPath
                Dim mediaIndex As Integer = SubFileCollection.Items.IndexOf(mediaPath)
                lines.Add("[subselected]" & mediaIndex.ToString)
            Else
                lines.Add("[subselected]" & sub_lastSelectedMedia.ToString)
            End If
            lines.Add("[playTime]" & _mediaPlayer.Time.ToString)
            lines.Add("[volume]" & _mediaPlayer.Volume.ToString)
        Else
            Return
        End If
        Try
            Using sw As New StreamWriter(record, False)
                sw.WriteLine(String.Join(Environment.NewLine, lines))
            End Using
        Catch ex As Exception
            Console.WriteLine(record & "輸出失敗" & ex.Message)
        End Try
    End Sub
    Public Sub WmcRefresh()
        Refresh_listbox_numbers()
        'If Not subFileCollection.Items.Cast(Of String).ToList Is backup Then
        '    backupUndo = backup
        '    backup = subFileCollection.Items.Cast(Of String).ToList
        'End If
        'If CheckedListBox1_isCheck("Auto Save") Then
        '    Form3.RadioButton15.Checked = True
        '    Form3.Form3_Load(New Object, New EventArgs)
        '    Form3.Button1_Click(New Object, New EventArgs)
        'End If
        For Each webOrder As String In FileCollection.Items
            If Ext_judge(webOrder, "網頁") Then

                Call Form2.WebMode(webOrder)

            ElseIf webOrder.Contains("crawler") Then
                Dim keyword As String = webOrder.Split("]")(1)
                Call Form2.WaitForLoad()
                Form2.WmcSearch(keyword)
            ElseIf webOrder.Contains("clickByClass") Then
                Dim elementClass As String = webOrder.Split("]")(1)
                Call Form2.WaitForLoad()

            End If

        Next

    End Sub
    Public Sub Refresh_listbox_numbers()
        Dim label13Text As New StringBuilder("")
        label13Text.Append(FileCollection.SelectedIndex + 1 & Path.AltDirectorySeparatorChar & FileCollection.Items.Count.ToString & "項")
        Label13.Text = label13Text.ToString
        Dim label17Text As New StringBuilder("")
        label17Text.Append(SubFileCollection.SelectedIndex + 1 & Path.AltDirectorySeparatorChar & SubFileCollection.Items.Count.ToString & "項")
        Label17.Text = label17Text.ToString
    End Sub
    Private Sub Ext_select(sender As Object, e As EventArgs) Handles CheckBox6.Click
        For sel = 0 To ExtList.Items.Count - 1
            If CheckBox6.Checked = True Then
                ExtList.SetItemChecked(sel, True)
            Else
                ExtList.SetItemChecked(sel, False)
            End If
        Next
    End Sub
    Private Function Ext_table(ext_type As String) As IEnumerable(Of String)
        Dim ext_explain As IEnumerable(Of String)
        If ext_type = "Images" Then
            ext_explain = {".jpg", ".png", ".jfif", ".gif", ".jpeg"}.Concat(Ext_table("其他圖片"))
        ElseIf ext_type = "其他圖片" Then
            ext_explain = {".clip", ".psd", ".webp"}
        ElseIf ext_type = "ASMR" Then
            ext_explain = Ext_table("Images").Concat(Ext_table("Media"))
        ElseIf ext_type = "Media" Then
            ext_explain = Ext_table("Videos").Concat(Ext_table("Sounds"))
        ElseIf ext_type = "Videos" Then
            ext_explain = {".mp4", ".mkv", ".wmv", ".flac"}
        ElseIf ext_type = "Sounds" Then
            ext_explain = {".mp3", ".m4a", ".mid", ".wav", ".ogg"}
        ElseIf ext_type = "壓縮檔" Then
            ext_explain = {".rar", ".7z", ".zip"}
        ElseIf ext_type = "光盤" Then
            ext_explain = {".mdf", ".iso", ".mds"}
        ElseIf ext_type = "遊戲" Then
            ext_explain = {".exe"}
        ElseIf ext_type = "網頁" Then
            ext_explain = {"http"}
        ElseIf ext_type = "文件" Then
            ext_explain = {".pdf"}.Concat(Ext_table("文字")).Concat(Ext_table("簡報")).Concat(Ext_table("word檔"))
        ElseIf ext_type = "文字" Then
            ext_explain = {".txt", ".ini", ".inf", ".trf", ".apr", ".py", ".vb", ".js", ".ovpn"}.Concat(Ext_table("字幕")).Concat(Ext_table("easyRecord"))
        ElseIf ext_type = "easyRecord" Then
            ext_explain = {".trf", ".wmc"}
        ElseIf ext_type = "字幕" Then
            ext_explain = {".ass", ".srt"}
        ElseIf ext_type = "簡報" Then
            ext_explain = {".xls", ".xlsx"}
        ElseIf ext_type = "word檔" Then
            ext_explain = {".doc", ".docx"}
        Else
            ext_explain = {ext_type}
        End If
        Return ext_explain
    End Function
    Public Function Ext_judge(ext As String, ext_type As String)
        If ext Is Nothing Then Return False
        Dim reverse = False
        If ext.StartsWith("!") Then
            reverse = True
            ext = ext.Remove(0, 1)
        End If
        Dim result = False
        If ext IsNot Nothing Then ext = ext.ToLower
        '容許.url .txt檔
        If Not {"網頁", "目錄", "程序"}.Contains(ext_type) Then
            Dim table As IEnumerable = Ext_table(ext_type)
            result = table.Cast(Of String).Contains(ext)
        ElseIf ext_type = "網頁" And Not File.Exists(ext) Then
            If ext IsNot Nothing Then result = ext.Contains("http")
        ElseIf ext_type = "目錄" Then
            If Directory.Exists(ext) Then result = True
        ElseIf ext_type = "程序" Then
            If RadioButton15.Checked Then result = True
        End If
        Return If(reverse, Not result, result)
    End Function
    Public Sub Next_Page(sender As Object, e As EventArgs)
        If targetList Is Nothing Then Return
        If targetList.SelectedIndex = -1 Then Return
        If targetList.SelectedItem.contains(".pdf") Then
            Next_Page_PDF(sender, e)
            Return
        End If
        Dim btn_name = sender.text
        Dim tarindex As Integer = 0
        If btn_name = "next" Then
            If Not targetList.SelectedIndex = targetList.Items.Count - 1 Then
                tarindex = targetList.SelectedIndex + 1
            Else
                Return
            End If
        ElseIf btn_name = "previous" Then
            If targetList.SelectedIndex <> 0 Then
                tarindex = targetList.SelectedIndex - 1
            Else
                Return
            End If
        End If
        targetList.SetSelected(targetList.SelectedIndex, False)
        targetList.SetSelected(tarindex, True)
    End Sub
    Private isButtonPressed As Boolean = False
    Public Async Sub Quick_page(sender As Object, e As EventArgs) Handles nextButton.MouseDown, previousButton.MouseDown
        isButtonPressed = True
        While isButtonPressed
            Await Task.Delay(500) ' 等待 500 毫秒（非同步，不阻塞 UI）
            Next_Page(sender, e)

        End While
    End Sub
    ' 鬆開按鈕時觸發
    Private Sub Button1_MouseUp(sender As Object, e As MouseEventArgs) Handles nextButton.MouseUp, previousButton.MouseUp
        isButtonPressed = False        ' 停止計時器
    End Sub
    Public Sub Next_Page_PDF(sender As Object, e As EventArgs)
        Dim btn_name = sender.text
        Dim rgx As New Regex("頁數：(\d+)/")
        If btn_name = "next" Then
            Dim now_page As Short = rgx.Match(Label10.Text).Groups(1).Value
            If doc.Pages.Count <= now_page Then Return
            now_page += 1
            Dim bmp As Image = doc.SaveAsImage(now_page - 1)
            Label10.Text = "頁數：" & now_page.ToString & Path.AltDirectorySeparatorChar & doc.Pages.Count.ToString
            PictureBox1.Image = CType(bmp, Image)
        ElseIf btn_name = "previous" Then
            Dim now_page As Short = rgx.Match(Label10.Text).Groups(1).Value
            If now_page <= 1 Then Return
            now_page -= 1
            Dim bmp As Image = doc.SaveAsImage(now_page - 1)
            Label10.Text = "頁數：" & now_page.ToString & Path.AltDirectorySeparatorChar & doc.Pages.Count.ToString
            PictureBox1.Image = CType(bmp, Image)
        End If
    End Sub

    Public Sub Next_item()
        If targetList.SelectedIndex + 1 >= targetList.Items.Count Then Return
        targetList.SelectedIndex += 1
        targetList.SetSelected(targetList.SelectedIndex, False)
    End Sub
    Private Sub Add_new_password(sender As Object, e As EventArgs) Handles Button5.Click
        PasswordBox.Items.Add(TextBox4.Text)
        Dim record As String = recordPath & "passward.txt"
        If File.Exists(record) = False Then
            Dim fs As FileStream = File.Create(record)
            fs.Close()
        End If
        Using sr As New StreamWriter(record, True)
            sr.WriteLine(TextBox4.Text)
        End Using
        TextBox4.Text = ""
    End Sub
    Dim isCodeChange As Boolean = False
    Private Sub Add_new_path(target As ComboBox, saveName As String)

        Dim savePath As String = recordPath & saveName
        Dim saveItems As New List(Of String)
        Dim newItem = target.Text.ToString()
        Console.WriteLine(newItem)
        ' 如果新項目為空則返回
        If String.IsNullOrEmpty(newItem) Then Return

        ' 避免多次 I/O，僅讀取一次文件
        If System.IO.File.Exists(savePath) Then
            saveItems = System.IO.File.ReadAllLines(savePath, Encoding.UTF8).ToList()
        End If

        ' 檢查是否為 .trf 檔案類型
        'If Ext_judge(saveName, ".trf") Then Return

        ' 如果已存在，則移動到列表頂部
        If saveItems.Contains(newItem) Then
            saveItems.Remove(newItem)
        End If

        ' 控制最大項目數量
        Dim MaxItemCount = 30
        If saveItems.Count >= MaxItemCount Then
            saveItems.RemoveAt(saveItems.Count - 1) ' 移除最舊的項目（最後一個）
        End If

        ' 在列表開頭插入新項目
        saveItems.Insert(0, newItem)
        Console.WriteLine(saveItems(0))
        ' 更新 ComboBox 項目
        target.Items.Clear()
        target.Items.AddRange(saveItems.ToArray())

        ' 一次性寫入更新後的內容到文件
        System.IO.File.WriteAllLines(savePath, saveItems, Encoding.UTF8)

        ' 設置狀態並更新 Text
        isCodeChange = True
        target.Text = newItem
        isCodeChange = False
    End Sub
    Private Sub Add_Record(target As String)
        Dim savePath As String = recordPath & "add_record.txt"
        Dim saveItems As New List(Of String)
        Console.WriteLine(target)
        ' 如果新項目為空則返回
        If String.IsNullOrEmpty(target) Then Return

        ' 避免多次 I/O，僅讀取一次文件
        If System.IO.File.Exists(savePath) Then
            saveItems = System.IO.File.ReadAllLines(savePath, Encoding.UTF8).ToList()
        End If

        ' 如果已存在，則移動到列表頂部
        If saveItems.Contains(target) Then
            saveItems.Remove(target)
        End If

        ' 控制最大項目數量
        Dim MaxItemCount = 30
        If saveItems.Count >= MaxItemCount Then
            saveItems.RemoveAt(saveItems.Count - 1) ' 移除最舊的項目（最後一個）
        End If

        ' 在列表開頭插入新項目
        saveItems.Insert(0, target)
        Console.WriteLine(saveItems(0))

        ' 一次性寫入更新後的內容到文件
        System.IO.File.WriteAllLines(savePath, saveItems, Encoding.UTF8)
    End Sub

    Private Sub ViewAllFilesInIso(Directorie As Object)

        For Each CDentry As DiscUtils.DiscFileInfo In Directorie.GetFiles
            SubFileCollection.Items.Add(CDentry.FullName.Replace(";1", "").ToLower)
        Next
        For Each Directories In Directorie.GetDirectories
            ViewAllFilesInIso(Directories)
        Next
    End Sub
    Private Sub Delete_password(sender As Object, e As EventArgs) Handles Button7.Click
        PasswordBox.Items.Remove(PasswordBox.SelectedItem)

        Dim desk As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\Documents"
        Dim record As String = desk & "passward.txt"
        FileSystem.DeleteFile(record)
        If Not File.Exists(record) Then
            Dim fs As FileStream = File.Create(record)
            fs.Close()
        End If
        Using sr As New StreamWriter(record, True)
            For Each item In PasswordBox.Items
                sr.WriteLine(item)
            Next
        End Using


    End Sub



    Public Sub FindFirstBitmap(tarlist As ListBox)
        Dim data As String() = Array.FindAll(Of String)(tarlist.Items.Cast(Of String).ToArray, Function(x)
                                                                                                   Dim ext As String = Path.GetExtension(x.ToLower)
                                                                                                   Return Ext_judge(ext, "Images")
                                                                                               End Function)
        Dim findEffectiveKeyword As Boolean = False
        Dim selected As New List(Of String)
        For Each keyword In Constants.GetValue("asmrSubnailKeyword")
            selected.AddRange(Array.FindAll(Of String)(data, Function(x) x.IndexOf(keyword) > 0))
            Console.WriteLine(selected.Count)
        Next
        If data.Length > 0 Then
            If PathComboBox.Text.IndexOf("asmr", StringComparison.OrdinalIgnoreCase) <= 0 Then
                SubFileCollection.SelectedItem = data(0)
            ElseIf selected.Count > 0 Then
                SubFileCollection.SelectedItem = selected(0)
            Else
                Dim horizontallyImage As IEnumerable(Of String) = data.Where(Function(x) Image_horizontally(x))
                If horizontallyImage.Count > 0 Then
                    SubFileCollection.SelectedItem = horizontallyImage(0)
                Else
                    SubFileCollection.SelectedItem = data(0)
                End If
            End If
        End If
        Transform_index_status_to_Form2()

    End Sub
    Public Async Function FindFirstSound(curitem As String) As Task
        Dim recordfile As String = curitem & "/playRecord.apr"
        If File.Exists(recordfile) And newMediaList = False Then
            Await Read_text_data(recordfile, FileCollection, True)
            Return
            If SubFileCollection.SelectedIndex <> -1 Then Return
        End If
        Dim data As String() = find_all_asmr_sound(SubFileCollection.Items.Cast(Of String).ToArray)
        If data IsNot Nothing Then SubFileCollection.SelectedItem = data(0)
        newMediaList = False
    End Function
    Public Function SearchFileInFolder(ori_folder As String, name As String)
        Dim filesCollections As List(Of String) = Getallfiles(ori_folder, name)
        'Dim filesCollections As ObjectModel.ReadOnlyCollection(Of String) = FileSystem.GetFiles(ori_folder, FileIO.SearchOption.SearchAllSubDirectories, name)
        'Console.WriteLine(filesCollections.Count)
        Dim filesArray As String() = filesCollections.Cast(Of String).ToArray
        Array.Sort(filesArray, AddressOf CompareNatural)
        Array.Reverse(filesArray)
        If filesCollections.Count <> 0 Then name = filesArray(0)
        Return name
    End Function
    Public Function SearchFolderInFolder(ori_folder As String, name As String)
        Dim foldersCollections As String() = Directory.GetDirectories(ori_folder, $"*{name}*", System.IO.SearchOption.AllDirectories)
        Console.WriteLine(foldersCollections.Count)
        Array.Sort(foldersCollections, AddressOf CompareNatural)
        Array.Reverse(foldersCollections)
        If foldersCollections.Count <> 0 Then name = foldersCollections(0)
        Return name
    End Function
    Sub OutputHandler(sender As Object, e As DataReceivedEventArgs)
        If Not String.IsNullOrEmpty(e.Data) Then
            lineCount += 1

            ' Add the text to the collected output.
            output.Append(Environment.NewLine + "[" + lineCount.ToString() + "]: " + e.Data)
        End If
    End Sub
    'path_debug
    Public Function GetMovedFolder(curitem As String)
        'Console.WriteLine(CheckedListBox1_isCheck("關閉自動修正"))
        'Console.WriteLine(Not curitem.Contains(":\"))
        If CheckedListBox1_isCheck("關閉自動修正") Then Return curitem
        If SearchTextBox.Text.Length > 0 Then Return curitem
        If Not curitem.Contains(":\") Then Return curitem
        Dim desk As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        Dim name As String = Path.GetFileNameWithoutExtension(curitem) + ".*"
        Dim name_with_ext = Path.GetFileName(curitem)
        Dim ori_curitem As String = curitem
        Dim last_folder As String = ""
        Dim searchResult As New List(Of String)
        If Directory.Exists(last_folder) Then
            '-------附檔名更改的情況-----------
            curitem = SearchFileInFolder(last_folder, name)
            If File.Exists(curitem) Then Return curitem
            '--------重複檔案(1)的情況---------
            'Dim rgx As Regex = New Regex("(.*)\(\d{1}\)")
            'Dim fn As String = rgx.Match(curitem).Groups(1).Value
            'curitem = fn
            'name = Path.GetFileNameWithoutExtension(curitem) + ".*"
            'curitem = searchFolder(last_folder, name)
            'If File.Exists(curitem) Then Return curitem
            last_folder = Path.GetFileName(last_folder)
        Else
            Dim directoryNames As String() = curitem.Split(Path.DirectorySeparatorChar)
            Dim directoryName = directoryNames(directoryNames.Count - 2)
            Console.WriteLine("Directory name: " & directoryName)
            last_folder = directoryName
        End If

        If PathComboBox.Text.Contains("ASMR") Then
            Dim RJnumber As Match = dlsiteRgx.Match(curitem)
            Console.WriteLine(RJnumber.Value)
            If RJnumber.Success Then
                searchResult = EverythingSearcher.SearchFiles(RJnumber.Value)
                'curitem = searchFolderInFolder("E:\ASMR", RJnumber.Groups(1).Value)
                If searchResult.Count > 0 Then
                    SubFileCollection.Items.Add($"找到檔案{searchResult.Count}個")
                    For Each foundfile In searchResult
                        SubFileCollection.Items.Add(foundfile)
                    Next
                    curitem = searchResult(0)
                End If
                Return curitem
            End If
        ElseIf PathComboBox.Text.Contains("遊戲") Or curitem.Contains("Commonly_used") Then
            'last_folder = Path.GetDirectoryName(curitem)
            'curitem = searchFolderInFolder("F:\Commonly_used", last_folder) & Path.DirectorySeparatorChar & name_with_ext
            Dim searchtarget = last_folder & Path.DirectorySeparatorChar & name_with_ext
            Console.WriteLine(searchtarget)
            searchResult = EverythingSearcher.SearchFiles(name_with_ext)
            searchResult = searchResult.FindAll(Function(n) n.Contains(last_folder))
            If searchResult.Count > 0 Then
                SubFileCollection.Items.Add($"找到檔案{searchResult.Count}個")
                For Each foundfile In searchResult
                    SubFileCollection.Items.Add(foundfile)
                Next
                curitem = searchResult(0)
            End If
        ElseIf curitem.Contains("download") Then
            searchResult = EverythingSearcher.SearchFiles(last_folder)
            If searchResult.Count > 0 Then
                SubFileCollection.Items.Add($"找到檔案{searchResult.Count}個")
                For Each foundfile In searchResult
                    SubFileCollection.Items.Add(foundfile)
                Next
                curitem = searchResult(0)
            End If
        End If
        Console.WriteLine(curitem)
        If File.Exists(curitem) Or Directory.Exists(curitem) Then Return curitem
        Console.WriteLine("last_folder: " & last_folder)
        Console.WriteLine($"last_folder is desk = {last_folder = desk} ")

        If last_folder = "desktop" Then
            Dim directories As String() = Directory.EnumerateDirectories(desk).ToArray
            Array.Sort(directories, AddressOf CompareNatural)
            'Array.Reverse(directories)
            For Each folder In directories
                Console.WriteLine(folder)
                '改成可選擇搜尋到的結果
                Dim sb As New StringBuilder("")
                sb.Append(folder & Path.DirectorySeparatorChar & name_with_ext)
                If File.Exists(sb.ToString) Then
                    curitem = sb.ToString
                End If
            Next
            Console.WriteLine(curitem)
            If File.Exists(curitem) Then Return curitem
        End If
        If File.Exists(curitem) Or Directory.Exists(curitem) Then Return curitem
        'last_folder = Path.GetDirectoryName(curitem)
        'Console.WriteLine(last_folder + name)

        Dim rgx As New Regex("(20\d{2})")
        Dim result = rgx.Match(last_folder)
        Dim draw_path As String = ""
        If result.Success Then
            draw_path = $"E:\(畫\{result.Groups(1).Value}整合"
        End If

        '改成可選擇搜尋到的結果
        Dim tar_file As New StringBuilder("")
        tar_file.Append(draw_path & Path.DirectorySeparatorChar & last_folder & Path.DirectorySeparatorChar & name_with_ext)

        Console.WriteLine(tar_file.ToString)
        If File.Exists(tar_file.ToString) Then curitem = tar_file.ToString

        If File.Exists(curitem) Then Return curitem
        searchResult = EverythingSearcher.SearchFiles($"""{name}""", "")
        If searchResult.Count > 0 Then
            curitem = searchResult(0)
            If File.Exists(curitem) Then Return curitem
        End If
        Return ori_curitem
    End Function

    '這個委託永遠不會被調用
    Public Function ThumbnailCallback() As Boolean
        Return False
    End Function
    Dim last_selectedItem As String = Nothing
    Dim sub_last_selectedItem As String = Nothing
    Dim sub_lastSelectedMedia As Integer = 0
    Dim isListboxOperating As Boolean = False
    Private Sub FileDoubleClick(sender As Object, e As EventArgs) Handles FileCollection.DoubleClick
        File_Preview(False)
    End Sub
    Public playover As Boolean = False
    Private Function CheckFileSelected()
        targetList = FileCollection
        Dim fcsl As String = Remove_label(FileCollection.SelectedItem)
        If Remove_label(last_selectedItem) <> fcsl Then
            If playover = False Then
                SaveASMRPlayRecord(last_selectedItem)
            Else
                Dim rls As String = Remove_label(last_selectedItem & "/playRecord.apr")
                File.Delete(rls)
                playover = False
            End If
            last_selectedItem = FileCollection.SelectedItem
            sub_last_selectedItem = Nothing
        ElseIf isListboxOperating And Not PathComboBox.Text.Contains("ASMR") Then
        Else
            Return True
        End If
        Return False
    End Function
    Private isAutoClick As Boolean = False
    Private Function Sub_checkFileSelected()

        targetList = SubFileCollection
        If sub_last_selectedItem <> SubFileCollection.SelectedItem Then
            If isAutoClick = False Then
                SaveASMRPlayRecord(FileCollection.SelectedItem)
            End If
            sub_last_selectedItem = SubFileCollection.SelectedItem
            Dim ext As String = Path.GetExtension(SubFileCollection.SelectedItem)
            If Ext_judge(ext, "Media") Then
                sub_lastSelectedMedia = SubFileCollection.SelectedIndex
            End If
            Return False
        End If
        Return True
    End Function
    Public Sub FileIndexChange(sender As Object, e As EventArgs) Handles FileCollection.SelectedIndexChanged, FileCollection.Click
        File_Preview(True)
    End Sub

    Public Async Sub File_Preview(checkFileSelect As Boolean)
        If savingTrf Then Return

        If FileCollection.Items.Count = 0 Or FileCollection.SelectedIndex = -1 Then Return
        If checkFileSelect Then
            If CheckFileSelected() Then Return
        End If
        Refresh_backup()
        If CheckedListBox1_isCheck("鎖定") Then Return
        release_all_process()
        If CheckedListBox1_isCheck("Close Load File") Then Return
        'subFileCollection.Items.Clear()
        Dim sel_ind As Short = FileNameComboBox.SelectedIndex
        If FileNameComboBox.Items.Count <> 0 Then FileNameComboBox.Items.Clear()
        Dim curitem As String = FileCollection.SelectedItem
        'Dim curitem As String = FileCollection.SelectedItems(FileCollection.SelectedItems.Count - 1).ToString
        '去標籤處理
        curitem = Remove_label(curitem)
        Dim ext As String = ""
        Dim last_write As String = ""
        Dim createTime As String = ""
        Dim installs() As String = New String() {}
        If Not Ext_judge(curitem, "網頁") AndAlso Not Ext_judge(curitem, "目錄") Then
            If Not File.Exists(curitem) AndAlso FileCollection.Items.Count <> 0 Then
                Dim newName As String = GetMovedFolder(curitem)
                If newName = curitem Then
                    If curitem.Contains(":\") And isListboxOperating = False Then
                        MsgBox("找不到檔案")
                    End If
                    Return
                Else
                    FileCollection.Items(FileCollection.SelectedIndex) = newName
                    curitem = newName
                End If

            End If
            ext = Path.GetExtension(curitem).ToLower()
            FileNameComboBox.Text = Path.GetFileNameWithoutExtension(curitem)
            ExtInput.Text = ext
            last_write = File.GetLastWriteTime(curitem).ToString("yyyy_MM_dd-HHmmss")
            createTime = File.GetCreationTime(curitem).ToString("yyyy_MMdd")
            installs = New String() {FileNameComboBox.Text, last_write, "刪去第N個字元 string:n", "orderbox+index"}

        ElseIf Ext_judge(curitem, "目錄") Then

            Try
                FileNameComboBox.Text = Path.GetFileNameWithoutExtension(curitem)
                last_write = Directory.GetLastWriteTime(curitem).ToString("yyyy_MM_dd-HHmmss")
                createTime = Directory.GetCreationTime(curitem).ToString("yyyy_MMdd")
                installs = New String() {FileNameComboBox.Text, createTime, "Custom"}
            Catch
            End Try
        ElseIf Ext_judge(curitem, "網頁") Then
            Try
                FileNameComboBox.Text = Path.GetFileNameWithoutExtension(curitem)
                installs = New String() {FileNameComboBox.Text}
            Catch
            End Try
        End If

        FileNameComboBox.Items.AddRange(installs)
        FileNameComboBox.SelectedIndex = sel_ind
        If Ext_judge(curitem, "目錄") Then
            SubFileCollection.Items.Clear()
            If CheckedListBox1_isCheck("ASMR播放模式") Then
                '清除圖片選取
                SubDirCheckBox.Checked = True
                isAutoClick = True
                VideoView1.Hide()
                'PictureBox1.Dock = DockStyle.Top
                'PictureBox1.Height -= 57
                videoPanel.Dock = DockStyle.Bottom
                videoPanel.Height = 57
            End If
            Dim files As String() = Getallfiles(curitem, "").
                Where(Function(f) Not f.EndsWith(".apr")).
                ToArray()
            Array.Sort(files, AddressOf CompareNatural)
            SubFileCollection.Items.AddRange(files)
            FindFirstBitmap(SubFileCollection)
            GroupBox3.Visible = curitem.Contains("テグラユウキ")
            If CheckedListBox1_isCheck("ASMR播放模式") Then
                SubFileCollection.ClearSelected()
                Await FindFirstSound(curitem)
                isAutoClick = False
            End If
            FileNameComboBox.Text = Path.GetFileNameWithoutExtension(curitem)
        End If
        If CheckedListBox1_isCheck("Close Preview") Then Return
        If Ext_judge(ext, "程序") Then
            Dim setProcess As Process = allProcesses(FileCollection.SelectedIndex)
            Console.WriteLine(setProcess.StandardInput)

        ElseIf Ext_judge(ext, "其他圖片") Then
            specialPicture(curitem)
        ElseIf Ext_judge(ext, "Images") Then
            Dim MyImage As Bitmap = Nothing
            PictureBox1.Visible = True
            RichTextBox1.Hide()
            Label15.Show()

            If CheckedListBox1_isCheck("ASMR播放模式") Then
                CType(PictureBox1.Image, Bitmap).MakeTransparent(Color.White)

            End If
            Try
                Try
                    MyImage = New Bitmap(curitem)
                    Label15.Text = MyImage.Size.Width & "px " & MyImage.Size.Height & "px"
                    PictureBox1.Image = CType(MyImage.Clone, Image)
                    Transfer_image_toForm2()
                    MyImage.Dispose()
                Catch ex As Exception
                    Dim bmp = New Bitmap(PictureBox1.Width, PictureBox1.Height)
                    Dim g = Graphics.FromImage(bmp)
                    g.DrawString("檔案毀損", New Font("新細明體", 9), Brushes.Black, New PointF(10, 10))
                    PictureBox1.Image = bmp
                End Try
            Catch ex As Exception
                MsgBox("無法預覽檔案" & curitem & "無法開啟" & ex.GetType().FullName & ex.Message)
            End Try
        ElseIf ext = ".pdf" Then
            doc.LoadFromFile(curitem)
            '遍歷PDF每一頁
            Dim bmp As Image = doc.SaveAsImage(0)
            Try
                PictureBox1.Image = CType(bmp, Image)
                Label10.Text = "頁數：1/" & doc.Pages.Count.ToString
                If Form2.Visible = True Then Form2.PictureBox1.Image = PictureBox1.Image
            Catch ex As Exception
                MsgBox("無法預覽檔案" & curitem & "無法開啟" & ex.GetType().FullName & ex.Message)
            End Try
        ElseIf ext = ".xls" Or ext = ".xlsx" Then
            Me.DataGridView1.Show()
            Dim strConn As String = “Provider=Microsoft.ACE.OLEDB.12.0;” & “Data Source=” & curitem & “;” & “Extended Properties=Excel 12.0;”
            Dim conn As New OleDbConnection(strConn)
            conn.Open()
            Dim strExcel As String = “”
            Dim myCommand As OleDbDataAdapter
            Dim ds As DataSet
            Dim excelShema As DataTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
            Dim firstSheetName As String = excelShema.Rows(0)("TABLE_NAME").ToString()
            strExcel = String.Format("Select * from [{0}]", firstSheetName)
            myCommand = New OleDbDataAdapter(strExcel, strConn)
            ds = New DataSet()
            myCommand.Fill(ds, "111")
            DataGridView1.DataSource = ds.Tables(0)
            'ElseIf Ext_judge(ext, ".wav") Then
            '    My.Computer.Audio.Play(curitem)
        ElseIf Ext_judge(ext, "Media") Then

            PictureBox1.Visible = False
            Label15.Hide()
            ComboBox5.Show()
            DeviceComboBox.Show()
            videoPanel.Show()
            VideoView1.Visible = If(CheckedListBox1_isCheck("ASMR播放模式"), False, True)
            RichTextBox1.Hide()
            Dim sameMedia As Boolean = checkSameMedia(curitem)
            If sameMedia Then
                _mediaPlayer.Play()
                Button24.Text = "暫停"
                Return
            End If
            Console.WriteLine(lastPosition)

            'Task.Delay(200).Wait() ' 避免 VLC 线程冲突
            Me.BeginInvoke(New Action(Sub()
                                          Dim media As New Media(_libvlc, curitem, FromType.FromPath)
                                          ' 加載字幕
                                          _mediaPlayer.Media = media

                                          'addsubtitle(curitem)

                                          _mediaPlayer.Play()

                                          If lastPosition <> 0 Then

                                              Try
                                                  _mediaPlayer.Time = lastPosition
                                              Catch ex As Exception
                                                  Debug.WriteLine("Error setting media time: " & ex.Message)
                                              End Try


                                              lastPosition = 0
                                          End If

                                          AddHandler _mediaPlayer.TimeChanged, AddressOf OnTimeChanged
                                      End Sub))
            Dim volume As Integer = _mediaPlayer.Volume
            TrackBar1.Value = If(volume >= 0, volume, 0)
            ' 让音量更新在后台进行，避免 UI 线程卡死
            Button24.Text = "暫停"

        ElseIf Ext_judge(ext, "文字") Then
            PictureBox1.Hide()
            RichTextBox1.Show()
            Using fileReader As New StreamReader(curitem)
                Dim stringReader As String = fileReader.ReadToEnd
                RichTextBox1.Text = stringReader
            End Using
        ElseIf Ext_judge(ext, "光盤") Then
            SubFileCollection.Items.Clear()
            Using isoStream As New FileStream(curitem, FileMode.Open)
                Dim cd As New CDReader(isoStream, True)
                For Each CDentry As DiscUtils.DiscFileInfo In cd.Root.GetFiles()
                    SubFileCollection.Items.Add(CDentry.FullName.Replace(";1", "").ToLower)
                Next
                For Each Directories In cd.Root.GetDirectories
                    ViewAllFilesInIso(Directories)
                Next
            End Using
        ElseIf Ext_judge(ext, "壓縮檔") Then
            SubFileCollection.Items.Clear()
            Try
                archive = ArchiveFactory.OpenArchive(curitem)
                For Each entry As Object In archive.Entries
                    SubFileCollection.Items.Add(entry.key)
                Next
            Catch
                Dim passward = try_password(curitem, 0)
                archive = ArchiveFactory.OpenArchive(curitem, New SharpCompress.Readers.ReaderOptions With {.Password = passward})
                For Each entry As Object In archive.Entries
                    SubFileCollection.Items.Add(entry.key)
                Next
            End Try
            Task.WaitAll()
            FindFirstBitmap(SubFileCollection)
        ElseIf Ext_judge(curitem, "網頁") Then
            Dim rgx As New Regex("(http.*)")
            curitem = rgx.Match(curitem).Groups(1).Value
            Transform_index_status_to_Form2()
            Call Form2.WebMode(curitem)
            Form2.Show()
        ElseIf Ext_judge(curitem, ".url") Then
            Dim fileReader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(curitem)
            Dim stringReader As String = fileReader.ReadToEnd
            Dim rgx As New Regex("(http.*)")
            curitem = rgx.Match(stringReader).Groups(1).Value
            Transform_index_status_to_Form2()
            Call Form2.WebMode(curitem)
            Form2.Show()
        End If
    End Sub

    Private Sub SpecialPicture(curitem As String)
        If CheckedListBox1_isCheck("使用nconvert") Then
            ' 替換檔案名稱中的特殊字元，例如inputjpg
            Console.WriteLine(curitem)

            Dim fileName = Make_Character_Valid(curitem, "png")
            ' 組合檔案名稱和副檔名，例如inputjpg.png
            Dim preImagePath = recordPath & "clip圖像\" & fileName
            'Shell("cmd.exe /c nconvert -out png -o output.png " & curitem, AppWinStyle.NormalFocus, True, 1000)
            Console.WriteLine(preImagePath)
            If File.Exists(preImagePath) Then
                Dim preview_image As Image = Image.FromFile(preImagePath)
                '將 image 物件指派給 PictureBox 控制項的 Image 屬性
                PictureBox1.Image = CType(preview_image.Clone, Image)
                Label15.Text = preview_image.Size.Width & "px " & preview_image.Size.Height & "px"
                '釋放 image 物件的資源
                preview_image.Dispose()
                If Form2.Visible = True Then Form2.PictureBox1.Image = PictureBox1.Image
                Transform_index_status_to_Form2()
            Else
                Dim psi As New ProcessStartInfo With {
                    .FileName = "cmd.exe",
                    .Arguments = "/c nconvert -out png -q 50 -o """ & preImagePath & """ """ & curitem & """",
                    .WindowStyle = ProcessWindowStyle.Hidden
                }
                Process.Start(psi)
            End If

            ' 建立一個 image 物件來讀取圖片檔案
            ' 假設圖片建立需要2.5秒鐘
            ' Dim watcher1 = get_file_watch(recordPath & "clip圖像")
            ' 尋找最新圖片是否存在

        Else
            Try
                Dim img As New MagickImage(curitem)
                Using memStream As New MemoryStream()
                    img.Format = MagickFormat.Bmp
                    img.Write(memStream)
                    PictureBox1.Image = New Bitmap(memStream)
                    Label15.Text = PictureBox1.Image.Size.Width & "px " & PictureBox1.Image.Size.Height & "px"
                End Using
                If Form2.Visible = True Then Form2.PictureBox1.Image = PictureBox1.Image
                Transform_index_status_to_Form2()
            Catch ex As Exception
                Dim shellFile1 As ShellFile = ShellFile.FromFilePath(curitem)
                shellFile1.Thumbnail.FormatOption = ShellThumbnailFormatOption.Default
                Dim shellThumb As Bitmap = shellFile1.Thumbnail.Bitmap
                PictureBox1.Image = CType(shellThumb, Image)
                If Form2.Visible = True Then Form2.PictureBox1.Image = PictureBox1.Image
                Transform_index_status_to_Form2()
            End Try
        End If

    End Sub
    Private Sub Addsubtitle(curitem As String)

        Dim itemDirectory = Path.GetDirectoryName(curitem)
        Dim subtitlePath As String = Path.Combine(itemDirectory, Path.GetFileNameWithoutExtension(curitem))
        'Console.WriteLine(File.Exists(subtitlePath))
        If File.Exists(subtitlePath + ".ass") Or File.Exists(subtitlePath + ".srt") Then
            Return
            '_mediaPlayer.AddSlave(MediaSlaveType.Subtitle, subtitlePath, True)
        End If
        Dim rgx As New Regex("(.*\W*)\W(\d{1}|\d{2})(\W.*)")
        Dim match_result As Match = rgx.Match(FileCollection.SelectedItem)
        Dim episode As Integer = Integer.Parse(match_result.Groups(2).Value)
        For Each files In Directory.GetFiles(itemDirectory)
            If Not Ext_judge(Path.GetExtension(files), "字幕") Then Continue For
            If rgx.Match(files).Groups(2).Value = episode Then
                'Console.WriteLine(files)
                _mediaPlayer.AddSlave(MediaSlaveType.Subtitle, files, True)
                'Return
            End If
        Next
    End Sub

    <Flags>
    Public Enum SIIGBF
        SIIGBF_RESIZETOFIT = &H0
        SIIGBF_BIGGERSIZEOK = &H1
        SIIGBF_MEMORYONLY = &H2
        SIIGBF_ICONONLY = &H4
        SIIGBF_THUMBNAILONLY = &H8
        SIIGBF_INCACHEONLY = &H10
        SIIGBF_CROPTOSQUARE = &H20
        SIIGBF_WIDETHUMBNAILS = &H40
        SIIGBF_ICONBACKGROUND = &H80
        SIIGBF_SCALEUP = &H100
    End Enum
    <DllImport(“shell32.dll”, CharSet:=CharSet.Unicode)>
    Private Shared Function SHCreateItemFromParsingName(ByVal path As String, ByVal pbc As IntPtr, <MarshalAs(UnmanagedType.LPStruct)> ByVal riid As Guid, <Out> ByRef factory As IShellItemImageFactory) As Integer
    End Function

    <ComImport>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    <Guid(“bcc18b79-ba16-442f-80c4-8a59c30c463b”)>
    Private Interface IShellItemImageFactory
        <PreserveSig>
        Function GetImage(ByVal size As Size, ByVal flags As SIIGBF, <Out> ByRef phbm As IntPtr) As Integer
    End Interface

    Private Sub Transfer_image_toForm2()
        Form2.PictureMode()
        Transform_index_status_to_Form2()
        Form2.Adjustment_Form()
        Form2.Label15.Text = Label15.Text
    End Sub
    Private Shared Sub DisplayPropertyValue(ByVal prop As IShellProperty)
        Dim value As String
        value = If(prop.ValueAsObject Is Nothing, "", prop.FormatForDisplay(PropertyDescriptionFormatOptions.None))
        Debug.WriteLine(prop.CanonicalName & " " & value)
    End Sub

    Public Function Try_password(curitem As String, index As Short)
        Console.WriteLine(PasswordBox.Items(index))
        PasswordBox.SetSelected(index, True)
        Dim passward As String = PasswordBox.Items(index).ToString
        Try
            archive = ArchiveFactory.OpenArchive(curitem, New SharpCompress.Readers.ReaderOptions With {.Password = passward})
        Catch ex As Exception
            If Not index = PasswordBox.Items.Count - 1 Then
                passward = Try_password(curitem, index + 1)
            Else
                MsgBox("讀取檔案失敗" & curitem & "可能有密碼或密碼錯誤")
                archive.Dispose()
            End If
        End Try
        Return passward
    End Function
    Public Sub Archive_Preview(sender As Object, e As EventArgs) Handles SubFileCollection.SelectedIndexChanged, SubFileCollection.Click
        'CheckedListBox1_setCheck("隱藏壓縮檔案集合", False)
        If Sub_checkFileSelected() Then Return
        If CheckedListBox1_isCheck("鎖定") Then Return
        release_all_process(False)
        'targetList = subFileCollection
        If CheckedListBox1_isCheck("Close Preview") Then Return
        If SubFileCollection.Items.Count = 0 Or SubFileCollection.SelectedIndex = -1 Then Return
        Refresh_backup()
        Dim curitem As String = SubFileCollection.SelectedItem
        If curitem.Contains("download in") Then
            Dim sort As String() = curitem.Split(" download in ")
            curitem = sort(3) & sort(0)
            curitem = GetMovedFolder(curitem)
            Console.WriteLine(curitem)
        End If
        If Ext_judge(curitem, "網頁") Then
            Call Form2.WebMode(curitem)
            Form2.Show()
            Return
        End If
        Dim ext = Path.GetExtension(curitem)
        Dim fileTrueName As String = FileCollection.SelectedItem.ToString.Replace("★", "")
        Dim FileExt As String = Path.GetExtension(fileTrueName)
        Console.WriteLine(FileExt)
        Dim foundfile As String = ""
        If Not Ext_judge(fileTrueName, "網頁") Then
            foundfile = Path.GetFullPath(fileTrueName)
        End If

        ExtInput.Text = ext
        FileNameComboBox.Text = Path.GetFileNameWithoutExtension(curitem)

        Using curitemStream As Stream = GetPreviewStream(curitem, foundfile, FileExt)

            If curitemStream Is Nothing Then Return

            If Ext_judge(ext, "光盤") Then
                instObj = curitemStream
            ElseIf Ext_judge(ext, "其他圖片") Then
                If File.Exists(curitem) Then
                    specialPicture(curitem)
                Else
                    Try
                        Dim img As New MagickImage(curitemStream)
                        Using memStream As New MemoryStream()
                            img.Format = MagickFormat.Bmp
                            img.Write(memStream)
                            PictureBox1.Image = New Bitmap(memStream)

                        End Using
                        If Form2.Visible = True Then Form2.PictureBox1.Image = PictureBox1.Image
                    Catch ex As Exception
                        MsgBox("無法預覽檔案" & curitem & "無法開啟" & ex.GetType().FullName & ex.Message)
                    End Try
                End If

            ElseIf Ext_judge(ext, "Images") Then
                PictureBox1.Visible = True
                RichTextBox1.Hide()
                Label15.Show()
                Try
                    Dim MyImage As New Bitmap(curitemStream)
                    Label15.Text = MyImage.Size.Width & "px " & MyImage.Size.Height & "px"
                    PictureBox1.Image = CType(MyImage.Clone, Image)
                    Transfer_image_toForm2()
                    MyImage.Dispose()
                Catch ex As Exception
                    MsgBox("無法顯示圖片" & ex.GetType().FullName & ex.Message)
                End Try
                curitemStream.Dispose()
            ElseIf Ext_judge(ext, "Media") Then
                Label15.Hide()
                RichTextBox1.Hide()
                PictureBox1.Visible = If(CheckedListBox1_isCheck("ASMR播放模式"), True, False)
                ComboBox5.Show()
                DeviceComboBox.Show()
                videoPanel.Show()
                VideoView1.Visible = If(CheckedListBox1_isCheck("ASMR播放模式"), False, True)
                Dim sameMedia As Boolean = checkSameMedia(curitem)
                If sameMedia Then
                    _mediaPlayer.Play()
                    Button24.Text = "暫停"
                    Return
                End If
                If Ext_judge(foundfile, "目錄") Or Ext_judge(foundfile, ".trf") Then
                    Me.BeginInvoke(New Action(Sub()
                                                  Dim media As New Media(_libvlc, curitem, FromType.FromPath)
                                                  ' 加載字幕

                                                  _mediaPlayer.Media = media
                                                  'addsubtitle(curitem)

                                                  _mediaPlayer.Play()
                                                  If lastPosition <> 0 Then _mediaPlayer.Time = lastPosition
                                                  lastPosition = 0

                                              End Sub))
                End If
                Dim volume As Integer = _mediaPlayer.Volume
                TrackBar1.Value = If(volume >= 0, volume, 0)
                Button24.Text = "暫停"
            ElseIf Ext_judge(ext, "文字") Then
                PictureBox1.Hide()
                RichTextBox1.Show()
                Using fileReader As New StreamReader(curitemStream)
                    Dim stringReader As String = fileReader.ReadToEnd
                    RichTextBox1.Text = stringReader
                End Using
            End If

        End Using
    End Sub

    Private Function GetPreviewStream(curitem As String, foundfile As String, fileExt As String) As Stream

        If File.Exists(curitem) Then
            Return File.OpenRead(curitem)
        End If

        If Ext_judge(fileExt, "光盤") Then
            Dim isoStream = File.OpenRead(foundfile)
            Dim cd = New CDReader(isoStream, True)
            Return cd.OpenFile(curitem, FileMode.Open)
        End If

        If Ext_judge(fileExt, "壓縮檔") Then
            Return archive.Entries(SubFileCollection.SelectedIndex).OpenEntryStream()
        End If

        Return Nothing
    End Function
    Private Sub ExtrackAllFilesInIso(Directorie As DiscUtils.DiscDirectoryInfo, cd As Object, savefilename As String)
        Dim directname As New StringBuilder("")

        ' Append more strings using + operator
        directname.Append(savefilename & Path.DirectorySeparatorChar & Directorie.FullName)

        If Not FileSystem.DirectoryExists(directname.ToString) Then
            FileSystem.CreateDirectory(directname.ToString)
        End If
        For Each CDentry In Directorie.GetFiles
            MessageLabel.Text = "解壓" & CDentry.Name
            Dim newfilePath As New StringBuilder("")
            newfilePath.Append(directname.ToString & Path.DirectorySeparatorChar & CDentry.Name.Replace(";1", "").ToLower)
            Dim newfile = File.Create(newfilePath.ToString)
            Dim cdpath As Stream = cd.OpenFile(CDentry.FullName, FileMode.Open)
            cdpath.CopyTo(newfile)
            newfile.Close()
        Next
        For Each CDentry In Directorie.GetDirectories
            ExtrackAllFilesInIso(CDentry, cd, savefilename)
        Next
    End Sub
    Public Function Checkindexlast(selected_list As Array, maxItem As Short)
        'selected_list need reverse
        If selected_list.Length = 0 Then
            Return Nothing
        ElseIf selected_list(0) = maxItem - 2 Then
            Array.Clear(selected_list, 0, 1)
            maxItem -= 1
            Checkindexlast(selected_list, maxItem)
        ElseIf selected_list.Length <> 1 Then
            selected_list(0) -= 1
            Array.Reverse(selected_list)
            Array.Clear(selected_list, 0, 1)
        Else
            Array.Reverse(selected_list)
            Return selected_list(0)
        End If
        Return Nothing
    End Function
    Public Sub Release_all_process(Optional archive_close As Boolean = True)


        If CheckedListBox1_isCheck("鎖定") Then Return

        'If Not isReadingData Then lastPosition = 0
        If CheckedListBox1_isCheck("ASMR播放模式") Then Return
        If videoPanel.Visible Then
            videoPanel.Hide()
        End If
        ComboBox5.Hide()
        DeviceComboBox.Hide()
        DataGridView1.Hide()
        'RichTextBox1.Hide()
        If archive IsNot Nothing And archive_close Then
            archive.Dispose()
            archive = Nothing
        End If
        'If PictureBox1.Image IsNot Nothing Then PictureBox1.Image.Dispose()
        'PictureBox1.Image = New Bitmap(1, 1)
        'If Form2.PictureBox1.Image IsNot Nothing Then Form2.PictureBox1.Image.Dispose()
        'Form2.PictureBox1.Image = New Bitmap(1, 1)
        ' 清圖片
        If PictureBox1.Image IsNot Nothing Then
            PictureBox1.Image.Dispose()
            PictureBox1.Image = Nothing
        End If

        If Form2.PictureBox1.Image IsNot Nothing Then
            Form2.PictureBox1.Image.Dispose()
            Form2.PictureBox1.Image = Nothing
        End If

    End Sub
    Private Function CheckSameMedia(curitem As String) As Boolean
        If ComboBox5.SelectedItem IsNot Nothing Then
            If ComboBox5.SelectedItem.ToString() = "自動重播(單首)" And Not _mediaPlayer.Media.State = VLCState.Playing Then
                Return False
            End If
        End If
        If _mediaPlayer.Media IsNot Nothing Then
            If curitem = New Uri(_mediaPlayer.Media.Mrl).LocalPath Then
                Return True
            End If
        End If
        Return False
    End Function


    Public Async Sub OnButton4Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim control As String = GroupBox2.Controls.OfType(Of RadioButton)().FirstOrDefault(Function(n) n.Checked).Name
        If control = "SpecialButton" AndAlso targetList Is SubFileCollection Then
            targetList = FileCollection
            OnButton4Click(sender, e)
            Return
        End If
        isListboxOperating = True
        Add_new_path(PathComboBox, "ori_pathrecord.txt")
        Add_new_path(TargetComboBox, "des_pathrecord.txt")
        release_all_process()
        If TargetComboBox.Text.Length = 0 Then TargetComboBox.Text = PathComboBox.Text
        Dim arguments As String = ""
        Dim counter = -1
        Dim collections As String() = New String() {""}
        Dim collections_number As New List(Of Integer)



        If targetList Is Nothing Then targetList = FileCollection
        If CheckBox5.Checked = False Then
            collections = targetList.SelectedItems.Cast(Of String).ToArray
            collections_number = targetList.SelectedIndices.Cast(Of Integer).ToList
        Else
            collections = targetList.Items.Cast(Of String).ToArray
            Dim indecis As New List(Of Integer)
            For i = 0 To targetList.Items.Count - 1
                indecis.Add(i)
            Next
            collections_number = indecis.Cast(Of Integer).ToList
        End If
        ' If ComboBox5.SelectedItem = "紀錄開啟檔案" = True Then Call Output_file_label(sender, e)

        Console.WriteLine(collections(0))
        Select Case control
            Case "OpenButton"
                '允許.torrent通過
                If collections.Count > 5 And Array.Find(Of String)(collections, Function(x) Not Ext_judge(Path.GetExtension(x), ".torrent")) IsNot Nothing Then
                    collections = collections.Take(5).ToArray()
                End If
                Button4openFile(collections, collections_number)
            Case "MoveButton"
                For Each foundfile As String In collections
                    counter += 1
                    Dim filename As String = Path.GetFileName(foundfile)
                    Dim ext As String = Path.GetExtension(foundfile).ToLower
                    If Ext_judge(foundfile, "目錄") Then
                        Dim directname As String = FileSystem.GetDirectoryInfo(foundfile).Name
                        Try
                            My.Computer.FileSystem.MoveDirectory(foundfile, TargetComboBox.Text & Path.DirectorySeparatorChar & directname, UIOption.AllDialogs)
                            FileCollection.Items.Remove(foundfile)
                        Catch ex As Exception
                            MsgBox("移動目錄失敗" & filename & ex.GetType().FullName & ex.Message)
                        End Try
                    Else
                        Try
                            My.Computer.FileSystem.MoveFile(foundfile, TargetComboBox.Text & Path.DirectorySeparatorChar & filename, UIOption.AllDialogs)
                            FileCollection.Items.Remove(foundfile)
                        Catch ex As Exception
                            MsgBox("移動檔案失敗" & filename & ex.GetType().FullName & ex.Message)
                        End Try
                    End If
                Next
            Case "CopyButton"
                For Each foundfile As String In collections
                    counter += 1
                    Dim filename As String = Path.GetFileNameWithoutExtension(foundfile)
                    Dim ext As String = Path.GetExtension(foundfile)
                    Dim fn As Match = dlsiteRgx.Match(foundfile)
                    If fn.Success Then filename = fn.Value + "_" + filename
                    Dim dir = If(Directory.Exists(TargetComboBox.Text), TargetComboBox.Text, Path.GetDirectoryName(foundfile))
                    Dim newFileName As String = Path.Combine(dir, filename & ext)

                    If Not fn.Success Then
                        Dim i = 1
                        While File.Exists(newFileName) Or i > 50
                            newFileName = Path.Combine(dir, $"{filename}_複製{i}{ext}")
                            i += 1
                        End While
                    End If

                    My.Computer.FileSystem.CopyFile(foundfile, newFileName, UIOption.AllDialogs)
                Next
            Case "DeleteButton"
                Dim namecount As Integer = If(collections.Length < 11, collections.Length, 10)
                Dim names As String = String.Join(", ", collections.Take(namecount))
                Dim result As DialogResult = MessageBox.Show($"確定要刪除{names}嗎？", "確認刪除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

                ' 根據用戶的選擇來執行刪除操作
                If result = DialogResult.Yes Then
                    ' 執行刪除操作

                    For Each foundfile As String In collections
                        counter += 1
                        Delete_Anothor_Thread(foundfile)
                    Next
                    MessageBox.Show("已刪除", "通知", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    ' 用戶選擇不刪除
                    MessageBox.Show("操作已取消", "通知", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If

            Case "SpecialButton"
                Select Case ComboBox7.Text
                    Case "Rename"
                        For Each foundfile As String In collections
                            counter += 1
                            Dim checkIndex = collections_number(counter)
                            targetList.Items(checkIndex) = Rename_Anothor_Thread(foundfile)
                        Next
                    Case "Set Up"
                        For Each foundfile As String In collections
                            If TargetComboBox.Text.Contains(".trf") Or String.IsNullOrEmpty(TargetComboBox.Text) Then TargetComboBox.Text = Path.GetDirectoryName(foundfile)
                            counter += 1
                            Dim filename As String = Path.GetFileName(foundfile)
                            Dim ext As String = Path.GetExtension(foundfile).ToLower
                            Dim filestream1 As FileStream = Nothing
                            If Ext_judge(ext, "光盤") Then
                                filestream1 = File.OpenRead(foundfile)
                            Else
                                filestream1 = New FileStream(TargetComboBox.Text & "file", FileMode.Create)
                                instObj.CopyTo(filestream1)

                            End If

                            Using vhdStream As FileStream = filestream1

                                Dim cd As New CDReader(vhdStream, True)
                                Dim pattern As String = "[AaUuTtOoRrUuNn]{7}\.[infINF]{3}"
                                Dim rgx As New Regex(pattern)
                                Dim infor As String = "AUTORUN.INF"
                                Dim newfile As FileStream
                                Dim path1 As Stream
                                MessageLabel.Visible = True
                                Dim saveFolder As String = TargetComboBox.Text & Path.DirectorySeparatorChar & Path.GetFileNameWithoutExtension(foundfile)
                                Directory.CreateDirectory(saveFolder)

                                For Each CDentry In cd.Root.GetFiles
                                    newfile = File.Create(saveFolder & Path.DirectorySeparatorChar & CDentry.Name.Replace(";1", "").ToLower)
                                    path1 = cd.OpenFile(CDentry.FullName, FileMode.Open)
                                    path1.CopyTo(newfile)
                                    newfile.Close()
                                    For Each match As Match In rgx.Matches(CDentry.FullName.Replace(";1", "").ToLower)
                                        Console.WriteLine("Found '{0}' at position {1}", match.Value, match.Index)
                                        infor = match.Value
                                    Next

                                Next
                                For Each Director1 In cd.Root.GetDirectories
                                    ExtrackAllFilesInIso(Director1, cd, saveFolder)
                                Next
                                Console.WriteLine(infor)
                                path1 = cd.OpenFile(infor, FileMode.Open)
                                Dim tr As TextReader = New StreamReader(path1)
                                MessageLabel.Text = "尋找啟動程序.."
                                Dim sentence As String = tr.ReadToEnd
                                pattern = "(?<=open[\s]*=[\s]*)[\w.\\]+"
                                rgx = New Regex(pattern)
                                For Each match As Match In rgx.Matches(sentence)
                                    Console.WriteLine("Found '{0}' at position {1}", match.Value, match.Index)
                                    Process.Start(saveFolder & Path.DirectorySeparatorChar & match.Value)
                                Next

                            End Using

                            If CheckedListBox1_isCheck("壓縮/解壓後刪除") Then
                                Try
                                    My.Computer.FileSystem.DeleteFile(foundfile)
                                Catch ex As Exception
                                    MsgBox("刪除檔案失敗" & filename & "可能已經刪除" & ex.GetType().FullName & ex.Message)
                                End Try
                            End If
                        Next
                    Case "Compression"
                        For Each foundfile As String In collections
                            counter += 1
                            Dim filename As String = Path.GetFileName(foundfile)
                            Dim ext As String = Path.GetExtension(foundfile).ToLower
                            Using zipToOpen As New FileStream(foundfile, FileMode.Open)
                                Using archive As New ZipArchive(zipToOpen, ZipArchiveMode.Update)
                                    Dim readmeEntry As ZipArchiveEntry = archive.CreateEntry(foundfile)
                                End Using
                            End Using
                            If CheckedListBox1_isCheck("Delete After Compress/Decompress") Then
                                Try
                                    My.Computer.FileSystem.DeleteFile(foundfile)
                                Catch ex As Exception
                                    MsgBox("刪除檔案失敗" & filename & "可能已經刪除" & ex.GetType().FullName & ex.Message)
                                End Try
                            End If
                        Next
                    Case "Decompression"
                        For Each foundfile As String In collections
                            counter += 1
                            Dim filename As String = Path.GetFileName(foundfile)
                            Dim ext As String = Path.GetExtension(foundfile).ToLower
                            Dim rename As String = Path.GetFileNameWithoutExtension(foundfile)
                            Dim entryCount As Short = 0
                            Dim extractedCount As Short = 0
                            Try
                                archive = ArchiveFactory.OpenArchive(foundfile)
                            Catch
                                Dim passward = try_password(foundfile, 0)
                                archive = ArchiveFactory.OpenArchive(foundfile, New SharpCompress.Readers.ReaderOptions With {.Password = passward})
                            End Try
                            'ProgressBar1.Maximum = archive.Entries.Where(Function(ex) Not ex.IsDirectory).Sum(Function(ex) ex.Size)
                            'ProgressBar1.Value = 0
                            For Each entry In archive.Entries
                                entryCount += 1
                                If Not entry.IsDirectory Then
                                    Console.WriteLine(entry.Key)
                                    entry.WriteToDirectory(TargetComboBox.Text & Path.DirectorySeparatorChar & rename, New ExtractionOptions With
                                  {.ExtractFullPath = True, .Overwrite = True})
                                    extractedCount += 1
                                    Dim percent = CInt((extractedCount / entryCount) * 100)
                                    ProgressBar1.Value = percent
                                End If
                            Next
                            archive.Dispose()
                            If CheckedListBox1_isCheck("Delete After Compress/Decompress") Then
                                Try
                                    My.Computer.FileSystem.DeleteFile(foundfile)
                                Catch ex As Exception
                                    MsgBox("刪除檔案失敗" & filename & "可能已經刪除" & ex.GetType().FullName & ex.Message)
                                End Try
                            End If
                            MessageLabel.Text = rename & "解壓完成"
                        Next
                    Case "檢查路徑"
                        For Each foundfile As String In collections
                            counter += 1
                            Dim checkIndex = collections_number(counter)
                            If File.Exists(foundfile) Then Continue For
                            FileCollection.Items(checkIndex) = GetMovedFolder(foundfile)
                        Next
                    Case "檢查網址"
                        If targetList IsNot FileCollection Then Return
                        For Each foundfile As String In collections
                            counter += 1
                            Dim checkIndex = collections_number(counter)
                            Dim original As String = FileCollection.Items(checkIndex).ToString()
                            Dim standardized As String = Standardized_denomination(original)
                            If standardized.Contains("x.com/i/status/") Then
                                standardized = Await ResolveIToUsername(standardized)
                            End If
                            FileCollection.Items(checkIndex) = Standardized_denomination(standardized)
                        Next
                    Case "加標籤"
                        If targetList IsNot FileCollection Then Return
                        For Each foundfile As String In collections
                            counter += 1
                            Dim checkIndex = collections_number(counter)
                            Handle_label(checkIndex)
                        Next
                    Case "自動分類"
                        Form3.AutoSort(collections.Cast(Of String).ToList, FileCollection)
                        CheckBox5.Checked = False
                    Case "計算"
                        Dim totalSum As Integer = 0
                        Dim rgx As New Regex("(\d+)")
                        For Each foundfile As String In collections
                            Dim matches = rgx.Matches(foundfile)
                            If rgx.IsMatch(foundfile) Then
                                For Each match In matches
                                    totalSum += Convert.ToInt32(match.Groups(1).Value)
                                    Console.WriteLine(match.Groups(1).Value)
                                Next
                            End If
                        Next
                        Dim bmp As New Bitmap(PictureBox1.Width, PictureBox1.Height)
                        Using g As Graphics = Graphics.FromImage(bmp)
                            Dim message As String = $"總共{totalSum}元"
                            g.DrawString(message, New Font("新細明體", 9), Brushes.Black, New PointF(10, 10))
                        End Using
                        PictureBox1.Image = bmp
                    Case "imgToPdf"
                        ' 取得預設檔名
                        Dim defaultName As String = Path.GetFileNameWithoutExtension(collections(0))

                        ' 跳出輸入框來輸入pdfname
                        Dim pdfname As String = InputBox("請輸入PDF檔名:", "PDF名稱輸入", defaultName)
                        If String.IsNullOrEmpty(pdfname) Then Return

                        Dim pdfPath As String = $"{TargetComboBox.Text}\{pdfname}.pdf"
                        Dim doc As New PdfSharp.Pdf.PdfDocument()

                        ' 檢查畫面上特定的 CheckBox 是否被勾選 (請將 CheckBox3 替換成你實際使用的 CheckBox 名稱)
                        Dim isPrintName As Boolean = CheckedListBox1_isCheck("imgToPdf打印文件名")

                        For Each foundfile As String In collections
                            If Ext_judge(Path.GetExtension(foundfile), "Images") Then
                                Dim imgPath As String = foundfile
                                Dim fileName As String = Path.GetFileName(foundfile)

                                Dim imgStream As New MemoryStream()

                                ' 使用 GDI+ 處理透明度與高相容性白底
                                Using originalBmp As New Bitmap(imgPath)
                                    Using newBmp As New Bitmap(originalBmp.Width, originalBmp.Height, PixelFormat.Format24bppRgb)
                                        Using g As Graphics = Graphics.FromImage(newBmp)
                                            ' 將背景填滿白色
                                            g.Clear(Color.White)
                                            ' 繪製原圖
                                            g.DrawImage(originalBmp, 0, 0, originalBmp.Width, originalBmp.Height)

                                            ' 【新增判斷】如果使用者有勾選 CheckBox，才在左上角畫上檔案名稱
                                            If isPrintName Then
                                                ' 使用新細明體，並用紅色顯眼標註，以利 AI 與肉眼辨識
                                                g.DrawString(fileName, New Font("新細明體", 16), Brushes.Black, New PointF(5, 10))
                                            End If
                                        End Using

                                        ' 轉存為高相容性的 JPEG 格式至記憶體
                                        newBmp.Save(imgStream, ImageFormat.Jpeg)
                                    End Using
                                End Using

                                ' 倒回串流起點讓 PDFsharp 讀取
                                imgStream.Position = 0

                                ' 將處理好的乾淨圖片載入並塞入 PDF 頁面
                                Using img As XImage = XImage.FromStream(imgStream)
                                    Dim page As PdfPage = doc.AddPage()
                                    page.MediaBox = New PdfSharp.Pdf.PdfRectangle(New XPoint(0, 0), New XPoint(img.PointWidth, img.PointHeight))

                                    Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
                                    gfx.DrawImage(img, 0, 0, img.PointWidth, img.PointHeight)
                                End Using

                                imgStream.Dispose()
                            End If
                        Next

                        ' 儲存與關閉 (已成功安裝 System.ValueTuple 套件，可直接使用原生 Save)
                        doc.Save(pdfPath)
                        doc.Close()
                    Case "sendtoDL"
                        Await SendToDownloadService(collections.ToList())
                End Select

        End Select
        Dim app As String = ""
        'If CheckedListBox1_isCheck("clip啟動") Then
        '    app = "C:\Program Files\CELSYS\CLIP STUDIO 1.5\CLIP STUDIO PAINT\CLIPStudioPaint.exe"
        'ElseIf CheckedListBox1_isCheck("MuseScore啟動") Then
        '    app = "C:\Program Files\MuseScore 3\bin\MuseScore3.exe"
        'End If
        If CheckedListBox1_isCheck("啟動") Then
            Dim proinfo As New ProcessStartInfo With {
        .FileName = app,
         .Arguments = arguments,
         .StandardErrorEncoding = Encoding.UTF8,
        .UseShellExecute = False,
         .RedirectStandardError = True,
         .RedirectStandardOutput = True
           }

            Dim prostart As New Process With {
             .StartInfo = proinfo
}
            prostart.Start()

            '方案1 使用cmd啟動程式
            '方案2 修改副檔名啟動的登機碼
            'SetAssociation(ext, ext & "File", app, ext & " File")
            ' Process.Start(foundfile)
        End If

        If Not CheckBox5.Checked And targetList IsNot SubFileCollection And FileCollection.Items.Count > 0 Then
            If collections_number.Count > 10 Then
                collections_number.RemoveRange(10, collections_number.Count - 10)
            End If
            For Each i As Integer In collections_number
                If i < FileCollection.Items.Count Then
                    FileCollection.SetSelected(i, True)
                End If
            Next
            If collections_number.Count > 0 And FileCollection.SelectedIndex = -1 Then
                FileCollection.SelectedIndex = FileCollection.Items.Count - 1
            End If
        End If
        Dim excluded = {"sendtoDL", "Rename"}

        If CheckBox5.Checked AndAlso Not excluded.Contains(ComboBox7.Text) Then
            OpenButton.Checked = True
        End If
        Refresh_backup()
        'CheckBox12_CheckedChanged(sender, e)
        isListboxOperating = False
    End Sub
    Private Async Function SendToDownloadService(urls As List(Of String)) As Task
        Using client As New Net.Http.HttpClient()
            For Each url In urls
                Try
                    Dim payload As String = $"{{""weburl"":""{url.Replace("""", "\""")}""}}"
                    Dim content As New Net.Http.StringContent(payload, Encoding.UTF8, "application/json")
                    Dim response As Net.Http.HttpResponseMessage =
                    Await client.PostAsync("http://localhost:7999/check-websites", content)
                    response.EnsureSuccessStatusCode()
                    Dim resultText As String = Await response.Content.ReadAsStringAsync()
                    SubFileCollection.Items.Add($"成功: {url}")
                    Console.WriteLine($"成功: {url} -> {resultText}")
                Catch ex As Exception
                    SubFileCollection.Items.Add($"失敗: {url}")
                    Console.WriteLine($"失敗: {url} ({ex.Message})")
                End Try
            Next
        End Using
    End Function
    Private Sub Button4openFile(Collections As String(), collections_number As List(Of Integer))

        For Each foundfile As String In Collections
            Dim fileLabel As String = Remove_label(foundfile)
            Dim filename As String = Path.GetFileName(fileLabel)
            Dim ext As String = Path.GetExtension(fileLabel).ToLower
            Console.WriteLine(ext)
            If ext = ".exe" Then
                If CheckedListBox1_isCheck("locale remulator") Then
                    Dim exeName As String = fileLabel
                    Dim guidNumber As String = "33a7041d-94fa-472c-aeb2-fd49fcc1cfff"
                    'If exeName.IndexOf("CHS", StringComparison.OrdinalIgnoreCase) > 0 Then
                    '    guidNumber = "c87d06b2-5079-48c2-b658-d92175880bf6"
                    'End If
                    Dim p As New ProcessStartInfo()
                    Dim lePath As String = "E:\download\Locale_Remulator.1.5.3-beta.1\Locale_Remulator.1.5.3-beta.1\LRProc.exe"
                    p.FileName = lePath
                    p.Arguments = $"{guidNumber} ""{exeName}"""
                    p.Verb = "runas" ' 要求以系統管理員身份執行
                    p.UseShellExecute = True
                    p.WorkingDirectory = Path.GetDirectoryName(lePath)
                    Dim res As Process = Process.Start(p)
                ElseIf CheckedListBox1_isCheck("locale emulator") Then
                    Dim exeName As String = fileLabel
                    Dim guidNumber As String = "8b37cfbd-4613-4e63-b3ba-d24c4c85c0ff"
                    If exeName.IndexOf("CHS", StringComparison.OrdinalIgnoreCase) > 0 Then
                        guidNumber = "c87d06b2-5079-48c2-b658-d92175880bf6"
                    End If
                    Dim p As New ProcessStartInfo()
                    Dim lePath As String = "C:\Locale.Emulator.2.5.0.1\LEProc.exe"
                    p.FileName = lePath
                    p.Arguments = $"-runas {guidNumber} ""{exeName}"""
                    p.UseShellExecute = False
                    p.WorkingDirectory = Path.GetDirectoryName(lePath)
                    Dim res As Process = Process.Start(p)
                    res.WaitForInputIdle(5000)
                End If
            Else
                Dim psi As New ProcessStartInfo With {
                    .FileName = fileLabel,
                    .WorkingDirectory = Path.GetDirectoryName(fileLabel),
                    .UseShellExecute = True
                }
                Process.Start(psi)
            End If
            If CheckedListBox1_isCheck("執行後移出") Then
                Dim ori_index = FileCollection.SelectedItems.IndexOf(foundfile)
                collections_number.RemoveAt(ori_index)
                FileCollection.Items.Remove(foundfile)
            End If
            If CheckedListBox1_isCheck("執行後移到最後一項") Then
                Dim ori_index = FileCollection.SelectedItems.IndexOf(foundfile)
                FileCollection.Items.Remove(foundfile)
                FileCollection.Items.Add(foundfile)
                If Not CheckedListBox1_isCheck("移動時選取原項") Then
                    collections_number.RemoveAt(ori_index)
                    collections_number.Add(FileCollection.Items.Count - 1)
                End If

            End If

        Next
    End Sub
    Public Function Rename_Anothor_Thread(foundfile As String)
        If (Me.InvokeRequired) Then
            Dim del As New DelShowMessage(AddressOf Rename_Anothor_Thread)
            Me.Invoke(del, foundfile)
        Else
            Console.WriteLine(foundfile)

            Dim folder As String = Path.GetDirectoryName(foundfile)
            Dim newname As String = Path.GetFileNameWithoutExtension(foundfile)
            Dim rgx As New Regex("string:(\d+)")
            If Ext_judge(foundfile, "目錄") Then
                Select Case FileNameComboBox.SelectedIndex
                    Case 1
                        newname = File.GetCreationTime(foundfile).ToString("yyyy_MMdd")
                    Case 2

                        Dim result = rgx.Match(orderComboBox.Text)
                        If result.Success Then
                            Dim targetindex As Integer = CInt(result.Groups(1).Value)
                            newname = newname.Remove(targetindex - 1, 1)
                        End If
                End Select
                Try
                    My.Computer.FileSystem.RenameDirectory(foundfile, newname)
                    Console.WriteLine(folder & newname)
                    Return folder & newname
                Catch ex As IOException
                    Console.WriteLine("命名檔案失敗" & newname & "可能被開啟" & ex.Message)
                Catch ex As UnauthorizedAccessException
                    Console.WriteLine("命名檔案失敗" & newname & "需要系統權限")
                End Try
            Else

                Dim ext As String = Path.GetExtension(foundfile).ToLower
                Select Case FileNameComboBox.SelectedIndex
                    Case 1
                        newname = File.GetLastWriteTime(foundfile).ToString("yyyy_MM_dd-HHmmss")
                    Case 2
                        Dim result = rgx.Match(orderComboBox.Text)
                        If result.Success Then
                            Dim targetindex As Integer = CInt(result.Groups(1).Value)
                            newname = newname.Remove(targetindex - 1, 1)
                        End If
                    Case 3
                        Dim index As Integer = FileCollection.Items.IndexOf(foundfile) + 1
                        newname = orderComboBox.Text + index.ToString()
                End Select
                If CheckedListBox1_isCheck("更改附檔名") Then
                    newname += ExtInput.Text
                Else
                    newname += ext
                End If
                Try
                    release_all_process()
                    Console.WriteLine(foundfile & newname)
                    My.Computer.FileSystem.RenameFile(foundfile, newname)
                    Console.WriteLine(folder & "\" & newname)
                    Return folder & "\" & newname
                    '目錄檔案集合.Items(collections_number(counter)) = Path.GetDirectoryName(foundfile) & "/" & newname
                Catch ex As IOException
                    '換執行緒
                    Console.WriteLine(ex.Message)
                Catch ex As UnauthorizedAccessException
                    Console.WriteLine("命名檔案失敗" & "需要系統權限")
                End Try
            End If
            Return newname
        End If
        Return foundfile
    End Function
    Public Function Delete_Anothor_Thread(foundfile As String)
        If (Me.InvokeRequired) Then
            Dim del As New DelShowMessage(AddressOf Delete_Anothor_Thread)
            Me.Invoke(del, foundfile)
        Else
            If Not Directory.Exists(foundfile) And Not File.Exists(foundfile) Then
                Return True
            End If
            If Ext_judge(foundfile, "目錄") Then
                Dim directname = FileSystem.GetDirectoryInfo(foundfile).FullName
                FileSystem.DeleteDirectory(directname, DeleteDirectoryOption.DeleteAllContents)
                targetList.Items.Remove(foundfile)
            Else
                My.Computer.FileSystem.DeleteFile(foundfile)
                targetList.Items.Remove(foundfile)
            End If
        End If
        Return True
    End Function




    Public Shared Sub SetAssociation(ByVal Extension As String, ByVal KeyName As String, ByVal OpenWith As String, ByVal FileDescription As String)
        Dim BaseKey As RegistryKey
        Dim OpenMethod As RegistryKey
        Dim Shell As RegistryKey

        Dim CurrentUser = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & Extension, True)
        CurrentUser.DeleteSubKey("UserChoice", False)
        CurrentUser.Close()
        SHChangeNotify(&H8000000, &H0, IntPtr.Zero, IntPtr.Zero)
        BaseKey = Registry.ClassesRoot.CreateSubKey(Extension)
        BaseKey.SetValue("", KeyName)

        OpenMethod = Registry.ClassesRoot.CreateSubKey(KeyName)
        OpenMethod.SetValue("", FileDescription)
        OpenMethod.CreateSubKey("DefaultIcon").SetValue("", "\' + OpenWith + " \ ",0")
        Shell = OpenMethod.CreateSubKey("Shell")
        Shell.CreateSubKey("edit").CreateSubKey("command").SetValue("", """" & OpenWith & """" & " ""%1""")
        Shell.CreateSubKey("open").CreateSubKey("command").SetValue("", """" & OpenWith & """" & " ""%1""")
        BaseKey.Close()
        OpenMethod.Close()
        Shell.Close()

        CurrentUser = Registry.CurrentUser.CreateSubKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" + Extension)
        CurrentUser = CurrentUser.OpenSubKey("UserChoice", RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.FullControl)
        CurrentUser.SetValue("Progid", KeyName, RegistryValueKind.String)
        CurrentUser.Close()


    End Sub

    <DllImport("shell32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Sub SHChangeNotify(ByVal wEventId As UInteger, ByVal uFlags As UInteger, ByVal dwItem1 As IntPtr, ByVal dwItem2 As IntPtr)

    End Sub

    Public Function Get_file_watch(path As String)
        Dim watcher = New FileSystemWatcher(path) With {
           .NotifyFilter = (NotifyFilters.LastAccess _
Or NotifyFilters.LastWrite _
Or NotifyFilters.FileName _
Or NotifyFilters.DirectoryName),
           .Filter = "*.*",
           .IncludeSubdirectories = True 'CheckBox3.Checked
           }
        ' .Filter = "*.*" 可匹配所有檔案
        ' Add event handlers.
        AddHandler watcher.Changed, AddressOf OnChanged
        AddHandler watcher.Created, AddressOf OnCreated
        AddHandler watcher.Deleted, AddressOf OnChanged
        AddHandler watcher.Renamed, AddressOf OnRenamed
        AddHandler watcher.Disposed, AddressOf OnDisposed
        'add filter
        watcher.EnableRaisingEvents = True
        Return watcher
        ' Begin watching.

        '--------------------------------------------------------------------------
    End Function
    Private Sub RadioButton2_Watch(sender As Object, e As EventArgs) Handles RadioButton2.Click
        Dim watcher1 = get_file_watch("C:\")
        Dim watcher2 = get_file_watch("E:\")
        Dim watcher3 = get_file_watch("F:\")
    End Sub
    Private Sub Game_Watch(sender As Object, e As EventArgs)
        Dim watcher1 = Get_file_watch("F:\galgame")
        Dim watcher2 = Get_file_watch("F:\工口rpg")
    End Sub


    Public Delegate Sub ExampleCallback(lineCount As Integer)
    Private Sub OnChanged(source As Object, e As FileSystemEventArgs)
        'check_order(e.FullPath)

    End Sub
    Private Sub OnDisposed(source As Object, e As FileSystemEventArgs)

    End Sub

    Public Function Index_Anothor_Thread(name As String)
        If (Me.InvokeRequired) Then
            Dim del As New DelShowMessage(AddressOf Index_Anothor_Thread)
            Me.Invoke(del, name)
        Else
            Thread.Sleep(5000)
            Dim preview_image As Image = Image.FromFile(name)
            '將 image 物件指派給 PictureBox 控制項的 Image 屬性
            PictureBox1.Image = CType(preview_image.Clone, Image)
            Label15.Text = preview_image.Size.Width & "px " & preview_image.Size.Height & "px"
            '釋放 image 物件的資源
            preview_image.Dispose()
            If Form2.Visible = True Then Form2.PictureBox1.Image = PictureBox1.Image
            Transform_index_status_to_Form2()
        End If
        Return True
    End Function
    Public Function Combobox3AnothorThread(matchtext As String)
        If (Me.InvokeRequired) Then
            Dim del As New DelShowMessage(AddressOf Combobox3AnothorThread)
            Return Me.Invoke(del, matchtext)
        Else
            Return PathComboBox.Text.Contains(matchtext)
        End If
    End Function
    Private Sub FileFilter(name As String)
        'If isopening Then Return
        Dim ext As String = Path.GetExtension(name).ToLower
        If Ext_judge(ext, ".clip") And Combobox3AnothorThread("clip.trf") Then
            Console.WriteLine($"File: created in {name}")
            AddItemListbox(name)
            Delete_thumbnail(name)
        ElseIf Ext_judge(ext, "遊戲") And Combobox3AnothorThread("遊戲") Then
            If name.IndexOf("config", StringComparison.OrdinalIgnoreCase) > 0 Then Return
            If name.IndexOf("uninstall", StringComparison.OrdinalIgnoreCase) > 0 Then Return
            If name.IndexOf("UnityCrashHandler", StringComparison.OrdinalIgnoreCase) > 0 Then Return
            AddItemListbox(name)
            'ElseIf Ext_judge(ext, "目錄") And combobox3AnothorThread("ASMR") Then
            '    If Directory.GetParent(name).ToString = "E:\ASMR" Then

            '    End If
        ElseIf Ext_judge(ext, "Media") And Combobox3AnothorThread("ASMR") Then
            AddItemListbox(Directory.GetParent(name).ToString)
        ElseIf Combobox3AnothorThread(name) Then
            CheckedListBox1_setCheck("Auto Save", False)
            Button13_Click(New Object, New EventArgs)
            PathComboBox.Name = name
        End If
    End Sub
    Private Sub OnCreated(source As Object, e As FileSystemEventArgs)
        ' Specify what is done when a file is created.
        fileFilter(e.FullPath)
    End Sub
    Public Shared Sub ResultCallback(lineCount As Integer)
        Console.WriteLine(
            "Independent task printed {0} lines.", lineCount)
    End Sub


    Private Delegate Function DelShowMessage(sMessage As String)
    Public Function AddItemListbox(s As String)
        If (Me.InvokeRequired) Then
            Dim del As New DelShowMessage(AddressOf AddItemListbox)
            Me.Invoke(del, s)
        Else
            If s.Contains("CELSYSUserData") Or s.Contains("previewthumb") Then
                Return False
            End If
            If FileCollection.SelectedItem = s Then
                Dim selectIndex = FileCollection.SelectedIndex
                FileCollection.Items.Remove(s)
                FileCollection.Items.Add(s)
                FileCollection.SelectedIndex = selectIndex
            Else
                FileCollection.Items.Remove(s)
                FileCollection.Items.Add(s)
            End If
            Refresh_backup()
        End If
        Return True
    End Function
    Public Sub OnRenamed(source As Object, e As RenamedEventArgs)
        ' Specify what is done when a file is renamed.
        fileFilter(e.FullPath)
        'check_order(e.FullPath)

    End Sub

    Dim WithEvents Client As New WebClient()
    Public Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If Form2.WebView21.Visible = False Then Return
        Add_new_path(TargetComboBox, "des_pathrecord.txt")
        Form2.CheckWebSite()
        'refresh_listbox_numbers()
    End Sub
    Private Sub Client_DownloadProgressChanged(ByVal sender As Object, ByVal e As DownloadProgressChangedEventArgs) Handles Client.DownloadProgressChanged '當Client正在下載時
        ProgressBar1.PerformStep()

    End Sub
    Private Sub Client_DownloadFileCompleted(ByVal sender As Object, ByVal e As AsyncCompletedEventArgs) Handles Client.DownloadFileCompleted  '當Clien結束下載
        'e.Cancelled 判斷是否為中斷(取消)下載
        'e.Error '判斷下載過程是否因發生錯誤而停止下載
    End Sub

    '----form1.vb-------
    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            Button4.Text = "執行全部文件"
        Else
            Button4.Text = "執行選擇文件"
        End If
    End Sub
    Private Sub Form_click(sender As Object, e As EventArgs) Handles Me.Activated
        If Form2.Visible And CheckedListBox1_isCheck("form1與form2連動") Then
            Form2.TopMost = True
            Form2.TopMost = False
            Me.TopMost = True
            Me.TopMost = False
        End If
    End Sub

    Private Sub PictureBox1_DoubleClick(sender As Object, e As EventArgs) Handles PictureBox1.DoubleClick
        Form2.Show()
        Hide()
        Form2.PictureMode()
        Transform_index_status_to_Form2()
    End Sub
    Private Sub VideoView1_DoubleClick(sender As Object, e As EventArgs) Handles VideoView1.DoubleClick
        Form2.Show()
        Hide()
        Form2.MediaMode()
        Transform_index_status_to_Form2()
    End Sub
    Private Function Transform_index_status_to_Form2()
        If targetList Is Nothing Then Return False
        Form2.Index = targetList.SelectedIndex
        Form2.TextBox2.Text = targetList.SelectedIndex + 1
        Form2.Label2.Text = "/" & targetList.Items.Count
        Form2.Text = targetList.SelectedItem
        Return True
    End Function

    Public Sub Remove_Button(sender As Object, e As EventArgs) Handles Button11.Click, Button3.Click
        Dim close_preview_checked As Boolean = CheckedListBox1_isCheck("Close Preview")
        If CheckBox5.Checked = True Then
            Dim delete_items As New List(Of String)
            delete_items.AddRange(FileCollection.Items.Cast(Of String).ToArray())
            For Each i In delete_items
                If Not String.IsNullOrEmpty(SearchTextBox.Text) Then backup.Remove(i)
            Next
            FileCollection.Items.Clear()
            If Not ComboBox7.Text = "sendtoDL" Then CheckBox5.Checked = False
            If Not String.IsNullOrEmpty(SearchTextBox.Text) Then
                SearchTextBox.Text = ""
            End If
        Else
            Dim now_select As Integer
            Dim list_backup As New List(Of String)
            list_backup.AddRange(FileCollection.SelectedItems.Cast(Of String).ToArray())
            now_select = FileCollection.SelectedIndex
            '     MsgBox("backup=" & list_backup.Count &
            'vbCrLf &
            '"selected=" & FileCollection.SelectedItems.Count)
            'FileCollection.ClearSelected()
            For Each i In list_backup
                'MsgBox(i)
                FileCollection.Items.Remove(i)
                If Not String.IsNullOrEmpty(SearchTextBox.Text) Then backup.Remove(i)
            Next
            'MsgBox("移除完成")
            Dim itemCount As Integer = FileCollection.Items.Count
            If itemCount > 0 Then
                FileCollection.SelectedIndex = Math.Min(now_select, itemCount - 1)
            End If
        End If
        CheckedListBox1_setCheck("Close Preview", close_preview_checked)
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        SubFileCollection.Items.RemoveAt(SubFileCollection.SelectedIndex)
    End Sub

    Public Sub CheckedListBox1_setCheck(name As String, value As Boolean)
        For i = 0 To OptionCheckedListBox.Items.Count - 1
            If String.Compare(OptionCheckedListBox.Items(i), name, True) = 0 Then
                OptionCheckedListBox.SetItemChecked(i, value)
                Return
            End If
        Next
    End Sub

    Public Function CheckedListBox1_isCheck(name As String)
        If OptionCheckedListBox.CheckedItems.Count = 0 Then Return False
        Return OptionCheckedListBox.CheckedItems.Contains(name)
    End Function

    Public Sub Add_additional_operation()
        ' 定義 CheckedListBox 和 ComboBox 的選項
        Dim checkedListBoxItems As String() = {
        "Close Preview", "Close Load File",
        "隱藏預覽", "隱藏壓縮檔案集合",
        "固定在最上層", "紀錄關閉視窗",
         "重複移到最底", "插入選取的下一項",
         "更改附檔名", "Auto Save", "下載在TargetBox",
         "移動時選取原項", "下載時不按讚", "關閉自動修正",
        "執行後移出", "執行後移到最後一項",
         "取消自動刪除重複", "Combine trf files",
         "locale emulator", "locale remulator",
         "鎖定", "ASMR播放模式", "使用nconvert", "imgToPdf打印文件名"
    }

        Dim comboBox7Items As String() = {
        "Rename", "sendtoDL", "自動分類", "加標籤", "imgToPdf",
        "Compression", "Decompression", "Set Up",
        "檢查路徑", "檢查網址", "計算"
    }

        Dim comboBox5Items As String() = {
        "不重播", "自動重播(單首)", "自動重播(全部)"
    }
        Dim comboBox1Items As String() = {
        "Select~", "Expandcount startNumber", "Movelistitem", "rename string:(\d+)", "DownloadCount"
    }
        ' 新增 CheckedListBox 選項
        OptionCheckedListBox.Items.AddRange(checkedListBoxItems)

        ' 新增 ComboBox7 選項
        ComboBox7.Items.AddRange(comboBox7Items)

        ' 新增 ComboBox5 選項
        ComboBox5.Items.AddRange(comboBox5Items)

        ' 新增 ComboBox1 選項
        orderComboBox.Items.AddRange(comboBox1Items)
    End Sub

    Private Sub Form1_text1_DragDrop(sender As Object, e As DragEventArgs) Handles PathComboBox.DragDrop, PathComboBox.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        PathComboBox.Text = files(0)
        'Add_new_path(ComboBox3, "ori_pathrecord.txt")
    End Sub
    Private Sub Form1_text2_DragDrop(sender As Object, e As DragEventArgs) Handles TargetComboBox.DragDrop, TargetComboBox.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        TargetComboBox.Text = files(0)
        'Add_new_path(ComboBox4, "des_pathrecord.txt")
    End Sub


    Private Sub CheckBox10_Click(sender As Object, e As EventArgs) Handles CheckBox10.Click
        If CheckBox10.Checked = False Then
            Label6.Text = "目錄類型"
        Else
            Label6.Text = "副檔名"
        End If
        ReadDirectory()
        'PathComboBox_TextChanged(sender, e)
    End Sub

    Private Sub RadioButton15_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton15.CheckedChanged
        allProcesses = Process.GetProcesses()
        For Each p As Process In allProcesses
            FileCollection.Items.Add(p.ProcessName)

        Next
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        If Not String.IsNullOrEmpty(PathComboBox.Text) Then
            PathComboBox.Text = Directory.GetParent(PathComboBox.Text).ToString
        End If
    End Sub
    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        If Not String.IsNullOrEmpty(TargetComboBox.Text) Then
            TargetComboBox.Text = Directory.GetParent(TargetComboBox.Text).ToString
        End If
    End Sub
    Private Sub EditItem()
        Dim input As String = InputBox("編輯地址：", DefaultResponse:=FileCollection.SelectedItem)
        FileCollection.Items(FileCollection.SelectedIndex) = input
    End Sub
    Public Sub AddItem_from_clipboard(sender As Object, e As EventArgs)

        Dim control As ListBox = GetSourceControl(Of ListBox)(sender)
        If control Is Nothing Then Return

        Console.WriteLine($"add item from {control}")
        AddItem(Setdata_from_clipboard())
    End Sub
    Public Sub AddItem_from_textbox(sender As Object, e As EventArgs)

        Dim control As ListBox = GetSourceControl(Of ListBox)(sender)
        If control Is Nothing Then Return
        Console.WriteLine($"add item from {control}")
        targetList = control
        Dim item As String
        '跳出含textbox的訊息框，有確定和取消，按確定就把item存入listbox
        item = InputBox("輸入新項目：", DefaultResponse:=String.Empty)
        If Not String.IsNullOrEmpty(item) Then
            AddItem({Standardized_denomination(item)})
        End If
    End Sub
    Public Sub AddItem(targets As String())

        If Not CheckedListBox1_isCheck("插入選取的下一項") Then
            ' 不需要插入，直接處理
            For Each item In targets
                AddLink(Standardized_denomination(item))
            Next
            MessageLabel.Text = "已加入 " & targets.Length & " 個項目"
        Else
            InsertLink(targets)
        End If
    End Sub
    Public Sub AddItem(target As String)
        AddItem({target})
    End Sub

    Private Function GetControlWithFallback(sender As ToolStripMenuItem) As ListBox
        If sender IsNot Nothing Then
            Dim strip As ContextMenuStrip = TryCast(sender.Owner, ContextMenuStrip)
            If strip IsNot Nothing Then
                Return CType(strip.SourceControl, ListBox)
            End If
        End If
        ' Fallback 邏輯
        If targetList Is Nothing Or PathComboBox.Text.Contains("ASMR") Then
            Return FileCollection
        Else
            Return targetList
        End If
    End Function

    Public Sub Clipboard_SetText(sender, e)
        Dim menuItem As ToolStripMenuItem = TryCast(sender, ToolStripMenuItem)
        Dim control As ListBox = GetControlWithFallback(menuItem)
        Dim allString = ""
        Dim tar_array = control.SelectedItems.Cast(Of String).ToArray
        If CheckBox5.Checked Then
            tar_array = control.Items.Cast(Of String).ToArray
        End If
        If tar_array.Count = 0 Then Return
        MessageLabel.Text = $"已複製 {tar_array.Count} 個項目"
        If tar_array.Count = 1 Then

            Clipboard.SetText(tar_array(0))
        Else
            For Each item In tar_array
                allString += item & vbCrLf
            Next
            Clipboard.SetText(allString)
        End If
        '加1個把數組轉成換行文字

    End Sub

    Private Sub Clipboard_SetIImage()
        '加1個把數組轉成換行文字
        Clipboard.SetData(DataFormats.Bitmap, PictureBox1.Image)
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Form1_Close(sender, New CancelEventArgs)
        FileCollection.Items.Clear()
        CheckedListBox1_setCheck("Auto Save", False)
        '----------造成嚴重bug--------
        PathComboBox.Text = ""
        'OpenFileDialog1.FileName = ""
        Form3.FileNameTextBox.Text = ""
        Form3.DirectoryNameTextBox.Text = ""
        '--------------

        Refresh_listbox_numbers()
        dataHasRead = False
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Process.Start("https://www.buymeacoffee.com/b200077")
    End Sub


    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        SpecialButton.Checked = True
        Select Case ComboBox7.Text
            Case "Synchronize"
                SubFileCollection.Items.Clear()
                For Each foundFile As String In EverythingSearcher.SearchFiles(
                        TargetComboBox.Text, "")
                    SubFileCollection.Items.Add(foundFile)
                Next
                For Each foundFile As String In EverythingSearcher.SearchFiles(
                        PathComboBox.Text, "")
                    FileCollection.Items.Add(foundFile)
                Next
            Case "計算", "自動分類", "imgToPdf", "sendtoDL"
                CheckBox5.Checked = True
            Case Else
                CheckBox5.Checked = False
        End Select
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Form2.Show()
        Call Form2.WebMode("https://www.google.com/")
        WmcRefresh()
    End Sub


    Private Sub GetOutputDevices() Handles DeviceComboBox.Click
        'If ComboBox8.Items.Count <> 0 Then Return
        Dim originselectitem As String = DeviceComboBox.SelectedItem
        DeviceComboBox.Items.Clear()
        If _mediaPlayer.Media IsNot Nothing Then
            Dim nowplayfile As String = New Uri(_mediaPlayer.Media.Mrl).LocalPath
            Dim nowposition As Long = _mediaPlayer.Time
            _mediaPlayer.Dispose()
            _mediaPlayer = Nothing
            _mediaPlayer = New MediaPlayer(_libvlc)
            AddHandler _mediaPlayer.EndReached, AddressOf MediaPlayer_EndReached
            AddHandler _mediaPlayer.TimeChanged, AddressOf OnTimeChanged
            Dim media As New Media(_libvlc, nowplayfile, FromType.FromPath)
            ' 加載字幕
            _mediaPlayer.Media = media
            'addsubtitle(curitem)
            VideoView1.MediaPlayer = _mediaPlayer
            _mediaPlayer.Play()
            _mediaPlayer.Time = nowposition

        End If
        For Each device In _mediaPlayer.AudioOutputDeviceEnum
            DeviceComboBox.Items.Add(device.Description)
            Console.WriteLine($"   子設備: {device.DeviceIdentifier} - {device.Description}")
        Next
        If String.IsNullOrEmpty(originselectitem) Then
            DeviceComboBox.SelectedIndex = 0
        Else
            DeviceComboBox.SelectedItem = originselectitem
        End If
    End Sub
    ' 設置指定的音訊輸出設備

    Private Sub DeviceComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DeviceComboBox.SelectedIndexChanged
        Dim deviceIdentifier As String = _mediaPlayer.AudioOutputDeviceEnum(DeviceComboBox.SelectedIndex).DeviceIdentifier ' 獲取設備標識符
        _mediaPlayer.SetOutputDevice(deviceIdentifier)
    End Sub

    Private Sub MeasureDevice(deviceName As String)
        If DeviceComboBox.SelectedItem.contains(deviceName) Then Return
        If DeviceComboBox.FindString(deviceName) = -1 Then Return
        DeviceComboBox.SelectedIndex = DeviceComboBox.FindString(deviceName)
    End Sub
    Private Sub Form1_UnActiveControl(sender As Object, e As EventArgs)
        ActiveControl = Nothing
    End Sub
    Private Sub PlayOrPause() Handles Button24.Click
        'Console.WriteLine("State change")
        '
        If _mediaPlayer.Media Is Nothing Then Return
        Select Case _mediaPlayer.Media.State
            Case VLCState.Playing
                _mediaPlayer.Pause()
                Button24.Text = "播放"
            Case VLCState.Paused, VLCState.Stopped
                _mediaPlayer.Play()
                Button24.Text = "暫停"
        End Select
    End Sub
    Public Sub Form1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = Chr(32) Then
            If Me.ActiveControl Is Nothing Then
                PlayOrPause()
            Else
                If Me.ActiveControl.Name <> Button24.Name Then ' 空白鍵的 ASCII 代碼是 32
                    PlayOrPause()
                End If
            End If
        End If


    End Sub
    Private Sub RichTextBox__KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox1.KeyDown
        If e.Control And e.KeyCode = Keys.S Then
            RichTextBox_saveText()
        End If
    End Sub
    Public Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If Me.ActiveControl Is RichTextBox1 Then Return
        Select Case e.KeyCode
            Case Keys.End
                FileCollection.SetSelected(FileCollection.SelectedIndex, False)
                FileCollection.SetSelected(FileCollection.Items.Count - 1, True)
            Case Keys.Delete
                If TypeOf Me.ActiveControl Is ComboBox OrElse TypeOf Me.ActiveControl Is TextBox Then
                    Return
                End If
                Remove_Button(sender, New EventArgs)
            Case Keys.Down
                'If 目錄檔案集合.CausesValidation Then Return
                Dim newIndex = targetList.SelectedIndex + 1
                If newIndex = targetList.Items.Count Then Return
                targetList.SelectedIndex = newIndex
                If Not e.Control And Not e.Shift Then
                    targetList.SetSelected(targetList.SelectedIndex, False)
                End If
            Case Keys.Up
                'If 目錄檔案集合.CausesValidation Then Return
                Dim newIndex = targetList.SelectedIndex - 1
                If newIndex = -1 Then Return
                If Not e.Control And Not e.Shift Then
                    targetList.SetSelected(targetList.SelectedIndex, False)
                End If
                targetList.SetSelected(newIndex, True)
            Case Keys.Left
                If Not videoPanel.Visible Then Return
                _mediaPlayer.Time -= 5000
            Case Keys.Right
                If Not videoPanel.Visible Then Return
                _mediaPlayer.Time += 5000
        End Select
        '判斷是否按下Control+方向上下
        If Not e.Control Then Return
        ' 檢查是否正在輸入 ComboBox 或 TextBox
        If TypeOf Me.ActiveControl Is ComboBox OrElse TypeOf Me.ActiveControl Is TextBox Then
            Return
        End If
        Select Case e.KeyCode
            Case Keys.D
                Button9_Click(sender, New EventArgs)
            'Case Keys.R
            '    Button11_Click_1(sender, New EventArgs)
            Case Keys.A
                CheckBox5.Checked = Not CheckBox5.Checked
            Case Keys.Enter
                OnButton4Click(sender, New EventArgs)
            Case Keys.Z

                listUndo()
            Case Keys.C
                Clipboard_SetText(targetList, e)
            Case Keys.X
                Clipboard_SetText(targetList, e)
                Remove_Button(sender, New EventArgs)
            Case Keys.V
                'If targetList Is Nothing Or PathComboBox.Text.Contains("ASMR") Then 
                targetList = FileCollection
                AddItem(Setdata_from_clipboard())
            Case Keys.End
                listRearrange(FileCollection.Items.Count)
        End Select
    End Sub
    Private Sub RichTextBox_saveText()
        If CheckedListBox1_isCheck("鎖定") Then
            MsgBox("請解除鎖定再操作")
            Return
        End If
        Dim savefile As String = targetList.SelectedItem.ToString
        Try
            Using writer As New StreamWriter(savefile, False)
                writer.Write(RichTextBox1.Text)
            End Using
            MsgBox("文件已保存")
        Catch ex As Exception
            MsgBox("保存失敗" & ex.Message)
        End Try
        'RichTextBox1.SaveFile(targetList.SelectedItem.ToString, RichTextBoxStreamType.PlainText)

    End Sub



    Private Sub Button19_Click_1(sender As Object, e As EventArgs) Handles Button19.Click
        'Form5.Show()
        tryBrowser.Show()
    End Sub



    Private Sub CheckBox11_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox11.CheckedChanged
        ListBox_Click(sender, e)
    End Sub



    Public Sub AddLink(item As String, Optional islabel As Boolean = False)

        '先判斷在列表中的有無標記
        Dim compare As String
        If targetList.Items.Contains(item) Then
            compare = "nolabel"
        ElseIf targetList.Items.Contains("★" + item) Then
            compare = "haslabel"
        Else
            compare = "noContain"
        End If
        item = item.Replace(Environment.NewLine, "")
        If compare = "haslabel" Then item = "★" + item
        If Not CheckedListBox1_isCheck("重複移到最底") Then
            If compare = "noContain" Or CheckedListBox1_isCheck("取消自動刪除重複") Then
                targetList.Items.Add(item)
                backup.Add(item)
            Else
                PasswordBox.Items.Add($"{item}已加入{targetList.Items.IndexOf(item) + 1}項")
                MessageLabel.Text = $"{item}已加入{targetList.Items.IndexOf(item) + 1}項"
                Form2.Label3.Text = $"{item}已加入{targetList.Items.IndexOf(item) + 1}項"
            End If
        Else

            Dim selected_item As String = targetList.SelectedItem
            targetList.Items.Remove(item)
            backup.Remove(item)
            targetList.Items.Add(item)
            backup.Add(item)
            If targetList.SelectedIndex = -1 Then
                targetList.SelectedItem = selected_item
            End If
        End If
        Refresh_backup()
    End Sub
    Public Sub InsertLink(targets As String())
        ' 需要插入選定索引的下一項
        Dim selectedItem As String = targetList.SelectedItem.ToString
        Dim oriForm As List(Of String) = targetList.Items.Cast(Of String).ToList()
        ' 使用 HashSet 优化查找和去重操作
        Dim targetSet As New List(Of String)
        Dim subindex As Integer = 0
        For Each target In targets                          ' 如果插入的項目在index後面移除
            If oriForm.Contains(target) Then
                If oriForm.IndexOf(target) < oriForm.IndexOf(selectedItem) Then
                    subindex += 1
                End If
                oriForm.Remove(target)
            End If
        Next
        Dim pattern As New Regex("[Mm](-?\d+)")
        Dim match As Match = pattern.Match(orderComboBox.Text)
        Dim indexValue As Integer = 1
        If match.Success Then
            indexValue = Integer.Parse(match.Groups(1).Value)
            'Form1.ComboBox1.Text = ""
        End If
        Dim newIndex = targetList.Items.IndexOf(selectedItem) + indexValue - subindex
        ' 插入新項目
        targetSet.AddRange(oriForm.Take(newIndex))       ' 插入前面的部分
        targetSet.AddRange(targets)                          ' 插入目標項目
        targetSet.AddRange(oriForm.Skip(newIndex))      ' 插入後面的部分
        ' 更新顯示
        targetList.BeginUpdate()
        targetList.Items.Clear()
        targetList.Items.AddRange(targetSet.ToArray())
        targetList.EndUpdate()
        targetList.SetSelected(newIndex - indexValue, True)
        MessageLabel.Text = "已加入 " & targets.Length & " 個項目"
    End Sub

    Private Sub Quick_Rename(sender As Object, e As KeyEventArgs) Handles FileNameComboBox.KeyDown, ExtInput.KeyDown
        If e.KeyCode <> Keys.Enter Then Return
        Dim checkIndex As Integer = targetList.SelectedIndex
        Dim foundfile As String = targetList.SelectedItem
        targetList.Items(checkIndex) = Rename_Anothor_Thread(foundfile)
        Next_item()
    End Sub

    Public Function Make_Character_Valid(invalidString As String, ext As String)
        ' 替換檔案名稱中的特殊字元，例如inputjpg
        Dim validFilename = Regex.Replace(invalidString, "[\:\\\.]", "")
        ' 組合檔案名稱和副檔名，例如inputjpg.png
        validFilename = String.Format("{0}." & ext, validFilename)
        Return validFilename
    End Function


    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        RadioButton4.Checked = False
        RadioButton5.Checked = False
        RadioButton6.Checked = False
    End Sub



    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles orderComboBox.KeyDown, orderComboBox.KeyDown
        If e.KeyCode <> Keys.Enter Then Return
        selectOperate()
        MoveOperate()
        expansionOperate()
    End Sub



    Private Sub PasswordBox_DoubleClick(sender As Object, e As EventArgs) Handles PasswordBox.DoubleClick
        ' 檢查 ListBox 是否有項目
        If PasswordBox.Items.Count = 0 Then Return

        ' 指定一個暫存檔案
        Dim tempFilePath As String = Path.Combine(Path.GetTempPath(), "ListBoxItems.txt")

        ' 寫入所有項目到檔案
        Using writer As New StreamWriter(tempFilePath, False, Encoding.UTF8)
            For Each item In PasswordBox.Items
                writer.WriteLine(item.ToString())
            Next
        End Using

        ' 使用 Notepad 開啟檔案
        Process.Start("notepad.exe", tempFilePath)
    End Sub

    ' MouseMove 事件：檢查是否為拖曳


    Private Sub Form1_MouseWheel(sender As Object, e As MouseEventArgs) ' Handles Me.MouseWheel
        Dim ext As String = Path.GetExtension(targetList.SelectedItem)
        If Not Ext_judge(ext, "Image") Then Return
        If PictureBox1.ClientRectangle.Contains(PictureBox1.PointToClient(Cursor.Position)) Then
            If e.Delta > 0 Then
                Debug.Print("滚轮向上")
                sender = previousButton
            Else
                Debug.Print("滚轮向下")
                sender = nextButton
            End If
            Next_Page(sender, e)
        End If
    End Sub

    Private Sub SelectOperate()

        ' 使用一個正則表達式來同時匹配單個數字和數字區間（可選的第二個數字）
        's15 s14~25
        Dim pattern As New Regex("[sS](\d+)(?:\~(\d+))?")
        Dim match As Match = pattern.Match(orderComboBox.Text)

        If match.Success Then
            FileCollection.SelectedIndex = -1  ' 清除選擇

            Dim startIndex As Integer = Integer.Parse(match.Groups(1).Value)
            ' 如果有匹配到第二個數字，就當作區間處理
            If match.Groups(2).Success Then
                Dim endIndex As Integer = Integer.Parse(match.Groups(2).Value)
                For i As Integer = startIndex - 1 To endIndex - 1
                    FileCollection.SetSelected(i, True)
                Next
            Else
                FileCollection.SelectedIndex = Math.Min(startIndex - 1, FileCollection.Items.Count)
            End If
        End If
    End Sub
    Private Sub ExpansionOperate()
        'expandcount startnumber
        Dim pattern As New Regex("[eE](\d+)\s[Nn](\S+)")
        Dim match As Match = pattern.Match(orderComboBox.Text)
        If match.Success Then
            Dim expandCount As Integer = Integer.Parse(match.Groups(1).Value)
            Dim nowNumber As Integer = CInt(match.Groups(2).Value)
            Dim stringPattern As New Regex($"(.*){nowNumber}(.*)")
            Dim result As Match = stringPattern.Match(FileCollection.SelectedItem.ToString)
            Dim context As String() = {result.Groups(1).Value, result.Groups(2).Value}
            Dim expandList As New List(Of String)
            For i = nowNumber + 1 To nowNumber + expandCount - 1
                FileCollection.Items.Add(context(0) & i.ToString & context(1))
            Next
        End If
    End Sub

    Private Sub MoveOperate()
        Dim inputText As String = orderComboBox.Text

        Dim pattern1 As New Regex("[Mm](-?\d+)", RegexOptions.Compiled)
        Dim pattern2 As New Regex("[Mm][Tt](\d+)", RegexOptions.Compiled)

        Dim match1 As Match = pattern1.Match(inputText)
        Dim match2 As Match = pattern2.Match(inputText)

        If match2.Success Then
            Dim targetIndex As Integer = Integer.Parse(match2.Groups(1).Value)
            Dim currentIndex As Integer = FileCollection.SelectedIndex
            listRearrange(targetIndex - currentIndex - 1)

        ElseIf match1.Success Then
            Dim moveValue As Integer = Integer.Parse(match1.Groups(1).Value)
            Console.WriteLine(moveValue)
            listRearrange(moveValue)
        End If

        orderComboBox.Text = ""
    End Sub

    Dim savingTrf As Boolean = False
    Public Sub SaveTrf(dirpath As String, filename As String, append As Boolean, extformat As String)
        If String.IsNullOrEmpty(dirpath) Or String.IsNullOrEmpty(filename) Then Return
        Dim desk As String = dirpath
        'Dim ext_format As String = GroupBox4.Controls.OfType(Of RadioButton)().FirstOrDefault(Function(n) n.Checked).Text
        Dim record As String = desk & "\" & filename & extformat
        If Not record.EndsWith(".trf") Then Directory.CreateDirectory(record)
        'output_time += 1
        Dim lines As New List(Of String)
        Dim files As List(Of String) = FileCollection.Items.Cast(Of String).ToList
        If Not CheckedListBox1_isCheck("Combine trf files") Then
            lines.AddRange(files)
        Else
            For Each foundfile In files
                If Path.GetExtension(foundfile) = ".trf" Then
                    Trf_backup(foundfile)
                    Dim fileContent As String = File.ReadAllText(foundfile)
                    ' 將內容寫入合併後的檔案
                    lines.Add(fileContent)
                End If
            Next
        End If
        lines.Add("[selected]" & FileCollection.SelectedIndex.ToString)
        If Not CheckedListBox1_isCheck("ASMR播放模式") And SubFileCollection.SelectedIndex <> -1 Then
            lines.Add("[subselected]" & SubFileCollection.SelectedIndex.ToString)
        End If
        If Form2.Visible And Form2.PictureBox1.Visible Then lines.Add("[Form2_open]")
        Try
            Using sw As New StreamWriter(record, append)
                sw.WriteLine(String.Join(Environment.NewLine, lines))
            End Using
        Catch ex As Exception
            Console.WriteLine(record & "輸出失敗" & ex.Message)
        End Try

        If String.IsNullOrEmpty(SearchTextBox.Text) Then
            If PathComboBox.Text <> record Then
                '-----------導致速度變慢--------------
                PathComboBox.Text = record
            End If
        End If
        If Form3.Visible Then
            MsgBox(record & "輸出完成")
            Form3.Close()
        End If
    End Sub
    Public Function PathRename()
        Dim filename As String
        Dim directoryName As String
        If Not File.Exists(PathComboBox.Text) Then Return False
        directoryName = Path.GetDirectoryName(PathComboBox.Text)
        'If OriginalNameRadioButton.Checked Then filename = Path.GetFileNameWithoutExtension(OpenFileDialog1.FileName)
        filename = Path.GetFileNameWithoutExtension(PathComboBox.Text)
        If Not String.IsNullOrEmpty(SearchTextBox.Text) Then

            directoryName = commonUsed
            filename = SearchTextBox.Text
            directoryName = Path.Combine(directoryName, "個別網站")
        End If
        If CheckedListBox1_isCheck("Combine trf files") Then

            directoryName = commonUsed
            filename = DateTime.Now.ToString("yyyy_MM_dd")
            filename += "網頁整合"
        End If
        Return {directoryName, filename}
    End Function

    Private Sub FileCollection_MouseMove(sender As Object, e As MouseEventArgs) Handles FileCollection.MouseMove, SubFileCollection.MouseMove, TargetComboBox.MouseMove, PathComboBox.MouseMove
        'If multiArranging Then Return
        If e.Button <> MouseButtons.Left Then Return
        dragSourceControl = DirectCast(sender, Control)  ' 記錄來源
        If TypeOf sender Is ListBox Then
            'If sender.items.count = 0 Then Return
            'If sender.SelectedItems.count > 1 Then Return
            Dim lb As ListBox = DirectCast(sender, ListBox)
            '' 把螢幕座標轉換成控制項內的座標
            Dim pt As Point = lb.PointToClient(New Point(e.X, e.Y))
            Dim fileList As String() = lb.SelectedItems.Cast(Of String).
    Select(Function(x) Remove_label(x)).ToArray()
            '' 判斷滑鼠是否在控制項範圍內
            If sender Is FileCollection Then File_Preview(True)
            If sender Is SubFileCollection Then Archive_Preview(sender, e)
            'If fileCollection.ClientRectangle.Contains(pt) Then targetList = fileCollection
            'If subFileCollection.ClientRectangle.Contains(pt) Then targetList = subFileCollection
            Dim dataObj As New DataObject()
            Dim isAllFile As Boolean = True
            For Each item In fileList
                If Not File.Exists(item) Then
                    isAllFile = False
                    Exit For
                End If
            Next
            Console.WriteLine($"isAllFile={isAllFile}")
            ' 檢查是否為檔案路徑
            If isAllFile Then
                For i = 0 To fileList.Count - 1
                    ' 拖曳檔案
                    Dim filename As String = Path.GetFileNameWithoutExtension(fileList(i))
                    Dim ext As String = Path.GetExtension(fileList(i))
                    Dim fn As Match = dlsiteRgx.Match(fileList(i))
                    Console.WriteLine(fn.Success)
                    If fn.Success Then
                        Dim tempFolder As String = Path.Combine(Path.GetTempPath(), "RenamedDragDrop")
                        Directory.CreateDirectory(tempFolder)
                        filename = fn.Value + "_" + filename + ext
                        Dim newPath = Path.Combine(tempFolder, filename)
                        File.Copy(fileList(i), newPath, True)
                        fileList(i) = newPath
                    End If
                Next
                dataObj.SetData(DataFormats.FileDrop, fileList)
            Else
                ' 拖曳純文字
                Console.WriteLine(String.Join(Environment.NewLine, fileList))
                dataObj.SetData(DataFormats.Text, String.Join(Environment.NewLine, fileList))
            End If
            Dim control As String = GroupBox2.Controls.OfType(Of RadioButton)().FirstOrDefault(Function(n) n.Checked).Name
            Select Case control
                Case "MoveButton"
                    sender.DoDragDrop(dataObj, DragDropEffects.Move)
                Case Else
                    sender.DoDragDrop(dataObj, DragDropEffects.Copy)
            End Select

        End If
        If TypeOf sender Is ComboBox Then
            ' 获取 ComboBox 的位置和大小
            Dim comboBoxRect As New Rectangle(sender.ClientRectangle.Location, sender.ClientRectangle.Size)
            ' 获取下拉箭头区域的宽度
            Dim dropDownButtonWidth As Integer = SystemInformation.VerticalScrollBarArrowHeight ' 下拉按钮宽度的近似值
            If e.X > comboBoxRect.Width - dropDownButtonWidth Then Return

            Dim selectedText As String = If(sender.SelectedItem, sender.SelectedItem.ToString(), "")

            Dim dataObj As New DataObject()

            ' 檢查是否為檔案或資料夾路徑
            If File.Exists(selectedText) Or Directory.Exists(selectedText) Then
                ' 拖曳檔案
                dataObj.SetData(DataFormats.FileDrop, selectedText)
            Else
                ' 拖曳純文字
                dataObj.SetData(DataFormats.Text, selectedText)
            End If
            sender.DoDragDrop(dataObj, DragDropEffects.Copy)
        End If
        dragSourceControl = Nothing  ' 拖曳結束後清除
    End Sub
    Private Sub Form1_scale() Handles Me.Resize
        '6:7
        If Not CheckedListBox1_isCheck("隱藏預覽") Then
            Dim proportion As Double = 11.0 / 26.0
            Panel12.Size = New Size(CInt(Size.Width * proportion), CInt(Panel12.Size.Height))
            Panel9.Size = New Size(CInt(Panel12.Size.Width), CInt(Panel9.Size.Height))
            Panel5.Size = New Size(CInt(Panel16.Size.Width), CInt(Panel5.Size.Height))
        Else
            Panel12.Size = New Size(CInt(Size.Width - Panel6.Size.Width - 20), CInt(Panel12.Size.Height))
            Panel9.Size = New Size(CInt(Panel12.Size.Width), CInt(Panel9.Size.Height))
            Panel5.Size = New Size(CInt(Panel16.Size.Width), CInt(Panel5.Size.Height))
        End If
    End Sub
    Private dragSourceControl As Control = Nothing
    Private Sub Form1_Index_DragDrop(sender As ListBox, e As DragEventArgs) Handles FileCollection.DragDrop, SubFileCollection.DragDrop
        Dim lb As ListBox = DirectCast(sender, ListBox) ' 接收拖曳的 ListBox
        ' 指定正確的目標
        targetList = lb
        ' 支援拖曳檔案
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            AddItem(files)
            ' 支援拖曳 URL
        ElseIf e.Data.GetDataPresent(DataFormats.Text) Then
            Dim url As String = CType(e.Data.GetData(DataFormats.Text), String)

            AddItem(url)
            ' 只接受 http(s) 開頭的連結
            'If url.StartsWith("http://") OrElse url.StartsWith("https://") Then

            'End If
        End If

    End Sub

    Public Function Setdata_from_clipboard()
        Dim items As New List(Of String)

        If Clipboard.ContainsFileDropList Then
            Dim files As Specialized.StringCollection = Clipboard.GetFileDropList
            For Each clipfile As String In files
                items.Add(clipfile)
            Next

        ElseIf Clipboard.GetText.Contains(vbCrLf) Then
            Dim text As String = Clipboard.GetText
            For Each line In text.Split({vbCrLf}, StringSplitOptions.None)
                Dim trimmed As String = line.Trim()
                If String.IsNullOrEmpty(trimmed) Then Continue For
                items.Add(trimmed)
            Next

        Else
            Dim trimmed As String = Clipboard.GetText.Trim
            If Not String.IsNullOrEmpty(trimmed) Then
                items.Add(trimmed)
            End If
        End If

        Return items.ToArray()
    End Function

    Private dragfrom = Nothing

    Private Sub Form1_DragEnter(sender As Object, e As DragEventArgs) Handles FileCollection.DragEnter, SubFileCollection.DragEnter, PathComboBox.DragEnter, TargetComboBox.DragEnter
        If sender Is dragSourceControl Then
            e.Effect = DragDropEffects.None
            Return
        End If
        If e.Data.GetDataPresent(DataFormats.FileDrop) OrElse e.Data.GetDataPresent(DataFormats.Text) Then
            Dim control As String = GroupBox2.Controls.OfType(Of RadioButton)().FirstOrDefault(Function(n) n.Checked).Name
            Select Case control
                Case "MoveButton"
                    e.Effect = DragDropEffects.Move
                Case Else
                    e.Effect = DragDropEffects.Copy
            End Select

        End If
    End Sub
End Class


Public Class FolderNameEditor
    Inherits UITypeEditor

    Public Overrides Function GetEditStyle(ByVal context As ITypeDescriptorContext) As UITypeEditorEditStyle
        Return UITypeEditorEditStyle.Modal
    End Function

    Public Overrides Function EditValue(ByVal context As ITypeDescriptorContext, ByVal provider As IServiceProvider, ByVal value As Object) As Object
        Dim browser As New FolderBrowserDialog()
        If value IsNot Nothing Then
            browser.DirectoryPath = String.Format("{0}", value)
        End If
        If browser.ShowDialog(Nothing) = DialogResult.OK Then
            Return browser.DirectoryPath
        End If
        Return value
    End Function
End Class

''' <summary>Vista 样式的选择文件对话框的基类</summary>
<Description("提供一个Vista样式的选择文件对话框")>
<Editor(GetType(FolderNameEditor), GetType(UITypeEditor))>
Public Class FolderBrowserDialog
    Inherits Component

    <DllImport("shell32.dll")>
    Private Shared Function SHILCreateFromPath(<MarshalAs(UnmanagedType.LPWStr)> ByVal pszPath As String, ByRef ppIdl As IntPtr, ByRef rgflnOut As UInteger) As Integer
    End Function
    <DllImport("shell32.dll")>
    Private Shared Function SHCreateShellItem(ByVal pidlParent As IntPtr, ByVal psfParent As IntPtr, ByVal pidl As IntPtr, ByRef ppsi As IShellItem) As Integer
    End Function
    <DllImport("user32.dll")>
    Private Shared Function GetActiveWindow() As IntPtr
    End Function
    Private Const ERROR_CANCELLED As UInteger = &H800704C7UI

    ''' <summary>初始化 FolderBrowser 的新实例</summary>
    Public Sub New()
    End Sub

    Private mainFrm As Form
    ''' <summary>初始化 FolderBrowser 的新实例</summary>
    ''' <param name="frm">依附的主窗体</param>
    Public Sub New(ByVal frm As Form)
        If frm IsNot Nothing Then mainFrm = frm
    End Sub

    ''' <summary>依附的主窗体</summary>
    Property MainForm As Form = mainFrm

    ''' <summary>获取在 FolderBrowser 中选择的文件夹路径</summary>
    Public Property DirectoryPath() As String

    ''' <summary>向用户显示 FolderBrowser 的对话框</summary>
    Public Function ShowDialog() As DialogResult
        Return ShowDialog(mainFrm)
    End Function
    ''' <summary>向用户显示 FolderBrowser 的对话框</summary>
    ''' <param name="owner">任何实现 System.Windows.Forms.IWin32Window（表示将拥有模式对话框的顶级窗口）的对象。</param>
    Public Function ShowDialog(ByVal owner As IWin32Window) As DialogResult
        Dim hwndOwner As IntPtr = If(owner IsNot Nothing, owner.Handle, GetActiveWindow())
        Dim dialog As IFileOpenDialog = CType(New FileOpenDialog(), IFileOpenDialog)
        Try
            Dim item As IShellItem = Nothing
            If Not String.IsNullOrEmpty(DirectoryPath) Then
                Dim idl As IntPtr
                Dim atts As UInteger = 0
                If SHILCreateFromPath(DirectoryPath, idl, atts) = 0 Then
                    If SHCreateShellItem(IntPtr.Zero, IntPtr.Zero, idl, item) = 0 Then
                        dialog.SetFolder(item)
                    End If
                End If
            End If
            dialog.SetOptions(FOS.FOS_PICKFOLDERS Or FOS.FOS_FORCEFILESYSTEM)
            Dim hr As UInteger = dialog.Show(hwndOwner)
            If hr = ERROR_CANCELLED Then
                Return DialogResult.Cancel
            End If

            If hr <> 0 Then
                Return DialogResult.Abort
            End If
            dialog.GetResult(item)
            Dim path As String = Nothing
            item.GetDisplayName(SIGDN.SIGDN_FILESYSPATH, path)
            DirectoryPath = path
            Return DialogResult.OK
        Finally
            Marshal.ReleaseComObject(dialog)
        End Try
    End Function

    <ComImport()>
    <Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")>
    Private Class FileOpenDialog
    End Class

    <ComImport()>
    <Guid("42f85136-db7e-439c-85f1-e4075d135fc8")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface IFileOpenDialog
        <PreserveSig()>
        Function Show(<[In]()> ByVal parent As IntPtr) As UInteger ' IModalWindow
        Sub SetFileTypes() ' not fully defined
        Sub SetFileTypeIndex(<[In]()> ByVal iFileType As UInteger)
        Sub GetFileTypeIndex(ByRef piFileType As UInteger)
        Sub Advise() ' not fully defined
        Sub Unadvise()
        Sub SetOptions(<[In]()> ByVal fos As FOS)
        Sub GetOptions(ByRef pfos As FOS)
        Sub SetDefaultFolder(ByVal psi As IShellItem)
        Sub SetFolder(ByVal psi As IShellItem)
        Sub GetFolder(ByRef ppsi As IShellItem)
        Sub GetCurrentSelection(ByRef ppsi As IShellItem)
        Sub SetFileName(<[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal pszName As String)
        Sub GetFileName(<MarshalAs(UnmanagedType.LPWStr)> ByRef pszName As String)
        Sub SetTitle(<[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal pszTitle As String)
        Sub SetOkButtonLabel(<[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal pszText As String)
        Sub SetFileNameLabel(<[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal pszLabel As String)
        Sub GetResult(ByRef ppsi As IShellItem)
        Sub AddPlace(ByVal psi As IShellItem, ByVal alignment As Integer)
        Sub SetDefaultExtension(<[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal pszDefaultExtension As String)
        Sub Close(ByVal hr As Integer)
        Sub SetClientGuid()
        ' not fully defined
        Sub ClearClientData()
        Sub SetFilter(<MarshalAs(UnmanagedType.[Interface])> ByVal pFilter As IntPtr)
        Sub GetResults(<MarshalAs(UnmanagedType.[Interface])> ByRef ppenum As IntPtr) ' not fully defined
        Sub GetSelectedItems(<MarshalAs(UnmanagedType.[Interface])> ByRef ppsai As IntPtr) ' not fully defined
    End Interface

    <ComImport()>
    <Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface IShellItem
        Sub BindToHandler() ' not fully defined
        Sub GetParent() ' not fully defined
        Sub GetDisplayName(<[In]()> ByVal sigdnName As SIGDN, <MarshalAs(UnmanagedType.LPWStr)> ByRef ppszName As String)
        Sub GetAttributes() ' not fully defined
        Sub Compare() ' not fully defined
    End Interface
    Private Enum SIGDN As UInteger
        SIGDN_DESKTOPABSOLUTEEDITING = &H8004C000UI
        SIGDN_DESKTOPABSOLUTEPARSING = &H80028000UI
        SIGDN_FILESYSPATH = &H80058000UI
        SIGDN_NORMALDISPLAY = 0
        SIGDN_PARENTRELATIVE = &H80080001UI
        SIGDN_PARENTRELATIVEEDITING = &H80031001UI
        SIGDN_PARENTRELATIVEFORADDRESSBAR = &H8007C001UI
        SIGDN_PARENTRELATIVEPARSING = &H80018001UI
        SIGDN_URL = &H80068000UI
    End Enum

    <Flags()>
    Private Enum FOS
        FOS_ALLNONSTORAGEITEMS = &H80
        FOS_ALLOWMULTISELECT = &H200
        FOS_CREATEPROMPT = &H2000
        FOS_DEFAULTNOMINIMODE = &H20000000
        FOS_DONTADDTORECENT = &H2000000
        FOS_FILEMUSTEXIST = &H1000
        FOS_FORCEFILESYSTEM = &H40
        FOS_FORCESHOWHIDDEN = &H10000000
        FOS_HIDEMRUPLACES = &H20000
        FOS_HIDEPINNEDPLACES = &H40000
        FOS_NOCHANGEDIR = 8
        FOS_NODEREFERENCELINKS = &H100000
        FOS_NOREADONLYRETURN = &H8000
        FOS_NOTESTFILECREATE = &H10000
        FOS_NOVALIDATE = &H100
        FOS_OVERWRITEPROMPT = 2
        FOS_PATHMUSTEXIST = &H800
        FOS_PICKFOLDERS = &H20
        FOS_SHAREAWARE = &H4000
        FOS_STRICTFILETYPES = 4
    End Enum
End Class


Public NotInheritable Class Constants
    Private Shared ReadOnly data As Dictionary(Of String, String())

    Shared Sub New()
        ' 讀取 .INI 文件
        Dim filePath As String = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            "Documents", "easyPreview", "paramaters.ini"
        )
        data = ReadIniFile(filePath)
    End Sub

    ' 對外提供共用函數，依 key 拿值
    Public Shared Function GetValue(key As String) As String()
        If data.ContainsKey(key) Then
            Return data(key)
        Else
            Return New String() {}
        End If
    End Function

    Private Shared Function ReadIniFile(filePath As String) As Dictionary(Of String, String())
        Dim d As New Dictionary(Of String, String())()
        For Each line As String In File.ReadLines(filePath, Encoding.UTF8)
            If String.IsNullOrWhiteSpace(line) OrElse line.StartsWith(";") Then Continue For

            Dim parts As String() = line.Split(New String() {"="}, StringSplitOptions.RemoveEmptyEntries)
            If parts.Length = 2 Then
                Dim key As String = parts(0).Trim()
                Dim values As String() =
                    parts(1).Trim().Trim("{"c, "}"c).
                    Split(New String() {""","}, StringSplitOptions.None).
                    Select(Function(x) x.Trim().Trim(""""c)).ToArray()
                d(key) = values
            End If
        Next
        Return d
    End Function
End Class

Module NamedPipeUtils
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Structure WIN32_FIND_DATA
        Public dwFileAttributes As UInteger
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=8)>
        Public ftCreationTime As Byte()
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=8)>
        Public ftLastAccessTime As Byte()
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=8)>
        Public ftLastWriteTime As Byte()
        Public nFileSizeHigh As UInteger
        Public nFileSizeLow As UInteger
        Public dwReserved0 As UInteger
        Public dwReserved1 As UInteger
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
        Public cFileName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=14)>
        Public cAlternateFileName As String
    End Structure

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Function FindFirstFile(lpFileName As String, ByRef lpFindFileData As WIN32_FIND_DATA) As IntPtr
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Function FindNextFile(hFindFile As IntPtr, ByRef lpFindFileData As WIN32_FIND_DATA) As Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Function FindClose(hFindFile As IntPtr) As Boolean
    End Function

    Public Function ListNamedPipes(Optional prefix As String = "") As List(Of String)
        Dim pipes As New List(Of String)()
        Dim findData As New WIN32_FIND_DATA()
        Dim pipePath As String = "\\.\pipe\*"
        Dim hFind As IntPtr = FindFirstFile(pipePath, findData)

        If hFind = IntPtr.Zero Then
            Return pipes ' 如果找不到任何命名管道，返回空列表
        End If

        Try
            Do
                Dim pipeName As String = findData.cFileName
                If Not String.IsNullOrEmpty(pipeName) AndAlso pipeName.StartsWith(prefix) Then
                    pipes.Add(pipeName)
                End If
            Loop While FindNextFile(hFind, findData)
        Finally
            FindClose(hFind)
        End Try

        Return pipes
    End Function
End Module
