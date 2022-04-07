Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Drawing.Design
Imports System.IO
Imports System.IO.Compression
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports AngleSharp
Imports AngleSharp.Io
Imports DiscUtils.Iso9660
Imports Microsoft.VisualBasic.FileIO
Imports Microsoft.Win32
Imports Microsoft.WindowsAPICodePack.Shell
Imports NAudio.Wave
Imports SharpCompress.Archives
Imports SharpCompress.Archives.Rar
Imports SharpCompress.Archives.SevenZip
Imports SharpCompress.Common
Imports SharpCompress.Readers
Imports Spire.Pdf



''' <summary>FolderBrowser 的设计器基类</summary>
''' 
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




Public Class Form1
    Private Class MyControl
        Inherits Form
        Implements IMessageFilter

        Private InterfaceClassGuid As Guid = New Guid(&H4D1E55B2, &HF16F, &H11CF, &H88, &HCB, &H0, &H11, &H11, &H0, &H0, &H30)
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
            Dim DeviceBroadcastHeader As DEV_BROADCAST_DEVICEINTERFACE = New DEV_BROADCAST_DEVICEINTERFACE()
            DeviceBroadcastHeader.dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE
            DeviceBroadcastHeader.dbcc_size = CUInt(Marshal.SizeOf(DeviceBroadcastHeader))
            DeviceBroadcastHeader.dbcc_reserved = 0
            DeviceBroadcastHeader.dbcc_classguid = InterfaceClassGuid
            Dim pDeviceBroadcastHeader As IntPtr = IntPtr.Zero
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

    Dim textbox2_out As String

    Public doc As PdfDocument = New PdfDocument()
    Public instObj As Stream = Stream.Null
    Public backup As Array = Nothing
    Public archive As IArchive = Nothing
    Public backup_e As EventArgs = New EventArgs


    Private Sub RadioButton14_Click(sender As Object, e As EventArgs) Handles RadioButton14.Click
        'TextBox1.Text = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile).ToString & "\Downloads"
        ComboBox3.Text = "E:\Download"
        ChecklistBox_Click(sender, e)
    End Sub
    Private Sub RadioButton1_Click(sender As Object, e As EventArgs) Handles RadioButton1.Click
        ComboBox3.Text = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString
        ChecklistBox_Click(sender, e)
    End Sub
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.AllowDrop = True
        目錄檔案集合.AllowDrop = True
        ComboBox3.AllowDrop = True
        ComboBox4.AllowDrop = True
        load_passward()
        load_past_path()
        Add_additional_operation()
        backup_e = e
    End Sub
    Private Sub ASMR_Auto_Play(sender As Object, e As AxWMPLib._WMPOCXEvents_PlayStateChangeEvent) Handles AxWindowsMediaPlayer1.PlayStateChange
        If AxWindowsMediaPlayer1.playState = WMPLib.WMPPlayState.wmppsStopped Then
            If ComboBox5.Text = "自動重播(全部)" Then
                AxWindowsMediaPlayer1.settings.setMode("loop", False)
                '放完壓縮檔案目錄後切換到下一個目錄
                Dim list As ListBox = New ListBox
                If 壓縮檔案集合.Items.Count = 0 Then
                    list = 目錄檔案集合
                Else
                    list = 壓縮檔案集合
                End If
                If Not Find_Next_ASMR(sender, backup_e, list) And 目錄檔案集合.SelectedIndex <> 目錄檔案集合.Items.Count - 1 Then
                    If 壓縮檔案集合.Items.Count <> 0 Then
                        目錄檔案集合.SetSelected(目錄檔案集合.SelectedIndex + 1, True)
                        目錄檔案集合.SetSelected(目錄檔案集合.SelectedIndex, False)

                    End If
                Else
                    目錄檔案集合.SetSelected(目錄檔案集合.SelectedIndex, False)
                    目錄檔案集合.SetSelected(0, True)

                End If


                AxWindowsMediaPlayer1.Ctlcontrols.play()
            ElseIf ComboBox5.Text = "自動重播(單首)" Then
                AxWindowsMediaPlayer1.settings.setMode("loop", True)
                AxWindowsMediaPlayer1.Ctlcontrols.play()
            Else
                AxWindowsMediaPlayer1.settings.setMode("loop", False)
            End If
        End If
    End Sub
    Public Function Find_Next_ASMR(sender As Object, e As EventArgs, list As ListBox)
        For i = list.SelectedIndex + 1 To list.Items.Count - 1
            Dim ext = Path.GetExtension(list.Items(i)).ToLower()
            If Ext_judge(ext, "影片") And Not Ext_judge(ext, "exception") Then
                list.SetSelected(list.SelectedIndex, False)
                list.SetSelected(i + 1, True)

                Return True
            End If
        Next
        Return False
    End Function
    Private Sub Button1_Dialog(sender As Object, e As EventArgs) Handles Button1.Click
        Using fbd As New FolderBrowserDialog(Me)
            If fbd.ShowDialog = DialogResult.OK Then
                ComboBox3.Text = fbd.DirectoryPath
            End If
        End Using
        If ComboBox3.Text = "" Then
            ComboBox3.Text = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString
        End If
        RadioButton3.Checked = True
    End Sub

    Private Sub Button2_Dialog(sender As Object, e As EventArgs) Handles Button2.Click
        Using fbd As New FolderBrowserDialog(Me)
            If fbd.ShowDialog = DialogResult.OK Then
                ComboBox4.Text = fbd.DirectoryPath
            End If
        End Using
        If String.IsNullOrEmpty(ComboBox3.Text.ToString) Then
            ComboBox3.Text = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString
        End If
        RadioButton11.Checked = True
        If CheckBox2.Checked = True Then
            ComboBox4.Text += "/" & DateTime.Now.ToString("yyyy_MM_dd")

        End If
        textbox2_out = ComboBox4.Text
    End Sub

    Private Sub Checkbox7_Dialog(sender As Object, e As EventArgs) Handles Button6.Click
        OpenFileDialog1.Filter = "txt files (*.txt)|*.txt"
        OpenFileDialog1.ShowDialog()
        If Not OpenFileDialog1.FileName.Length = 0 Then
            Using sr As StreamReader = New StreamReader(OpenFileDialog1.FileName)

                While (sr.Peek() >= 0)
                    目錄檔案集合.Items.Add(sr.ReadLine())
                End While
            End Using
        End If
        backup = 目錄檔案集合.Items.Cast(Of String).ToArray
        Console.WriteLine(OpenFileDialog1.FileName)
        If OpenFileDialog1.FileName.Contains("影片") Then
            CheckBox1.Checked = True
        ElseIf OpenFileDialog1.FileName.Contains("畫") Then
            ComboBox5.SelectedItem = "clip啟動"
        End If
        refresh_listbox_item_numbers()
    End Sub
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        OpenFileDialog1.Filter = "All files (*.*)|*.*"
        OpenFileDialog1.Multiselect = True
        OpenFileDialog1.ShowDialog()
        For Each sr As String In OpenFileDialog1.FileNames
            目錄檔案集合.Items.Add(sr)
        Next

    End Sub
    Public Sub Output_file_label(sender As Object, e As EventArgs) Handles Button8.Click
        Form3.Visible = True

    End Sub
    Private Sub RadioButton2_Text(sender As Object, e As EventArgs) Handles RadioButton13.Click, RadioButton11.Click, CheckBox2.Click
        If RadioButton13.Checked = True Then
            ComboBox4.Text = ComboBox3.Text

        ElseIf RadioButton11.Checked = True Then

            If ComboBox3.Text.Length = 0 Then
                Using fbd As New FolderBrowserDialog(Me)
                    If fbd.ShowDialog = DialogResult.OK Then
                        MsgBox(fbd.DirectoryPath)
                    End If
                End Using
            End If
            ComboBox4.Text = ComboBox3.Text
        End If
        If CheckBox2.Checked = True Then
            ComboBox4.Text += "\" & DateTime.Now.ToString("yyyy_MM_dd")

        End If
        textbox2_out = ComboBox4.Text
    End Sub
    Private Sub Find_Episode(sender As Object, e As EventArgs) Handles Button18.Click, Button19.Click
        If 目錄檔案集合.SelectedItem Is Nothing Then
            Return
        End If
        Dim rgx As Regex = New Regex("(.*\W)(\d{1}|\d{2})(\W.*)")
        Dim episode As Integer = Integer.Parse(rgx.Match(目錄檔案集合.SelectedItem).Groups(2).Value)
        Console.WriteLine(episode.ToString)
        Dim btn_name As String = CType(sender, Button).Name
        If btn_name = "Button18" Then
            episode += 1
        ElseIf btn_name = "Button19" Then
            episode -= 1
        End If
        For Each files In Directory.GetFiles(Path.GetDirectoryName(目錄檔案集合.SelectedItem))
            If episode < 10 Then
                If rgx.Match(files).Groups(2).Value = episode Then
                    目錄檔案集合.Items(目錄檔案集合.SelectedIndex) = files
                    file_Preview(sender, e)
                    Exit For
                ElseIf files.Contains("0" & episode.ToString) Then  '補0
                    目錄檔案集合.Items(目錄檔案集合.SelectedIndex) = files
                    file_Preview(sender, e)
                    Exit For
                End If
            Else
                If rgx.Match(files).Groups(2).Value = episode Then
                    目錄檔案集合.Items(目錄檔案集合.SelectedIndex) = files
                    file_Preview(sender, e)
                    Exit For
                End If
            End If

        Next
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        Dim cal As String = "intersection"
        Dim keyWord As String = TextBox5.Text
        If Not String.IsNullOrEmpty(TextBox5.Text) Then
            If TextBox5.Text(0) = "-" Then
                cal = "subsection"
                keyWord = keyWord.Remove(0, 1)
            End If
        End If
        Dim new_array As ArrayList = New ArrayList()

        For Each i As String In backup
            If i.Contains(keyWord) And cal = "intersection" Then
                new_array.Add(i.ToString)
            End If
            If Not i.Contains(keyWord) And cal = "subsection" Then
                new_array.Add(i.ToString)
            End If
        Next
        目錄檔案集合.Items.Clear()
        For Each i In new_array
            目錄檔案集合.Items.Add(i)
        Next
        refresh_listbox_item_numbers()
    End Sub
    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        Dim obj As ListBox = Nothing
        Dim new_array As ArrayList = New ArrayList()
        If ComboBox5.SelectedItem = "取樣" Then
            If CheckBox11.Checked = True Then
                obj = 壓縮檔案集合
            Else
                obj = 目錄檔案集合
            End If

            For i = 0 To obj.Items.Count - 2
                Dim item = Path.GetFileNameWithoutExtension(obj.Items(i).ToString)
                Dim rgx As Regex = New Regex(item.Substring(0, item.Length - 1) & ".")
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
        Else

        End If

    End Sub
    Private Async Sub ChecklistBox_Click(sender As Object, e As EventArgs) Handles Button1.Click, RadioButton3.Click
        副檔名列表.Items.Clear()
        Dim sOption As System.IO.SearchOption
        If CheckBox3.Checked = True Then
            sOption = FileIO.SearchOption.SearchAllSubDirectories
        Else
            sOption = FileIO.SearchOption.SearchTopLevelOnly
        End If
        Dim rgx As Regex = New Regex("http")
        If Not rgx.IsMatch(ComboBox3.Text) Then
            If CheckBox10.Checked = True Then
                For Each foundFile As String In My.Computer.FileSystem.GetFiles(ComboBox3.Text, searchOptionCheck)
                    Dim ext As String = Path.GetExtension(foundFile)
                    Dim number As Integer = 0
                    For Each 副檔名 As String In 副檔名列表.Items
                        If 副檔名 = ext Then
                            Exit For
                        End If
                        number += 1
                    Next
                    If number = 副檔名列表.Items.Count Then
                        副檔名列表.Items.Add(ext)
                    End If
                Next
            Else
                副檔名列表.Items.Add("圖片")
                副檔名列表.Items.Add("影音")
                副檔名列表.Items.Add("壓縮檔")
                副檔名列表.Items.Add("光盤")
                副檔名列表.Items.Add("遊戲")
                副檔名列表.Items.Add("文件")
            End If
        Else
            Dim my_headers = "cookie': 'FANBOXSESSID=22147253_jU1IYcPHBaY2HLN5HQj4w1DknppbKiwI'"
            Dim httpClient As System.Net.Http.HttpClient = New System.Net.Http.HttpClient()

            Dim responseMessage = Await httpClient.GetAsync(ComboBox3.Text)


            If (responseMessage.StatusCode = System.Net.HttpStatusCode.OK) Then

                Dim responseResult As String = responseMessage.Content.ReadAsStringAsync().Result

                'Console.WriteLine(responseResult)
                Dim config = Configuration.Default.WithDefaultLoader().WithDefaultCookies
                Dim context = BrowsingContext.[New](config)
                Dim document = context.OpenAsync(
                                     Sub(Res)
                                         Res.Content(responseResult).Address(ComboBox3.Text).Header(HeaderNames.SetCookie, my_headers)
                                     End Sub).Result
                Dim celle = document.QuerySelector("img").GetAttribute("src")
                Console.WriteLine(document.TextContent)

                'QuerySelector(".entry-content")找出class="entry-content"的所有元素

                目錄檔案集合.Items.Add(celle)

            End If

            'Console.ReadKey()
        End If
        'load passward

    End Sub
    Public Function load_passward()
        Dim path1 = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\Documents\passward.txt"
        If File.Exists(path1) And ListBox1.Items.Count = 0 Then
            Using sr As StreamReader = New StreamReader(File.OpenRead(path1))

                While (sr.Peek() >= 0)
                    Try
                        ListBox1.Items.Add(sr.ReadLine())
                    Catch ex As Exception
                        Console.WriteLine(ex.Message)
                    End Try
                End While
            End Using
        End If
        Return True
    End Function
    Public Sub load_past_path()
        Dim ori_path = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\Documents\ori_pathrecord.txt"
        Dim des_path = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\Documents\des_pathrecord.txt"
        If File.Exists(ori_path) Then
            Using sr As StreamReader = New StreamReader(File.OpenRead(ori_path))

                While (sr.Peek() >= 0)
                    Try
                        ComboBox3.Items.Add(sr.ReadLine())
                    Catch ex As Exception
                        Console.WriteLine(ex.Message)
                    End Try
                End While
            End Using
        End If
        If File.Exists(des_path) Then
            Using sr As StreamReader = New StreamReader(File.OpenRead(des_path))

                While (sr.Peek() >= 0)
                    Try
                        ComboBox4.Items.Add(sr.ReadLine())
                    Catch ex As Exception
                        Console.WriteLine(ex.Message)
                    End Try
                End While
            End Using
        End If
    End Sub
    Public Function searchOptionCheck()
        Dim sOption As System.IO.SearchOption

        If CheckBox3.Checked = True Then
            sOption = FileIO.SearchOption.SearchAllSubDirectories
        Else
            sOption = FileIO.SearchOption.SearchTopLevelOnly
        End If
        Return sOption
    End Function
    Private Function getallDirectories(path As String) As IEnumerable(Of String)

        Return System.IO.Directory.EnumerateDirectories(path).Union(System.IO.Directory.EnumerateDirectories(path).SelectMany(Function(d)

                                                                                                                                  Try
                                                                                                                                      Return FileSystem.GetDirectories(path)
                                                                                                                                  Catch e As UnauthorizedAccessException
                                                                                                                                      Return Enumerable.Empty(Of String)()
                                                                                                                                  End Try
                                                                                                                              End Function))
    End Function
    Private Function getallfiles(path As String) As IEnumerable(Of String)

        Try
            Return System.IO.Directory.EnumerateFiles(path).Union(System.IO.Directory.EnumerateDirectories(path).SelectMany(Function(d)

                                                                                                                                Try
                                                                                                                                    Return FileSystem.GetFiles(path, FileIO.SearchOption.SearchAllSubDirectories)
                                                                                                                                Catch e As UnauthorizedAccessException
                                                                                                                                    Return Enumerable.Empty(Of String)()
                                                                                                                                End Try
                                                                                                                            End Function))

        Catch
            Return Nothing
        End Try
    End Function
    Private Sub checkedListBox1_ItemCheck(ByVal sender As Object, ByVal e As ItemCheckEventArgs) Handles 副檔名列表.ItemCheck
        Me.BeginInvoke(CType(Function()
                                 ListBox_Click(sender, e)
                                 Return True
                             End Function, MethodInvoker))
    End Sub

    Private Sub ListBox_Click(sender As Object, e As EventArgs)
        目錄檔案集合.Items.Clear()
        壓縮檔案集合.Items.Clear()
        Dim result = False
        If CheckBox11.Checked = True Then
            For Each cont As String In 副檔名列表.CheckedItems
                Dim Directories As IEnumerable = getallDirectories(ComboBox3.Text)
                For Each foundDirectory As String In Directories
                    Dim files As IEnumerable = getallfiles(foundDirectory)
                    If files Is Nothing Then
                        Continue For
                    End If
                    If FileSystem.GetFiles(foundDirectory, FileIO.SearchOption.SearchAllSubDirectories).Count = 0 Then
                        FileSystem.DeleteDirectory(foundDirectory, DeleteDirectoryOption.ThrowIfDirectoryNonEmpty)
                        Continue For
                    End If
                    result = True
                    If cont = "遊戲" Or cont = "影音" Then
                        result = False
                    End If
                    For Each foundFile As String In files
                        Console.WriteLine(foundFile)
                        Dim ext As String = Path.GetExtension(foundFile).ToLower()
                        If Ext_judge(ext, cont) And cont = "遊戲" Then
                            result = True
                        ElseIf Ext_judge(ext, cont) And cont = "影音" Then
                            result = True
                        ElseIf Not Ext_judge(ext, cont) Then
                            If cont = "影音" Or cont = "遊戲" Then
                            Else
                                result = False
                            End If
                        End If
                        Console.WriteLine(result)
                    Next
                    Dim directoryName = Path.GetFileName(Path.GetDirectoryName(foundDirectory))
                    If result = True And Not 目錄檔案集合.Items.Contains(directoryName) Then
                        目錄檔案集合.Items.Add(foundDirectory)
                    End If
                Next

            Next
        Else
            For Each cont As Object In 副檔名列表.CheckedItems
                If CheckBox10.Checked = True Then
                    For Each foundFile As String In My.Computer.FileSystem.GetFiles(
                     ComboBox3.Text, searchOptionCheck(), "*" & cont)
                        目錄檔案集合.Items.Add(foundFile)
                    Next
                Else
                    For Each foundFile As String In My.Computer.FileSystem.GetFiles(
                     ComboBox3.Text, searchOptionCheck())
                        Dim ext As String = Path.GetExtension(foundFile).ToLower()
                        If Ext_judge(ext, cont) Then
                            目錄檔案集合.Items.Add(foundFile)
                        End If
                    Next

                End If

            Next
        End If
        backup = 目錄檔案集合.Items.Cast(Of String).ToArray
        If Not String.IsNullOrEmpty(TextBox5.Text) Then
            TextBox5_TextChanged(sender, e)
        End If
        refresh_listbox_item_numbers()
    End Sub
    Public Sub refresh_listbox_item_numbers()
        Label13.Text = 目錄檔案集合.Items.Count.ToString & "項"
    End Sub
    Private Sub Ext_select(sender As Object, e As EventArgs) Handles CheckBox6.Click

        For sel = 0 To 副檔名列表.Items.Count - 1
            If CheckBox6.Checked = True Then
                副檔名列表.SetItemChecked(sel, True)
            Else
                副檔名列表.SetItemChecked(sel, False)
            End If
        Next
    End Sub
    Private Function Ext_judge(ext As String, ext_type As String)
        Dim result = False
        If ext IsNot Nothing Then
            ext = ext.ToLower
        End If

        '容許.url .txt檔
        If CheckBox11.Checked = True And ext_type IsNot "exception" Then
            If ext = ".url" Or ext = ".txt" Then
                result = True
            End If
        End If
        If ext_type = "exception" Then
            If ext = ".url" Or ext = ".txt" Then
                result = True
            End If
        ElseIf ext_type = "圖片" Then
            If ext = ".jpg" Or ext = ".png" Or ext = ".jfif" Or ext = ".gif" Or ext = ".jpeg" Or Ext_judge(ext, "其他圖片") Then
                result = True
            End If
            If Ext_judge(ext, "遊戲") And CheckBox11.Checked Then
                result = False
            End If
        ElseIf ext_type = "其他圖片" Then
            If ext = ".clip" Or ext = ".psd" Then
                result = True
            End If
        ElseIf ext_type = "ASMR" Then
            If Ext_judge(ext, "圖片") Or Ext_judge(ext, "影音") Then
                result = True
            End If
        ElseIf ext_type = "影音" Then
            If ext = ".mp4" Or ext = ".mkv" Or ext = ".mp3" Or ext = ".mid" Or ext = ".wav" Or ext = ".wmv" Or ext = ".flac" Or ext = ".flv" Then
                result = True
            End If
        ElseIf ext_type = "壓縮檔" Then
            If ext = ".rar" Or ext = ".7z" Or ext = ".zip" Then
                result = True
            End If
        ElseIf ext_type = "光盤" Then
            If ext = ".mdf" Or ext = ".iso" Or ext = ".mds" Then
                result = True
            End If
        ElseIf ext_type = "遊戲" Then
            If ext = ".exe" Then
                result = True
            End If

        ElseIf ext_type = "文件" Then
            If ext = ".pdf" Or Ext_judge(ext, "文字") Or Ext_judge(ext, "簡報") Or Ext_judge(ext, "word檔") Then
                result = True
            End If
        ElseIf ext_type = "文字" Then
            If ext = ".txt" Or ext = ".ini" Or ext = ".inf" Or Ext_judge(ext, "字幕") Then
                result = True
            End If
        ElseIf ext_type = "字幕" Then
            If ext = ".ass" Or ext = ".srt" Then
                result = True
            End If
        ElseIf ext_type = "簡報" Then
            If ext = ".xls" Or ext = ".xlsx" Then
                result = True
            End If
        ElseIf ext_type = "word檔" Then
            If ext = ".doc" Or ext = ".docx" Then
                result = True
            End If
        ElseIf ext_type = "目錄" Then
            If Directory.Exists(ext) = True Then
                result = True
            End If
        ElseIf ext_type = "程序" Then
            If RadioButton15.Checked = True Then
                result = True
            End If
        ElseIf ext_type = ext Then
            result = True
        End If
        Return result

    End Function
    Public Sub Next_Page(sender As Object, e As EventArgs) Handles Button10.Click
        Dim rgx As Regex = New Regex("頁數：(\d+)/")
        Dim now_page As Integer = rgx.Match(Label10.Text).Groups(1).Value
        If doc.Pages.Count > now_page Then
            now_page += 1
            Dim bmp As Image = doc.SaveAsImage(now_page - 1)
            Label10.Text = "頁數：" & now_page.ToString & "/" & doc.Pages.Count.ToString

            PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
            PictureBox1.Image = CType(bmp, Image)
        End If
    End Sub
    Public Sub Privious_Page(sender As Object, e As EventArgs) Handles Button3.Click
        Dim rgx As Regex = New Regex("頁數：(\d+)/")
        Dim now_page As Integer = rgx.Match(Label10.Text).Groups(1).Value

        If now_page > 1 Then
            now_page -= 1
            Dim bmp As Image = doc.SaveAsImage(now_page - 1)
            Label10.Text = "頁數：" & now_page.ToString & "/" & doc.Pages.Count.ToString


            PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
            PictureBox1.Image = CType(bmp, Image)
        End If
    End Sub
    Private Sub Add_new_password(sender As Object, e As EventArgs) Handles Button5.Click
        ListBox1.Items.Add(TextBox4.Text)
        Dim desk As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\Documents"
        Dim record As String = desk & "\passward.txt"
        If File.Exists(record) = False Then
            Dim fs As FileStream = File.Create(record)
            fs.Close()
        End If
        Using sr As StreamWriter = New StreamWriter(record, True)
            sr.WriteLine(TextBox4.Text)
        End Using
        TextBox4.Text = ""
    End Sub
    Private Sub Add_new_path()
        If ComboBox3.Items.Contains(ComboBox3.Text) = False And Not String.IsNullOrEmpty(ComboBox3.Text) Then
            ComboBox3.Items.Add(ComboBox3.Text)
        End If
        If ComboBox4.Items.Contains(ComboBox4.Text) = False And Not String.IsNullOrEmpty(ComboBox4.Text) Then
            ComboBox4.Items.Add(ComboBox4.Text)
        End If
        Dim ori_path As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\Documents\ori_pathrecord.txt"
        Dim des_path As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\Documents\des_pathrecord.txt"
        File.Delete(ori_path)
        File.Delete(des_path)
        Using fs As FileStream = File.Create(ori_path)
            fs.Close()
        End Using
        Using fs As FileStream = File.Create(des_path)
            fs.Close()
        End Using
        Using sr As StreamWriter = New StreamWriter(ori_path)
            For Each path As String In ComboBox3.Items
                sr.WriteLine(path)
            Next
        End Using
        Using sr As StreamWriter = New StreamWriter(des_path)
            For Each path As String In ComboBox4.Items
                sr.WriteLine(path)
            Next
        End Using

    End Sub
    Private Sub viewAllFilesInIso(Directorie As Object)

        For Each CDentry As DiscUtils.DiscFileInfo In Directorie.GetFiles
            壓縮檔案集合.Items.Add(CDentry.FullName.Replace(";1", "").ToLower)
        Next
        For Each Directories In Directorie.GetDirectories
            viewAllFilesInIso(Directories)
        Next
    End Sub
    Private Sub Delete_password(sender As Object, e As EventArgs) Handles Button7.Click
        ListBox1.Items.Remove(ListBox1.SelectedItem)

        Dim desk As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\Documents"
        Dim record As String = desk & "\passward.txt"
        FileSystem.DeleteFile(record)
        If Not File.Exists(record) Then
            Dim fs As FileStream = File.Create(record)
            fs.Close()
        End If
        Using sr As StreamWriter = New StreamWriter(record, True)
            For Each item In ListBox1.Items
                sr.WriteLine(item)
            Next
        End Using


    End Sub
    Private Sub file_Preview_open_directory(sender As Object, e As EventArgs) Handles 目錄檔案集合.DoubleClick
        If Directory.Exists(目錄檔案集合.SelectedItem) Then
            ComboBox3.Text = 目錄檔案集合.SelectedItem
            CheckBox11.Checked = False
            RadioButton3.Checked = True
            ChecklistBox_Click(sender, e)
            ListBox_Click(sender, e)
        End If
    End Sub
    Public Function FindFirstBitmap(sender As Object, e As EventArgs, list As ListBox)
        For i = 0 To list.Items.Count - 1
            Dim ext = Path.GetExtension(list.Items(i)).ToLower()
            If Ext_judge(ext, "圖片") And Not Ext_judge(ext, "exception") And Not Ext_judge(ext, "其他圖片") Then
                壓縮檔案集合.SelectedIndex = i
                Archive_Preview(sender, e)
                Exit For
            End If
        Next
        Return True
    End Function
    Public Sub file_Preview(sender As Object, e As EventArgs) Handles 目錄檔案集合.SelectedIndexChanged

        If Not (目錄檔案集合.Items.Count = 0 Or 目錄檔案集合.SelectedItems.Count = 0) Then

            If AxWindowsMediaPlayer1.Visible = True Then
                AxWindowsMediaPlayer1.Hide()
                AxWindowsMediaPlayer1.Controls.Clear()
            End If
            PictureBox1.Image = New Bitmap(1, 1)
            Me.DataGridView1.Hide()
            If CheckBox1.Checked Then
                Return
            End If
            Dim sel_ind As Integer = ComboBox2.SelectedIndex
            ComboBox2.Items.Clear()
            Dim curitem As String = 目錄檔案集合.SelectedItem
            If Not File.Exists(curitem) Then
                Dim name = Path.GetFileName(curitem)
                Dim last_folder = Path.GetFileName(Path.GetDirectoryName(curitem))
                Console.WriteLine(last_folder)
                Dim desk = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString
                Console.WriteLine(Path.GetFileName(desk))
                If last_folder = Path.GetFileName(desk) Then

                    For Each folder In Directory.EnumerateDirectories(desk)
                        Console.WriteLine(folder & "\" & name)
                        '改成可選擇搜尋到的結果
                        If File.Exists(folder & "\" & name) Then
                            目錄檔案集合.Items(目錄檔案集合.SelectedIndex) = folder & "\" & name
                            curitem = folder & "\" & name

                        End If
                    Next
                Else
                    Dim draw_path = "E:\(畫\2022整合"
                    Console.WriteLine(draw_path & "\" & last_folder & "\" & name)
                    '改成可選擇搜尋到的結果
                    If File.Exists(draw_path & "\" & last_folder & "\" & name) Then
                        目錄檔案集合.Items(目錄檔案集合.SelectedIndex) = draw_path & "\" & last_folder & "\" & name
                        curitem = draw_path & "\" & last_folder & "\" & name

                    End If

                End If

            End If
            Dim ext As String = Path.GetExtension(curitem).ToLower()
            ComboBox2.Text = Path.GetFileNameWithoutExtension(curitem)
            TextBox3.Text = ext
            Dim last_write As String = ""
            Dim rgx As Regex = New Regex("http")
            If Not rgx.IsMatch(curitem) Then
                last_write = File.GetLastWriteTime(curitem).ToString("yyyy_MM_dd-HHmmss")
            End If

            Dim installs() As String = New String() _
            {ComboBox2.Text,
             last_write,
             "Custom"}
            ' ComboBox2.Text.Substring(4, ComboBox2.Text.Length - 4),

            ComboBox2.Items.AddRange(installs)
            ComboBox2.SelectedIndex = sel_ind
            If Ext_judge(curitem, "目錄") Then
                壓縮檔案集合.Items.Clear()
                Dim sOption = searchOptionCheck()
                For Each foundFile As String In My.Computer.FileSystem.GetFiles(
                        curitem, FileIO.SearchOption.SearchAllSubDirectories)

                    壓縮檔案集合.Items.Add(foundFile)

                Next

                FindFirstBitmap(sender, e, 壓縮檔案集合)
                ComboBox2.Text = Path.GetFileNameWithoutExtension(curitem)
            ElseIf Ext_judge(ext, "程序") Then
                壓縮檔案集合.Items.Clear()
                For Each files As String In Process.GetProcessById(目錄檔案集合.SelectedIndex).StandardInput.ToString
                    壓縮檔案集合.Items.Add(files)
                Next
            ElseIf Ext_judge(ext, "其他圖片") Then
                Try
                    Dim shellFile As ShellFile = ShellFile.FromFilePath(curitem)
                    Dim shellThumb As Bitmap = shellFile.Thumbnail.Bitmap
                    PictureBox1.Image = CType(shellThumb, Image)
                    If Form2.Visible = True Then
                        Form2.PictureBox1.Image = PictureBox1.Image
                    End If

                Catch ex As Exception
                    MsgBox("無法預覽檔案" & curitem & "無法開啟" & ex.GetType().FullName & ex.Message)
                End Try
            ElseIf Ext_judge(ext, "圖片") Then
                PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
                Dim MyImage As Bitmap = New Bitmap(curitem)
                Try

                    PictureBox1.Image = CType(MyImage.Clone, Image)

                    Form2.PictureBox1.Image = PictureBox1.Image.Clone

                Catch ex As Exception
                    MsgBox("無法預覽檔案" & curitem & "無法開啟" & ex.GetType().FullName & ex.Message)
                End Try
                MyImage.Dispose()


            ElseIf ext = ".pdf" Then

                doc.LoadFromFile(curitem)
                '遍歷PDF每一頁
                Dim bmp As Image = doc.SaveAsImage(0)

                Try
                    PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
                    PictureBox1.Image = CType(bmp, Image)
                    Label10.Text = "頁數：1/" & doc.Pages.Count.ToString
                    If Form2.Visible = True Then
                        Form2.PictureBox1.Image = PictureBox1.Image
                    End If
                Catch ex As Exception
                    MsgBox("無法預覽檔案" & curitem & "無法開啟" & ex.GetType().FullName & ex.Message)
                End Try
            ElseIf ext = ".xls" Or ext = ".xlsx" Then
                Me.DataGridView1.Show()
                Dim strConn As String = “Provider=Microsoft.ACE.OLEDB.12.0;” & “Data Source=” & curitem & “;” & “Extended Properties=Excel 12.0;”
                Dim conn As OleDbConnection = New OleDbConnection(strConn)
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
            ElseIf Ext_judge(ext, "影音") Then
                AxWindowsMediaPlayer1.Show()
                AxWindowsMediaPlayer1.URL = curitem
                AxWindowsMediaPlayer1.Ctlcontrols.play()
            ElseIf Ext_judge(ext, "文字") Then
                Dim fileReader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(curitem)
                Dim stringReader As String = fileReader.ReadToEnd
                Dim bmp = New Bitmap(PictureBox1.Width, PictureBox1.Height)
                Dim g = Drawing.Graphics.FromImage(bmp)
                g.DrawString(stringReader, New Font("新細明體", 9), Brushes.Black, New PointF(10, 10))
                PictureBox1.Image = bmp
                If Form2.Visible = True Then
                    Form2.PictureBox1.Image = PictureBox1.Image
                End If
                fileReader.Close()
            ElseIf Ext_judge(ext, ".mdf") Then
            ElseIf Ext_judge(ext, ".iso") Then
                壓縮檔案集合.Items.Clear()
                Using isoStream As FileStream = New FileStream(curitem, FileMode.Open)
                    Dim cd As CDReader = New CDReader(isoStream, True)
                    For Each CDentry As DiscUtils.DiscFileInfo In cd.Root.GetFiles()
                        壓縮檔案集合.Items.Add(CDentry.FullName.Replace(";1", "").ToLower)
                    Next
                    For Each Directories In cd.Root.GetDirectories
                        viewAllFilesInIso(Directories)
                    Next
                End Using
            ElseIf Ext_judge(ext, "壓縮檔") Then
                壓縮檔案集合.Items.Clear()
                Try
                    archive = ArchiveFactory.Open(curitem)
                    For Each entry As Object In archive.Entries
                        壓縮檔案集合.Items.Add(entry.key)
                    Next
                Catch
                    try_password(curitem, 0)
                End Try
                Task.WaitAll()
                FindFirstBitmap(sender, e, 壓縮檔案集合)
            End If
        End If

    End Sub
    Public Sub try_password(curitem As String, index As Integer)
        Try
            archive = ArchiveFactory.Open(curitem, New SharpCompress.Readers.ReaderOptions With {.Password = ListBox1.Items(index)})
            For Each entry As Object In archive.Entries
                壓縮檔案集合.Items.Add(entry.key)
            Next
        Catch ex As Exception
            If Not index = ListBox1.Items.Count - 1 Then
                try_password(curitem, index + 1)
            Else
                MsgBox("讀取檔案失敗" & curitem & "可能有密碼或密碼錯誤")
                archive.Dispose()
            End If
        End Try
    End Sub
    Public Sub Archive_Preview(sender As Object, e As EventArgs) Handles 壓縮檔案集合.SelectedIndexChanged
        If AxWindowsMediaPlayer1.Visible = True Then
            AxWindowsMediaPlayer1.Hide()
            AxWindowsMediaPlayer1.Controls.Clear()
        End If
        PictureBox1.Image = New Bitmap(1, 1)
        Dim ext As String = Path.GetExtension(目錄檔案集合.SelectedItem).Replace(";1", "").ToLower()
        Dim foundfile As String = Path.GetFullPath(目錄檔案集合.SelectedItem)
        Dim curitem As Stream = Stream.Null
        If Ext_judge(foundfile, "目錄") Then
            curitem = File.OpenRead(壓縮檔案集合.SelectedItem)
        ElseIf Ext_judge(ext, "光盤") Then
            Dim isoStream As FileStream = File.OpenRead(foundfile)
            Dim cd = New CDReader(isoStream, True).OpenFile(壓縮檔案集合.SelectedItem, FileMode.Open)
            cd.CopyTo(curitem)

        ElseIf Ext_judge(ext, "壓縮檔") Then
            curitem = archive.Entries(壓縮檔案集合.SelectedIndex).OpenEntryStream
        End If
        ext = Path.GetExtension(壓縮檔案集合.SelectedItem)
        TextBox3.Text = ext
        ComboBox2.Text = Path.GetFileNameWithoutExtension(壓縮檔案集合.SelectedItem)
        If Ext_judge(ext, "光盤") Then
            instObj = curitem
        ElseIf Ext_judge(ext, "圖片") Then
            Dim MyImage As Bitmap = New Bitmap(curitem)
            PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
            Try


                PictureBox1.Image = CType(MyImage.Clone, Image)
                If Form2.Visible = True Then
                    Form2.PictureBox1.Image = PictureBox1.Image
                End If

            Catch ex As Exception
                MsgBox("無法顯示圖片" & ex.GetType().FullName & ex.Message)
            End Try
            MyImage.Dispose()
            curitem.Dispose()
        ElseIf Ext_judge(ext, "影音") Then

            Dim path = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile).ToString & "\AppData\Local\Temp\temporVideo.mp4"
            If Not Ext_judge(foundfile, "目錄") Then
                archive.Entries(壓縮檔案集合.SelectedIndex).WriteToFile(path)
            Else
                path = 壓縮檔案集合.SelectedItem
            End If
            AxWindowsMediaPlayer1.Show()
            AxWindowsMediaPlayer1.URL = path
            AxWindowsMediaPlayer1.Ctlcontrols.play()
        ElseIf Ext_judge(ext, "文字") Then
            Dim tr As TextReader = New StreamReader(curitem)

            Dim stringReader As String = tr.ReadToEnd
            Dim bmp = New Bitmap(PictureBox1.Width, PictureBox1.Height)
            Dim g = Drawing.Graphics.FromImage(bmp)
            g.DrawString(stringReader, New Font("新細明體", 9), Brushes.Black, New PointF(10, 10))
            PictureBox1.Image = bmp

        End If
        If instObj Is Nothing Then
            Console.WriteLine("close curitem")
            curitem.Close()
            Task.WaitAll()
        End If
    End Sub
    Private Sub ExtrackAllFilesInIso(Directorie As DiscUtils.DiscDirectoryInfo, cd As Object)
        Dim directname = ComboBox4.Text & "\temporaryfile\" & "\" & Directorie.FullName
        If Not FileSystem.DirectoryExists(directname) Then
            FileSystem.CreateDirectory(directname)
        End If
        For Each CDentry In Directorie.GetFiles
            Label11.Text = "解壓" & CDentry.Name
            Dim newfile = File.Create(directname & "\" & CDentry.Name.Replace(";1", "").ToLower)
            Dim path As Stream = cd.OpenFile(CDentry.FullName, FileMode.Open)
            path.CopyTo(newfile)
            newfile.Close()
        Next
        For Each CDentry In Directorie.GetDirectories
            ExtrackAllFilesInIso(CDentry, cd)
        Next
    End Sub
    Public Function checkindexlast(selected_list As Array, maxItem As Integer)
        'selected_list need reverse
        If selected_list.Length = 0 Then
            Return Nothing
        ElseIf selected_list(0) = maxItem - 2 Then
            Array.Clear(selected_list, 0, 1)
            maxItem -= 1
            checkindexlast(selected_list, maxItem)
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

    Private Sub Button4_Action(sender As Object, e As EventArgs) Handles Button4.Click
        ''把picturebox1釋放
        Add_new_path()
        If archive IsNot Nothing And RadioButton12.Checked = False Then
            archive.Dispose()
        End If
        Form2.PictureBox1.Image = New Bitmap(1, 1)
        PictureBox1.Image = New Bitmap(1, 1)
        PictureBox1.Image.Dispose()
        PictureBox1.Image = Nothing
        If archive IsNot Nothing Then
            archive.Dispose()
        End If
        If Form2.Visible = True Then
            Form2.PictureBox1.Image = New Bitmap(1, 1)
            Form2.PictureBox1.Image = Nothing
            Form2.PictureBox1.Image.Dispose()
        End If

        If ComboBox4.Text.Length = 0 Then
            ComboBox4.Text = ComboBox3.Text
        End If
        Dim select_indeces = 副檔名列表.CheckedIndices.Cast(Of Integer)().ToArray()
        Dim selected_list = Nothing
        Dim item_copy = 目錄檔案集合.Items.Cast(Of String)().ToArray()
        If CheckBox5.Checked = False Then
            selected_list = 目錄檔案集合.SelectedIndices.Cast(Of Integer)().ToArray()
            selected_list = checkindexlast(selected_list, 目錄檔案集合.Items.Count)
            Console.WriteLine(selected_list)

            Dim newLabel = 目錄檔案集合.SelectedItems.Cast(Of String)().ToArray()
            目錄檔案集合.Items.Clear()
            For Each item In newLabel
                目錄檔案集合.Items.Add(item)
            Next
        End If
        If ComboBox5.SelectedItem = "紀錄開啟檔案" = True Then
            Call Output_file_label(sender, e)
        End If
        Dim arguments As String = ""
        Dim counter = -1
        For Each foundfile As String In 目錄檔案集合.Items.Cast(Of String).ToArray
            counter += 1
            Dim filename As String = Path.GetFileName(foundfile)
            Dim ext As String = Path.GetExtension(foundfile).ToLower
            Dim rename As String = Path.GetFileNameWithoutExtension(foundfile)
            Console.WriteLine(ComboBox2.SelectedIndex)
            Select Case ComboBox2.SelectedIndex
                Case 1
                    rename = File.GetLastWriteTime(foundfile).ToString("yyyy_MM_dd-HHmmss")
                Case 2
                    rename = rename.Substring(4, rename.Length - 4)
            End Select

            For Each act As RadioButton In GroupBox2.Controls
                If act.Checked = True Then
                    If act.Text = "開啟" Then

                        If ComboBox5.SelectedItem = "clip啟動" Then

                            arguments += $" ""{foundfile}"""

                        ElseIf ComboBox5.SelectedItem = "對上層執行動作" Then
                            Dim argument As String = "/select, """ & foundfile & """"
                            Process.Start("explorer.exe", argument)
                        ElseIf ComboBox5.SelectedItem = "locale emulator" Then
                            Dim exeName As String = foundfile
                            Dim p As ProcessStartInfo = New ProcessStartInfo()
                            Dim lePath As String = "C:\Locale.Emulator.2.5.0.1\LEProc.exe"
                            p.FileName = lePath
                            p.Arguments = $"-run ""{exeName}"""
                            p.UseShellExecute = False
                            p.WorkingDirectory = Path.GetDirectoryName(lePath)
                            Dim res As Process = Process.Start(p)
                            res.WaitForInputIdle(5000)
                        ElseIf ext = ".exe" Then
                            Try
                                Dim proinfo As ProcessStartInfo = New ProcessStartInfo With {
                    .FileName = foundfile,
                    .UseShellExecute = False,
                    .WorkingDirectory = Path.GetDirectoryName(foundfile),
                    .CreateNoWindow = True
                }

                                Dim prostart As Process = New Process With {
                        .StartInfo = proinfo
                    }
                                prostart.Start()
                            Catch
                                Dim proc As Process = New Process()
                                proc.StartInfo.FileName = foundfile
                                proc.StartInfo.UseShellExecute = True
                                proc.StartInfo.Verb = "runas"
                                proc.Start()
                            End Try

                        Else
                            Process.Start(foundfile)

                        End If


                    End If
                    If act.Text = "移動" Then


                        If CheckBox11.Checked = True Then
                            Dim directname As String = FileSystem.GetDirectoryInfo(foundfile).Name
                            Try
                                My.Computer.FileSystem.MoveDirectory(foundfile, ComboBox4.Text & "\" & directname, UIOption.AllDialogs)
                            Catch ex As Exception
                                MsgBox("移動檔案失敗" & filename & ex.GetType().FullName & ex.Message)
                            End Try
                        Else
                            Try
                                My.Computer.FileSystem.MoveFile(foundfile, ComboBox4.Text & "\" & filename, UIOption.AllDialogs)
                            Catch ex As Exception
                                MsgBox("移動檔案失敗" & filename & ex.GetType().FullName & ex.Message)
                            End Try
                        End If
                        'If 目錄檔案集合.Items.Count - 1 = selected_list(0) Then
                        'selected_list(0) = 目錄檔案集合.Items.Count - 2
                        'End If
                    End If
                    If act.Text = "複製" Then

                        My.Computer.FileSystem.CopyFile(foundfile, ComboBox4.Text & "\" & filename)
                    End If
                    If act.Text = "刪除" Then

                        Delete_Anothor_Thread(foundfile)
                        'If 目錄檔案集合.Items.Count - 1 = selected_list(0) Then
                        'selected_list(0) = 目錄檔案集合.Items.Count - 2
                        'End If
                    End If
                    If act.Text = "重新命名" Then


                        If ComboBox5.SelectedItem = "更改附檔名" = True Then
                            rename += TextBox3.Text
                        Else
                            rename += ext
                        End If
                        Try
                            Rename_Anothor_Thread(foundfile & "_" & rename)
                        Catch ex As IOException
                            '換執行緒
                            MsgBox("命名檔案失敗" & rename & "可能被開啟" & ex.Message)
                        Catch ex As UnauthorizedAccessException
                            MsgBox("命名檔案失敗" & rename & "需要系統權限")
                        End Try
                    End If
                    If act.Text = "安裝" Then
                        Dim filestream1 As FileStream = Nothing
                        If Ext_judge(ext, ".mdf") Then
                            open_mdf_file(foundfile)
                        ElseIf Ext_judge(ext, ".iso") Then
                            filestream1 = File.OpenRead(foundfile)
                        Else
                            filestream1 = File.Create(ComboBox4.Text & "file")
                            instObj.CopyTo(filestream1)
                            If 壓縮檔案集合.SelectedItem = "" Then
                                open_mdf_file(ComboBox4.Text & "file")
                            End If
                        End If
                        Using vhdStream As FileStream = filestream1

                            Dim cd As CDReader = New CDReader(vhdStream, True)
                            Dim pattern As String = "[AaUuTtOoRrUuNn]{7}\.[infINF]{3}"
                            Dim rgx As New Regex(pattern)
                            Dim infor As String = "AUTORUN.INF"
                            Dim newfile As FileStream
                            Dim path As Stream
                            Label11.Visible = True
                            Directory.CreateDirectory(ComboBox4.Text & "\temporaryfile")

                            For Each CDentry In cd.Root.GetFiles
                                Label11.Text = "解壓" & CDentry.Name
                                newfile = File.Create(ComboBox4.Text & "\temporaryfile\" & CDentry.Name.Replace(";1", "").ToLower)
                                path = cd.OpenFile(CDentry.FullName, FileMode.Open)
                                path.CopyTo(newfile)
                                newfile.Close()
                                For Each match As Match In rgx.Matches(CDentry.FullName.Replace(";1", "").ToLower)
                                    Console.WriteLine("Found '{0}' at position {1}", match.Value, match.Index)
                                    infor = match.Value
                                Next

                            Next
                            For Each Director1 In cd.Root.GetDirectories
                                ExtrackAllFilesInIso(Director1, cd)
                            Next
                            Console.WriteLine(infor)
                            path = cd.OpenFile(infor, FileMode.Open)
                            Dim tr As TextReader = New StreamReader(path)
                            Label11.Text = "尋找啟動程序.."
                            Dim sentence As String = tr.ReadToEnd
                            pattern = "(?<=open[\s]*=[\s]*)[\w.\\]+"
                            rgx = New Regex(pattern)
                            For Each match As Match In rgx.Matches(sentence)
                                Console.WriteLine("Found '{0}' at position {1}", match.Value, match.Index)
                                Process.Start(ComboBox4.Text & "\temporaryfile\" & match.Value)
                            Next

                        End Using

                        If ComboBox5.SelectedItem = "壓縮/解壓後刪除" Then
                            Try
                                My.Computer.FileSystem.DeleteFile(foundfile)
                            Catch ex As Exception
                                MsgBox("刪除檔案失敗" & filename & "可能已經刪除" & ex.GetType().FullName & ex.Message)
                            End Try
                        End If
                    End If
                    If act.Text = "壓縮" Then


                        Using zipToOpen As FileStream = New FileStream(foundfile, FileMode.Open)
                            Using archive As ZipArchive = New ZipArchive(zipToOpen, ZipArchiveMode.Update)
                                Dim readmeEntry As ZipArchiveEntry = archive.CreateEntry(foundfile)
                            End Using
                        End Using
                        If ComboBox5.SelectedItem = "壓縮/解壓後刪除" Then
                            Try
                                My.Computer.FileSystem.DeleteFile(foundfile)
                            Catch ex As Exception
                                MsgBox("刪除檔案失敗" & filename & "可能已經刪除" & ex.GetType().FullName & ex.Message)
                            End Try
                        End If
                    End If
                    If act.Text = "解壓" Then
                        Dim archive As IArchive = ArchiveFactory.Open(foundfile)
                        For Each entry In archive.Entries
                            If Not entry.IsDirectory Then
                                Console.WriteLine(entry.Key)
                                entry.WriteToDirectory(ComboBox4.Text & "\" & rename, New ExtractionOptions With
                              {.ExtractFullPath = True, .Overwrite = True})
                            End If
                        Next
                        archive.Dispose()
                        If ComboBox5.SelectedItem = "壓縮/解壓後刪除" = True Then
                            Try
                                My.Computer.FileSystem.DeleteFile(foundfile)
                            Catch ex As Exception
                                MsgBox("刪除檔案失敗" & filename & "可能已經刪除" & ex.GetType().FullName & ex.Message)
                            End Try
                        End If
                        MsgBox("解壓完成")
                    End If
                End If
            Next
        Next
        If ComboBox5.SelectedItem = "clip啟動" Then
            Console.WriteLine(arguments)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
            Dim app As String = "C:\Program Files\CELSYS\CLIP STUDIO 1.5\CLIP STUDIO PAINT\CLIPStudioPaint.exe"
            Dim proinfo As ProcessStartInfo = New ProcessStartInfo With {
            .FileName = app,
             .Arguments = arguments,
             .StandardErrorEncoding = System.Text.Encoding.UTF8,
            .UseShellExecute = False,
             .RedirectStandardError = True,
             .RedirectStandardOutput = True
               }

            Dim prostart As Process = New Process With {
             .StartInfo = proinfo
              }
            prostart.Start()

            '方案1 使用cmd啟動程式
            '方案2 修改副檔名啟動的登機碼
            'SetAssociation(ext, ext & "File", app, ext & " File")
            ' Process.Start(foundfile)
        End If
        If Not String.IsNullOrEmpty(ComboBox3.Text) Then
            Call ChecklistBox_Click(sender, e)
            For Each sel In select_indeces
                副檔名列表.SetItemChecked(sel, True)
            Next
            Call ListBox_Click(sender, e)
            If selected_list IsNot Nothing Then
                '讓listbox顯示到點選物件的page
                Try
                    目錄檔案集合.SetSelected(selected_list, True)
                Catch ex As Exception
                    目錄檔案集合.SetSelected(selected_list - 1, True)
                End Try

            End If
        Else
            目錄檔案集合.Items.Clear()
            For Each sel In item_copy
                目錄檔案集合.Items.Add(sel)
            Next
            If selected_list = 目錄檔案集合.Items.Count Then
                selected_list -= 1
            End If
            目錄檔案集合.SetSelected(selected_list, True)
        End If
    End Sub
    Public Function Rename_Anothor_Thread(foundfile_rename As String)
        If (Me.InvokeRequired) Then
            Dim del As DelShowMessage = New DelShowMessage(AddressOf Rename_Anothor_Thread)
            Me.Invoke(del, foundfile_rename)
        Else
            Dim subs As String() = foundfile_rename.Split("_"c)
            My.Computer.FileSystem.RenameFile(subs(0), subs(1))
        End If
        Return True
    End Function
    Public Function Delete_Anothor_Thread(foundfile As String)
        If (Me.InvokeRequired) Then
            Dim del As DelShowMessage = New DelShowMessage(AddressOf Delete_Anothor_Thread)
            Me.Invoke(del, foundfile)
        Else
            If CheckBox11.Checked = True Then
                Dim directname = FileSystem.GetDirectoryInfo(foundfile).FullName
                FileSystem.DeleteDirectory(directname, DeleteDirectoryOption.DeleteAllContents)
            Else
                Try
                    FileSystem.DeleteFile(foundfile)
                Catch ex As Exception
                    MsgBox("刪除檔案失敗" & foundfile & "可能已經刪除" & ex.GetType().FullName & ex.Message)
                End Try
            End If
        End If
        Return True
    End Function

    Private con As System.Data.SqlClient.SqlConnection

    Private Sub open_mdf_file(ByVal foundfile As Object)
        con = New System.Data.SqlClient.SqlConnection()
        con.ConnectionString = $"Data Source=.\SQLEXPRESS;
                          AttachDbFilename={foundfile};
                          Integrated Security=True;
                          Connect Timeout=30;
                          User Instance=True"
        con.Open()
        MessageBox.Show("Connection opened")
        con.Close()
        MessageBox.Show("Connection closed")
    End Sub
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
        OpenMethod.CreateSubKey("DefaultIcon").SetValue("", "\"" + OpenWith + " \ ",0")
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
   
    Public Function get_file_watch(path As String)
        Dim watcher = New FileSystemWatcher(path) With {
           .NotifyFilter = (NotifyFilters.LastAccess _
Or NotifyFilters.LastWrite _
Or NotifyFilters.FileName _
Or NotifyFilters.DirectoryName),
           .Filter = "*.*",
           .IncludeSubdirectories = True 'CheckBox3.Checked
           }
        ' Add event handlers.
        AddHandler watcher.Changed, AddressOf OnChanged
        AddHandler watcher.Created, AddressOf OnChanged
        AddHandler watcher.Deleted, AddressOf OnChanged
        AddHandler watcher.Renamed, AddressOf OnRenamed
        'add filter
        watcher.EnableRaisingEvents = True
        Return watcher
        ' Begin watching.

        '--------------------------------------------------------------------------
    End Function
    Private Sub RadioButton2_Watch(sender As Object, e As EventArgs) Handles RadioButton2.Click
        'Dim notePad As Process = Process.Start("C: \Program Files\CELSYS\CLIP STUDIO 1.5\CLIP STUDIO PAINT\CLIPStudioPaint.exe")
        ' DetectOpenFiles.GetOpenFilesEnumerator(notePad.Id)


        ' Watch for changes in LastAccess and LastWrite times, and
        ' the renaming of files or directories. 
        ' Only watch text files.
        Dim watcher1 = get_file_watch("C:\")
        Dim watcher2 = get_file_watch("E:\")
        Dim watcher3 = get_file_watch("F:\")
        '





    End Sub


    Public Class ThreadWithState
        ' State information used in the task.
        Private ReadOnly boilerplate As String
        Private ReadOnly numberValue As Integer

        ' Delegate used to execute the callback method when the
        ' task is complete.
        Private ReadOnly callback As ExampleCallback

        ' The constructor obtains the state information and the
        ' callback delegate.
        Public Sub New(text As String, number As Integer,
        callbackDelegate As ExampleCallback)
            boilerplate = text
            numberValue = number
            callback = callbackDelegate
        End Sub

        ' The thread procedure performs the task, such as
        ' formatting and printing a document, and then invokes
        ' the callback delegate with the number of lines printed.
        Public Sub ThreadProc()
            Console.WriteLine(boilerplate, numberValue)
            If Not (callback Is Nothing) Then
                callback(1)
            End If
        End Sub
    End Class
    Public Delegate Sub ExampleCallback(lineCount As Integer)
    Private Sub OnChanged(source As Object, e As FileSystemEventArgs)

    End Sub
    Public Shared Sub ResultCallback(lineCount As Integer)
        Console.WriteLine(
            "Independent task printed {0} lines.", lineCount)
    End Sub


    Private Delegate Function DelShowMessage(sMessage As String)
    Public Function AddItemListbox(s As String)
        If (Me.InvokeRequired) Then
            Dim del As DelShowMessage = New DelShowMessage(AddressOf AddItemListbox)
            Me.Invoke(del, s)
        Else
            If Not 目錄檔案集合.Items.Contains(s) And Not s.Contains("CELSYSUserData") Then
                目錄檔案集合.Items.Add(s)
            End If

        End If
        Return True
    End Function
    Public Function OnRenamed(source As Object, e As RenamedEventArgs)

        ' Specify what is done when a file is renamed.

        Dim ext As String = Path.GetExtension(e.FullPath).ToLower
        If Ext_judge(ext, "圖片") Then
            Console.WriteLine($"File: {e.OldFullPath} renamed to {e.FullPath}")
            Dim desk As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            Dim record As String = desk & "\" & DateTime.Now.ToString("yyyy_MM_dd") & ".txt"
            AddItemListbox(e.FullPath)

            ' Dim ori_content As String = ""
            ' Using sr As StreamReader = New StreamReader(record)
            '  ori_content = sr.ReadToEnd()
            '   End Using
            '  Using sr As StreamWriter = New StreamWriter(record)
            ' SR.WriteLine(ori_content & e.FullPath)
            'End Using

        End If

        Return True

    End Function

    Dim WithEvents Client As New System.Net.WebClient()
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        ProgressBar1.Visible = True
        ' Set Minimum to 1 to represent the first file being copied.
        ProgressBar1.Minimum = 1
        ProgressBar1.Maximum = 100
        Client.DownloadFileAsync(New Uri(ComboBox3.Text), ComboBox4.Text)

        ' Set Maximum to the total number of files to copy.

    End Sub
    Private Sub Client_DownloadProgressChanged(ByVal sender As Object, ByVal e As System.Net.DownloadProgressChangedEventArgs) Handles Client.DownloadProgressChanged '當Client正在下載時
        ProgressBar1.PerformStep()

    End Sub
    Private Sub Client_DownloadFileCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs) Handles Client.DownloadFileCompleted  '當Clien結束下載
        'e.Cancelled 判斷是否為中斷(取消)下載
        'e.Error '判斷下載過程是否因發生錯誤而停止下載
    End Sub

    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents Label9 As Label
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Button3 As Button
    Friend WithEvents Button10 As Button
    Friend WithEvents Label10 As Label
    Friend WithEvents RadioButton14 As RadioButton
    Friend WithEvents Label11 As Label
    Friend WithEvents Button11 As Button
    Friend WithEvents CheckBox10 As CheckBox
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents FolderBrowserDialog2 As Windows.Forms.FolderBrowserDialog
    Friend WithEvents CheckBox11 As CheckBox
    Friend WithEvents Button12 As Button
    Friend WithEvents CheckBox9 As CheckBox
    Friend WithEvents TextBox5 As TextBox
    Friend WithEvents TextBox6 As TextBox
    Friend WithEvents RadioButton15 As RadioButton
    Friend WithEvents Button13 As Button
    Friend WithEvents Button14 As Button
    Friend WithEvents Button15 As Button
    Friend WithEvents Button16 As Button
    Friend WithEvents Button17 As Button
    Friend WithEvents AxWindowsMediaPlayer1 As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents ComboBox3 As ComboBox
    Friend WithEvents ComboBox4 As ComboBox
    Friend WithEvents ComboBox5 As ComboBox
    Friend WithEvents Label12 As Label
    Friend WithEvents Button18 As Button
    Friend WithEvents Button19 As Button
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents Label13 As Label
End Class