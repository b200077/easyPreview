Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Drawing.Design
Imports System.Drawing.Imaging
Imports System.IO
Imports System.IO.Compression
Imports System.Net
Imports System.Reflection
Imports System.Runtime
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Menu
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports AngleSharp
Imports AngleSharp.Dom
Imports AngleSharp.Html.Dom
Imports AngleSharp.Io
Imports Aspose.Slides
Imports Aspose.Slides.Export.Web
Imports CefSharp
Imports CefSharp.Handler
Imports CefSharp.WinForms
Imports DiscUtils
Imports DiscUtils.Iso9660
Imports GLib
Imports ImageMagick
Imports Microsoft.UI.Xaml.Controls
Imports Microsoft.VisualBasic.FileIO
Imports Microsoft.Web.WebView2.WinForms
Imports Microsoft.Win32
Imports Microsoft.WindowsAPICodePack.Shell
Imports Microsoft.WindowsAPICodePack.Shell.PropertySystem
Imports NAudio.Wave
Imports Pango
Imports SharpCompress.Archives
Imports SharpCompress.Archives.Rar
Imports SharpCompress.Archives.SevenZip
Imports SharpCompress.Common
Imports SharpCompress.Compressors.Xz
Imports Spire.Pdf
Imports VisioForge.Core.VideoEdit.Timeline.Timeline
Imports VisioForge.Libs.MediaFoundation.OPM
Imports VisioForge.Libs.NDI


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

Public Class MyCustomMenuHandler
    Implements IContextMenuHandler
    Public Sub OnBeforeContextMenu(ByVal browserControl As IWebBrowser, ByVal browser As IBrowser, ByVal frame As IFrame, ByVal parameters As IContextMenuParams, ByVal model As IMenuModel) Implements IContextMenuHandler.OnBeforeContextMenu
        If model.Count > 0 Then
            model.AddSeparator()
        End If

        model.AddItem(CType(26501, CefMenuCommand), "Show DevTools")
        model.AddSeparator()
        model.AddItem(CType(26503, CefMenuCommand), "加入連結")
        model.AddItem(CType(26505, CefMenuCommand), "開啟連結")
        model.AddItem(CType(26504, CefMenuCommand), "儲存圖片")
        model.AddItem(CType(26507, CefMenuCommand), "複製連結")
        model.AddItem(CType(26506, CefMenuCommand), "取代連結")
        'model.AddItem(CType(26505, CefMenuCommand), "Run Script..")
    End Sub
    Public Sub SaveImage(ByVal imageUrl As String, ByVal filename As String, ByVal format As ImageFormat)
        Dim client As System.Net.WebClient = New WebClient()

        Dim stream As System.IO.Stream
        Try
            stream = client.OpenRead(imageUrl)
        Catch
            'If imageUrl.Contains("i.pximg.net") Then

            'End If
            Dim tarPanel As Panel = System.Windows.Forms.Application.OpenForms("Form2").Controls("Panel1")
            Dim tarList As ChromiumWebBrowser = tarPanel.Controls.OfType(Of ChromiumWebBrowser)().FirstOrDefault()
            Dim tarUrl As String = getBrowserUrl(tarList)
            client.Headers.Add("Referer", tarUrl)
            stream = client.OpenRead(imageUrl)
        End Try
        Dim bitmap As Bitmap
        bitmap = New Bitmap(stream)

        If bitmap IsNot Nothing Then
            bitmap.Save(filename, format)
        End If

        stream.Flush()
        stream.Close()
        client.Dispose()
    End Sub

    Private Delegate Function DelShowMessage(parameters As String, controls As ListBox)
    Private Delegate Function DelRefreshForm(tarForm As Form1)
    Private Delegate Function DelGetSelect(controls As ChromiumWebBrowser) As String
    Private Function getBrowserUrl(controls As ChromiumWebBrowser)
        If controls Is Nothing Then Return Nothing
        If (controls.InvokeRequired) Then
            Dim del As DelGetSelect = New DelGetSelect(AddressOf getBrowserUrl)
            controls.Invoke(del, controls)
        Else
            Return controls.Address
        End If
        Return Nothing
    End Function
    Public Function add_form1_list(parameters As String, controls As ListBox)
        If (controls.InvokeRequired) Then
            Dim del As DelShowMessage = New DelShowMessage(AddressOf add_form1_list)
            controls.Invoke(del, parameters, controls)
        Else
            If Not controls.Items.Contains(parameters) Then controls.Items.Add(parameters)
        End If
        Return True
    End Function
    Public Function replace_form1_item(parameters As String, controls As ListBox)
        If (controls.InvokeRequired) Then
            Dim del As DelShowMessage = New DelShowMessage(AddressOf replace_form1_item)
            controls.Invoke(del, parameters, controls)
        Else
            If controls.SelectedIndex = -1 Then Return False
            controls.Items(controls.SelectedIndex) = parameters
        End If
        Return True
    End Function
    Public Function form1_refresh(tarForm As Form1)
        If (tarForm.InvokeRequired) Then
            Dim del As DelRefreshForm = New DelRefreshForm(AddressOf form1_refresh)
            tarForm.Invoke(del, tarForm)
        Else
            tarForm.refresh_backup()
        End If
        Return True
    End Function
    Public Function getFileName(oriName As String)
        Dim rgx As Regex = New Regex("([^\/]*)$")
        Dim fn As String = rgx.Match(oriName).Groups(1).Value

        If fn.Contains("?") Then

            Dim subs As String() = fn.Split("?")
            If subs(0).Contains(".") Then
                fn = subs(0)
            Else
                fn = subs(0) & ".jfif"
            End If
        ElseIf fn.Contains(".") Then
            Dim subs As String() = fn.Split(".")
            fn = subs(0) & "." & subs(1).Split(":")(0)
        End If
        Return fn
    End Function
    Public Function OnContextMenuCommand(ByVal browserControl As IWebBrowser, ByVal browser As IBrowser, ByVal frame As IFrame, ByVal parameters As IContextMenuParams, ByVal commandId As CefMenuCommand, ByVal eventFlags As CefEventFlags) As Boolean Implements IContextMenuHandler.OnContextMenuCommand
        If commandId = CType(26503, CefMenuCommand) Then

            If (parameters.LinkUrl.Length > 0) Then
                '寫入輸入檔案
                Dim tarForm As Form1 = System.Windows.Forms.Application.OpenForms("Form1")
                add_form1_list(parameters.LinkUrl, tarForm.Controls("目錄檔案集合"))
                form1_refresh(tarForm)
            End If
        End If
        If commandId = CType(26505, CefMenuCommand) Then

            If (parameters.LinkUrl.Length > 0) Then
                '寫入輸入檔案
                System.Diagnostics.Process.Start(parameters.LinkUrl)
            End If
        End If
        If commandId = CType(26504, CefMenuCommand) And parameters.MediaType = ContextMenuMediaType.Image Then

            'Clipboard.SetText(parameters.SourceUrl)

            Dim save_dialog As SaveFileDialog = New SaveFileDialog
            save_dialog.FileName = getFileName(parameters.SourceUrl)
            Dim ext As String = save_dialog.FileName.ToString.Split(".")(1)
            save_dialog.Filter = ext & " files (*." & ext & ")|*." & ext & "|All files (*.*)|*.*"

            If save_dialog.ShowDialog = DialogResult.OK Then
                Select Case ext
                    Case "jpg"
                        SaveImage(parameters.SourceUrl, save_dialog.FileName, ImageFormat.Jpeg)
                    Case "jfif"
                        SaveImage(parameters.SourceUrl, save_dialog.FileName, ImageFormat.Jpeg)
                    Case "png"
                        SaveImage(parameters.SourceUrl, save_dialog.FileName, ImageFormat.Png)
                    Case "gif"
                        SaveImage(parameters.SourceUrl, save_dialog.FileName, ImageFormat.Gif)
                End Select

            End If





            ' Process.Start(fn)

        End If
        If commandId = CType(26506, CefMenuCommand) Then

            If (parameters.LinkUrl.Length > 0) Then
                '寫入輸入檔案
                Dim tarForm As Form1 = System.Windows.Forms.Application.OpenForms("Form1")
                replace_form1_item(parameters.LinkUrl, tarForm.Controls("目錄檔案集合"))
                form1_refresh(tarForm)
            End If
        End If
        If commandId = CType(26507, CefMenuCommand) Then
            Clipboard.SetText(parameters.LinkUrl)
        End If
        If commandId = CType(26501, CefMenuCommand) Then
            browser.ShowDevTools(Nothing, parameters.XCoord, parameters.YCoord)

        End If
        Return False
    End Function

    Public Sub OnContextMenuDismissed(ByVal browserControl As IWebBrowser, ByVal browser As IBrowser, ByVal frame As IFrame) Implements IContextMenuHandler.OnContextMenuDismissed

    End Sub


    Public Function RunContextMenu(ByVal browserControl As IWebBrowser, ByVal browser As IBrowser, ByVal frame As IFrame, ByVal parameters As IContextMenuParams, ByVal model As IMenuModel, ByVal callback As IRunContextMenuCallback) As Boolean Implements IContextMenuHandler.RunContextMenu
        Return False
    End Function


End Class
Public Class MyExtHandler
    Implements IExtensionHandler

    Public LoadExtensionPopup As Action(Of String)
    Private Function CanAccessBrowser(ByVal extension As IExtension, ByVal browser As IBrowser, ByVal includeIncognito As Boolean, ByVal targetBrowser As IBrowser) As Boolean Implements IExtensionHandler.CanAccessBrowser
        Return True
    End Function

    Private Function GetActiveBrowser(ByVal extension As IExtension, ByVal browser As IBrowser, ByVal includeIncognito As Boolean) As IBrowser Implements IExtensionHandler.GetActiveBrowser

        Return browser
    End Function

    Private Function GetExtensionResource(ByVal extension As IExtension, ByVal browser As IBrowser, ByVal file As String, ByVal callback As IGetExtensionResourceCallback) As Boolean Implements IExtensionHandler.GetExtensionResource
        Return True
    End Function

    Private Function OnBeforeBackgroundBrowser(ByVal extension As IExtension, ByVal url As String, ByVal settings As IBrowserSettings) As Boolean Implements IExtensionHandler.OnBeforeBackgroundBrowser
        Return True
    End Function

    Private Function OnBeforeBrowser(ByVal extension As IExtension, ByVal browser As IBrowser, ByVal activeBrowser As IBrowser, ByVal index As Integer, ByVal url As String, ByVal active As Boolean, ByVal windowInfo As IWindowInfo, ByVal settings As IBrowserSettings) As Boolean Implements IExtensionHandler.OnBeforeBrowser
        Return True
    End Function

    Private Sub OnExtensionLoaded(ByVal extension As IExtension) Implements IExtensionHandler.OnExtensionLoaded
        Dim manifest = extension.Manifest
        Dim browserAction = manifest("browser_action").GetDictionary()

        If browserAction.ContainsKey("default_popup") Then
            Dim popupUrl = browserAction("default_popup").GetString()
            popupUrl = "chrome-extension://" & extension.Identifier & "/" + popupUrl
            LoadExtensionPopup?.Invoke(popupUrl)
        End If
    End Sub

    Private Sub OnExtensionLoadFailed(ByVal errorCode As CefErrorCode) Implements IExtensionHandler.OnExtensionLoadFailed
    End Sub

    Private Sub OnExtensionUnloaded(ByVal extension As IExtension) Implements IExtensionHandler.OnExtensionUnloaded
    End Sub
    Private Sub Dispose() Implements IExtensionHandler.Dispose

    End Sub
End Class

Public Class CustomResourceRequestHandler
    Inherits CefSharp.Handler.ResourceRequestHandler
    Private Delegate Function DelShowMessage(controls As ListBox) As String
    Private Function getListSelected(controls As ListBox)
        If (controls.InvokeRequired) Then
            Dim del As DelShowMessage = New DelShowMessage(AddressOf getListSelected)
            controls.Invoke(del, controls)
        Else
            Return controls.SelectedItem
        End If
        Return Nothing
    End Function
    Protected Overrides Function OnBeforeResourceLoad(ByVal chromiumWebBrowser As IWebBrowser, ByVal browser As IBrowser, ByVal frame As IFrame, ByVal request As IRequest, ByVal callback As IRequestCallback) As CefReturnValue
        'request.SetReferrer("http://xxx.xx/", ReferrerPolicy.Default)
        If request.Url.StartsWith("https://i.pximg.net/img-original/img/") Then
            Dim tarList As ListBox = System.Windows.Forms.Application.OpenForms("Form1").Controls("目錄檔案集合")
            Dim tarUrl As String = getListSelected(tarList)
            request.SetReferrer(tarUrl, ReferrerPolicy.Default)
        End If
        If request.Url.Contains("adserver") Then
            callback.Cancel()
            Return CefReturnValue.Cancel
        End If
        Return CefReturnValue.Continue
    End Function

End Class
Public Class CustomRequestHandler
    Inherits CefSharp.Handler.RequestHandler

    Protected Overrides Function GetResourceRequestHandler(ByVal chromiumWebBrowser As IWebBrowser, ByVal browser As IBrowser, ByVal frame As IFrame, ByVal request As IRequest, ByVal isNavigation As Boolean, ByVal isDownload As Boolean, ByVal requestInitiator As String, ByRef disableDefaultHandling As Boolean) As IResourceRequestHandler
        Return New CustomResourceRequestHandler
    End Function
End Class

Public Class RenderProcessMessageHandler
    Implements IRenderProcessMessageHandler

    Private Sub OnContextCreated(ByVal browserControl As IWebBrowser, ByVal browser As IBrowser, ByVal frame As IFrame) Implements IRenderProcessMessageHandler.OnContextCreated
        Const script As String = "document.addEventListener('DOMContentLoaded', function(){  });"
        'Const script As String = "document.addEventListener('DOMContentLoaded', function(){ alert('DomLoaded'); });"
        frame.ExecuteJavaScriptAsync(script)
    End Sub
    Private Sub OnContextReleased(ByVal browserControl As IWebBrowser, ByVal browser As IBrowser, ByVal frame As IFrame) Implements IRenderProcessMessageHandler.OnContextReleased

    End Sub
    Private Sub OnFocusedNodeChanged(ByVal browserControl As IWebBrowser, ByVal browser As IBrowser, ByVal frame As IFrame, ByVal node As IDomNode) Implements IRenderProcessMessageHandler.OnFocusedNodeChanged

    End Sub
    Private Sub OnUncaughtException(ByVal browserControl As IWebBrowser, ByVal browser As IBrowser, ByVal frame As IFrame, ByVal exception As JavascriptException) Implements IRenderProcessMessageHandler.OnUncaughtException

    End Sub
End Class


Public Class DownloadHandler
    Implements IDownloadHandler

    Public Sub OnBeforeDownload(chromiumWebBrowser As IWebBrowser, browser As IBrowser, downloadItem As DownloadItem, callback As IBeforeDownloadCallback) Implements IDownloadHandler.OnBeforeDownload
        If Not callback.IsDisposed Then
            Using callback
                callback.Continue(downloadItem.SuggestedFileName, showDialog:=False)
            End Using
        End If
    End Sub
    Public Function CanDownload(chromiumWebBrowser As IWebBrowser, browser As IBrowser, url As String, requestMethod As String) As Boolean Implements IDownloadHandler.CanDownload
        Return True
    End Function
    Public Sub OnDownloadUpdated(chromiumWebBrowser As IWebBrowser, browser As IBrowser, downloadItem As DownloadItem, callback As IDownloadItemCallback) Implements IDownloadHandler.OnDownloadUpdated
        If downloadItem.IsComplete Or downloadItem.IsCancelled Then
            MessageBox.Show("The file " & downloadItem.FullPath & " has been downloaded.")
        End If
    End Sub
End Class

Namespace ImageDimensions
    Module ImageHelper
        Const errorMessage As String = "Could not recognize image format."
        Private imageFormatDecoders As Dictionary(Of Byte(), Func(Of BinaryReader, Size)) = New Dictionary(Of Byte(), Func(Of BinaryReader, Size))() From {
            {New Byte() {&H42, &H4D}, AddressOf DecodeBitmap},
            {New Byte() {&H47, &H49, &H46, &H38, &H37, &H61}, AddressOf DecodeGif},
            {New Byte() {&H47, &H49, &H46, &H38, &H39, &H61}, AddressOf DecodeGif},
            {New Byte() {&H89, &H50, &H4E, &H47, &HD, &HA, &H1A, &HA}, AddressOf DecodePng},
            {New Byte() {&HFF, &HD8}, AddressOf DecodeJfif}
        }

        Function GetDimensions(ByVal path As String) As Size
            Using binaryReader As New BinaryReader(File.OpenRead(path))

                Try
                    Return GetDimensions(binaryReader)
                Catch e As ArgumentException

                    If e.Message.StartsWith(errorMessage) Then
                        Throw New ArgumentException(errorMessage, "path", e)
                    Else
                        Throw e
                    End If
                End Try
            End Using
        End Function

        Function GetDimensions(ByVal binaryReader As BinaryReader) As Size
            Dim maxMagicBytesLength As Integer = imageFormatDecoders.Keys.OrderByDescending(Function(x) x.Length).First().Length
            Dim magicBytes As Byte() = New Byte(maxMagicBytesLength - 1) {}

            For i As Integer = 0 To maxMagicBytesLength - 1
                magicBytes(i) = binaryReader.ReadByte()

                For Each kvPair In imageFormatDecoders

                    If magicBytes.StartsWith(kvPair.Key) Then
                        Return kvPair.Value(binaryReader)
                    End If
                Next
            Next

            Throw New ArgumentException(errorMessage, "binaryReader")
        End Function

        <Extension()>
        Private Function StartsWith(ByVal thisBytes As Byte(), ByVal thatBytes As Byte()) As Boolean
            For i As Integer = 0 To thatBytes.Length - 1

                If thisBytes(i) <> thatBytes(i) Then
                    Return False
                End If
            Next

            Return True
        End Function

        <Extension()>
        Private Function ReadLittleEndianInt16(ByVal binaryReader As BinaryReader) As Short
            Dim bytes As Byte() = New Byte(Len(New Short) - 1) {}
            For i = 0 To Len(New Short) - 2
                bytes(Len(New Short) - 2 - i) = binaryReader.ReadByte()
            Next
            Return BitConverter.ToInt16(bytes, 0)
        End Function

        <Extension()>
        Private Function ReadLittleEndianInt32(ByVal binaryReader As BinaryReader) As Integer
            Dim bytes As Byte() = New Byte(Len(New Integer) - 1) {}
            For i = 0 To Len(New Integer) - 1
                bytes(Len(New Integer) - 2 - i) = binaryReader.ReadByte()
            Next
            Return BitConverter.ToInt32(bytes, 0)
        End Function

        Private Function DecodeBitmap(ByVal binaryReader As BinaryReader) As Size
            binaryReader.ReadBytes(16)
            Dim width As Integer = binaryReader.ReadInt32()
            Dim height As Integer = binaryReader.ReadInt32()
            Return New Size(width, height)
        End Function

        Private Function DecodeGif(ByVal binaryReader As BinaryReader) As Size
            Dim width As Integer = binaryReader.ReadInt16()
            Dim height As Integer = binaryReader.ReadInt16()
            Return New Size(width, height)
        End Function

        Private Function DecodePng(ByVal binaryReader As BinaryReader) As Size
            binaryReader.ReadBytes(8)
            Dim width As Integer = binaryReader.ReadLittleEndianInt32()
            Dim height As Integer = binaryReader.ReadLittleEndianInt32()
            Return New Size(width, height)
        End Function

        Private Function DecodeJfif(ByVal binaryReader As BinaryReader) As Size
            While binaryReader.ReadByte() = &HFF
                Dim marker As Byte = binaryReader.ReadByte()
                Dim chunkLength As Short = binaryReader.ReadLittleEndianInt16()

                If marker = &HC0 Then
                    binaryReader.ReadByte()
                    Dim height As Integer = binaryReader.ReadLittleEndianInt16()
                    Dim width As Integer = binaryReader.ReadLittleEndianInt16()
                    Return New Size(width, height)
                End If

                binaryReader.ReadBytes(chunkLength - 2)
            End While

            Throw New ArgumentException(errorMessage)
        End Function
    End Module
End Namespace





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
    Public backup As New List(Of String)
    Public backupUndo As New List(Of String)
    Public archive As IArchive = Nothing
    Public allProcesses As System.Diagnostics.Process() = Nothing
    Dim lineCount As Integer = 0
    Dim output As StringBuilder = New StringBuilder()
    Dim Axplaytime As Double = 0

    Private Sub RadioButton14_Click(sender As Object, e As EventArgs) Handles RadioButton14.Click
        'TextBox1.Text = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile).ToString & "\Downloads"
        ComboBox3.Text = "E:\Download"
        ChecklistBox_Click(sender, e)

    End Sub
    Private Sub RadioButton1_Click(sender As Object, e As EventArgs) Handles RadioButton1.Click
        ComboBox3.Text = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString
        ChecklistBox_Click(sender, e)
    End Sub
    Public Sub Form1_Close(sender As System.Object, e As System.EventArgs) Handles MyBase.FormClosing
        Dim Ordname As String = "F:\desktop\常用\order" & Me.Text & ".ord"
        If File.Exists(Ordname) Then File.Delete(Ordname)
        If Not ComboBox3.Text.ToString.Contains("開啟.trf") Then
            Dim openPath As String = "F:\desktop\常用\" & System.DateTime.Now.ToString("yyyy_MM") & "開啟.trf"
            Using sr As New StreamWriter(openPath, True)
                sr.WriteLine(Me.Text)
            End Using
        End If
        If Not CheckBox4.Checked Then Return
        If String.IsNullOrEmpty(ComboBox3.Text) Then Return
        refresh_backup()
    End Sub
    Private Sub Form1_beforeShown(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.AllowDrop = True
        目錄檔案集合.AllowDrop = True
        ComboBox3.AllowDrop = True
        ComboBox4.AllowDrop = True
        load_passward()
        load_past_path()
        Add_additional_operation()
        load_stripMenuItem()
        PictureBox1.Image = New Bitmap(1, 1)
        Form2.PictureBox1.Image = New Bitmap(1, 1)
        AxWindowsMediaPlayer1.settings.autoStart = True
        chromiumWebBrowser_settings()
    End Sub
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Shown

        Dim commandArgs As String = "Form1"
        If Environment.GetCommandLineArgs().Length < 2 Then Return
        commandArgs = Environment.GetCommandLineArgs()(1)
        Dim ext = Path.GetExtension(commandArgs)
        If Ext_judge(ext, ".trf") Then
            If 目錄檔案集合.Items.Count <> 0 Then Return
            trf_backup(commandArgs)
            OpenFileDialog1.FileName = commandArgs
            read_text_data(commandArgs, 目錄檔案集合)
            delete_repeat_item()
            Add_new_path(ComboBox3, "ori_pathrecord.txt")
            Return
        End If
        If Ext_judge(ext, "Images") Then
            ComboBox3.Text = Path.GetDirectoryName(commandArgs)
            RadioButton3.Checked = True
            ChecklistBox_Click(sender, e)
            副檔名列表.SetItemChecked(0, True)
            ListBox_Click(sender, e)
            目錄檔案集合.SelectedItem = commandArgs

            Button11_Click(sender, e)
        ElseIf Ext_judge(ext, "Sounds") Then
            ComboBox3.Text = Path.GetDirectoryName(commandArgs)
            RadioButton3.Checked = True
            ChecklistBox_Click(sender, e)
            副檔名列表.SetItemChecked(1, True)
            ListBox_Click(sender, e)
            目錄檔案集合.SelectedItem = commandArgs
        ElseIf Ext_judge(ext, "目錄") Then
            ComboBox3.Text = commandArgs
            RadioButton3.Checked = True
            ChecklistBox_Click(sender, e)

        End If
        If Ext_judge(ext, ".wmc") Then
            trf_backup(commandArgs)
            OpenFileDialog1.FileName = commandArgs
            read_text_data(commandArgs, 壓縮檔案集合)
            Add_new_path(ComboBox3, "ori_pathrecord.txt")
            Return
        End If
        ' commandArgs = Path.GetFileNameWithoutExtension(commandArgs)

        '  File.Create("F:\desktop\常用\order\" & commandArgs & ".ord").Close()
        'Dim ordWatcher = get_file_watch("F:\desktop\常用\order")
    End Sub
    Public Sub chromiumWebBrowser_settings()
        Dim settings As CefSettings = New CefSettings()
        settings.CachePath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\cef"
        settings.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.110 Safari/537.36 /CefSharp Browser" & Cef.CefSharpVersion
        settings.CefCommandLineArgs.Add("enable-extension-activity-logging")
        settings.CefCommandLineArgs.Add("extensions-on-chrome-urls")
        settings.CefCommandLineArgs.Add("extensions-not-webstore")
        settings.CefCommandLineArgs.Add("show-component-extension-options")
        'settings.CefCommandLineArgs.Add("disable-web-security", "1")
        settings.CefCommandLineArgs.Add("disable-web-security")
        'settings.CefCommandLineArgs.Add("enable-media-stream")
        'settings.CefCommandLineArgs.Add("allow-running-insecure-content")
        ' settings.CefCommandLineArgs.Add("force-wave-audio")
        Cef.Initialize(settings)
        'ChromiumWebBrowser1.MenuHandler = New MyCustomMenuHandler()

    End Sub
    Private Sub Open_Directory()
        For Each item In 目錄檔案集合.SelectedItems
            Dim argument As String = "/select, """ & item & """"
            System.Diagnostics.Process.Start("explorer.exe", argument)
        Next
    End Sub

    Private Sub load_stripMenuItem()
        Dim rightHotKey As ToolStripMenuItem
        Dim subHotKey As ToolStripMenuItem
        '---------------------- ContextMenuStrip1---------------------------------
        rightHotKey = New ToolStripMenuItem("上一集")
        AddHandler rightHotKey.Click, AddressOf Find_Episode
        ContextMenuStrip1.Items.Add(rightHotKey)
        rightHotKey = New ToolStripMenuItem("下一集")
        AddHandler rightHotKey.Click, AddressOf Find_Episode
        ContextMenuStrip1.Items.Add(rightHotKey)
        rightHotKey = New ToolStripMenuItem("Open File Directory")
        AddHandler rightHotKey.Click, AddressOf Open_Directory
        ContextMenuStrip1.Items.Add(rightHotKey)
        rightHotKey = New ToolStripMenuItem("Copy Path")
        AddHandler rightHotKey.Click, AddressOf Clipboard_SetText
        ContextMenuStrip1.Items.Add(rightHotKey)
        rightHotKey = New ToolStripMenuItem("加入")
        AddHandler rightHotKey.Click, AddressOf addItem_from_clipboard
        ContextMenuStrip1.Items.Add(rightHotKey)
        rightHotKey = New ToolStripMenuItem("匯入")
        AddHandler rightHotKey.Click, AddressOf AddItem_from_filedialod
        ContextMenuStrip1.Items.Add(rightHotKey)
        rightHotKey = New ToolStripMenuItem("上一步")
        AddHandler rightHotKey.Click, AddressOf listUndo
        ContextMenuStrip1.Items.Add(rightHotKey)
        rightHotKey = New ToolStripMenuItem("操控列表")
        'AddHandler rightHotKey.Click, AddressOf delete_repeat_item
        ContextMenuStrip1.Items.Add(rightHotKey)
        subHotKey = New ToolStripMenuItem("移到最底")
        AddHandler subHotKey.Click, AddressOf to_the_bottom
        rightHotKey.DropDownItems.Add(subHotKey)
        subHotKey = New ToolStripMenuItem("上移")
        AddHandler subHotKey.Click, AddressOf list_rearrange
        rightHotKey.DropDownItems.Add(subHotKey)
        subHotKey = New ToolStripMenuItem("下移")
        AddHandler subHotKey.Click, AddressOf list_rearrange
        rightHotKey.DropDownItems.Add(subHotKey)
        subHotKey = New ToolStripMenuItem("Delete_Repeat_Item")
        AddHandler subHotKey.Click, AddressOf delete_repeat_item
        rightHotKey.DropDownItems.Add(subHotKey)
        subHotKey = New ToolStripMenuItem("編輯")
        AddHandler subHotKey.Click, AddressOf editItem
        rightHotKey.DropDownItems.Add(subHotKey)
        subHotKey = New ToolStripMenuItem("反轉列表")
        AddHandler subHotKey.Click, AddressOf reverse_item_list
        rightHotKey.DropDownItems.Add(subHotKey)
        rightHotKey = New ToolStripMenuItem("尋找資源")
        'AddHandler rightHotKey.Click, AddressOf findSource
        ContextMenuStrip1.Items.Add(rightHotKey)
        subHotKey = New ToolStripMenuItem("圖片")
        AddHandler subHotKey.Click, AddressOf findSource
        rightHotKey.DropDownItems.Add(subHotKey)
        subHotKey = New ToolStripMenuItem("ASMR")
        AddHandler subHotKey.Click, AddressOf findSource
        rightHotKey.DropDownItems.Add(subHotKey)
        subHotKey = New ToolStripMenuItem("遊戲")
        AddHandler subHotKey.Click, AddressOf findSource
        rightHotKey.DropDownItems.Add(subHotKey)
        subHotKey = New ToolStripMenuItem("漫畫")
        AddHandler subHotKey.Click, AddressOf findSource
        rightHotKey.DropDownItems.Add(subHotKey)
        '---------------------- ContextMenuStrip2---------------------------------
        rightHotKey = New ToolStripMenuItem("複製圖像")
        AddHandler rightHotKey.Click, AddressOf Clipboard_SetIImage
        ContextMenuStrip2.Items.Add(rightHotKey)
        '---------------------- ContextMenuStrip3---------------------------------
        rightHotKey = New ToolStripMenuItem("Copy Path")
        AddHandler rightHotKey.Click, AddressOf Clipboard_SetText_Compress
        ContextMenuStrip3.Items.Add(rightHotKey)
        rightHotKey = New ToolStripMenuItem("加入")
        AddHandler rightHotKey.Click, AddressOf addItem_from_clipboard_to_archive
        ContextMenuStrip3.Items.Add(rightHotKey)
    End Sub

    Private Sub reverse_item_list()
        '倒轉排列(項數除2，無條件捨去,i(2)=i(n-2))
        Dim theNumber As Double
        Dim items As ListBox.ObjectCollection = 目錄檔案集合.Items
        theNumber = 目錄檔案集合.Items.Count / 2
        Dim theRounded = Math.Sign(theNumber) * Math.Floor(Math.Abs(theNumber) * 100) / 100.0
        Dim n = items.Count
        For i = 0 To theRounded
            Dim temp = items(i).ToString
            items(i) = items(n - i)
            items(n - i) = temp
        Next
        refresh_backup()
    End Sub

    Private Async Sub findSource(sender As Object, e As EventArgs)
        Form2.WebMode("")
        Form2.Show()
        'Dim browser = Form2.ChromiumWebBrowser1
        Dim browser As WebView2 = Form2.WebView21

        Dim rgx As Regex = New Regex("(RJ\d+)")
        Dim fn As String = rgx.Match(目錄檔案集合.SelectedItem).Groups(1).Value
        ' Get the value of the element

        Dim title As String = ""
        Dim btn_name As String = CType(sender, ToolStripMenuItem).Text
        Dim search_site As String() = {}
        Console.WriteLine(btn_name)
        If btn_name = "圖片" Then
            search_site = {"get_address", "sauceNAO", "二次元画像詳細検索"}
        ElseIf btn_name = "ASMR" Then
            search_site = {"dlsite", "mikocon", "anime-sharing"}
        ElseIf btn_name = "遊戲" Then
            search_site = {"get_title", "dlsite", "mikocon", "anime-sharing", "eyny", "nyaa"}
        ElseIf btn_name = "漫畫" Then
            search_site = {"get_title", "nhentai", "nyaa", "F:\CG", "F:\Commonly_used\(H漫"}
        End If
        For Each i As String In search_site
            If i = "get_title" Then
                If fn.Contains("dlsite") Then
                    title = Await browser.ExecuteScriptAsync("document.getElementById('work_name').innerHTML;")
                ElseIf fn.Contains("melonbooks") Then
                    title = Await browser.ExecuteScriptAsync("document.getElementsByClassName('page-header').innerHTML;")
                End If
            ElseIf i = "get_address" Then
                fn = 目錄檔案集合.SelectedItem
            End If
            If i = "dlsite" Then
                壓縮檔案集合.Items.Add("https://www.dlsite.com/maniax/work/=/product_id/" & fn & ".html")
            ElseIf i = "mikocon" Then
                browser.CoreWebView2.Navigate("https://www.mikocon.com/search.php?Mod=forum")
                Await browser.ExecuteScriptAsync("document.getElementById('scform_srchtxt').value = '" & fn & "';")
                Await browser.ExecuteScriptAsync("document.getElementById('scform_submit').click();")
                壓縮檔案集合.Items.Add(browser.Source.ToString())
            ElseIf i = "sauceNAO" Then
                browser.CoreWebView2.Navigate("https://saucenao.com/index.php")
                Await browser.ExecuteScriptAsync("document.getElementById('fileInput').value = '" & fn & "';")
                Await browser.ExecuteScriptAsync("document.getElementById('fileInput').onchange();")
                壓縮檔案集合.Items.Add(browser.Source.ToString())
            ElseIf i = "二次元画像詳細検索" Then
                browser.CoreWebView2.Navigate("https://ascii2d.net/")
                Await browser.ExecuteScriptAsync("document.getElementById('file-form').value = '" & fn & "';")
                Await browser.ExecuteScriptAsync("document.getElementsByClassName('btn btn-secondary').click();")
                壓縮檔案集合.Items.Add(browser.Source.ToString())
            ElseIf i = "anime-sharing" Then
                browser.CoreWebView2.Navigate("http://www.anime-sharing.com/forum/")
                Await browser.ExecuteScriptAsync("document.getElementsByName('query')[0].value = '" & fn & "';")
                Await browser.ExecuteScriptAsync("document.getElementsByName('submit')[0].click();")
                壓縮檔案集合.Items.Add(browser.Source.ToString())
            ElseIf i = "nhentai" Then

                browser.CoreWebView2.Navigate("https://nhentai.net/")
                Await browser.ExecuteScriptAsync("document.getElementsByName('q')[0].value = '" & title & "';")
                Await browser.ExecuteScriptAsync("document.getElementsByClassName('fa fa-search fa-lg')[0].click();")
                壓縮檔案集合.Items.Add(browser.Source.ToString())
            ElseIf i = "nyaa" Then
                browser.CoreWebView2.Navigate("https://sukebei.nyaa.si/")
                Await browser.ExecuteScriptAsync("document.getElementsByClassName('form-control search-bar')[0].value = '" & title & "';")
                Await browser.ExecuteScriptAsync("document.getElementsByClassName('btn btn-primary form-control')[0].click();")
                壓縮檔案集合.Items.Add(browser.Source.ToString())
            ElseIf i.Contains("F:\") Or i.Contains("E:\") Then
                Dim directories As Array = FileSystem.GetDirectories(i, FileIO.SearchOption.SearchAllSubDirectories, fn).Cast(Of String).ToArray
                For Each directory In directories
                    壓縮檔案集合.Items.Add(directory)
                Next
            End If
        Next



        'search item name





    End Sub
    Private Function twitter_repeat(name As String)
        If Not name.Contains("twitter") Then Return name
        Dim rgx As Regex = New Regex("(.*)[\?]")
        Dim value = rgx.Match(name).Groups(1).Value
        If Not String.IsNullOrEmpty(value) Then name = value
        name = name.Replace("mobile.", "")
        Return name
    End Function
    Private Function judge_by_serial_number(checkArray As List(Of String), name As String)
        name = twitter_repeat(name)
        If Not checkArray.Contains(name) Then Return True
        Return False
    End Function
    Private Function union_serial_number(name As String, index As Integer)
        name = twitter_repeat(name)
        目錄檔案集合.Items(index) = name
        Return name
    End Function
    Public Sub delete_repeat_item()
        Dim new_arraylist As New List(Of String)
        Dim need_delete As New List(Of String)
        Dim processingIndex As Short = 0
        Dim oriArray As Array = 目錄檔案集合.Items.Cast(Of String).ToArray
        For Each item As String In oriArray
            If judge_by_serial_number(new_arraylist, item) Then
                '更改名稱:消除mobile,以及?後面網址
                item = union_serial_number(item, processingIndex)
                new_arraylist.Add(item)

            Else
                need_delete.Add(item)
            End If
            processingIndex += 1
        Next
        For Each item In need_delete
            目錄檔案集合.Items.Remove(item)
        Next
        refresh_backup()
    End Sub
    Private Sub to_the_bottom()
        Dim selectedItems As String() = 目錄檔案集合.SelectedItems.Cast(Of String).ToArray
        For Each selectedItem In selectedItems
            目錄檔案集合.Items.Remove(selectedItem)
            目錄檔案集合.Items.Add(selectedItem)
        Next
        目錄檔案集合.SetSelected(目錄檔案集合.Items.Count - 1, True)
        refresh_backup()
    End Sub
    Private Sub list_rearrange(sender As Object, e As EventArgs)
        Dim selectedIndex As Short = 目錄檔案集合.SelectedIndex
        Dim temp_name As String = 目錄檔案集合.SelectedItem.ToString

        Dim btn_name As String = CType(sender, ToolStripMenuItem).Text

        If btn_name = "下移" Then
            目錄檔案集合.Items(selectedIndex) = 目錄檔案集合.Items(selectedIndex + 1)
            目錄檔案集合.Items(selectedIndex + 1) = temp_name
            目錄檔案集合.SetSelected(selectedIndex + 1, True)

        ElseIf btn_name = "上移" Then
            目錄檔案集合.Items(selectedIndex) = 目錄檔案集合.Items(selectedIndex - 1)
            目錄檔案集合.Items(selectedIndex - 1) = temp_name
            目錄檔案集合.SetSelected(selectedIndex - 1, True)
        End If
        目錄檔案集合.SetSelected(selectedIndex, False)
        refresh_backup()
    End Sub
    Private Sub listUndo()
        Dim select_index = 目錄檔案集合.SelectedIndex
        Dim c = backup
        backup = backupUndo
        backupUndo = c
        目錄檔案集合.Items.Clear()
        For Each i In backup
            目錄檔案集合.Items.Add(i)
        Next
        目錄檔案集合.SetSelected(select_index, True)
        refresh_listbox_numbers()
    End Sub
    Private Sub ASMR_Auto_Play(sender As Object, e As AxWMPLib._WMPOCXEvents_PlayStateChangeEvent) Handles AxWindowsMediaPlayer1.PlayStateChange

        If AxWindowsMediaPlayer1.playState = WMPLib.WMPPlayState.wmppsMediaEnded Then
            If ComboBox5.Text = "自動重播(單首)" Then
                AxWindowsMediaPlayer1.settings.setMode("loop", True)
                Return
            End If
            Dim playList As ListBox = Nothing
            If False Then 'ComboBox3.Text.Contains("ASMR") And Directory.Exists(目錄檔案集合.SelectedItem) Then
                Dim presetList As String() = Form3.read_item_from_trf(目錄檔案集合.SelectedItem)
                Dim now_index = Array.IndexOf(presetList, 壓縮檔案集合.SelectedItem)
                If presetList.Count - 1 <> now_index Then
                    playList = 壓縮檔案集合
                    playList.SelectedItem = presetList(now_index + 1)
                    playList.SetSelected(playList.SelectedIndex, False)
                    Return
                Else
                    playList = 目錄檔案集合
                End If

            Else
                If Not 壓縮檔案集合.SelectedIndex > 0 Then
                    playList = 目錄檔案集合
                Else
                    playList = 壓縮檔案集合
                End If
            End If
            If Not playList.SelectedIndex = playList.Items.Count - 1 Then
                playList.SelectedIndex = playList.SelectedIndices(playList.SelectedIndices.Count - 1) + 1
                playList.SetSelected(playList.SelectedIndex, False)
                If False Then 'ComboBox3.Text.Contains("ASMR") And Directory.Exists(目錄檔案集合.SelectedItem) Then
                    Dim presetList As String() = Form3.read_item_from_trf(目錄檔案集合.SelectedItem)
                    壓縮檔案集合.SelectedItem = presetList(0)
                End If
            Else
                If ComboBox5.Text = "自動重播(全部)" Then
                    '放完壓縮檔案目錄後切換到下一個目錄
                    If playList.SelectedIndex = playList.Items.Count - 1 Then
                        playList.SetSelected(playList.SelectedIndex, False)
                        playList.SetSelected(0, True)
                    End If
                Else
                    AxWindowsMediaPlayer1.settings.setMode("loop", False)
                    playList.SetSelected(playList.SelectedIndex, False)
                End If
            End If

        End If
    End Sub
    Private Sub Button1_Dialog(sender As Object, e As EventArgs) Handles Button1.Click
        Using fbd As New FolderBrowserDialog(Me)
            fbd.DirectoryPath = ComboBox3.Text
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
            fbd.DirectoryPath = ComboBox4.Text
            If fbd.ShowDialog = DialogResult.OK Then
                ComboBox4.Text = fbd.DirectoryPath
            End If
        End Using
        If String.IsNullOrEmpty(ComboBox3.Text.ToString) Then
            ComboBox3.Text = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString
        End If
        RadioButton11.Checked = True
        If CheckBox2.Checked = True Then
            ComboBox4.Text += Path.DirectorySeparatorChar & System.DateTime.Now.ToString("yyyy_MM_dd")

        End If
        textbox2_out = ComboBox4.Text
    End Sub

    Private Sub Checkbox7_Dialog(sender As Object, e As EventArgs) Handles Button6.Click
        OpenFileDialog1.Filter = "trf files (*.trf)|*.trf|wmc files (*.wmc)|*.trf|All files (*.*)|*.*"
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            trf_backup(OpenFileDialog1.FileName)
            Dim ext = Path.GetExtension(OpenFileDialog1.FileName)
            For Each control As RadioButton In Form3.GroupBox4.Controls
                If control.Text = ext Then control.Checked = True
            Next
            If Not OpenFileDialog1.FileName.Contains(".wmc") Then
                read_text_data(OpenFileDialog1.FileName, 目錄檔案集合)
            Else
                CheckBox1.Checked = False
                read_text_data(OpenFileDialog1.FileName, 壓縮檔案集合)
                'WmcRefresh()
            End If
        End If

    End Sub
    Public Sub trf_backup(filename As String)
        Dim saveFolder As String = "F:\desktop\常用\backup\" & System.DateTime.Now.ToString("yyyy_MM_dd") & Path.DirectorySeparatorChar
        If Not Directory.Exists(saveFolder) Then Directory.CreateDirectory(saveFolder)
        Dim backupFile As String = saveFolder & Path.GetFileName(filename)
        If File.Exists(backupFile) Then Return
        My.Computer.FileSystem.CopyFile(filename, backupFile)
        Label11.Text = Path.GetFileName(backupFile) & "備份成功"
    End Sub
    Public Sub read_text_data(oripath As String, targetList As ListBox)
        Dim item As String = ""
        Dim rgx As Regex = New Regex("(http.*)")
        Dim tempListbox As New ArrayList
        Using sr As New StreamReader(oripath)
            While (sr.Peek() >= 0)
                '去掉網址以外字元
                If oripath.Contains("網頁") Or oripath.Contains("web") Then
                    item = rgx.Match(sr.ReadLine()).Groups(1).Value
                Else
                    item = sr.ReadLine()

                End If
                If String.IsNullOrEmpty(item) Then Continue While
                If item.Contains("[selected]") Then
                    If targetList.Items.Count = 0 Then Return
                    If Not item.Contains("[selected]-1") Then targetList.SetSelected(item.Split("]")(1), True)
                    Continue While
                End If
                If item.Contains("[playTime]") Then
                    AxWindowsMediaPlayer1.Ctlcontrols.currentPosition = item.Split("]")(1)
                    Continue While
                End If
                If item.Contains("[opened]") Then

                    Continue While
                End If
                If item.Contains("[Form2_open]") Then
                    Button11_Click(New Object, New EventArgs)
                    Continue While
                End If
                targetList.Items.Add(item.Trim)
                ' tempListbox.Add(item.Trim)

                '自創分頁檔 tab record file .trf
                'If Not Ext_judge(Path.GetExtension(item), ".trf") Then

                ' Else
                'Dim tab_form As Form1 = New Form1
                'tab_form.read_text_data(item.Trim)
                'tab_form.Show()
                'End If
            End While

        End Using

        For Each additem As String In tempListbox
            targetList.Items.Add(additem)
        Next

        If targetList IsNot 目錄檔案集合 Then Return
        additional_operate(oripath)
        refresh_backup()
    End Sub
    Private Sub additional_operate(oripath As String)
        If oripath.Contains("影片") Or oripath.Contains("清單") Then
            CheckBox1.Checked = True
            CheckBox4.Checked = True
        ElseIf oripath.Contains("clip") Then
            ComboBox5.SelectedItem = "clip啟動"
            CheckBox4.Checked = True
            RadioButton2.Checked = True
            RadioButton2_Watch(New Object, New EventArgs)
        ElseIf oripath.Contains("畫") Then

            CheckBox4.Checked = True
        ElseIf oripath.Contains("開啟") Then

            ComboBox5.SelectedItem = "執行後移到最後一項"
        ElseIf oripath.Contains("ASMR") Then

            CheckBox7.Checked = True
            If ComboBox6.MaxLength = 0 Then ComboBox6_click(New Object, New EventArgs)
            If Not ComboBox6.SelectedItem.contains("耳機") Then ComboBox6.SelectedIndex = ComboBox6.FindString("耳機")
        ElseIf oripath.Contains("遊戲") Then
            If ComboBox6.MaxLength = 0 Then ComboBox6_click(New Object, New EventArgs)
            If Not ComboBox6.SelectedItem.contains("喇叭") Then ComboBox6.SelectedIndex = ComboBox6.FindString("喇叭")
        End If
        ComboBox3.Text = oripath
        Form3.TextBox1.Text = Path.GetFileName(oripath)
        Me.Text = oripath
    End Sub
    Private Sub AddItem_from_filedialod()
        OpenFileDialog1.Filter = "All files (*.*)|*.*"
        OpenFileDialog1.Multiselect = True
        OpenFileDialog1.ShowDialog()
        For Each sr As String In OpenFileDialog1.FileNames
            目錄檔案集合.Items.Add(sr)
        Next

    End Sub
    Public Sub Output_file_label(sender As Object, e As EventArgs) Handles Button8.Click
        Form3.Show()
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
            ComboBox4.Text += Path.DirectorySeparatorChar & System.DateTime.Now.ToString("yyyy_MM_dd")

        End If
        textbox2_out = ComboBox4.Text
    End Sub
    Private Sub Find_Episode(sender As Object, e As EventArgs)
        If 目錄檔案集合.SelectedItem Is Nothing Then
            Return
        End If
        Dim rgx As Regex = New Regex("(.*\W*)\W(\d{1}|\d{2})(\W.*)")
        Dim ext As String = Path.GetExtension(目錄檔案集合.SelectedItem)
        Dim episode As Short = Integer.Parse(rgx.Match(目錄檔案集合.SelectedItem).Groups(2).Value)

        Dim btn_name As String = CType(sender, ToolStripMenuItem).Text
        Console.WriteLine(btn_name)
        If btn_name = "下一集" Then
            episode += 1
        ElseIf btn_name = "上一集" Then
            episode -= 1
        End If
        For Each files In Directory.GetFiles(Path.GetDirectoryName(目錄檔案集合.SelectedItem))
            If Path.GetExtension(files) <> ext Then
                Continue For
            End If
            If rgx.Match(files).Groups(2).Value = episode Then
                目錄檔案集合.Items(目錄檔案集合.SelectedIndex) = files
                file_Preview(sender, e)
                Exit For
            End If

        Next
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        Dim cal As String = "intersection"
        Dim keyWord As String = TextBox5.Text
        Dim target As String = "name"
        Dim selectitems = 目錄檔案集合.SelectedItems.Cast(Of String).ToArray
        If Not String.IsNullOrEmpty(TextBox5.Text) Then
            If TextBox5.Text(0) = "-" Then
                cal = "subsection"
                keyWord = keyWord.Remove(0, 1)
            End If
            If TextBox5.Text.Contains("*") Then target = "pixel"
            If TextBox5.Text.Contains("desktopPicture") Then target = "desktopPicture"

        End If
        Select Case target
            Case "name"
                Dim new_array As New List(Of String)
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
            Case "pixel"
                Dim new_array As New List(Of String)
                For Each i As String In backup
                    If Not Ext_judge(Path.GetExtension(i).ToLower(), "Images") Then
                        Continue For
                    End If
                    Try
                        Using image As System.Drawing.Image = System.Drawing.Image.FromFile(i)
                            If image.Width.ToString & "*" & image.Height.ToString = keyWord Then
                                new_array.Add(i.ToString)
                            End If
                        End Using

                        '------------Bitmap------------
                        'Dim MyImage As Bitmap = New Bitmap(i)
                        'If MyImage.Width.ToString & "*" & MyImage.Height.ToString = keyWord Then
                        '    new_array.Add(i.ToString)
                        'End If
                        'MyImage.Dispose()


                        '------------ImageDimensions------------
                        'Dim picture_size As Size = ImageDimensions.GetDimensions(i)
                        'If picture_size.Width & "*" & picture_size.Height = keyWord Then
                        'new_array.Add(i.ToString)
                        'End If
                    Catch ex As Exception
                        Console.WriteLine(ex.Message)
                    End Try
                Next
                目錄檔案集合.Items.Clear()
                For Each i In new_array
                    目錄檔案集合.Items.Add(i)
                Next
            Case "desktopPicture"
                Dim new_array As New List(Of String)
                For Each i As String In backup
                    If Not Ext_judge(Path.GetExtension(i).ToLower(), "Images") Then
                        Continue For
                    End If
                    Try
                        Using image As System.Drawing.Image = System.Drawing.Image.FromFile(i)
                            If image.Width > image.Height Then
                                new_array.Add(i.ToString)
                            End If
                        End Using


                    Catch ex As Exception
                        Console.WriteLine(ex.Message)
                    End Try
                Next
                目錄檔案集合.Items.Clear()
                For Each i In new_array
                    目錄檔案集合.Items.Add(i)
                Next
        End Select
        For Each item In selectitems
            If 目錄檔案集合.Items.Contains(item) Then
                Dim index = 目錄檔案集合.Items.IndexOf(item)
                目錄檔案集合.SetSelected(index, True)
            End If
        Next
        refresh_listbox_numbers()
    End Sub
    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        Dim obj As ListBox = Nothing
        Dim new_array As New List(Of String)
        Select Case ComboBox5.SelectedItem
            Case "取樣"
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
            Case "Combine trf files"
                CheckBox4.Checked = False

        End Select


    End Sub
    Private Sub ChecklistBox_Click(sender As Object, e As EventArgs) Handles ComboBox3.TextChanged
        Button13_Click(sender, e)
        副檔名列表.Items.Clear()
        Dim files As IEnumerable
        If ComboBox3.Text.Contains(".trf") Then
            '-----------------------
            Dim oriCount As IEnumerable = 目錄檔案集合.Items.Cast(Of String)
            OpenFileDialog1.FileName = ComboBox3.Text
            read_text_data(ComboBox3.Text, 目錄檔案集合)
            files = 目錄檔案集合.Items.Cast(Of String)
            If files.Equals(oriCount) Then Return
            '--------------------------
            trf_backup(ComboBox3.Text)


            delete_repeat_item()

        Else
            files = My.Computer.FileSystem.GetFiles(ComboBox3.Text, searchOptionCheck)
        End If

        Dim rgx As Regex = New Regex("http")
        If Not rgx.IsMatch(ComboBox3.Text) Then
            If CheckBox10.Checked = True Then
                For Each foundFile As String In files
                    Dim ext As String = Path.GetExtension(foundFile)
                    Dim number As Short = 0
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
                副檔名列表.Items.Add("Images")
                副檔名列表.Items.Add("Media")
                副檔名列表.Items.Add("壓縮檔")
                副檔名列表.Items.Add("光盤")
                副檔名列表.Items.Add("遊戲")
                副檔名列表.Items.Add("文件")
                If ComboBox3.Text.Contains("音樂") Then
                    副檔名列表.SetItemChecked(1, True)
                    ListBox_Click(sender, e)
                ElseIf ComboBox3.Text.Contains("H漫") Then
                    副檔名列表.SetItemChecked(0, True)
                    ListBox_Click(sender, e)
                End If
            End If
        Else
            web_simuler(ComboBox3.Text)
            'If ComboBox6.Text <> "twitter" Then
            ' Await web_crawler(ComboBox3.Text)
            ' Else
            '
            ' End If
            'load passward
        End If
        Me.Text = ComboBox3.Text

        Add_new_path(ComboBox3, "ori_pathrecord.txt")
        Add_new_path(ComboBox4, "des_pathrecord.txt")
    End Sub
    'Public Async Function web_crawler(webUrl As String) As Task(Of Task)
    '    Dim httpClient As System.Net.Http.HttpClient = New System.Net.Http.HttpClient()
    '    httpClient.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "http: //developer.github.com/v3/#user-agent-required")
    '    Dim responseMessage = Await httpClient.GetAsync(webUrl)

    '    Console.WriteLine(responseMessage.StatusCode)
    '    Console.WriteLine(responseMessage.Content.ReadAsStringAsync().Result)

    '    If (responseMessage.StatusCode = System.Net.HttpStatusCode.OK) Then

    '        Dim responseResult As String = responseMessage.Content.ReadAsStringAsync().Result

    '        'Console.WriteLine(responseResult)
    '        Dim config = Configuration.Default.WithDefaultLoader().WithDefaultCookies
    '        Dim context = BrowsingContext.[New](config)
    '        Dim document = context.OpenAsync(
    '                                 Sub(Res)
    '                                     Res.Content(responseResult).Address(webUrl).Header("User-Agent", "http://developer.github.com/v3/#user-agent-required")
    '                                 End Sub).Result
    '        Dim cells As IHtmlCollection(Of IElement) = document.QuerySelectorAll(web_form(0))

    '        For Each item In cells
    '            Dim image_url = Nothing
    '            If ComboBox6.Text = "ehentai" Then
    '                If cells.Length <> 1 Then
    '                    image_url = item.QuerySelector("a").GetAttribute("herf")
    '                    Console.WriteLine(image_url)
    '                    Await web_crawler(image_url)
    '                Else
    '                    image_url = document.QuerySelectorAll(".sni")(0).QuerySelector("img").GetAttribute("src")
    '                    Console.WriteLine(image_url)
    '                    目錄檔案集合.Items.Add(web_form(1) & image_url)
    '                End If
    '            Else
    '                image_url = item.QuerySelector("img").GetAttribute("src")
    '                目錄檔案集合.Items.Add(web_form(1) & image_url)
    '            End If
    '            'QuerySelector(".entry-content")找出class="entry-content"的所有元素


    '        Next

    '    End If

    '    'Console.ReadKey()
    'End Function
    Private Function GetAuthCredentials(ByVal browser As IWebBrowser, ByVal isProxy As Boolean, ByVal host As String, ByVal port As Integer, ByVal realm As String, ByVal scheme As String, ByRef username As String, ByRef password As String) As Boolean
        If host.EndsWith("the-shire.org") Then
            username = "Frodo"
            password = "theR1nG"
            Return True
        End If

        Return False
    End Function
    Public Sub web_simuler(webUrl As String)
        ' WebBrowser1.Show()
        ' WebBrowser1.ScrollBarsEnabled = False
        ' WebBrowser1.ScriptErrorsSuppressed = True
        ' WebBrowser1.Navigate(webUrl)
        ' While WebBrowser1.ReadyState <> WebBrowserReadyState.Complete
        ' Application.DoEvents()
        'End While
        ' WebBrowser1.Document.DomDocument.ToString()

        'ChromiumWebBrowser1.Show()
        'ChromiumWebBrowser1.Load(webUrl)


    End Sub
    'Public Function web_form(index As Integer)
    '    '{class,imgAddress before link}
    '    Dim pattern As String() = New String() {"", "", ""}
    '    If ComboBox6.Text = "kemono.party" Then
    '        pattern(0) = ".post__thumbnail"
    '        pattern(1) = "https://kemono.party"
    '    ElseIf ComboBox6.Text = "18comic" Then
    '        pattern(0) = ".panel-body"
    '    ElseIf ComboBox6.Text = "ehentai" Then
    '        pattern(0) = ".gdt"
    '    ElseIf ComboBox6.Text = "twitter" Then
    '        pattern(0) = ".css-1dbjc4n r-1p0dtai r-1mlwlqe r-1d2f490 r-11wrixw r-61z16t r-1udh08x r-u8s1d r-zchlnj r-ipm5af r-417010"
    '    End If
    '    Return pattern(index)
    'End Function
    Public Function load_passward()
        Dim path1 = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\Documents\passward.txt"
        If File.Exists(path1) And ListBox1.Items.Count = 0 Then
            Using sr As New StreamReader(File.OpenRead(path1))

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
            Using sr As New StreamReader(File.OpenRead(ori_path))

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
            Using sr As New StreamReader(File.OpenRead(des_path))

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
    Public Function searchOptionCheck() As Microsoft.VisualBasic.FileIO.SearchOption
        If CheckBox3.Checked Then
            Return Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories
        Else
            Return Microsoft.VisualBasic.FileIO.SearchOption.SearchTopLevelOnly
        End If

    End Function
    Public Function searchOptionCheck_SystemIO() As System.IO.SearchOption
        If CheckBox3.Checked Then
            Return System.IO.SearchOption.AllDirectories
        Else
            Return System.IO.SearchOption.TopDirectoryOnly
        End If

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
    Private Function getallfiles(path As String, keyword As String) As Short
        Try
            Return FileSystem.GetFiles(path, searchOptionCheck(), keyword).Count

        Catch e As UnauthorizedAccessException
            Return Enumerable.Empty(Of String)().Count
        End Try
    End Function
    Private Sub checkedListBox1_ItemCheck(ByVal sender As Object, ByVal e As ItemCheckEventArgs) Handles 副檔名列表.ItemCheck
        Me.BeginInvoke(CType(Function()
                                 ListBox_Click(sender, e)
                                 Return True
                             End Function, MethodInvoker))
    End Sub

    Private Sub ListBox_Click(sender As Object, e As EventArgs)
        Dim record As String() = {"", ""}
        If 目錄檔案集合.SelectedIndex <> -1 Then
            record(0) = 目錄檔案集合.SelectedItem
        End If
        If 壓縮檔案集合.SelectedIndex <> -1 Then
            record(1) = 壓縮檔案集合.SelectedItem
        End If
        If Axplaytime <> 0 Then
            Axplaytime = AxWindowsMediaPlayer1.Ctlcontrols.currentPosition
        End If

        目錄檔案集合.Items.Clear()
        壓縮檔案集合.Items.Clear()
        If CheckBox11.Checked Then
            Dim newTable As String() = 副檔名列表.CheckedItems.Cast(Of String).ToArray

            If CheckBox10.Checked = False Then
                For Each cont As Object In newTable
                    Dim Directories As IEnumerable = getallDirectories(ComboBox3.Text)
                    For Each foundDirectory As String In Directories
                        Dim all_files_count As Short = getallfiles(foundDirectory, "*")
                        Console.WriteLine(all_files_count)
                        Dim target_count As Short = 0
                        For Each fileExt In Ext_table(cont)
                            target_count += getallfiles(foundDirectory, "*" & fileExt)
                            Console.WriteLine(target_count)
                        Next
                        If target_count > all_files_count - target_count Then
                            If Path.GetDirectoryName(foundDirectory) <> "E:\download" Then 目錄檔案集合.Items.Add(foundDirectory)
                        ElseIf cont = "media" Or cont = "games" Then
                            If Path.GetDirectoryName(foundDirectory) <> "E:\download" Then 目錄檔案集合.Items.Add(foundDirectory)
                        End If
                    Next
                Next
            Else
                For Each cont As Object In newTable
                    Dim Directories As IEnumerable = getallDirectories(ComboBox3.Text)
                    For Each foundDirectory As String In Directories
                        Dim all_files_count As Short = getallfiles(foundDirectory, "")
                        Dim target_count As Short = getallfiles(foundDirectory, cont)
                        If target_count > all_files_count - target_count Then
                            目錄檔案集合.Items.Add(foundDirectory)

                        End If
                    Next
                Next
            End If

        Else
            For Each cont As Object In 副檔名列表.CheckedItems
                If CheckBox10.Checked Then
                    '搜尋ext
                    Dim searchCondition = My.Computer.FileSystem.GetFiles(
                     ComboBox3.Text, searchOptionCheck(), "*" & cont)
                    For Each foundFile As String In searchCondition
                        目錄檔案集合.Items.Add(foundFile)
                    Next
                Else
                    'Dim di As DirectoryInfo = New DirectoryInfo(ComboBox3.Text)
                    'Dim files As System.IO.FileSystemInfo() = di.GetFiles("*", searchOptionCheck_SystemIO)
                    'Dim orderedFiles = files.ToList
                    'Dim searchCondition = My.Computer.FileSystem.GetFiles(
                    ' ComboBox3.Text, searchOptionCheck())
                    Dim files As String() = Directory.GetFiles(ComboBox3.Text, "*.*", searchOptionCheck_SystemIO())
                    Array.Sort(files, AddressOf CompareNatural)
                    For Each foundFile In files
                        Dim ext As String = Path.GetExtension(foundFile).ToLower()
                        If Ext_judge(ext, cont) Then
                            目錄檔案集合.Items.Add(foundFile)
                        End If
                    Next

                End If

            Next
        End If
        目錄檔案集合.SelectedItem = record(0)
        壓縮檔案集合.SelectedItem = record(1)
        If Not String.IsNullOrEmpty(TextBox5.Text) Then
            TextBox5_TextChanged(sender, e)
        End If
        refresh_backup()
    End Sub

    Public Sub refresh_backup()
        refresh_listbox_numbers()
        If 目錄檔案集合.Items.Count = 0 Then Return
        If ComboBox3.Text.Contains(".wmc") Then Return
        If ComboBox3.Text.Contains("backup") Then Return
        If Not String.IsNullOrEmpty(TextBox5.Text) Then Return

        If Not 目錄檔案集合.Items.Cast(Of String).ToList.SequenceEqual(backup) Then
            backupUndo = backup
            backup = 目錄檔案集合.Items.Cast(Of String).ToList
        End If
        If Directory.Exists(ComboBox3.Text) Then Return
        If Not String.IsNullOrEmpty(ComboBox3.Text) Then trf_backup(ComboBox3.Text)
        If CheckBox4.Checked Then
            Form3.Form3_Load(New Object, New EventArgs)
            Form3.Button1_Click(New Object, New EventArgs)
        End If


    End Sub
    Public Async Sub WmcRefresh()
        refresh_listbox_numbers()
        If Not 壓縮檔案集合.Items.Cast(Of String).ToList Is backup Then
            backupUndo = backup
            backup = 壓縮檔案集合.Items.Cast(Of String).ToList
        End If
        If CheckBox4.Checked Then
            Form3.RadioButton15.Checked = True
            Form3.Form3_Load(New Object, New EventArgs)
            Form3.Button1_Click(New Object, New EventArgs)
        End If
        For Each webOrder As String In 壓縮檔案集合.Items
            If Ext_judge(webOrder, "網頁") Then

                Form2.WebMode(webOrder)
                Await Form2.waitForLoad()
            ElseIf webOrder.Contains("crawler") Then
                Dim keyword As String = webOrder.Split("]")(1)
                Await Form2.waitForLoad()
                Form2.wmcSearch(keyword)
            ElseIf webOrder.Contains("clickByClass") Then
                Dim elementClass As String = webOrder.Split("]")(1)
                Await Form2.waitForLoad()

            End If

        Next

    End Sub
    Public Sub refresh_listbox_numbers()
        Dim label13Text As New StringBuilder("")
        label13Text.Append(目錄檔案集合.SelectedIndex + 1 & Path.AltDirectorySeparatorChar & 目錄檔案集合.Items.Count.ToString & "項")
        Label13.Text = label13Text.ToString
        Dim label17Text As New StringBuilder("")
        label17Text.Append(壓縮檔案集合.SelectedIndex + 1 & Path.AltDirectorySeparatorChar & 壓縮檔案集合.Items.Count.ToString & "項")
        Label17.Text = label17Text.ToString
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
    Private Function Ext_table(ext_type As String) As IEnumerable(Of String)
        Dim ext_explain As IEnumerable(Of String) = {}
        If ext_type = "Images" Then
            ext_explain = {".jpg", ".png", ".jfif", ".gif", ".jpeg"}.Concat(Ext_table("其他圖片"))
        ElseIf ext_type = "其他圖片" Then
            ext_explain = {".clip", ".psd", ".webp"}
        ElseIf ext_type = "ASMR" Then
            ext_explain = Ext_table("Images").Concat(Ext_table("Media"))
        ElseIf ext_type = "Media" Then
            ext_explain = Ext_table("Videos").Concat(Ext_table("Sounds"))
        ElseIf ext_type = "Videos" Then
            ext_explain = {".mp4", ".mkv", ".wmv", ".flac", ".flv"}
        ElseIf ext_type = "Sounds" Then
            ext_explain = {".mp3", ".m4a", ".mid", ".wav"}

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
            ext_explain = {".txt", ".ini", ".inf", ".trf"}.Concat(Ext_table("字幕"))
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
    Private Function Ext_judge(ext As String, ext_type As String)
        Dim result = False
        If ext IsNot Nothing Then ext = ext.ToLower
        '容許.url .txt檔
        If Not {"網頁", "目錄", "程序"}.Contains(ext_type) Then
            Dim table As IEnumerable = Ext_table(ext_type)
            result = table.Cast(Of String).Contains(ext)
        ElseIf ext_type = "網頁" Then
            result = ext.Contains("http")
        ElseIf ext_type = "目錄" Then
            If Directory.Exists(ext) = True Then result = True
        ElseIf ext_type = "程序" Then
            If RadioButton15.Checked = True Then result = True
        End If
        Return result

    End Function
    Public Sub Next_Page(sender As Object, e As EventArgs) Handles Button10.Click
        Dim rgx As Regex = New Regex("頁數：(\d+)/")
        Dim now_page As Short = rgx.Match(Label10.Text).Groups(1).Value
        If doc.Pages.Count > now_page Then
            now_page += 1
            Dim bmp As Image = doc.SaveAsImage(now_page - 1)
            Label10.Text = "頁數：" & now_page.ToString & Path.AltDirectorySeparatorChar & doc.Pages.Count.ToString

            PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
            PictureBox1.Image = CType(bmp, Image)
        End If
    End Sub
    Public Sub Privious_Page(sender As Object, e As EventArgs) Handles Button3.Click
        Dim rgx As Regex = New Regex("頁數：(\d+)/")
        Dim now_page As Short = rgx.Match(Label10.Text).Groups(1).Value

        If now_page > 1 Then
            now_page -= 1
            Dim bmp As Image = doc.SaveAsImage(now_page - 1)
            Label10.Text = "頁數：" & now_page.ToString & Path.AltDirectorySeparatorChar & doc.Pages.Count.ToString


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
        Using sr As New StreamWriter(record, True)
            sr.WriteLine(TextBox4.Text)
        End Using
        TextBox4.Text = ""
    End Sub
    Private Sub Add_new_path(target As ComboBox, saveName As String)
        Dim savePath As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\Documents\" & saveName
        Dim saveItems As New List(Of String)
        Dim newItem = target.Text.ToString
        Using sr As New StreamReader(savePath, Encoding.UTF8)
            While (sr.Peek() >= 0)
                saveItems.Add(sr.ReadLine())
            End While
        End Using
        If String.IsNullOrEmpty(newItem) Then Return
        If Ext_judge(saveName, ".trf") Then Return
        If saveItems.Contains(newItem) Then
            Return
            '重選會報錯
            saveItems.Remove(newItem)
        End If
        Dim MaxItemCount = 15
        If saveItems.Count > MaxItemCount Then
            ' Remove the last item from the arraylist


            ' Move all items in the arraylist one position backwards
            For i As Short = saveItems.Count - 1 To 1 Step -1
                'If i - 1 < 0 Then Exit For
                saveItems(i) = saveItems(i - 1)
            Next
            saveItems.RemoveAt(0)
        End If

        ' Insert the new item at the beginning of the arraylist
        saveItems.Insert(0, newItem)

        '---------------deprecated code----------------------
        'saveItems.Add(target.Text)
        'Dim saveBeginIndex As Integer = saveItems.Count - MaxItemCount + 1
        'Dim newArray As New ArrayList

        'For i = saveBeginIndex To saveBeginIndex + MaxItemCount
        '    If i >= 0 And i < saveItems.Count Then
        '        newArray.Add(saveItems(i))

        '    End If
        'Next
        'newArray.Reverse()
        '---------------------------------------------------

        target.Items.Clear()
        Using sr As New StreamWriter(savePath, False, Encoding.UTF8)
            For Each item In saveItems
                sr.WriteLine(item)
                target.Items.Add(item)
            Next
        End Using
        'target.Text = newItem
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
        Using sr As New StreamWriter(record, True)
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
        Dim data As String() = Array.FindAll(list.Items.Cast(Of String).ToArray, Function(x) Ext_judge(Path.GetExtension(x.ToLower), "Images") And Not Ext_judge(Path.GetExtension(x.ToLower), "exception") And Not Ext_judge(Path.GetExtension(x.ToLower), "其他圖片"))
        If Array.Find(data, Function(x) x.Contains("メイン") Or x.Contains("表紙") Or x.Contains("あり") Or x.Contains("有り")) IsNot Nothing Then
            壓縮檔案集合.SelectedItem = Array.Find(data, Function(x) x.Contains("メイン"))
        ElseIf data.Count > 0 Then
            壓縮檔案集合.SelectedItem = data(0)
        End If
        Return True
    End Function
    Public Function searchFolder(ori_folder As String, name As String)
        Dim filesCollections As ObjectModel.ReadOnlyCollection(Of String) = FileSystem.GetFiles(ori_folder, Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories, name)
        Console.WriteLine(filesCollections.Count)
        If filesCollections.Count <> 0 Then name = filesCollections(0).ToString
        Return name
    End Function
    Sub OutputHandler(sender As Object, e As DataReceivedEventArgs)
        If Not String.IsNullOrEmpty(e.Data) Then
            lineCount += 1

            ' Add the text to the collected output.
            output.Append(Environment.NewLine + "[" + lineCount.ToString() + "]: " + e.Data)
        End If
    End Sub
    Public Function path_debug(curitem As String)

        Dim name = Path.GetFileNameWithoutExtension(curitem) + ".*"
        Dim name_with_ext = Path.GetFileName(curitem)
        Dim ori_curitem As String = curitem
        Dim last_folder As String = Path.GetDirectoryName(curitem)
        Console.WriteLine(last_folder + name)
        If Directory.Exists(last_folder) Then
            '-------附檔名更改的情況-----------
            curitem = searchFolder(last_folder, name)
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
            'Dim lastSeparatorIndex As Integer = curitem.LastIndexOf(Path.DirectorySeparatorChar)
            'If lastSeparatorIndex >= 0 AndAlso lastSeparatorIndex < curitem.Length - 1 Then

            'Else
            '    Console.WriteLine("Invalid path: " & curitem)
            'End If
        End If


        If ComboBox3.Text.Contains("碧藍檔案") Or curitem.Contains("碧藍檔案") Then
            ' name = last_folder & Path.DirectorySeparatorChar & name
            curitem = searchFolder("F:\CG\碧藍檔案", name)
            'curitem = searchFolder(Path.GetDirectoryName(curitem), name)

        ElseIf ComboBox3.Text.Contains("原神") Then
            ' name = last_folder & Path.DirectorySeparatorChar & name

            curitem = searchFolder("F:\CG\原神", name)
            'curitem = searchFolder(Path.GetDirectoryName(curitem), name)
        ElseIf ComboBox3.Text.Contains("ASMR") Then
            curitem = curitem.Replace("F:\", "E:\")
            'curitem = searchFolder("E:\ASMR", name)

        End If
        Console.WriteLine(curitem)
        If File.Exists(curitem) Then Return curitem
        Dim desk = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString
        If last_folder = desk Then
            For Each folder In Directory.EnumerateDirectories(desk)

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

        Dim draw_path = "E:\(畫\2023整合"
        If last_folder.Contains("2022") Then
            draw_path = "E:\(畫\2022整合"
        End If
        Console.WriteLine(last_folder)
        Console.WriteLine(draw_path)
        'For Each tar_dir In Directory.GetDirectories("E:\(畫")
        '    If tar_dir.Contains(DateTime.Now.ToString("yyyy")) Then
        '        draw_path = tar_dir
        '        Exit For
        '    ElseIf last_folder.Contains("2022") Then
        '        draw_path = "E:\(畫\2022整合"
        '        Exit For
        '    End If
        'Next
        '改成可選擇搜尋到的結果
        Dim tar_file As New StringBuilder("")
        tar_file.Append(draw_path & Path.DirectorySeparatorChar & last_folder & Path.DirectorySeparatorChar & name_with_ext)

        Console.WriteLine(tar_file.ToString)
        If File.Exists(tar_file.ToString) Then curitem = tar_file.ToString
        'If last_folder = Path.GetFileName(desk) Then
        'Else
        'End If
        If File.Exists(curitem) Then Return curitem
        Return ori_curitem
    End Function
    Private Function GetBitmap(ByVal source As BitmapSource) As Bitmap
        Dim bmp As Bitmap = New Bitmap(source.PixelWidth, source.PixelHeight, System.Drawing.Imaging.PixelFormat.Format32bppPArgb)
        Dim data As BitmapData = bmp.LockBits(New System.Drawing.Rectangle(System.Drawing.Point.Empty, bmp.Size), ImageLockMode.[WriteOnly], System.Drawing.Imaging.PixelFormat.Format32bppPArgb)
        source.CopyPixels(Int32Rect.Empty, data.Scan0, data.Height * data.Stride, data.Stride)
        bmp.UnlockBits(data)
        Return bmp
    End Function
    Public Sub file_Preview(sender As Object, e As EventArgs) Handles 目錄檔案集合.SelectedIndexChanged
        release_all_process()
        refresh_backup()
        If CheckBox1.Checked Then Return
        If (目錄檔案集合.Items.Count = 0 Or 目錄檔案集合.SelectedItems.Count = 0) Then Return
        Dim sel_ind As Short = ComboBox2.SelectedIndex
        If ComboBox2.Items.Count <> 0 Then ComboBox2.Items.Clear()
        Dim curitem As String = 目錄檔案集合.SelectedItems(目錄檔案集合.SelectedItems.Count - 1)
        Dim ext As String = ""
        Dim last_write As String = ""
        Dim installs() As String = New String() {}
        If Not Ext_judge(curitem, "網頁") And Not Ext_judge(curitem, "目錄") Then
            If Not File.Exists(curitem) And 目錄檔案集合.Items.Count <> 0 Then
                Dim newName As String = path_debug(curitem)
                If newName = curitem Then
                    MsgBox("找不到檔案")
                    Return
                Else
                    目錄檔案集合.Items(目錄檔案集合.SelectedIndex) = newName
                    curitem = newName
                End If

            End If
            ext = Path.GetExtension(curitem).ToLower()
            ComboBox2.Text = Path.GetFileNameWithoutExtension(curitem)
            TextBox3.Text = ext
            last_write = File.GetLastWriteTime(curitem).ToString("yyyy_MM_dd-HHmmss")
            installs = New String() {ComboBox2.Text, last_write, "Custom"}
        Else
            Try
                ComboBox2.Text = Path.GetFileNameWithoutExtension(curitem)
            Catch
            End Try

        End If
        ' ComboBox2.Text.Substring(4, ComboBox2.Text.Length - 4),

        ComboBox2.Items.AddRange(installs)
        ComboBox2.SelectedIndex = sel_ind
        If Ext_judge(curitem, "目錄") Then
            壓縮檔案集合.Items.Clear()
            For Each foundFile As String In My.Computer.FileSystem.GetFiles(
                    curitem, searchOptionCheck())

                壓縮檔案集合.Items.Add(foundFile)

            Next
            FindFirstBitmap(sender, e, 壓縮檔案集合)
            ComboBox2.Text = Path.GetFileNameWithoutExtension(curitem)
        ElseIf Ext_judge(ext, ".trf") Then
            壓縮檔案集合.Items.Clear()
            read_text_data(curitem, 壓縮檔案集合)
        End If
        If CheckBox1.Checked Then Return
        If Ext_judge(ext, "程序") Then
            Dim setProcess As System.Diagnostics.Process = allProcesses(目錄檔案集合.SelectedIndex)
            Console.WriteLine(setProcess.StandardInput)
            '  Dim bmp = New Bitmap(PictureBox1.Width, PictureBox1.Height)
            ' Dim g = Drawing.Graphics.FromImage(bmp)
            ' g.DrawString(, New Font("新細明體", 9), Brushes.Black, New PointF(10, 10))
            ' PictureBox1.Image = bmp
        ElseIf Ext_judge(ext, "其他圖片") Then
            PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
            Try

                'Dim file1 As StorageFile = Await StorageFile.GetFileFromPathAsync(curitem)
                'Console.WriteLine(file1.Path)
                ' Dim Thumb1 As StorageItemThumbnail = Await file1.GetThumbnailAsync(ThumbnailMode.SingleItem)
                MagickNET.Initialize()

                Dim img As MagickImage = New MagickImage(curitem)
                Using memStream As New MemoryStream()
                    img.Format = MagickFormat.Bmp
                    img.Write(memStream)
                    PictureBox1.Image = New Bitmap(memStream)

                End Using


                'Dim shellFileProperty As ShellPropertyCollection = New ShellPropertyCollection(curitem)
                'Dim noNullProperty As Array = shellFileProperty.Where(Function(prop) prop.CanonicalName IsNot Nothing).OrderBy(Of String)(Function(prop) prop.CanonicalName).ToArray
                'Array.ForEach(Of IShellProperty)(noNullProperty, Sub(p)
                '                                                     DisplayPropertyValue(p)
                '                                                 End Sub)

                'shellFile1.Thumbnail.FormatOption = ShellThumbnailFormatOption.ThumbnailOnly
                'Dim shellThumb As Bitmap = shellFile1.Thumbnail.LargeBitmap
                'PictureBox1.Image = CType(shellThumb, Image)
                If Form2.Visible = True Then Form2.PictureBox1.Image = PictureBox1.Image
            Catch ex As Exception
                Dim shellFile1 As ShellFile = ShellFile.FromFilePath(curitem)
                ' shellFile1.Thumbnail.FormatOption = ShellThumbnailFormatOption.ThumbnailOnly
                Dim shellThumb As Bitmap = shellFile1.Thumbnail.LargeBitmap
                PictureBox1.Image = CType(shellThumb, Image)
                If Form2.Visible = True Then Form2.PictureBox1.Image = PictureBox1.Image
                Console.WriteLine(ex.Message)

                'Dim p As New Process()
                'p.StartInfo.FileName = "F:\desktop\NConvert\nconvert.exe"
                'p.StartInfo.Arguments = "/view " & curitem
                'p.Start()
            End Try
        ElseIf Ext_judge(ext, "Images") Then
            PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
            Dim MyImage As Bitmap = Nothing
            Try
                Try
                    MyImage = New Bitmap(curitem)
                    Label15.Text = MyImage.Size.Width & "px " & MyImage.Size.Height & "px"
                    PictureBox1.Image = CType(MyImage.Clone, Image)
                    Form2.PictureMode()
                    Form2.PictureBox1.Image = CType(PictureBox1.Image.Clone, Image)
                    transform_index_status_to_Form2()
                    MyImage.Dispose()
                Catch ex As Exception
                    PictureBox1.Load(curitem)
                End Try
            Catch ex As Exception
                MsgBox("無法預覽檔案" & curitem & "無法開啟" & ex.GetType().FullName & ex.Message)
            End Try
        ElseIf ext = ".pdf" Then
            doc.LoadFromFile(curitem)
            '遍歷PDF每一頁
            Dim bmp As Image = doc.SaveAsImage(0)
            Try
                PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
                PictureBox1.Image = CType(bmp, Image)
                Label10.Text = "頁數：1/" & doc.Pages.Count.ToString
                If Form2.Visible = True Then Form2.PictureBox1.Image = PictureBox1.Image
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
            'ElseIf Ext_judge(ext, ".wav") Then
            '    My.Computer.Audio.Play(curitem)
        ElseIf Ext_judge(ext, "Media") Then
            AxWindowsMediaPlayer1.Show()
            Me.BeginInvoke(New Action(Sub()
                                          Me.AxWindowsMediaPlayer1.URL = curitem
                                          'If Axplaytime <> 0 Then
                                          '    AxWindowsMediaPlayer1.Ctlcontrols.currentPosition = Axplaytime
                                          'End If
                                      End Sub))

        ElseIf Ext_judge(ext, "文字") Then
            Dim fileReader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(curitem)
            Dim stringReader As String = fileReader.ReadToEnd
            Dim bmp = New Bitmap(PictureBox1.Width, PictureBox1.Height)
            Dim g = System.Drawing.Graphics.FromImage(bmp)
            g.DrawString(stringReader, New System.Drawing.Font("新細明體", 9), System.Drawing.Brushes.Black, New System.Drawing.PointF(10, 10))
            PictureBox1.Image = bmp
            fileReader.Close()
        ElseIf Ext_judge(ext, "光盤") Then
            壓縮檔案集合.Items.Clear()
            Using isoStream As New FileStream(curitem, FileMode.Open)
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
                Dim passward = try_password(curitem, 0)
                archive = ArchiveFactory.Open(curitem, New SharpCompress.Readers.ReaderOptions With {.Password = passward})
                For Each entry As Object In archive.Entries
                    壓縮檔案集合.Items.Add(entry.key)
                Next
            End Try
            System.Threading.Tasks.Task.WaitAll()
            FindFirstBitmap(sender, e, 壓縮檔案集合)
        ElseIf Ext_judge(curitem, "網頁") Then
            'Dim newWindow As Form2 = New Form2
            'newWindow.WebMode(curitem)
            'newWindow.Show()
            transform_index_status_to_Form2()
            Form2.WebMode(curitem)
            Form2.Show()
        End If

    End Sub
    Private Shared Sub DisplayPropertyValue(ByVal prop As IShellProperty)
        Dim value As String = String.Empty
        value = If(prop.ValueAsObject Is Nothing, "", prop.FormatForDisplay(PropertyDescriptionFormatOptions.None))
        Debug.WriteLine(prop.CanonicalName & " " & value)
    End Sub

    Public Function try_password(curitem As String, index As Short)
        Console.WriteLine(ListBox1.Items(index))
        ListBox1.SetSelected(index, True)
        Dim passward As String = ListBox1.Items(index).ToString
        Try
            archive = ArchiveFactory.Open(curitem, New SharpCompress.Readers.ReaderOptions With {.Password = passward})
        Catch ex As Exception
            If Not index = ListBox1.Items.Count - 1 Then
                passward = try_password(curitem, index + 1)
            Else
                MsgBox("讀取檔案失敗" & curitem & "可能有密碼或密碼錯誤")
                archive.Dispose()
            End If
        End Try
        Return passward
    End Function
    Public Sub Archive_Preview(sender As Object, e As EventArgs) Handles 壓縮檔案集合.SelectedIndexChanged
        release_all_process()
        If CheckBox1.Checked Then Return

        Dim curitem As String = 壓縮檔案集合.SelectedItem

        If OpenFileDialog1.FileName.Contains(".wmc") Then
            'Dim newWindow As Form2 = New Form2
            'newWindow.WebMode(curitem)
            'newWindow.Show()
            Form2.WebMode(curitem)
            Form2.Show()
            Return
        End If
        If Ext_judge(curitem, "網頁") Then
            'Dim newWindow As Form2 = New Form2
            'newWindow.WebMode(curitem)
            'newWindow.Show()
            Form2.WebMode(curitem)
            Form2.Show()
            Return
        End If
        Dim ext = Path.GetExtension(curitem)
        Dim FileExt As String = Path.GetExtension(目錄檔案集合.SelectedItem).Replace(";1", "").ToLower()
        Dim foundfile As String = Path.GetFullPath(目錄檔案集合.SelectedItem)
        Dim curitemStream As Stream = Stream.Null

        TextBox3.Text = ext
        ComboBox2.Text = Path.GetFileNameWithoutExtension(curitem)
        If Ext_judge(foundfile, "目錄") Or Ext_judge(FileExt, ".trf") Then
            If Not Ext_judge(ext, "網頁") Then
                curitemStream = File.OpenRead(curitem)
            End If
        ElseIf Ext_judge(FileExt, "光盤") Then
            Dim isoStream As FileStream = File.OpenRead(foundfile)
            Dim cd = New CDReader(isoStream, True).OpenFile(curitem, FileMode.Open)
            cd.CopyTo(curitemStream)
        ElseIf Ext_judge(FileExt, "壓縮檔") Then
            curitemStream = archive.Entries(壓縮檔案集合.SelectedIndex).OpenEntryStream
        End If


        If Ext_judge(ext, "光盤") Then
            instObj = curitemStream
            ' PictureBox1.Image = bmp
        ElseIf Ext_judge(ext, "其他圖片") Then
            Try
                MagickNET.Initialize()

                Dim img As MagickImage = New MagickImage(curitemStream)
                Using memStream As New MemoryStream()
                    img.Format = MagickFormat.Bmp
                    img.Write(memStream)
                    PictureBox1.Image = New Bitmap(memStream)

                End Using


                'Dim shellFileProperty As ShellPropertyCollection = New ShellPropertyCollection(curitem)
                'Dim noNullProperty As Array = shellFileProperty.Where(Function(prop) prop.CanonicalName IsNot Nothing).OrderBy(Of String)(Function(prop) prop.CanonicalName).ToArray
                'Array.ForEach(Of IShellProperty)(noNullProperty, Sub(p)
                '                                                     DisplayPropertyValue(p)
                '                                                 End Sub)

                'shellFile1.Thumbnail.FormatOption = ShellThumbnailFormatOption.ThumbnailOnly
                'Dim shellThumb As Bitmap = shellFile1.Thumbnail.LargeBitmap
                'PictureBox1.Image = CType(shellThumb, Image)
                If Form2.Visible = True Then Form2.PictureBox1.Image = PictureBox1.Image

            Catch ex As Exception
                MsgBox("無法預覽檔案" & curitem & "無法開啟" & ex.GetType().FullName & ex.Message)
            End Try
        ElseIf Ext_judge(ext, "Images") Then
            PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
            Try
                Dim MyImage As Bitmap = New Bitmap(curitemStream)
                PictureBox1.Image = CType(MyImage.Clone, Image)
                If Form2.Visible = True Then Form2.PictureBox1.Image = PictureBox1.Image
                MyImage.Dispose()
            Catch ex As Exception
                MsgBox("無法顯示圖片" & ex.GetType().FullName & ex.Message)
            End Try
            curitemStream.Dispose()
        ElseIf Ext_judge(ext, "Media") Then

            Dim path = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile).ToString & "\AppData\Local\Temp\temporVideo.mp4"
            Console.WriteLine(ext)
            If Ext_judge(foundfile, "目錄") Or Ext_judge(FileExt, ".trf") Then
                path = curitem
            Else
                archive.Entries(壓縮檔案集合.SelectedIndex).WriteToFile(path)
            End If
            AxWindowsMediaPlayer1.Show()
            If AxWindowsMediaPlayer1.Ctlcontrols.currentPosition = 0 Then AxWindowsMediaPlayer1.close()
            Me.BeginInvoke(New Action(Sub()
                                          Me.AxWindowsMediaPlayer1.URL = path
                                      End Sub))
        ElseIf Ext_judge(ext, "文字") Then
            Using tr As TextReader = New StreamReader(curitemStream)
                Dim stringReader As String = tr.ReadToEnd
                Dim bmp = New Bitmap(PictureBox1.Width, PictureBox1.Height)
                Dim g = System.Drawing.Graphics.FromImage(bmp)
                g.DrawString(stringReader, New System.Drawing.Font("新細明體", 9), System.Drawing.Brushes.Black, New System.Drawing.PointF(10, 10))
                PictureBox1.Image = bmp
            End Using
            FindFirstBitmap(sender, e, 壓縮檔案集合)

        End If
        If instObj Is Nothing Then
            Console.WriteLine("close curitem")
            curitemStream.Close()
            System.Threading.Tasks.Task.WaitAll()
        End If
    End Sub
    Private Sub ExtrackAllFilesInIso(Directorie As DiscUtils.DiscDirectoryInfo, cd As Object, savefilename As String)
        Dim directname As New StringBuilder("")

        ' Append more strings using + operator
        directname.Append(savefilename & Path.DirectorySeparatorChar & Directorie.FullName)

        If Not FileSystem.DirectoryExists(directname.ToString) Then
            FileSystem.CreateDirectory(directname.ToString)
        End If
        For Each CDentry In Directorie.GetFiles
            Label11.Text = "解壓" & CDentry.Name
            Dim newfilePath As New StringBuilder("")
            newfilePath.Append(directname.ToString & System.IO.Path.DirectorySeparatorChar & CDentry.Name.Replace(";1", "").ToLower)
            Dim newfile = File.Create(newfilePath.ToString)
            Dim path As Stream = cd.OpenFile(CDentry.FullName, FileMode.Open)
            path.CopyTo(newfile)
            newfile.Close()
        Next
        For Each CDentry In Directorie.GetDirectories
            ExtrackAllFilesInIso(CDentry, cd, savefilename)
        Next
    End Sub
    Public Function checkindexlast(selected_list As Array, maxItem As Short)
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
    Public Sub release_all_process()
        If CheckBox7.Checked Then Return
        If AxWindowsMediaPlayer1.Visible = True Then
            AxWindowsMediaPlayer1.Hide()
            AxWindowsMediaPlayer1.Controls.Clear()
        End If
        Me.DataGridView1.Hide()
        'Me.ChromiumWebBrowser1.Hide()
        If archive IsNot Nothing Then
            If ComboBox7.Text <> "Set Up" And 壓縮檔案集合.SelectedIndex = -1 Then
                archive.Dispose()
                archive = Nothing
            End If
        End If
        If PictureBox1.Image IsNot Nothing Then PictureBox1.Image.Dispose()
        PictureBox1.Image = New Bitmap(1, 1)
        If Form2.PictureBox1.Image IsNot Nothing Then Form2.PictureBox1.Image.Dispose()
        Form2.PictureBox1.Image = New Bitmap(1, 1)
    End Sub


    Private Sub Button4_Action(sender As Object, e As EventArgs) Handles Button4.Click
        ''把picturebox1釋放
        Add_new_path(ComboBox3, "ori_pathrecord.txt")
        Add_new_path(ComboBox4, "des_pathrecord.txt")
        release_all_process()
        If ComboBox4.Text.Length = 0 Then ComboBox4.Text = ComboBox3.Text

        Dim arguments As String = ""
        Dim counter = -1
        Dim collections As String() = New String() {""}
        Dim collections_number As New List(Of Integer)
        Dim targetList As ListBox = 目錄檔案集合
        If ComboBox5.Text = "執行壓縮檔案" Then
            targetList = 壓縮檔案集合
        End If
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
        If ComboBox5.SelectedItem = "紀錄開啟檔案" = True Then Call Output_file_label(sender, e)
        Dim control As String = GroupBox2.Controls.OfType(Of RadioButton)().FirstOrDefault(Function(n) n.Checked).Text
        Console.WriteLine(control)
        Select Case control
            Case "Open"
                For Each foundfile As String In collections
                    Dim filename As String = Path.GetFileName(foundfile)
                    Dim ext As String = Path.GetExtension(foundfile).ToLower
                    If ComboBox5.Text.ToString.Contains("啟動") Then

                        arguments += $" ""{foundfile}"""


                    ElseIf ComboBox5.Text = "locale emulator" Then
                        Dim exeName As String = foundfile
                        Dim p As ProcessStartInfo = New ProcessStartInfo()
                        Dim lePath As String = "C:\Locale.Emulator.2.5.0.1\LEProc.exe"
                        p.FileName = lePath
                        p.Arguments = $"-run ""{exeName}"""
                        p.UseShellExecute = False
                        p.WorkingDirectory = Path.GetDirectoryName(lePath)
                        Dim res As System.Diagnostics.Process = System.Diagnostics.Process.Start(p)
                        res.WaitForInputIdle(5000)
                    ElseIf ext = ".exe" Then
                        measureDevice(sender, e, "喇叭")
                        Try
                            Dim proinfo As ProcessStartInfo = New ProcessStartInfo With {
                .FileName = foundfile,
                .UseShellExecute = False,
                .WorkingDirectory = Path.GetDirectoryName(foundfile),
                .CreateNoWindow = True
            }

                            Dim prostart As System.Diagnostics.Process = New System.Diagnostics.Process With {
                    .StartInfo = proinfo
                }
                            prostart.Start()
                        Catch
                            Dim proc As System.Diagnostics.Process = New System.Diagnostics.Process()
                            proc.StartInfo.FileName = foundfile
                            proc.StartInfo.UseShellExecute = True
                            proc.StartInfo.Verb = "runas"
                            proc.Start()
                        End Try
                    ElseIf Ext_judge(foundfile, "目錄") Then
                        Try
                            Dim newform = New Form1
                            newform.Show()
                            newform.ComboBox3.Text = foundfile
                            newform.RadioButton3.Checked = True
                            newform.ChecklistBox_Click(sender, e)

                        Catch ex As Exception
                            MsgBox(ex.Message)
                            System.Diagnostics.Process.Start(foundfile)
                        End Try
                    Else
                        System.Diagnostics.Process.Start(foundfile)

                    End If
                    If ComboBox5.Text.Contains("執行後") Then
                        Dim ori_index = 目錄檔案集合.Items.IndexOf(foundfile)

                        目錄檔案集合.Items.Remove(foundfile)
                        If ComboBox5.Text.Contains("移到最後一項") Then
                            目錄檔案集合.Items.Add(foundfile)
                            collections_number.RemoveAt(ori_index)
                            collections_number.Add(目錄檔案集合.Items.Count - 1)
                        End If
                    End If


                Next
            Case "Move"
                For Each foundfile As String In collections
                    counter += 1
                    Dim filename As String = Path.GetFileName(foundfile)
                    Dim ext As String = Path.GetExtension(foundfile).ToLower
                    If Ext_judge(foundfile, "目錄") Then
                        Dim directname As String = FileSystem.GetDirectoryInfo(foundfile).Name
                        Try
                            My.Computer.FileSystem.MoveDirectory(foundfile, ComboBox4.Text & Path.DirectorySeparatorChar & directname, UIOption.AllDialogs)
                            目錄檔案集合.Items.Remove(foundfile)
                        Catch ex As Exception
                            MsgBox("移動檔案失敗" & filename & ex.GetType().FullName & ex.Message)
                        End Try
                    Else
                        Try
                            My.Computer.FileSystem.MoveFile(foundfile, ComboBox4.Text & Path.DirectorySeparatorChar & filename, UIOption.AllDialogs)
                            目錄檔案集合.Items.Remove(foundfile)
                        Catch ex As Exception
                            MsgBox("移動檔案失敗" & filename & ex.GetType().FullName & ex.Message)
                        End Try
                    End If
                Next
            Case "Copy"
                For Each foundfile As String In collections
                    counter += 1
                    Dim filename As String = Path.GetFileName(foundfile)
                    Dim ext As String = Path.GetExtension(foundfile).ToLower
                    My.Computer.FileSystem.CopyFile(foundfile, ComboBox4.Text & Path.DirectorySeparatorChar & filename, UIOption.AllDialogs)
                Next
            Case "Delete"
                For Each foundfile As String In collections
                    counter += 1
                    Dim filename As String = Path.GetFileName(foundfile)
                    Dim ext As String = Path.GetExtension(foundfile).ToLower
                    Delete_Anothor_Thread(foundfile)
                Next
            Case "Special"
                Select Case ComboBox7.Text
                    Case "Rename"
                        For Each foundfile As String In collections
                            counter += 1
                            Dim filename As String = Path.GetFileName(foundfile)
                            Dim ext As String = Path.GetExtension(foundfile).ToLower
                            Dim rename As String = Path.GetFileNameWithoutExtension(foundfile)
                            Select Case ComboBox2.SelectedIndex
                                Case 1
                                    rename = File.GetLastWriteTime(foundfile).ToString("yyyy_MM_dd-HHmmss")
                                Case 2
                                    rename = rename.Substring(4, rename.Length - 4)
                            End Select
                            If ComboBox5.SelectedItem = "更改附檔名" Then
                                rename += TextBox3.Text
                            Else
                                rename += ext
                            End If

                            Try
                                Rename_Anothor_Thread(foundfile & Path.AltDirectorySeparatorChar & rename)
                                '目錄檔案集合.Items(collections_number(counter)) = Path.GetDirectoryName(foundfile) & "/" & rename
                            Catch ex As IOException
                                '換執行緒
                                MsgBox("命名檔案失敗" & rename & "可能被開啟" & ex.Message)
                            Catch ex As UnauthorizedAccessException
                                MsgBox("命名檔案失敗" & rename & "需要系統權限")
                            End Try
                        Next
                    Case "Set Up"

                        For Each foundfile As String In collections
                            If ComboBox4.Text.Contains(".trf") Or String.IsNullOrEmpty(ComboBox4.Text) Then ComboBox4.Text = Path.GetDirectoryName(foundfile)
                            counter += 1
                            Dim filename As String = Path.GetFileName(foundfile)
                            Dim ext As String = Path.GetExtension(foundfile).ToLower
                            Dim filestream1 As FileStream = Nothing
                            If Ext_judge(ext, "光盤") Then
                                filestream1 = File.OpenRead(foundfile)
                            Else
                                filestream1 = New FileStream(ComboBox4.Text & "file", FileMode.Create)
                                instObj.CopyTo(filestream1)

                            End If
                            Using vhdStream As FileStream = filestream1

                                Dim cd As CDReader = New CDReader(vhdStream, True)
                                Dim pattern As String = "[AaUuTtOoRrUuNn]{7}\.[infINF]{3}"
                                Dim rgx As New Regex(pattern)
                                Dim infor As String = "AUTORUN.INF"
                                Dim newfile As FileStream
                                Dim path1 As Stream
                                Label11.Visible = True
                                Dim saveFolder As String = ComboBox4.Text & Path.DirectorySeparatorChar & Path.GetFileNameWithoutExtension(foundfile)
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
                                Label11.Text = "尋找啟動程序.."
                                Dim sentence As String = tr.ReadToEnd
                                pattern = "(?<=open[\s]*=[\s]*)[\w.\\]+"
                                rgx = New Regex(pattern)
                                For Each match As Match In rgx.Matches(sentence)
                                    Console.WriteLine("Found '{0}' at position {1}", match.Value, match.Index)
                                    System.Diagnostics.Process.Start(saveFolder & Path.DirectorySeparatorChar & match.Value)
                                Next

                            End Using

                            If ComboBox5.SelectedItem = "壓縮/解壓後刪除" Then
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
                            If ComboBox5.SelectedItem = "Delete After Compress/Decompress" Then
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
                                archive = ArchiveFactory.Open(foundfile)
                            Catch
                                Dim passward = try_password(foundfile, 0)
                                archive = ArchiveFactory.Open(foundfile, New SharpCompress.Readers.ReaderOptions With {.Password = passward})
                            End Try
                            'ProgressBar1.Maximum = archive.Entries.Where(Function(ex) Not ex.IsDirectory).Sum(Function(ex) ex.Size)
                            'ProgressBar1.Value = 0
                            For Each entry In archive.Entries
                                entryCount += 1
                                If Not entry.IsDirectory Then
                                    Console.WriteLine(entry.Key)
                                    entry.WriteToDirectory(ComboBox4.Text & Path.DirectorySeparatorChar & rename, New ExtractionOptions With
                                  {.ExtractFullPath = True, .Overwrite = True})
                                    extractedCount += 1
                                    Dim percent = CInt((extractedCount / entryCount) * 100)
                                    ProgressBar1.Value = percent
                                End If
                            Next
                            archive.Dispose()
                            If ComboBox5.SelectedItem = "Delete After Compress/Decompress" = True Then
                                Try
                                    My.Computer.FileSystem.DeleteFile(foundfile)
                                Catch ex As Exception
                                    MsgBox("刪除檔案失敗" & filename & "可能已經刪除" & ex.GetType().FullName & ex.Message)
                                End Try
                            End If
                            Label11.Text = rename & "解壓完成"
                        Next
                    Case "Synchronize"
                        Dim folder1 As System.Collections.Generic.List(Of String) = New System.Collections.Generic.List(Of String)
                        Dim folder2 As System.Collections.Generic.List(Of String) = New System.Collections.Generic.List(Of String)
                        目錄檔案集合.Items.Cast(Of String).ToList.ForEach(Sub(i) folder1.Add(i.Replace(ComboBox3.Text, "")))
                        壓縮檔案集合.Items.Cast(Of String).ToList.ForEach(Sub(i) folder2.Add(i.Replace(ComboBox4.Text, "")))

                        Dim noMove As IEnumerable(Of String) = folder1.Cast(Of String).Intersect(folder2.Cast(Of String))
                        Dim addToFolder1 As IEnumerable(Of String) = folder2.Cast(Of String).Except(noMove.Cast(Of String))
                        Dim addToFolder2 As IEnumerable(Of String) = folder1.Cast(Of String).Except(noMove.Cast(Of String))
                        'folder1 to folder2
                        For Each file In addToFolder2
                            Dim filename As String = Path.GetFileName(file)
                            My.Computer.FileSystem.CopyFile(ComboBox3.Text & file, ComboBox4.Text & Path.DirectorySeparatorChar & filename, UIOption.AllDialogs)
                        Next
                        'folder2 to folder1
                        For Each file In addToFolder1
                            Dim filename As String = Path.GetFileName(file)
                            My.Computer.FileSystem.CopyFile(ComboBox4.Text & file, ComboBox3.Text & Path.DirectorySeparatorChar & filename, UIOption.AllDialogs)
                        Next
                    Case "檢查路徑"
                        For Each foundfile As String In collections
                            counter += 1
                            Dim checkIndex = collections_number(counter)
                            If File.Exists(foundfile) Then Continue For
                            目錄檔案集合.Items(checkIndex) = path_debug(foundfile)
                        Next
                    Case "檢查網址"
                        For Each foundfile As String In collections
                            counter += 1
                            Dim checkIndex = collections_number(counter)
                            If 目錄檔案集合.Items(checkIndex).contains("RJ") Then Continue For
                            Form2.WebView21.CoreWebView2.Navigate(目錄檔案集合.Items(checkIndex))
                            目錄檔案集合.Items(checkIndex) = Form2.WebView21.Source.AbsoluteUri
                        Next
                    Case "執行爬蟲指令"
                    Case "建立ASMR清單"
                End Select

        End Select
        Dim app As String = ""
        If ComboBox5.SelectedItem = "clip啟動" Then
            app = "C:\Program Files\CELSYS\CLIP STUDIO 1.5\CLIP STUDIO PAINT\CLIPStudioPaint.exe"
        ElseIf ComboBox5.SelectedItem = "MuseScore啟動" Then
            app = "C:\Program Files\MuseScore 3\bin\MuseScore3.exe"
        End If
        If ComboBox5.Text.ToString.Contains("啟動") Then
            Dim proinfo As ProcessStartInfo = New ProcessStartInfo With {
        .FileName = app,
         .Arguments = arguments,
         .StandardErrorEncoding = System.Text.Encoding.UTF8,
        .UseShellExecute = False,
         .RedirectStandardError = True,
         .RedirectStandardOutput = True
           }

            Dim prostart As System.Diagnostics.Process = New System.Diagnostics.Process With {
             .StartInfo = proinfo
              }
            prostart.Start()

            '方案1 使用cmd啟動程式
            '方案2 修改副檔名啟動的登機碼
            'SetAssociation(ext, ext & "File", app, ext & " File")
            ' Process.Start(foundfile)
        End If
        If CheckBox5.Checked = False Then
            For Each i As Short In collections_number
                目錄檔案集合.SetSelected(i, True)
            Next
        End If
        RadioButton4.Checked = True
        refresh_backup()
        'CheckBox12_CheckedChanged(sender, e)
    End Sub

    Public Function Rename_Anothor_Thread(foundfile_rename As String)
        If (Me.InvokeRequired) Then
            Dim del As DelShowMessage = New DelShowMessage(AddressOf Rename_Anothor_Thread)
            Me.Invoke(del, foundfile_rename)
        Else
            Dim subs As String() = foundfile_rename.Split("/"c)
            My.Computer.FileSystem.RenameFile(subs(0), subs(1))
        End If
        Return True
    End Function
    Public Function Delete_Anothor_Thread(foundfile As String)
        If (Me.InvokeRequired) Then
            Dim del As DelShowMessage = New DelShowMessage(AddressOf Delete_Anothor_Thread)
            Me.Invoke(del, foundfile)
        Else
            If Ext_judge(foundfile, "目錄") Then
                Dim directname = FileSystem.GetDirectoryInfo(foundfile).FullName
                FileSystem.DeleteDirectory(directname, DeleteDirectoryOption.DeleteAllContents)
                目錄檔案集合.Items.Remove(foundfile)
            Else
                My.Computer.FileSystem.DeleteFile(foundfile)
                目錄檔案集合.Items.Remove(foundfile)
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

    Public Function get_file_watch(path As String)
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
        check_order(e.FullPath)

    End Sub
    Private Sub OnDisposed(source As Object, e As FileSystemEventArgs)

    End Sub
    Private Sub check_order(name As String)
        Dim ext As String = Path.GetExtension(name).ToLower
        If Not Ext_judge(ext, ".ord") Then Return
        Console.WriteLine($"File: created in {name}")
        Using sr As New StreamReader(name)
            While (sr.Peek() >= 0)
                Dim order = sr.ReadLine()
                Select Case order
                    Case "stop"
                        AxWindowsMediaPlayer1.Ctlcontrols.stop()
                    Case "play"
                        AxWindowsMediaPlayer1.Ctlcontrols.play()
                    Case "next"
                        目錄檔案集合.SelectedIndex += 1
                        目錄檔案集合.SetSelected(目錄檔案集合.SelectedIndex, False)
                    Case "previous"
                        目錄檔案集合.SelectedIndex -= 1
                        目錄檔案集合.SetSelected(目錄檔案集合.SelectedIndex + 1, False)
                End Select
            End While
        End Using
        Using sr As StreamWriter = New StreamWriter(name, False)

        End Using
    End Sub
    Private Sub add_clip(name As String)
        Dim ext As String = Path.GetExtension(name).ToLower
        If Ext_judge(ext, "其他圖片") Then
            Console.WriteLine($"File: created in {name}")
            Dim desk As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            AddItemListbox(name)
        End If
    End Sub
    Private Sub OnCreated(source As Object, e As FileSystemEventArgs)
        ' Specify what is done when a file is created.
        add_clip(e.FullPath)
        check_order(e.FullPath)

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
            If Not s.Contains("CELSYSUserData") And Not s.Contains("previewthumb") Then
                目錄檔案集合.Items.Remove(s)
                目錄檔案集合.Items.Add(s)
                refresh_backup()
            End If

        End If
        Return True
    End Function
    Public Sub OnRenamed(source As Object, e As RenamedEventArgs)
        ' Specify what is done when a file is renamed.
        add_clip(e.FullPath)
        check_order(e.FullPath)

    End Sub

    Dim WithEvents Client As New System.Net.WebClient()
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        'ProgressBar1.Visible = True
        ' Set Minimum to 1 to represent the first file being copied.
        'ProgressBar1.Minimum = 1
        'ProgressBar1.Maximum = 100
        'Client.DownloadFileAsync(New Uri(ComboBox3.Text), ComboBox4.Text)

        ' Set Maximum to the total number of files to copy
        'Dim i = 0
        'For Each imageUrl As String In 目錄檔案集合.SelectedItems
        '    Dim localPath = ComboBox4.Text & Path.DirectorySeparatorChar & i & ".jpg"
        '    Dim config = Configuration.Default.WithDefaultLoader().WithDefaultCookies
        '    Dim context = BrowsingContext.[New](config)
        '    Dim Download = context.GetService(Of IDocumentLoader)().FetchAsync(New DocumentRequest(New Url(imageUrl)))

        '    Using response = Await Download.Task

        '        Using target = File.OpenWrite(localPath)

        '            Await response.Content.CopyToAsync(target)
        '        End Using
        '    End Using
        '    i += 1
        'Next
        Form2.checkWebSite()
        refresh_listbox_numbers()
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
    Friend WithEvents FolderBrowserDialog2 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents CheckBox11 As CheckBox
    Friend WithEvents Button12 As Button
    Friend WithEvents CheckBox9 As CheckBox
    Friend WithEvents TextBox5 As TextBox
    Friend WithEvents TextBox6 As TextBox
    Friend WithEvents RadioButton15 As RadioButton
    Friend WithEvents Button14 As Button
    Friend WithEvents Button15 As Button
    Friend WithEvents Button16 As Button
    Friend WithEvents Button17 As Button
    Friend WithEvents AxWindowsMediaPlayer1 As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents ComboBox3 As ComboBox
    Friend WithEvents ComboBox4 As ComboBox
    Friend WithEvents ComboBox5 As ComboBox
    Friend WithEvents Label12 As Label
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents Label13 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents ComboBox6 As ComboBox
    Friend WithEvents Label15 As Label
    Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
    Friend WithEvents CheckBox4 As CheckBox
    Friend WithEvents ContextMenuStrip2 As ContextMenuStrip
    Friend WithEvents Button13 As Button
    Friend WithEvents Button18 As Button
    Friend WithEvents RadioButton16 As RadioButton
    Friend WithEvents Label16 As Label
    Friend WithEvents ComboBox7 As ComboBox
    Friend WithEvents Label17 As Label
    Friend WithEvents ContextMenuStrip3 As ContextMenuStrip
    Friend WithEvents Button20 As Button
    Friend WithEvents CheckBox7 As CheckBox
    Friend WithEvents TreeView1 As Forms.TreeView
    Friend WithEvents Button19 As Button
End Class

