Imports System.ComponentModel
Imports System.Media
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports AudioSwitcher.AudioApi.CoreAudio
Imports AxWMPLib
Imports CefSharp
Imports CefSharp.Handler
Imports Microsoft.VisualBasic.Devices
Imports Microsoft.WindowsAPICodePack.Shell.PropertySystem.SystemProperties.System
Imports VisioForge.Libs.WindowsMediaLib

Public Class Form5

    Private DevicesComboBox As New ComboBox
    Private DefaultDeviceLabel As New Label
    Private WithEvents SetDefaultButton As New Button
    Private Const DRVM_MAPPER_PREFERRED_GET As Integer = &H2015
    Private Const DRVM_MAPPER_PREFERRED_SET As Integer = &H2016

    'Private WAVE_MAPPER As New IntPtr(-1)

    ' This just brings together a device ID and a WaveOutCaps so 
    ' that we can store them in a combobox.
    'Private Structure WaveOutDevice

    '    Private m_id As Integer
    '    Public Property Id() As Integer
    '        Get
    '            Return m_id
    '        End Get
    '        Set(ByVal value As Integer)
    '            m_id = value
    '        End Set
    '    End Property

    '    Private m_caps As WAVEOUTCAPS
    '    Public Property WaveOutCaps() As WAVEOUTCAPS
    '        Get
    '            Return m_caps
    '        End Get
    '        Set(ByVal value As WAVEOUTCAPS)
    '            m_caps = value
    '        End Set
    '    End Property

    '    Sub New(ByVal id As Integer, ByVal caps As WAVEOUTCAPS)
    '        m_id = id
    '        m_caps = caps
    '    End Sub

    '    Public Overrides Function ToString() As String
    '        Return WaveOutCaps.szPname
    '    End Function

    'End Structure
    <DllImport("winmm.dll")>
    Private Shared Function waveOutSetVolume(hwo As IntPtr, dwVolume As UInteger) As Integer
    End Function

    Declare Function waveOutMessage Lib "winmm.dll" (ByVal hwo As IntPtr, ByVal uMsg As Integer, ByVal dw1 As IntPtr, ByVal dw2 As IntPtr) As Integer

    ' Declare a function to enumerate devices
    Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Integer

    ' Declare a function to get device capabilities
    Declare Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveOutGetDevCapsA" (ByVal uDeviceID As Integer, ByRef lpCaps As WAVEOUTCAPS, ByVal uSize As Integer) As Integer

    ' Define a structure to store device capabilities

    Public Structure WAVEOUTCAPS
        Public wMid As Short
        Public wPid As Short
        Public vDriverVersion As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)>
        Public szPname As String
        Public dwFormats As Integer
        Public wChannels As Short
        Public wReserved1 As Short
        Public dwSupport As Integer
    End Structure

    ' Declare a function to open a waveOut device

    ' Define some constants



    ' Declare a structure to store wave format information

    ' Declare a function to open a waveOut device
    Declare Function waveOutOpen Lib "winmm.dll" (ByRef phwo As IntPtr, ByVal uDeviceID As Integer, ByVal pwfx As WAVEFORMATEX, ByVal dwCallback As IntPtr, ByVal dwInstance As IntPtr, ByVal fdwOpen As Integer) As Integer

    ' Define some constants
    Const WAVE_MAPPER = -1
    Const WAVE_FORMAT_QUERY = &H1

    ' Declare a structure to store wave format information with correct alignment
    <StructLayout(LayoutKind.Sequential)>
    Public Structure WAVEFORMATEX
        <MarshalAs(UnmanagedType.U2)> Public wFormatTag As Short
        <MarshalAs(UnmanagedType.U2)> Public nChannels As Short
        <MarshalAs(UnmanagedType.U4)> Public nSamplesPerSec As Integer
        <MarshalAs(UnmanagedType.U4)> Public nAvgBytesPerSec As Integer
        <MarshalAs(UnmanagedType.U2)> Public nBlockAlign As Short
        <MarshalAs(UnmanagedType.U2)> Public wBitsPerSample As Short
        <MarshalAs(UnmanagedType.U2)> Public cbSize As Short
    End Structure
    Private Sub Form5_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' I do use the IDE for this stuff normally... (in case anyone is wondering)
        Me.Controls.AddRange(New Control() {DevicesComboBox, DefaultDeviceLabel, SetDefaultButton})
        DevicesComboBox.Location = New Point(5, 5)
        DevicesComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        DevicesComboBox.Width = Me.ClientSize.Width - 10
        DevicesComboBox.Anchor = AnchorStyles.Left Or AnchorStyles.Right
        DefaultDeviceLabel.Location = New Point(DevicesComboBox.Left, DevicesComboBox.Bottom + 5)
        DefaultDeviceLabel.AutoSize = True
        SetDefaultButton.Location = New Point(DefaultDeviceLabel.Left, DefaultDeviceLabel.Bottom + 5)
        SetDefaultButton.Text = "Set Default"
        SetDefaultButton.AutoSize = True
        RefreshInformation()
    End Sub

    Private Sub RefreshInformation()
        PopulateDeviceComboBox()
        DisplayDefaultWaveOutDevice()
    End Sub

    Private Sub PopulateDeviceComboBox()
        DevicesComboBox.Items.Clear()
        ' How many wave out devices are there? WaveOutGetNumDevs API call.
        'Dim waveOutDeviceCount As Integer = VisioForge.Libs.WindowsMediaLib.waveOut.GetNumDevs()
        'For i As Integer = 0 To waveOutDeviceCount - 1
        '    Dim caps As New WaveOutCaps
        '    ' Get a name - its in a WAVEOUTCAPS structure. 
        '    ' The name is truncated to 31 chars by the api call. You probably have to 
        '    ' dig around in the registry to get the full name.
        '    Dim result As Integer = VisioForge.Libs.WindowsMediaLib.waveOut.GetDevCaps(i, caps, Marshal.SizeOf(caps))
        '    If result <> MMSYSERR.NoError Then
        '        Dim err As MMSYSERR = DirectCast(result, MMSYSERR)
        '        Throw New Win32Exception("GetDevCaps() error, Result: " & result.ToString("x8") & ", " & err.ToString)
        '    End If
        '    DevicesComboBox.Items.Add(New WaveOutDevice(i, caps))
        'Next

        Dim numDevs As Integer = waveOutGetNumDevs()

        ' Loop through each device ID and get its name and capabilities
        For i As Integer = 0 To numDevs - 1

            ' Declare a variable to store device capabilities
            Dim caps As New WAVEOUTCAPS()

            ' Call the function with device ID and caps structure as parameters
            Dim result As Integer = waveOutGetDevCaps(i, caps, Marshal.SizeOf(caps))

            ' Check if the function succeeded
            If result = 0 Then

                ' The function returned the device name and capabilities in caps structure 
                Console.WriteLine("Device ID: " & i)
                Console.WriteLine("Device Name: " & caps.szPname)
                Console.WriteLine("Device Channels: " & caps.wChannels)
                Console.WriteLine("Device Formats: " & caps.dwFormats)
                Console.WriteLine("Device Support: " & caps.dwSupport)
                DevicesComboBox.Items.Add(caps.szPname)
                ' You can add more code here to process or display the device information

            Else

                ' The function failed and returned an error code in result variable 
                Console.WriteLine("The function failed with error code " & result)

            End If

        Next i

        'Dim devices As IEnumerable(Of CoreAudioDevice) = New CoreAudioController().GetPlaybackDevices()
        'For Each device In devices
        '    DevicesComboBox.Items.Add(device.FullName)
        'Next
        'DevicesComboBox.SelectedIndex = 0


    End Sub

    Private Sub DisplayDefaultWaveOutDevice()
        'Dim currentDefault As Integer = GetIdOfDefaultWaveOutDevice()
        'Dim device As WaveOutDevice = DirectCast(DevicesComboBox.Items(currentDefault), WaveOutDevice)
        'DefaultDeviceLabel.Text = "Defualt: " & device.WaveOutCaps.szPname

        ' Declare a variable to store the ID of the device you want to get the name of (you can change it according to your needs)
        Dim deviceID As Integer = 0 'GetIdOfDefaultWaveOutDevice()

        ' Declare a variable to store device capabilities
        Dim caps As New WAVEOUTCAPS()

        ' Call the function with device ID and caps structure as parameters
        Dim result As Integer = waveOutGetDevCaps(deviceID, caps, Marshal.SizeOf(caps))

        ' Check if the function succeeded and returned MMSYSERR_NOERROR (0)
        If result = 0 Then
            DefaultDeviceLabel.Text = "Defualt: " & caps.szPname
            ' The function returned the device name and capabilities in caps structure 
            Console.WriteLine("Device ID: " & deviceID)
            Console.WriteLine("Device Name: " & caps.szPname)

        Else

            ' The function failed and returned an error code in result variable 
            Console.WriteLine("The function failed with error code " & result)

        End If

        'DefaultDeviceLabel.Text = New CoreAudioController().DefaultPlaybackDevice.FullName
        'DevicesComboBox.SelectedItem = DefaultDeviceLabel.Text
    End Sub

    Private Function GetIdOfDefaultWaveOutDevice() As Integer
        'Dim id As Integer = 0
        'Dim hId As IntPtr
        'Dim flags As Integer = 0
        'Dim hFlags As IntPtr
        'Dim result As Integer
        'Try
        '     It would be easier to declare a nice overload with ByRef Integers.
        '    hId = Marshal.AllocHGlobal(4)
        '    hFlags = Marshal.AllocHGlobal(4)
        '     http://msdn.microsoft.com/en-us/library/bb981557.aspx
        '    result = VisioForge.Libs.WindowsMediaLib.waveOut.Message(WAVE_MAPPER, DRVM_MAPPER_PREFERRED_GET, hId, hFlags)
        '    If result <> MMSYSERR.NoError Then
        '        Dim err As MMSYSERR = DirectCast(result, MMSYSERR)
        '        Throw New Win32Exception("waveOutMessage() error, Result: " & result.ToString("x8") & ", " & err.ToString)
        '    End If
        '    id = Marshal.ReadInt32(hId)
        '    flags = Marshal.ReadInt32(hFlags)
        'Finally
        '    Marshal.FreeHGlobal(hId)
        '    Marshal.FreeHGlobal(hFlags)
        'End Try
        ' There is only one flag, DRVM_MAPPER_PREFERRED_FLAGS_PREFERREDONLY, defined as 1
        ' "When the DRVM_MAPPER_PREFERRED_FLAGS_PREFERREDONLY flag bit is set, ... blah ...,  
        ' the waveIn and waveOut APIs use only the current preferred device and do not search 
        ' for other available devices if the preferred device is unavailable. 
        'Return id



        Dim defaultDeviceID As Integer = WAVE_MAPPER

        ' Declare a variable to store the wave format information
        Dim waveFormat As New WAVEFORMATEX()

        ' Set some values for the wave format structure (you can change them according to your needs)
        waveFormat.wFormatTag = 1 ' PCM format 
        waveFormat.nChannels = 2 ' stereo 
        waveFormat.nSamplesPerSec = 44100 ' sample rate 
        waveFormat.wBitsPerSample = 16 ' bits per sample 
        waveFormat.nBlockAlign = CShort(waveFormat.nChannels * (waveFormat.wBitsPerSample / 8)) ' block alignment 
        waveFormat.nAvgBytesPerSec = waveFormat.nSamplesPerSec * waveFormat.nBlockAlign ' average bytes per second 
        waveFormat.cbSize = 0 ' extra information 

        ' Declare a variable to store the handle of the default device
        Dim defaultDeviceHandle As IntPtr

        ' Call the function with WAVE_MAPPER as the device identifier and WAVE_FORMAT_QUERY as the flag
        'Dim result As Integer = waveOutOpen(defaultDeviceHandle, defaultDeviceID, waveFormat, IntPtr.Zero, IntPtr.Zero, WAVE_FORMAT_QUERY)
        Dim result As Integer = waveOutOpen(defaultDeviceHandle, defaultDeviceID, waveFormat, IntPtr.Zero, IntPtr.Zero, WAVE_FORMAT_QUERY Or CallingConvention.StdCall)
        ' Check if the function succeeded and returned MMSYSERR_NOERROR (0)
        If result = 0 Then

            ' The function returned a handle to the default device in defaultDeviceHandle variable 
            Console.WriteLine("The handle of the default device is " & defaultDeviceHandle.ToString())

        Else

            ' The function failed and returned an error code in result variable 
            Console.WriteLine("The function failed with error code " & result)

        End If

        Return defaultDeviceID
    End Function

    Private Sub SetDefaultButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SetDefaultButton.Click
        If DevicesComboBox.Items.Count = 0 Then Return

        'Dim selectedDevice As WaveOutDevice = DirectCast(DevicesComboBox.SelectedItem, WaveOutDevice)
        'SetDefault(selectedDevice.Id)
        SetDefault(DevicesComboBox.SelectedIndex)
        RefreshInformation()
    End Sub
    ' Import the waveOutSetVolume function from the Windows API


    ' Set the default audio device for output to a specific device

    Private Sub SetDefault(ByVal id As Integer)
        'Dim defaultId As Integer = GetIdOfDefaultWaveOutDevice()
        'If defaultId = id Then Return ' no change.

        '----------------main--------------------
        'Dim devices As IEnumerable(Of CoreAudioDevice) = New CoreAudioController().GetPlaybackDevices()
        'devices(id).SetAsDefault()
        '----------------------------------------

        'Form1.AxWindowsMediaPlayer1.settings.audioDevice = DevicesComboBox.SelectedItem

        ' so here we say "change the id of the device that has id id to 0", which makes it the default.

        'result = waveOut.Message(WAVE_MAPPER, DRVM_MAPPER_PREFERRED_SET, New IntPtr(id), IntPtr.Zero)
        'If result <> MMSYSERR.NoError Then
        '    Dim err As MMSYSERR = DirectCast(result, MMSYSERR)
        '    drvm_mapper_preferred_set 只支持xp以下版本
        '    Throw New Win32Exception("waveoutmessage() error, result: " & result.ToString("x8") & ", " & err.ToString)
        'End If
        ' Define some constants

        Const DRVM_MAPPER_PREFERRED_GET = &H2015

        ' Declare a variable to store the device ID
        Dim deviceID As Integer = id

        ' Call the function with WAVE_MAPPER as the handle and DRVM_MAPPER_PREFERRED_GET as the message
        Dim result As Integer = waveOutMessage(New IntPtr(WAVE_MAPPER), DRVM_MAPPER_PREFERRED_GET, deviceID, 0)

        ' Check if the function succeeded
        If result = 0 Then
            ' The function returned the preferred device ID in deviceID variable
            Console.WriteLine("The preferred device ID is " & deviceID)
        Else
            ' The function failed and returned an error code in result variable
            Console.WriteLine("The function failed with error code " & result)
        End If

    End Sub
    ' Declare a function to get the preferred device ID




End Class
