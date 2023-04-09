

Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Menu
Imports Aspose.Slides
Imports AudioSwitcher.AudioApi
Imports AudioSwitcher.AudioApi.CoreAudio
Imports CefSharp
Imports CefSharp.DevTools.BackgroundService
Imports CefSharp.WinForms
Imports Microsoft.VisualBasic.FileIO
Imports VisioForge.Libs.NDI
Imports VisioForge.Libs.WindowsMediaLib
Imports VisioForge.MediaFramework.GStreamer.Helpers

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。 
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer



    '注意: 以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請勿使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.副檔名列表 = New System.Windows.Forms.CheckedListBox()
        Me.目錄檔案集合 = New System.Windows.Forms.ListBox()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RadioButton15 = New System.Windows.Forms.RadioButton()
        Me.RadioButton14 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton3 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.RadioButton16 = New System.Windows.Forms.RadioButton()
        Me.RadioButton10 = New System.Windows.Forms.RadioButton()
        Me.RadioButton6 = New System.Windows.Forms.RadioButton()
        Me.RadioButton5 = New System.Windows.Forms.RadioButton()
        Me.RadioButton4 = New System.Windows.Forms.RadioButton()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.RadioButton11 = New System.Windows.Forms.RadioButton()
        Me.RadioButton13 = New System.Windows.Forms.RadioButton()
        Me.壓縮檔案集合 = New System.Windows.Forms.ListBox()
        Me.ContextMenuStrip3 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.Button9 = New System.Windows.Forms.Button()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.CheckBox5 = New System.Windows.Forms.CheckBox()
        Me.CheckBox6 = New System.Windows.Forms.CheckBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button10 = New System.Windows.Forms.Button()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Button11 = New System.Windows.Forms.Button()
        Me.CheckBox10 = New System.Windows.Forms.CheckBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.CheckBox11 = New System.Windows.Forms.CheckBox()
        Me.Button12 = New System.Windows.Forms.Button()
        Me.CheckBox9 = New System.Windows.Forms.CheckBox()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        Me.TextBox6 = New System.Windows.Forms.TextBox()
        Me.Button16 = New System.Windows.Forms.Button()
        Me.Button17 = New System.Windows.Forms.Button()
        Me.ComboBox3 = New System.Windows.Forms.ComboBox()
        Me.ComboBox4 = New System.Windows.Forms.ComboBox()
        Me.ComboBox5 = New System.Windows.Forms.ComboBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.ComboBox6 = New System.Windows.Forms.ComboBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.CheckBox4 = New System.Windows.Forms.CheckBox()
        Me.Button13 = New System.Windows.Forms.Button()
        Me.Button18 = New System.Windows.Forms.Button()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.ComboBox7 = New System.Windows.Forms.ComboBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Button20 = New System.Windows.Forms.Button()
        Me.CheckBox7 = New System.Windows.Forms.CheckBox()
        Me.TreeView1 = New System.Windows.Forms.TreeView()
        Me.AxWindowsMediaPlayer1 = New AxWMPLib.AxWindowsMediaPlayer()
        Me.Button19 = New System.Windows.Forms.Button()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxWindowsMediaPlayer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(449, 19)
        Me.Button1.Margin = New System.Windows.Forms.Padding(2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 20)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Open Dialog"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(20, 22)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 12)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Path："
        '
        '副檔名列表
        '
        Me.副檔名列表.FormattingEnabled = True
        Me.副檔名列表.Location = New System.Drawing.Point(10, 101)
        Me.副檔名列表.Margin = New System.Windows.Forms.Padding(2)
        Me.副檔名列表.Name = "副檔名列表"
        Me.副檔名列表.Size = New System.Drawing.Size(84, 106)
        Me.副檔名列表.TabIndex = 0
        '
        '目錄檔案集合
        '
        Me.目錄檔案集合.AllowDrop = True
        Me.目錄檔案集合.ContextMenuStrip = Me.ContextMenuStrip1
        Me.目錄檔案集合.FormattingEnabled = True
        Me.目錄檔案集合.HorizontalScrollbar = True
        Me.目錄檔案集合.ItemHeight = 12
        Me.目錄檔案集合.Location = New System.Drawing.Point(97, 101)
        Me.目錄檔案集合.Margin = New System.Windows.Forms.Padding(2)
        Me.目錄檔案集合.Name = "目錄檔案集合"
        Me.目錄檔案集合.ScrollAlwaysVisible = True
        Me.目錄檔案集合.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.目錄檔案集合.Size = New System.Drawing.Size(336, 112)
        Me.目錄檔案集合.TabIndex = 3
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(61, 4)
        '
        'PictureBox1
        '
        Me.PictureBox1.ContextMenuStrip = Me.ContextMenuStrip2
        Me.PictureBox1.Location = New System.Drawing.Point(436, 90)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(379, 267)
        Me.PictureBox1.TabIndex = 4
        Me.PictureBox1.TabStop = False
        '
        'ContextMenuStrip2
        '
        Me.ContextMenuStrip2.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.ContextMenuStrip2.Name = "ContextMenuStrip2"
        Me.ContextMenuStrip2.Size = New System.Drawing.Size(61, 4)
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RadioButton15)
        Me.GroupBox1.Controls.Add(Me.RadioButton14)
        Me.GroupBox1.Controls.Add(Me.RadioButton2)
        Me.GroupBox1.Controls.Add(Me.RadioButton3)
        Me.GroupBox1.Controls.Add(Me.RadioButton1)
        Me.GroupBox1.Location = New System.Drawing.Point(524, 9)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Size = New System.Drawing.Size(387, 34)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Path Options"
        '
        'RadioButton15
        '
        Me.RadioButton15.AutoSize = True
        Me.RadioButton15.Location = New System.Drawing.Point(284, 16)
        Me.RadioButton15.Margin = New System.Windows.Forms.Padding(2)
        Me.RadioButton15.Name = "RadioButton15"
        Me.RadioButton15.Size = New System.Drawing.Size(63, 16)
        Me.RadioButton15.TabIndex = 18
        Me.RadioButton15.Text = "Program"
        Me.RadioButton15.UseVisualStyleBackColor = True
        '
        'RadioButton14
        '
        Me.RadioButton14.AutoSize = True
        Me.RadioButton14.Location = New System.Drawing.Point(64, 15)
        Me.RadioButton14.Margin = New System.Windows.Forms.Padding(2)
        Me.RadioButton14.Name = "RadioButton14"
        Me.RadioButton14.Size = New System.Drawing.Size(71, 16)
        Me.RadioButton14.TabIndex = 10
        Me.RadioButton14.TabStop = True
        Me.RadioButton14.Text = "Download"
        Me.RadioButton14.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(201, 16)
        Me.RadioButton2.Margin = New System.Windows.Forms.Padding(2)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(79, 16)
        Me.RadioButton2.TabIndex = 9
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "FileWatcher"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Location = New System.Drawing.Point(138, 15)
        Me.RadioButton3.Margin = New System.Windows.Forms.Padding(2)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(59, 16)
        Me.RadioButton3.TabIndex = 8
        Me.RadioButton3.TabStop = True
        Me.RadioButton3.Text = "Custom"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(4, 14)
        Me.RadioButton1.Margin = New System.Windows.Forms.Padding(2)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(61, 16)
        Me.RadioButton1.TabIndex = 6
        Me.RadioButton1.Text = "Desktop"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Location = New System.Drawing.Point(823, 47)
        Me.CheckBox3.Margin = New System.Windows.Forms.Padding(2)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(90, 16)
        Me.CheckBox3.TabIndex = 17
        Me.CheckBox3.Text = "Subdirectories"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.RadioButton16)
        Me.GroupBox2.Controls.Add(Me.RadioButton10)
        Me.GroupBox2.Controls.Add(Me.RadioButton6)
        Me.GroupBox2.Controls.Add(Me.RadioButton5)
        Me.GroupBox2.Controls.Add(Me.RadioButton4)
        Me.GroupBox2.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.GroupBox2.Location = New System.Drawing.Point(816, 88)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Size = New System.Drawing.Size(104, 79)
        Me.GroupBox2.TabIndex = 8
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Action"
        '
        'RadioButton16
        '
        Me.RadioButton16.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadioButton16.AutoSize = True
        Me.RadioButton16.Location = New System.Drawing.Point(4, 58)
        Me.RadioButton16.Margin = New System.Windows.Forms.Padding(2)
        Me.RadioButton16.Name = "RadioButton16"
        Me.RadioButton16.Size = New System.Drawing.Size(56, 16)
        Me.RadioButton16.TabIndex = 7
        Me.RadioButton16.TabStop = True
        Me.RadioButton16.Text = "Special"
        Me.RadioButton16.UseVisualStyleBackColor = True
        '
        'RadioButton10
        '
        Me.RadioButton10.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadioButton10.AutoSize = True
        Me.RadioButton10.Location = New System.Drawing.Point(51, 20)
        Me.RadioButton10.Margin = New System.Windows.Forms.Padding(2)
        Me.RadioButton10.Name = "RadioButton10"
        Me.RadioButton10.Size = New System.Drawing.Size(52, 16)
        Me.RadioButton10.TabIndex = 6
        Me.RadioButton10.TabStop = True
        Me.RadioButton10.Text = "Delete"
        Me.RadioButton10.UseVisualStyleBackColor = True
        '
        'RadioButton6
        '
        Me.RadioButton6.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadioButton6.AutoSize = True
        Me.RadioButton6.Location = New System.Drawing.Point(51, 40)
        Me.RadioButton6.Margin = New System.Windows.Forms.Padding(2)
        Me.RadioButton6.Name = "RadioButton6"
        Me.RadioButton6.Size = New System.Drawing.Size(49, 16)
        Me.RadioButton6.TabIndex = 2
        Me.RadioButton6.TabStop = True
        Me.RadioButton6.Text = "Copy"
        Me.RadioButton6.UseVisualStyleBackColor = True
        '
        'RadioButton5
        '
        Me.RadioButton5.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadioButton5.AutoSize = True
        Me.RadioButton5.Location = New System.Drawing.Point(4, 39)
        Me.RadioButton5.Margin = New System.Windows.Forms.Padding(2)
        Me.RadioButton5.Name = "RadioButton5"
        Me.RadioButton5.Size = New System.Drawing.Size(50, 16)
        Me.RadioButton5.TabIndex = 1
        Me.RadioButton5.TabStop = True
        Me.RadioButton5.Text = "Move"
        Me.RadioButton5.UseVisualStyleBackColor = True
        '
        'RadioButton4
        '
        Me.RadioButton4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadioButton4.AutoSize = True
        Me.RadioButton4.Checked = True
        Me.RadioButton4.Location = New System.Drawing.Point(3, 20)
        Me.RadioButton4.Margin = New System.Windows.Forms.Padding(2)
        Me.RadioButton4.Name = "RadioButton4"
        Me.RadioButton4.Size = New System.Drawing.Size(48, 16)
        Me.RadioButton4.TabIndex = 0
        Me.RadioButton4.TabStop = True
        Me.RadioButton4.Text = "Open"
        Me.RadioButton4.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 61)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(47, 12)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Target："
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(449, 58)
        Me.Button2.Margin = New System.Windows.Forms.Padding(2)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 20)
        Me.Button2.TabIndex = 12
        Me.Button2.Text = "Open Dialog"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.CheckBox2)
        Me.GroupBox4.Controls.Add(Me.RadioButton11)
        Me.GroupBox4.Controls.Add(Me.RadioButton13)
        Me.GroupBox4.Location = New System.Drawing.Point(524, 47)
        Me.GroupBox4.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox4.Size = New System.Drawing.Size(233, 34)
        Me.GroupBox4.TabIndex = 9
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Target Options"
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(133, 15)
        Me.CheckBox2.Margin = New System.Windows.Forms.Padding(2)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(78, 16)
        Me.CheckBox2.TabIndex = 9
        Me.CheckBox2.Text = "New Folder"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'RadioButton11
        '
        Me.RadioButton11.AutoSize = True
        Me.RadioButton11.Location = New System.Drawing.Point(64, 14)
        Me.RadioButton11.Margin = New System.Windows.Forms.Padding(2)
        Me.RadioButton11.Name = "RadioButton11"
        Me.RadioButton11.Size = New System.Drawing.Size(59, 16)
        Me.RadioButton11.TabIndex = 8
        Me.RadioButton11.Text = "Custom"
        Me.RadioButton11.UseVisualStyleBackColor = True
        '
        'RadioButton13
        '
        Me.RadioButton13.AutoSize = True
        Me.RadioButton13.Location = New System.Drawing.Point(4, 14)
        Me.RadioButton13.Margin = New System.Windows.Forms.Padding(2)
        Me.RadioButton13.Name = "RadioButton13"
        Me.RadioButton13.Size = New System.Drawing.Size(61, 16)
        Me.RadioButton13.TabIndex = 6
        Me.RadioButton13.Text = "Original"
        Me.RadioButton13.UseVisualStyleBackColor = True
        '
        '壓縮檔案集合
        '
        Me.壓縮檔案集合.ContextMenuStrip = Me.ContextMenuStrip3
        Me.壓縮檔案集合.FormattingEnabled = True
        Me.壓縮檔案集合.HorizontalScrollbar = True
        Me.壓縮檔案集合.ItemHeight = 12
        Me.壓縮檔案集合.Location = New System.Drawing.Point(97, 269)
        Me.壓縮檔案集合.Margin = New System.Windows.Forms.Padding(2)
        Me.壓縮檔案集合.Name = "壓縮檔案集合"
        Me.壓縮檔案集合.ScrollAlwaysVisible = True
        Me.壓縮檔案集合.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.壓縮檔案集合.Size = New System.Drawing.Size(334, 112)
        Me.壓縮檔案集合.TabIndex = 13
        '
        'ContextMenuStrip3
        '
        Me.ContextMenuStrip3.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.ContextMenuStrip3.Name = "ContextMenuStrip3"
        Me.ContextMenuStrip3.Size = New System.Drawing.Size(61, 4)
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(820, 406)
        Me.Button4.Margin = New System.Windows.Forms.Padding(2)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(104, 24)
        Me.Button4.TabIndex = 15
        Me.Button4.Text = "執行選擇文件"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("PMingLiU", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label3.Location = New System.Drawing.Point(434, 362)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(103, 12)
        Me.Label3.TabIndex = 19
        Me.Label3.Text = "Change File Name："
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(10, 361)
        Me.Button5.Margin = New System.Windows.Forms.Padding(2)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(38, 25)
        Me.Button5.TabIndex = 22
        Me.Button5.Text = "Add"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(52, 362)
        Me.Button7.Margin = New System.Windows.Forms.Padding(2)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(40, 24)
        Me.Button7.TabIndex = 24
        Me.Button7.Text = "Delete"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'ComboBox2
        '
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(541, 359)
        Me.ComboBox2.Margin = New System.Windows.Forms.Padding(2)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(204, 20)
        Me.ComboBox2.TabIndex = 27
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(10, 251)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(48, 12)
        Me.Label4.TabIndex = 29
        Me.Label4.Text = "Password"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(94, 251)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(61, 12)
        Me.Label5.TabIndex = 30
        Me.Label5.Text = "Archives set"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(9, 83)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(25, 12)
        Me.Label6.TabIndex = 31
        Me.Label6.Text = "Exts"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(96, 83)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(43, 12)
        Me.Label7.TabIndex = 32
        Me.Label7.Text = "Files Set"
        '
        'ListBox1
        '
        Me.ListBox1.BackColor = System.Drawing.SystemColors.Window
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 12
        Me.ListBox1.Location = New System.Drawing.Point(10, 269)
        Me.ListBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(82, 64)
        Me.ListBox1.TabIndex = 33
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(97, 217)
        Me.Button6.Margin = New System.Windows.Forms.Padding(2)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(38, 25)
        Me.Button6.TabIndex = 34
        Me.Button6.Text = "Input"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Button8
        '
        Me.Button8.Location = New System.Drawing.Point(140, 217)
        Me.Button8.Margin = New System.Windows.Forms.Padding(2)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(46, 25)
        Me.Button8.TabIndex = 35
        Me.Button8.Text = "Output"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'Button9
        '
        Me.Button9.Location = New System.Drawing.Point(820, 374)
        Me.Button9.Margin = New System.Windows.Forms.Padding(2)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(104, 25)
        Me.Button9.TabIndex = 36
        Me.Button9.Text = "Download file"
        Me.Button9.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(10, 409)
        Me.ProgressBar1.Margin = New System.Windows.Forms.Padding(2)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(418, 18)
        Me.ProgressBar1.TabIndex = 37
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(435, 412)
        Me.Label8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(32, 12)
        Me.Label8.TabIndex = 38
        Me.Label8.Text = "0/0kb"
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(748, 359)
        Me.TextBox3.Margin = New System.Windows.Forms.Padding(2)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(62, 22)
        Me.TextBox3.TabIndex = 39
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(10, 337)
        Me.TextBox4.Margin = New System.Windows.Forms.Padding(2)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(82, 22)
        Me.TextBox4.TabIndex = 41
        '
        'CheckBox5
        '
        Me.CheckBox5.AutoSize = True
        Me.CheckBox5.Location = New System.Drawing.Point(356, 222)
        Me.CheckBox5.Margin = New System.Windows.Forms.Padding(2)
        Me.CheckBox5.Name = "CheckBox5"
        Me.CheckBox5.Size = New System.Drawing.Size(68, 16)
        Me.CheckBox5.TabIndex = 42
        Me.CheckBox5.Text = "Select All"
        Me.CheckBox5.UseVisualStyleBackColor = True
        '
        'CheckBox6
        '
        Me.CheckBox6.AutoSize = True
        Me.CheckBox6.Location = New System.Drawing.Point(59, 82)
        Me.CheckBox6.Margin = New System.Windows.Forms.Padding(2)
        Me.CheckBox6.Name = "CheckBox6"
        Me.CheckBox6.Size = New System.Drawing.Size(38, 16)
        Me.CheckBox6.TabIndex = 43
        Me.CheckBox6.Text = "All"
        Me.CheckBox6.UseVisualStyleBackColor = True
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(577, 409)
        Me.ComboBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(232, 20)
        Me.ComboBox1.TabIndex = 46
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(530, 413)
        Me.Label9.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(44, 12)
        Me.Label9.TabIndex = 47
        Me.Label9.Text = "Script："
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(436, 91)
        Me.DataGridView1.Margin = New System.Windows.Forms.Padding(2)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersWidth = 51
        Me.DataGridView1.RowTemplate.Height = 27
        Me.DataGridView1.Size = New System.Drawing.Size(375, 266)
        Me.DataGridView1.TabIndex = 48
        Me.DataGridView1.Visible = False
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(762, 43)
        Me.Button3.Margin = New System.Windows.Forms.Padding(2)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(56, 20)
        Me.Button3.TabIndex = 49
        Me.Button3.Text = "previous"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button10
        '
        Me.Button10.Location = New System.Drawing.Point(762, 67)
        Me.Button10.Margin = New System.Windows.Forms.Padding(2)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(56, 18)
        Me.Button10.TabIndex = 50
        Me.Button10.Text = "next"
        Me.Button10.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(821, 70)
        Me.Label10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(39, 12)
        Me.Label10.TabIndex = 52
        Me.Label10.Text = "Page："
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(435, 387)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(50, 12)
        Me.Label11.TabIndex = 54
        Me.Label11.Text = "待機中..."
        '
        'Button11
        '
        Me.Button11.Location = New System.Drawing.Point(755, 278)
        Me.Button11.Margin = New System.Windows.Forms.Padding(2)
        Me.Button11.Name = "Button11"
        Me.Button11.Size = New System.Drawing.Size(54, 21)
        Me.Button11.TabIndex = 56
        Me.Button11.Text = "Remove"
        Me.Button11.UseVisualStyleBackColor = True
        '
        'CheckBox10
        '
        Me.CheckBox10.AutoSize = True
        Me.CheckBox10.Location = New System.Drawing.Point(10, 233)
        Me.CheckBox10.Margin = New System.Windows.Forms.Padding(2)
        Me.CheckBox10.Name = "CheckBox10"
        Me.CheckBox10.Size = New System.Drawing.Size(74, 16)
        Me.CheckBox10.TabIndex = 59
        Me.CheckBox10.Text = "Search Ext"
        Me.CheckBox10.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        Me.OpenFileDialog1.Multiselect = True
        '
        'CheckBox11
        '
        Me.CheckBox11.AutoSize = True
        Me.CheckBox11.Location = New System.Drawing.Point(10, 217)
        Me.CheckBox11.Margin = New System.Windows.Forms.Padding(2)
        Me.CheckBox11.Name = "CheckBox11"
        Me.CheckBox11.Size = New System.Drawing.Size(68, 16)
        Me.CheckBox11.TabIndex = 60
        Me.CheckBox11.Text = "Directory"
        Me.CheckBox11.UseVisualStyleBackColor = True
        '
        'Button12
        '
        Me.Button12.Location = New System.Drawing.Point(298, 381)
        Me.Button12.Margin = New System.Windows.Forms.Padding(2)
        Me.Button12.Name = "Button12"
        Me.Button12.Size = New System.Drawing.Size(52, 25)
        Me.Button12.TabIndex = 57
        Me.Button12.Text = "Remove"
        Me.Button12.UseVisualStyleBackColor = True
        '
        'CheckBox9
        '
        Me.CheckBox9.AutoSize = True
        Me.CheckBox9.Location = New System.Drawing.Point(96, 386)
        Me.CheckBox9.Margin = New System.Windows.Forms.Padding(2)
        Me.CheckBox9.Name = "CheckBox9"
        Me.CheckBox9.Size = New System.Drawing.Size(68, 16)
        Me.CheckBox9.TabIndex = 55
        Me.CheckBox9.Text = "Select All"
        Me.CheckBox9.UseVisualStyleBackColor = True
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(191, 217)
        Me.TextBox5.Margin = New System.Windows.Forms.Padding(2)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(103, 22)
        Me.TextBox5.TabIndex = 62
        '
        'TextBox6
        '
        Me.TextBox6.Location = New System.Drawing.Point(191, 383)
        Me.TextBox6.Margin = New System.Windows.Forms.Padding(2)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(103, 22)
        Me.TextBox6.TabIndex = 63
        '
        'Button16
        '
        Me.Button16.Location = New System.Drawing.Point(425, 21)
        Me.Button16.Margin = New System.Windows.Forms.Padding(2)
        Me.Button16.Name = "Button16"
        Me.Button16.Size = New System.Drawing.Size(20, 18)
        Me.Button16.TabIndex = 69
        Me.Button16.Text = "↑"
        Me.Button16.UseVisualStyleBackColor = True
        '
        'Button17
        '
        Me.Button17.Location = New System.Drawing.Point(425, 58)
        Me.Button17.Margin = New System.Windows.Forms.Padding(2)
        Me.Button17.Name = "Button17"
        Me.Button17.Size = New System.Drawing.Size(20, 18)
        Me.Button17.TabIndex = 70
        Me.Button17.Text = "↑"
        Me.Button17.UseVisualStyleBackColor = True
        '
        'ComboBox3
        '
        Me.ComboBox3.FormattingEnabled = True
        Me.ComboBox3.Location = New System.Drawing.Point(63, 21)
        Me.ComboBox3.Margin = New System.Windows.Forms.Padding(2)
        Me.ComboBox3.Name = "ComboBox3"
        Me.ComboBox3.Size = New System.Drawing.Size(357, 20)
        Me.ComboBox3.TabIndex = 72
        '
        'ComboBox4
        '
        Me.ComboBox4.FormattingEnabled = True
        Me.ComboBox4.Location = New System.Drawing.Point(64, 57)
        Me.ComboBox4.Margin = New System.Windows.Forms.Padding(2)
        Me.ComboBox4.Name = "ComboBox4"
        Me.ComboBox4.Size = New System.Drawing.Size(356, 20)
        Me.ComboBox4.TabIndex = 73
        '
        'ComboBox5
        '
        Me.ComboBox5.FormattingEnabled = True
        Me.ComboBox5.Location = New System.Drawing.Point(821, 223)
        Me.ComboBox5.Margin = New System.Windows.Forms.Padding(2)
        Me.ComboBox5.Name = "ComboBox5"
        Me.ComboBox5.Size = New System.Drawing.Size(99, 20)
        Me.ComboBox5.TabIndex = 74
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(822, 206)
        Me.Label12.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(99, 12)
        Me.Label12.TabIndex = 75
        Me.Label12.Text = "Advanced operation"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(820, 278)
        Me.CheckBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(90, 16)
        Me.CheckBox1.TabIndex = 78
        Me.CheckBox1.Text = "Close Preview"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(174, 86)
        Me.Label13.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(23, 12)
        Me.Label13.TabIndex = 79
        Me.Label13.Text = "0項"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(822, 243)
        Me.Label14.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(76, 12)
        Me.Label14.TabIndex = 81
        Me.Label14.Text = "Output Devices"
        '
        'ComboBox6
        '
        Me.ComboBox6.FormattingEnabled = True
        Me.ComboBox6.Location = New System.Drawing.Point(821, 257)
        Me.ComboBox6.Margin = New System.Windows.Forms.Padding(2)
        Me.ComboBox6.Name = "ComboBox6"
        Me.ComboBox6.Size = New System.Drawing.Size(99, 20)
        Me.ComboBox6.TabIndex = 82
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(736, 339)
        Me.Label15.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(80, 12)
        Me.Label15.TabIndex = 84
        Me.Label15.Text = "0000px 0000px"
        '
        'CheckBox4
        '
        Me.CheckBox4.AutoSize = True
        Me.CheckBox4.Checked = True
        Me.CheckBox4.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox4.Location = New System.Drawing.Point(820, 297)
        Me.CheckBox4.Margin = New System.Windows.Forms.Padding(2)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(72, 16)
        Me.CheckBox4.TabIndex = 87
        Me.CheckBox4.Text = "Auto Save"
        Me.CheckBox4.UseVisualStyleBackColor = True
        '
        'Button13
        '
        Me.Button13.Location = New System.Drawing.Point(376, 80)
        Me.Button13.Margin = New System.Windows.Forms.Padding(2)
        Me.Button13.Name = "Button13"
        Me.Button13.Size = New System.Drawing.Size(56, 18)
        Me.Button13.TabIndex = 88
        Me.Button13.Text = "Close"
        Me.Button13.UseVisualStyleBackColor = True
        '
        'Button18
        '
        Me.Button18.Location = New System.Drawing.Point(689, 383)
        Me.Button18.Margin = New System.Windows.Forms.Padding(2)
        Me.Button18.Name = "Button18"
        Me.Button18.Size = New System.Drawing.Size(122, 23)
        Me.Button18.TabIndex = 89
        Me.Button18.Text = "Buy me a coffee"
        Me.Button18.UseVisualStyleBackColor = True
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(822, 172)
        Me.Label16.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(72, 12)
        Me.Label16.TabIndex = 91
        Me.Label16.Text = "Special Action"
        '
        'ComboBox7
        '
        Me.ComboBox7.FormattingEnabled = True
        Me.ComboBox7.Location = New System.Drawing.Point(821, 186)
        Me.ComboBox7.Margin = New System.Windows.Forms.Padding(2)
        Me.ComboBox7.Name = "ComboBox7"
        Me.ComboBox7.Size = New System.Drawing.Size(99, 20)
        Me.ComboBox7.TabIndex = 92
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(174, 251)
        Me.Label17.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(23, 12)
        Me.Label17.TabIndex = 93
        Me.Label17.Text = "0項"
        '
        'Button20
        '
        Me.Button20.Location = New System.Drawing.Point(376, 382)
        Me.Button20.Margin = New System.Windows.Forms.Padding(2)
        Me.Button20.Name = "Button20"
        Me.Button20.Size = New System.Drawing.Size(52, 25)
        Me.Button20.TabIndex = 94
        Me.Button20.Text = "Refresh"
        Me.Button20.UseVisualStyleBackColor = True
        '
        'CheckBox7
        '
        Me.CheckBox7.AutoSize = True
        Me.CheckBox7.Location = New System.Drawing.Point(820, 315)
        Me.CheckBox7.Margin = New System.Windows.Forms.Padding(2)
        Me.CheckBox7.Name = "CheckBox7"
        Me.CheckBox7.Size = New System.Drawing.Size(48, 16)
        Me.CheckBox7.TabIndex = 95
        Me.CheckBox7.Text = "多開"
        Me.CheckBox7.UseVisualStyleBackColor = True
        '
        'TreeView1
        '
        Me.TreeView1.Location = New System.Drawing.Point(95, 102)
        Me.TreeView1.Margin = New System.Windows.Forms.Padding(2)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.Size = New System.Drawing.Size(720, 250)
        Me.TreeView1.TabIndex = 96
        Me.TreeView1.Visible = False
        '
        'AxWindowsMediaPlayer1
        '
        Me.AxWindowsMediaPlayer1.Enabled = True
        Me.AxWindowsMediaPlayer1.Location = New System.Drawing.Point(436, 91)
        Me.AxWindowsMediaPlayer1.Margin = New System.Windows.Forms.Padding(2)
        Me.AxWindowsMediaPlayer1.Name = "AxWindowsMediaPlayer1"
        Me.AxWindowsMediaPlayer1.OcxState = CType(resources.GetObject("AxWindowsMediaPlayer1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxWindowsMediaPlayer1.Size = New System.Drawing.Size(378, 266)
        Me.AxWindowsMediaPlayer1.TabIndex = 71
        Me.AxWindowsMediaPlayer1.Visible = False
        '
        'Button19
        '
        Me.Button19.Location = New System.Drawing.Point(356, 243)
        Me.Button19.Name = "Button19"
        Me.Button19.Size = New System.Drawing.Size(64, 23)
        Me.Button19.TabIndex = 97
        Me.Button19.Text = "測試紐"
        Me.Button19.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AllowDrop = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(931, 436)
        Me.Controls.Add(Me.Button19)
        Me.Controls.Add(Me.CheckBox7)
        Me.Controls.Add(Me.Button20)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.ComboBox7)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.Button18)
        Me.Controls.Add(Me.Button13)
        Me.Controls.Add(Me.CheckBox4)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.CheckBox3)
        Me.Controls.Add(Me.ComboBox6)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.ComboBox5)
        Me.Controls.Add(Me.ComboBox4)
        Me.Controls.Add(Me.ComboBox3)
        Me.Controls.Add(Me.Button17)
        Me.Controls.Add(Me.Button16)
        Me.Controls.Add(Me.TextBox6)
        Me.Controls.Add(Me.CheckBox11)
        Me.Controls.Add(Me.TextBox5)
        Me.Controls.Add(Me.CheckBox10)
        Me.Controls.Add(Me.Button12)
        Me.Controls.Add(Me.Button11)
        Me.Controls.Add(Me.CheckBox9)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Button10)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.CheckBox6)
        Me.Controls.Add(Me.CheckBox5)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Button9)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.壓縮檔案集合)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.目錄檔案集合)
        Me.Controls.Add(Me.副檔名列表)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.AxWindowsMediaPlayer1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.TreeView1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "Form1"
        Me.Text = "EasyPreview_All-seeing file"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxWindowsMediaPlayer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents 副檔名列表 As CheckedListBox
    Friend WithEvents 目錄檔案集合 As ListBox
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents RadioButton10 As RadioButton
    Friend WithEvents RadioButton6 As RadioButton
    Friend WithEvents RadioButton5 As RadioButton
    Friend WithEvents RadioButton4 As RadioButton
    Friend WithEvents Label2 As Label
    Friend WithEvents Button2 As Button
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents CheckBox2 As CheckBox
    Friend WithEvents 壓縮檔案集合 As ListBox
    Friend WithEvents Button4 As Button
    Friend WithEvents CheckBox3 As CheckBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Button5 As Button
    Friend WithEvents Button7 As Button
    Friend WithEvents ComboBox2 As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents RadioButton2 As RadioButton
    Friend WithEvents Button6 As Button
    Friend WithEvents Button8 As Button
    Friend WithEvents Button9 As Button
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents Label8 As Label
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents TextBox4 As TextBox
    Friend WithEvents CheckBox5 As CheckBox
    Friend WithEvents RadioButton3 As RadioButton
    Friend WithEvents RadioButton1 As RadioButton
    Friend WithEvents RadioButton11 As RadioButton
    Friend WithEvents RadioButton13 As RadioButton
    Friend WithEvents CheckBox6 As CheckBox

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            Button4.Text = "執行全部文件"
        Else
            Button4.Text = "執行選擇文件"
        End If
    End Sub



    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles PictureBox1.DoubleClick
        Form2.PictureMode()
        Form2.PictureBox1.Image = CType(PictureBox1.Image.Clone, Image)
        transform_index_status_to_Form2()
        'For i = Form6.undoLabel + 1 To Form6.undoList.Count - 1
        '    Form6.ListBox1.Items.RemoveAt(i)

        'Next
        'Form6.ListBox1.Items.Remove(Form2.Index)
        'Form6.ListBox1.Items.Add(Form2.Index)
        'Form6.undoList.Add(Form6.ListBox1.Items.Cast(Of String).ToArray)
        Form2.Show()
        Hide()
    End Sub
    Private Sub transform_index_status_to_Form2()

        If 壓縮檔案集合.Items.Count = 0 And 壓縮檔案集合.SelectedItem = Nothing Then
            Form2.Index = 目錄檔案集合.SelectedIndex
            Form2.TextBox2.Text = 目錄檔案集合.SelectedIndex
            Form2.Label2.Text = "/" & 目錄檔案集合.Items.Count - 1
            Form2.Text = 目錄檔案集合.SelectedItem
        Else
            Form2.Index = 壓縮檔案集合.SelectedIndex
            Form2.TextBox2.Text = 壓縮檔案集合.SelectedIndex
            Form2.Label2.Text = "/" & 壓縮檔案集合.Items.Count - 1
            Form2.Text = 壓縮檔案集合.SelectedItem
        End If
    End Sub

    Public Sub Button11_Click_1(sender As Object, e As EventArgs) Handles Button11.Click
        Dim now_select = 0
        Dim list_backup As String() = New String() {0}

        If CheckBox5.Checked = True Then
            list_backup = 目錄檔案集合.Items.Cast(Of String).ToArray()


        Else
            list_backup = 目錄檔案集合.SelectedItems.Cast(Of String).ToArray()
            now_select = 目錄檔案集合.SelectedIndices(0)
        End If

        For Each i In list_backup
            目錄檔案集合.Items.Remove(i)
            If Not String.IsNullOrEmpty(TextBox5.Text) Then backup.Remove(i)
        Next

        If 目錄檔案集合.Items.Count <> 0 Then
            Try
                目錄檔案集合.SelectedIndex = now_select
            Catch
                目錄檔案集合.SelectedIndex = now_select - 1
            End Try
        Else
            'remove bug
            refresh_backup()




        End If


    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        壓縮檔案集合.Items.RemoveAt(壓縮檔案集合.SelectedIndex)
    End Sub


    Public Sub Add_additional_operation()
        '進階操作的部分
        ComboBox5.Items.Add("Do not perform advanced operations")
        ComboBox5.Items.Add("locale emulator")
        ComboBox5.Items.Add("自動重播(單首)")
        ComboBox5.Items.Add("自動重播(全部)")
        ComboBox5.Items.Add("執行後移出")
        ComboBox5.Items.Add("執行後移到最後一項")
        'ComboBox5.Items.Add("Delete After Compress/Decompress")
        ComboBox5.Items.Add("執行壓縮檔案")
        'ComboBox5.Items.Add("解壓分割")
        ComboBox5.Items.Add("更改附檔名")
        'ComboBox5.Items.Add("clip啟動")
        ComboBox5.Items.Add("MuseScore啟動")
        ComboBox5.Items.Add("紀錄開啟檔案")
        ComboBox5.Items.Add("取樣")


        ComboBox5.Items.Add("Combine trf files")
        ComboBox5.Items.Add("下載在combobox4中")
        'ComboBox5.Items.Add("顯示TreeView")
        '爬蟲網站的部分
        'ComboBox6.Items.Add("kemono.party")
        'ComboBox6.Items.Add("18comic")
        'ComboBox6.Items.Add("ehentai")
        'ComboBox6.Items.Add("twitter")
        'ComboBox6.Items.Add("pixiv")

        '可用裝置


        '特殊操作的部分
        ComboBox7.Items.Add("Rename")
        ComboBox7.Items.Add("Compression")
        ComboBox7.Items.Add("Decompression")
        ComboBox7.Items.Add("Set Up")
        ComboBox7.Items.Add("Synchronize")
        ComboBox7.Items.Add("檢查路徑")
        ComboBox7.Items.Add("檢查網址")
        ComboBox7.Items.Add("執行爬蟲指令")
        ComboBox7.Items.Add("建立ASMR清單")
    End Sub

    Public Sub Form1_Index_DragDrop(sender As System.Object, e As System.Windows.Forms.DragEventArgs) Handles 目錄檔案集合.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        Dim arguments As String = ""
        For Each sourcrPath In files
            If 目錄檔案集合.Items.Contains(sourcrPath) Then 目錄檔案集合.Items.Remove(sourcrPath)
            'If sourcrPath.Contains(".trf") Then
            'read_text_data(sourcrPath.Trim)
            'Continue For
            'End If
            目錄檔案集合.Items.Add(sourcrPath)
            If Not String.IsNullOrEmpty(TextBox5.Text) Then
                backup.Add(sourcrPath)
            End If
            arguments += $" ""{ sourcrPath}"""
            If RadioButton6.Checked Then
                My.Computer.FileSystem.CopyFile(sourcrPath, ComboBox4.Text & "\" & Path.GetFileName(sourcrPath), UIOption.AllDialogs)
            ElseIf RadioButton5.Checked Then
                My.Computer.FileSystem.MoveFile(sourcrPath, ComboBox4.Text & "\" & Path.GetFileName(sourcrPath), UIOption.AllDialogs)
            End If
        Next




        If ComboBox5.SelectedItem = "clip啟動" Then
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
        refresh_backup()
    End Sub
    Private Sub Form1_text1_DragDrop(sender As System.Object, e As System.Windows.Forms.DragEventArgs) Handles ComboBox3.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        ComboBox3.Text = files(0)
        RadioButton3.Checked = True
        Dim select_indeces = 副檔名列表.CheckedIndices.Cast(Of Integer)().ToArray()
        Call ChecklistBox_Click(sender, e)
        For Each sel In select_indeces
            副檔名列表.SetItemChecked(sel, True)
        Next
        Call ListBox_Click(sender, e)
        Add_new_path(ComboBox3, "ori_pathrecord.txt")
    End Sub
    Private Sub Form1_text2_DragDrop(sender As System.Object, e As System.Windows.Forms.DragEventArgs) Handles ComboBox4.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        ComboBox4.Text = files(0)
        Add_new_path(ComboBox4, "des_pathrecord.txt")
    End Sub

    Private Sub Form1_DragEnter(sender As System.Object, e As System.Windows.Forms.DragEventArgs) Handles 目錄檔案集合.DragEnter, ComboBox3.DragEnter, ComboBox4.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub
    Private Sub checkbox10_Click(sender As Object, e As EventArgs) Handles CheckBox10.Click
        If CheckBox10.Checked = False Then
            Label6.Text = "目錄類型"
        Else
            Label6.Text = "副檔名"
        End If
        ChecklistBox_Click(sender, e)
    End Sub

    Private Sub RadioButton15_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton15.CheckedChanged
        allProcesses = Process.GetProcesses()
        For Each p As Process In allProcesses
            目錄檔案集合.Items.Add(p.ProcessName)

        Next
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        If Not String.IsNullOrEmpty(ComboBox3.Text) Then
            ComboBox3.Text = Directory.GetParent(ComboBox3.Text).ToString
        End If
    End Sub
    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        If Not String.IsNullOrEmpty(ComboBox4.Text) Then
            ComboBox4.Text = Directory.GetParent(ComboBox4.Text).ToString
        End If
    End Sub
    Private Sub editItem()
        Form8.Show()
    End Sub
    Private Sub addItem_from_clipboard()
        Form4.Show()
        Form4.tarForm = 目錄檔案集合
    End Sub
    Private Sub addItem_from_clipboard_to_archive()
        Form4.Show()
        Form4.tarForm = 壓縮檔案集合
    End Sub

    Public Sub Clipboard_SetText()
        Dim allString = ""
        Dim tar_array = 目錄檔案集合.SelectedItems.Cast(Of String).ToArray
        If CheckBox5.Checked Then
            tar_array = 目錄檔案集合.Items.Cast(Of String).ToArray
        End If
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
    Private Sub Clipboard_SetText_Compress()

        '加1個把數組轉成換行文字
        Clipboard.SetText(壓縮檔案集合.SelectedItem.ToString)
    End Sub
    Private Sub Clipboard_SetIImage()
        '加1個把數組轉成換行文字
        Clipboard.SetData(DataFormats.Bitmap, PictureBox1.Image)
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Form1_Close(sender, e)
        目錄檔案集合.Items.Clear()
        'CheckBox4.Checked = False
        refresh_listbox_numbers()
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Process.Start("https://www.buymeacoffee.com/b200077")
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs)
        Form5.Show()
    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        RadioButton16.Checked = True
        If ComboBox7.Text = "Synchronize" Then
            壓縮檔案集合.Items.Clear()
            For Each foundFile As String In My.Computer.FileSystem.GetFiles(
                    ComboBox4.Text, searchOptionCheck())
                壓縮檔案集合.Items.Add(foundFile)
            Next
            For Each foundFile As String In My.Computer.FileSystem.GetFiles(
                    ComboBox3.Text, searchOptionCheck())
                目錄檔案集合.Items.Add(foundFile)
            Next
        End If
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Form2.Show()
        Form2.WebMode("https://www.google.com/")
        WmcRefresh()
    End Sub

    Private Sub ComboBox6_click(sender As Object, e As EventArgs) Handles ComboBox6.Click
        If ComboBox6.Items.Count <> 0 Then Return
        Using AudioController As New CoreAudioController()
            Dim devices As IEnumerable(Of CoreAudioDevice) = AudioController.GetPlaybackDevices()
            For Each device In devices
                ComboBox6.Items.Add(device.FullName)
            Next
            ComboBox6.SelectedItem = AudioController.DefaultPlaybackDevice.FullName
        End Using
    End Sub
    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        Using controller As New CoreAudioController()
            Dim devices As IEnumerable(Of CoreAudioDevice) = controller.GetPlaybackDevices()
            devices(ComboBox6.SelectedIndex).SetAsDefault()
        End Using
    End Sub

    Private Sub measureDevice(sender As Object, e As EventArgs, deviceName As String)
        Using controller As New CoreAudioController()
            If Not controller.DefaultPlaybackDevice.FullName.Contains(deviceName) Then
                If Not ComboBox6.SelectedItem.contains(deviceName) Then
                    ComboBox6.SelectedIndex = ComboBox6.FindString(deviceName)
                Else
                    ComboBox6_SelectedIndexChanged(sender, e)
                End If
            End If
        End Using
    End Sub
    Private Sub AxWindowsMediaPlayer1_Enter(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Space) Then
            If AxWindowsMediaPlayer1.playState = WMPLib.WMPPlayState.wmppsPlaying Then
                AxWindowsMediaPlayer1.Ctlcontrols.pause()
            ElseIf AxWindowsMediaPlayer1.playState = WMPLib.WMPPlayState.wmppsPaused Then
                measureDevice(sender, e, "耳機")
                AxWindowsMediaPlayer1.Ctlcontrols.play()
            End If
        End If
        If e.KeyChar = Convert.ToChar(Keys.Down) Then
            目錄檔案集合.SelectedIndex = 目錄檔案集合.SelectedIndex + 1
            目錄檔案集合.SetSelected(目錄檔案集合.SelectedIndex, False)
        End If
        If e.KeyChar = Convert.ToChar(Keys.Up) Then
            目錄檔案集合.SelectedIndex = 目錄檔案集合.SelectedIndex - 1
            目錄檔案集合.SetSelected(目錄檔案集合.SelectedIndex + 1, False)
        End If
        If e.KeyChar = Convert.ToChar(Keys.Enter) Then
            Button4_Action(sender, New EventArgs)
        End If
    End Sub

    Private Sub Button19_Click_1(sender As Object, e As EventArgs) Handles Button19.Click
        Form5.Show()
    End Sub
End Class

