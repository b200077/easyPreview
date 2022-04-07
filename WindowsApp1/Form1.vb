

Imports System.IO

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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.副檔名列表 = New System.Windows.Forms.CheckedListBox()
        Me.目錄檔案集合 = New System.Windows.Forms.ListBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RadioButton15 = New System.Windows.Forms.RadioButton()
        Me.RadioButton14 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton3 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.RadioButton12 = New System.Windows.Forms.RadioButton()
        Me.RadioButton10 = New System.Windows.Forms.RadioButton()
        Me.RadioButton9 = New System.Windows.Forms.RadioButton()
        Me.RadioButton8 = New System.Windows.Forms.RadioButton()
        Me.RadioButton7 = New System.Windows.Forms.RadioButton()
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
        Me.Button13 = New System.Windows.Forms.Button()
        Me.Button16 = New System.Windows.Forms.Button()
        Me.Button17 = New System.Windows.Forms.Button()
        Me.AxWindowsMediaPlayer1 = New AxWMPLib.AxWindowsMediaPlayer()
        Me.ComboBox3 = New System.Windows.Forms.ComboBox()
        Me.ComboBox4 = New System.Windows.Forms.ComboBox()
        Me.ComboBox5 = New System.Windows.Forms.ComboBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Button18 = New System.Windows.Forms.Button()
        Me.Button19 = New System.Windows.Forms.Button()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Label13 = New System.Windows.Forms.Label()
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
        Me.Button1.Location = New System.Drawing.Point(599, 24)
        Me.Button1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(100, 25)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "開啟資料夾"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(27, 28)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(52, 15)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "路徑："
        '
        '副檔名列表
        '
        Me.副檔名列表.FormattingEnabled = True
        Me.副檔名列表.Location = New System.Drawing.Point(13, 126)
        Me.副檔名列表.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.副檔名列表.Name = "副檔名列表"
        Me.副檔名列表.Size = New System.Drawing.Size(111, 124)
        Me.副檔名列表.TabIndex = 0
        '
        '目錄檔案集合
        '
        Me.目錄檔案集合.AllowDrop = True
        Me.目錄檔案集合.FormattingEnabled = True
        Me.目錄檔案集合.HorizontalScrollbar = True
        Me.目錄檔案集合.ItemHeight = 15
        Me.目錄檔案集合.Location = New System.Drawing.Point(129, 126)
        Me.目錄檔案集合.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.目錄檔案集合.Name = "目錄檔案集合"
        Me.目錄檔案集合.ScrollAlwaysVisible = True
        Me.目錄檔案集合.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.目錄檔案集合.Size = New System.Drawing.Size(447, 139)
        Me.目錄檔案集合.TabIndex = 3
        '
        'PictureBox1
        '
        Me.PictureBox1.Location = New System.Drawing.Point(581, 112)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(501, 334)
        Me.PictureBox1.TabIndex = 4
        Me.PictureBox1.TabStop = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RadioButton15)
        Me.GroupBox1.Controls.Add(Me.RadioButton14)
        Me.GroupBox1.Controls.Add(Me.RadioButton2)
        Me.GroupBox1.Controls.Add(Me.RadioButton3)
        Me.GroupBox1.Controls.Add(Me.RadioButton1)
        Me.GroupBox1.Controls.Add(Me.CheckBox3)
        Me.GroupBox1.Location = New System.Drawing.Point(699, 11)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(516, 42)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "路徑選項"
        '
        'RadioButton15
        '
        Me.RadioButton15.AutoSize = True
        Me.RadioButton15.Location = New System.Drawing.Point(316, 18)
        Me.RadioButton15.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RadioButton15.Name = "RadioButton15"
        Me.RadioButton15.Size = New System.Drawing.Size(73, 19)
        Me.RadioButton15.TabIndex = 18
        Me.RadioButton15.Text = "從程序"
        Me.RadioButton15.UseVisualStyleBackColor = True
        '
        'RadioButton14
        '
        Me.RadioButton14.AutoSize = True
        Me.RadioButton14.Location = New System.Drawing.Point(71, 16)
        Me.RadioButton14.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RadioButton14.Name = "RadioButton14"
        Me.RadioButton14.Size = New System.Drawing.Size(58, 19)
        Me.RadioButton14.TabIndex = 10
        Me.RadioButton14.TabStop = True
        Me.RadioButton14.Text = "下載"
        Me.RadioButton14.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(205, 16)
        Me.RadioButton2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(97, 19)
        Me.RadioButton2.TabIndex = 9
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "FileWatcher"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Location = New System.Drawing.Point(139, 16)
        Me.RadioButton3.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(58, 19)
        Me.RadioButton3.TabIndex = 8
        Me.RadioButton3.TabStop = True
        Me.RadioButton3.Text = "自訂"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(5, 18)
        Me.RadioButton1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(58, 19)
        Me.RadioButton1.TabIndex = 6
        Me.RadioButton1.Text = "桌面"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Location = New System.Drawing.Point(395, 18)
        Me.CheckBox3.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(104, 19)
        Me.CheckBox3.TabIndex = 17
        Me.CheckBox3.Text = "包含子目錄"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.RadioButton12)
        Me.GroupBox2.Controls.Add(Me.RadioButton10)
        Me.GroupBox2.Controls.Add(Me.RadioButton9)
        Me.GroupBox2.Controls.Add(Me.RadioButton8)
        Me.GroupBox2.Controls.Add(Me.RadioButton7)
        Me.GroupBox2.Controls.Add(Me.RadioButton6)
        Me.GroupBox2.Controls.Add(Me.RadioButton5)
        Me.GroupBox2.Controls.Add(Me.RadioButton4)
        Me.GroupBox2.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.GroupBox2.Location = New System.Drawing.Point(1088, 112)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox2.Size = New System.Drawing.Size(144, 152)
        Me.GroupBox2.TabIndex = 8
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "動作"
        '
        'RadioButton12
        '
        Me.RadioButton12.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadioButton12.AutoSize = True
        Me.RadioButton12.Location = New System.Drawing.Point(15, 125)
        Me.RadioButton12.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RadioButton12.Name = "RadioButton12"
        Me.RadioButton12.Size = New System.Drawing.Size(58, 19)
        Me.RadioButton12.TabIndex = 7
        Me.RadioButton12.TabStop = True
        Me.RadioButton12.Text = "安裝"
        Me.RadioButton12.UseVisualStyleBackColor = True
        '
        'RadioButton10
        '
        Me.RadioButton10.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadioButton10.AutoSize = True
        Me.RadioButton10.Location = New System.Drawing.Point(79, 25)
        Me.RadioButton10.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RadioButton10.Name = "RadioButton10"
        Me.RadioButton10.Size = New System.Drawing.Size(58, 19)
        Me.RadioButton10.TabIndex = 6
        Me.RadioButton10.TabStop = True
        Me.RadioButton10.Text = "刪除"
        Me.RadioButton10.UseVisualStyleBackColor = True
        '
        'RadioButton9
        '
        Me.RadioButton9.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadioButton9.AutoSize = True
        Me.RadioButton9.Location = New System.Drawing.Point(79, 100)
        Me.RadioButton9.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RadioButton9.Name = "RadioButton9"
        Me.RadioButton9.Size = New System.Drawing.Size(58, 19)
        Me.RadioButton9.TabIndex = 5
        Me.RadioButton9.TabStop = True
        Me.RadioButton9.Text = "解壓"
        Me.RadioButton9.UseVisualStyleBackColor = True
        '
        'RadioButton8
        '
        Me.RadioButton8.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadioButton8.AutoSize = True
        Me.RadioButton8.Location = New System.Drawing.Point(15, 100)
        Me.RadioButton8.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RadioButton8.Name = "RadioButton8"
        Me.RadioButton8.Size = New System.Drawing.Size(58, 19)
        Me.RadioButton8.TabIndex = 4
        Me.RadioButton8.TabStop = True
        Me.RadioButton8.Text = "壓縮"
        Me.RadioButton8.UseVisualStyleBackColor = True
        '
        'RadioButton7
        '
        Me.RadioButton7.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadioButton7.AutoSize = True
        Me.RadioButton7.Location = New System.Drawing.Point(16, 75)
        Me.RadioButton7.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RadioButton7.Name = "RadioButton7"
        Me.RadioButton7.Size = New System.Drawing.Size(88, 19)
        Me.RadioButton7.TabIndex = 3
        Me.RadioButton7.TabStop = True
        Me.RadioButton7.Text = "重新命名"
        Me.RadioButton7.UseVisualStyleBackColor = True
        '
        'RadioButton6
        '
        Me.RadioButton6.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadioButton6.AutoSize = True
        Me.RadioButton6.Location = New System.Drawing.Point(79, 50)
        Me.RadioButton6.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RadioButton6.Name = "RadioButton6"
        Me.RadioButton6.Size = New System.Drawing.Size(58, 19)
        Me.RadioButton6.TabIndex = 2
        Me.RadioButton6.TabStop = True
        Me.RadioButton6.Text = "複製"
        Me.RadioButton6.UseVisualStyleBackColor = True
        '
        'RadioButton5
        '
        Me.RadioButton5.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadioButton5.AutoSize = True
        Me.RadioButton5.Location = New System.Drawing.Point(15, 50)
        Me.RadioButton5.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RadioButton5.Name = "RadioButton5"
        Me.RadioButton5.Size = New System.Drawing.Size(58, 19)
        Me.RadioButton5.TabIndex = 1
        Me.RadioButton5.TabStop = True
        Me.RadioButton5.Text = "移動"
        Me.RadioButton5.UseVisualStyleBackColor = True
        '
        'RadioButton4
        '
        Me.RadioButton4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadioButton4.AutoSize = True
        Me.RadioButton4.Checked = True
        Me.RadioButton4.Location = New System.Drawing.Point(15, 25)
        Me.RadioButton4.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RadioButton4.Name = "RadioButton4"
        Me.RadioButton4.Size = New System.Drawing.Size(58, 19)
        Me.RadioButton4.TabIndex = 0
        Me.RadioButton4.TabStop = True
        Me.RadioButton4.Text = "開啟"
        Me.RadioButton4.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(27, 76)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(52, 15)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "目標："
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(599, 72)
        Me.Button2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(100, 25)
        Me.Button2.TabIndex = 12
        Me.Button2.Text = "開啟資料夾"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.CheckBox2)
        Me.GroupBox4.Controls.Add(Me.RadioButton11)
        Me.GroupBox4.Controls.Add(Me.RadioButton13)
        Me.GroupBox4.Location = New System.Drawing.Point(699, 59)
        Me.GroupBox4.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox4.Size = New System.Drawing.Size(229, 42)
        Me.GroupBox4.TabIndex = 9
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "目標選項"
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(139, 16)
        Me.CheckBox2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(89, 19)
        Me.CheckBox2.TabIndex = 9
        Me.CheckBox2.Text = "新資料夾"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'RadioButton11
        '
        Me.RadioButton11.AutoSize = True
        Me.RadioButton11.Location = New System.Drawing.Point(75, 18)
        Me.RadioButton11.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RadioButton11.Name = "RadioButton11"
        Me.RadioButton11.Size = New System.Drawing.Size(58, 19)
        Me.RadioButton11.TabIndex = 8
        Me.RadioButton11.Text = "自訂"
        Me.RadioButton11.UseVisualStyleBackColor = True
        '
        'RadioButton13
        '
        Me.RadioButton13.AutoSize = True
        Me.RadioButton13.Location = New System.Drawing.Point(5, 18)
        Me.RadioButton13.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RadioButton13.Name = "RadioButton13"
        Me.RadioButton13.Size = New System.Drawing.Size(73, 19)
        Me.RadioButton13.TabIndex = 6
        Me.RadioButton13.Text = "原路徑"
        Me.RadioButton13.UseVisualStyleBackColor = True
        '
        '壓縮檔案集合
        '
        Me.壓縮檔案集合.FormattingEnabled = True
        Me.壓縮檔案集合.HorizontalScrollbar = True
        Me.壓縮檔案集合.ItemHeight = 15
        Me.壓縮檔案集合.Location = New System.Drawing.Point(129, 336)
        Me.壓縮檔案集合.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.壓縮檔案集合.Name = "壓縮檔案集合"
        Me.壓縮檔案集合.ScrollAlwaysVisible = True
        Me.壓縮檔案集合.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.壓縮檔案集合.Size = New System.Drawing.Size(444, 139)
        Me.壓縮檔案集合.TabIndex = 13
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(1093, 508)
        Me.Button4.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(139, 30)
        Me.Button4.TabIndex = 15
        Me.Button4.Text = "執行選擇文件"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(579, 452)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 15)
        Me.Label3.TabIndex = 19
        Me.Label3.Text = "更改檔案名稱："
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(13, 451)
        Me.Button5.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(51, 31)
        Me.Button5.TabIndex = 22
        Me.Button5.Text = "新增"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(69, 452)
        Me.Button7.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(53, 30)
        Me.Button7.TabIndex = 24
        Me.Button7.Text = "刪除"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'ComboBox2
        '
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(684, 449)
        Me.ComboBox2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(308, 23)
        Me.ComboBox2.TabIndex = 27
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(13, 314)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(67, 15)
        Me.Label4.TabIndex = 29
        Me.Label4.Text = "壓縮密碼"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(125, 314)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(97, 15)
        Me.Label5.TabIndex = 30
        Me.Label5.Text = "壓縮檔案預覽"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(11, 108)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(52, 15)
        Me.Label6.TabIndex = 31
        Me.Label6.Text = "附檔名"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(125, 108)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(67, 15)
        Me.Label7.TabIndex = 32
        Me.Label7.Text = "文件集合"
        '
        'ListBox1
        '
        Me.ListBox1.BackColor = System.Drawing.SystemColors.Window
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 15
        Me.ListBox1.Location = New System.Drawing.Point(13, 336)
        Me.ListBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(108, 79)
        Me.ListBox1.TabIndex = 33
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(129, 271)
        Me.Button6.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(51, 31)
        Me.Button6.TabIndex = 34
        Me.Button6.Text = "輸入"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Button8
        '
        Me.Button8.Location = New System.Drawing.Point(187, 271)
        Me.Button8.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(51, 31)
        Me.Button8.TabIndex = 35
        Me.Button8.Text = "輸出"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'Button9
        '
        Me.Button9.Location = New System.Drawing.Point(1093, 468)
        Me.Button9.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(139, 31)
        Me.Button9.TabIndex = 36
        Me.Button9.Text = "下載檔案"
        Me.Button9.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(13, 511)
        Me.ProgressBar1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(557, 22)
        Me.ProgressBar1.TabIndex = 37
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(580, 515)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(39, 15)
        Me.Label8.TabIndex = 38
        Me.Label8.Text = "0/0kb"
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(997, 449)
        Me.TextBox3.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(81, 25)
        Me.TextBox3.TabIndex = 39
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(13, 421)
        Me.TextBox4.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(108, 25)
        Me.TextBox4.TabIndex = 41
        '
        'CheckBox5
        '
        Me.CheckBox5.AutoSize = True
        Me.CheckBox5.Location = New System.Drawing.Point(453, 278)
        Me.CheckBox5.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CheckBox5.Name = "CheckBox5"
        Me.CheckBox5.Size = New System.Drawing.Size(119, 19)
        Me.CheckBox5.TabIndex = 42
        Me.CheckBox5.Text = "執行所有文件"
        Me.CheckBox5.UseVisualStyleBackColor = True
        '
        'CheckBox6
        '
        Me.CheckBox6.AutoSize = True
        Me.CheckBox6.Location = New System.Drawing.Point(61, 105)
        Me.CheckBox6.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CheckBox6.Name = "CheckBox6"
        Me.CheckBox6.Size = New System.Drawing.Size(59, 19)
        Me.CheckBox6.TabIndex = 43
        Me.CheckBox6.Text = "全選"
        Me.CheckBox6.UseVisualStyleBackColor = True
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(769, 511)
        Me.ComboBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(308, 23)
        Me.ComboBox1.TabIndex = 46
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(696, 515)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(59, 15)
        Me.Label9.TabIndex = 47
        Me.Label9.Text = "cookie："
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(581, 114)
        Me.DataGridView1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersWidth = 51
        Me.DataGridView1.RowTemplate.Height = 27
        Me.DataGridView1.Size = New System.Drawing.Size(500, 332)
        Me.DataGridView1.TabIndex = 48
        Me.DataGridView1.Visible = False
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(935, 76)
        Me.Button3.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 31)
        Me.Button3.TabIndex = 49
        Me.Button3.Text = "上一頁"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button10
        '
        Me.Button10.Location = New System.Drawing.Point(1016, 75)
        Me.Button10.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(75, 32)
        Me.Button10.TabIndex = 50
        Me.Button10.Text = "下一頁"
        Me.Button10.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(1097, 84)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(52, 15)
        Me.Label10.TabIndex = 52
        Me.Label10.Text = "頁數："
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(580, 484)
        Me.Label11.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(64, 15)
        Me.Label11.TabIndex = 54
        Me.Label11.Text = "待機中..."
        '
        'Button11
        '
        Me.Button11.Location = New System.Drawing.Point(397, 271)
        Me.Button11.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button11.Name = "Button11"
        Me.Button11.Size = New System.Drawing.Size(51, 31)
        Me.Button11.TabIndex = 56
        Me.Button11.Text = "移出"
        Me.Button11.UseVisualStyleBackColor = True
        '
        'CheckBox10
        '
        Me.CheckBox10.AutoSize = True
        Me.CheckBox10.Location = New System.Drawing.Point(12, 291)
        Me.CheckBox10.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CheckBox10.Name = "CheckBox10"
        Me.CheckBox10.Size = New System.Drawing.Size(104, 19)
        Me.CheckBox10.TabIndex = 59
        Me.CheckBox10.Text = "顯示附檔名"
        Me.CheckBox10.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'CheckBox11
        '
        Me.CheckBox11.AutoSize = True
        Me.CheckBox11.Location = New System.Drawing.Point(13, 266)
        Me.CheckBox11.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CheckBox11.Name = "CheckBox11"
        Me.CheckBox11.Size = New System.Drawing.Size(89, 19)
        Me.CheckBox11.TabIndex = 60
        Me.CheckBox11.Text = "搜尋目錄"
        Me.CheckBox11.UseVisualStyleBackColor = True
        '
        'Button12
        '
        Me.Button12.Location = New System.Drawing.Point(397, 476)
        Me.Button12.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button12.Name = "Button12"
        Me.Button12.Size = New System.Drawing.Size(51, 31)
        Me.Button12.TabIndex = 57
        Me.Button12.Text = "移出"
        Me.Button12.UseVisualStyleBackColor = True
        '
        'CheckBox9
        '
        Me.CheckBox9.AutoSize = True
        Me.CheckBox9.Location = New System.Drawing.Point(453, 482)
        Me.CheckBox9.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CheckBox9.Name = "CheckBox9"
        Me.CheckBox9.Size = New System.Drawing.Size(119, 19)
        Me.CheckBox9.TabIndex = 55
        Me.CheckBox9.Text = "執行所有文件"
        Me.CheckBox9.UseVisualStyleBackColor = True
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(255, 271)
        Me.TextBox5.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(136, 25)
        Me.TextBox5.TabIndex = 62
        '
        'TextBox6
        '
        Me.TextBox6.Location = New System.Drawing.Point(255, 479)
        Me.TextBox6.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(136, 25)
        Me.TextBox6.TabIndex = 63
        '
        'Button13
        '
        Me.Button13.Location = New System.Drawing.Point(507, 102)
        Me.Button13.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button13.Name = "Button13"
        Me.Button13.Size = New System.Drawing.Size(71, 24)
        Me.Button13.TabIndex = 64
        Me.Button13.Text = "匯入"
        Me.Button13.UseVisualStyleBackColor = True
        '
        'Button16
        '
        Me.Button16.Location = New System.Drawing.Point(565, 25)
        Me.Button16.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button16.Name = "Button16"
        Me.Button16.Size = New System.Drawing.Size(27, 22)
        Me.Button16.TabIndex = 69
        Me.Button16.Text = "↑"
        Me.Button16.UseVisualStyleBackColor = True
        '
        'Button17
        '
        Me.Button17.Location = New System.Drawing.Point(565, 70)
        Me.Button17.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button17.Name = "Button17"
        Me.Button17.Size = New System.Drawing.Size(27, 22)
        Me.Button17.TabIndex = 70
        Me.Button17.Text = "↑"
        Me.Button17.UseVisualStyleBackColor = True
        '
        'AxWindowsMediaPlayer1
        '
        Me.AxWindowsMediaPlayer1.Enabled = True
        Me.AxWindowsMediaPlayer1.Location = New System.Drawing.Point(436, 91)
        Me.AxWindowsMediaPlayer1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.AxWindowsMediaPlayer1.Name = "AxWindowsMediaPlayer1"
        Me.AxWindowsMediaPlayer1.OcxState = CType(resources.GetObject("AxWindowsMediaPlayer1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxWindowsMediaPlayer1.Size = New System.Drawing.Size(378, 266)
        Me.AxWindowsMediaPlayer1.TabIndex = 71
        Me.AxWindowsMediaPlayer1.Visible = False
        '
        'ComboBox3
        '
        Me.ComboBox3.FormattingEnabled = True
        Me.ComboBox3.Location = New System.Drawing.Point(84, 26)
        Me.ComboBox3.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ComboBox3.Name = "ComboBox3"
        Me.ComboBox3.Size = New System.Drawing.Size(475, 23)
        Me.ComboBox3.TabIndex = 72
        '
        'ComboBox4
        '
        Me.ComboBox4.FormattingEnabled = True
        Me.ComboBox4.Location = New System.Drawing.Point(85, 71)
        Me.ComboBox4.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ComboBox4.Name = "ComboBox4"
        Me.ComboBox4.Size = New System.Drawing.Size(473, 23)
        Me.ComboBox4.TabIndex = 73
        '
        'ComboBox5
        '
        Me.ComboBox5.FormattingEnabled = True
        Me.ComboBox5.Location = New System.Drawing.Point(1101, 296)
        Me.ComboBox5.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ComboBox5.Name = "ComboBox5"
        Me.ComboBox5.Size = New System.Drawing.Size(131, 23)
        Me.ComboBox5.TabIndex = 74
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(1097, 278)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(67, 15)
        Me.Label12.TabIndex = 75
        Me.Label12.Text = "進階操作"
        '
        'Button18
        '
        Me.Button18.Location = New System.Drawing.Point(431, 102)
        Me.Button18.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button18.Name = "Button18"
        Me.Button18.Size = New System.Drawing.Size(71, 24)
        Me.Button18.TabIndex = 76
        Me.Button18.Text = "下一集"
        Me.Button18.UseVisualStyleBackColor = True
        '
        'Button19
        '
        Me.Button19.Location = New System.Drawing.Point(355, 102)
        Me.Button19.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button19.Name = "Button19"
        Me.Button19.Size = New System.Drawing.Size(71, 24)
        Me.Button19.TabIndex = 77
        Me.Button19.Text = "上一集"
        Me.Button19.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(1100, 325)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(89, 19)
        Me.CheckBox1.TabIndex = 78
        Me.CheckBox1.Text = "關閉預覽"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(252, 106)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(29, 15)
        Me.Label13.TabIndex = 79
        Me.Label13.Text = "0項"
        '
        'Form1
        '
        Me.AllowDrop = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1247, 550)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Button19)
        Me.Controls.Add(Me.Button18)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.ComboBox5)
        Me.Controls.Add(Me.ComboBox4)
        Me.Controls.Add(Me.ComboBox3)
        Me.Controls.Add(Me.AxWindowsMediaPlayer1)
        Me.Controls.Add(Me.Button17)
        Me.Controls.Add(Me.Button16)
        Me.Controls.Add(Me.Button13)
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
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
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
    Friend WithEvents RadioButton7 As RadioButton
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
    Friend WithEvents RadioButton9 As RadioButton
    Friend WithEvents RadioButton8 As RadioButton
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
    Friend WithEvents RadioButton12 As RadioButton
    Friend WithEvents CheckBox6 As CheckBox

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            Button4.Text = "執行全部文件"
        Else
            Button4.Text = "執行選擇文件"
        End If
    End Sub



    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Form2.PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
        Form2.Show()
        Form2.PictureBox1.Image = PictureBox1.Image
    End Sub

    Private Sub Button11_Click_1(sender As Object, e As EventArgs) Handles Button11.Click
        If CheckBox5.Checked = False Then
            Dim now_select = 目錄檔案集合.SelectedIndices(0)
            Dim list_backup = 目錄檔案集合.SelectedItems.Cast(Of String)().ToArray()
            For Each i In list_backup
                目錄檔案集合.Items.Remove(i)
            Next
            If 目錄檔案集合.Items.Count <> 0 Then
                Try
                    目錄檔案集合.SelectedIndex = now_select
                Catch
                    目錄檔案集合.SelectedIndex = now_select - 1
                End Try
            End If
        Else
            目錄檔案集合.Items.Clear()
        End If
        backup = 目錄檔案集合.Items.Cast(Of String).ToArray
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        壓縮檔案集合.Items.RemoveAt(壓縮檔案集合.SelectedIndex)
    End Sub

    Public Sub Add_additional_operation()
        ComboBox5.Items.Add("不執行進階操作")
        ComboBox5.Items.Add("壓縮/解壓後刪除")
        ComboBox5.Items.Add("解壓分割")
        ComboBox5.Items.Add("更改附檔名")
        ComboBox5.Items.Add("clip啟動")
        ComboBox5.Items.Add("紀錄開啟檔案")
        ComboBox5.Items.Add("取樣")
        ComboBox5.Items.Add("對上層執行動作")
        ComboBox5.Items.Add("locale emulator")
        ComboBox5.Items.Add("自動重播(單首)")
        ComboBox5.Items.Add("自動重播(全部)")
    End Sub


    Private Sub Form1_Index_DragDrop(sender As System.Object, e As System.Windows.Forms.DragEventArgs) Handles 目錄檔案集合.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        Dim arguments As String = ""
        For Each path In files
            目錄檔案集合.Items.Add(path)
            arguments += $" ""{path}"""
        Next
        backup = 目錄檔案集合.Items.Cast(Of String).ToArray
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
    End Sub
    Private Sub Form1_text2_DragDrop(sender As System.Object, e As System.Windows.Forms.DragEventArgs) Handles ComboBox4.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        ComboBox4.Text = files(0)
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
        For Each p As Process In Process.GetProcesses()
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


End Class

