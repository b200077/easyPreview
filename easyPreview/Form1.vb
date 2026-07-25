

Imports System.ComponentModel
Imports System.IO
Imports System.Text.RegularExpressions
Imports LibVLCSharp.Shared


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
        Me.ExtList = New System.Windows.Forms.CheckedListBox()
        Me.FileCollection = New System.Windows.Forms.ListBox()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.SpecialButton = New System.Windows.Forms.RadioButton()
        Me.DeleteButton = New System.Windows.Forms.RadioButton()
        Me.CopyButton = New System.Windows.Forms.RadioButton()
        Me.MoveButton = New System.Windows.Forms.RadioButton()
        Me.OpenButton = New System.Windows.Forms.RadioButton()
        Me.SubFileCollection = New System.Windows.Forms.ListBox()
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
        Me.PasswordBox = New System.Windows.Forms.ListBox()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.Button9 = New System.Windows.Forms.Button()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.CheckBox5 = New System.Windows.Forms.CheckBox()
        Me.CheckBox6 = New System.Windows.Forms.CheckBox()
        Me.orderComboBox = New System.Windows.Forms.ComboBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.CheckBox10 = New System.Windows.Forms.CheckBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.CheckBox11 = New System.Windows.Forms.CheckBox()
        Me.Button12 = New System.Windows.Forms.Button()
        Me.CheckBox9 = New System.Windows.Forms.CheckBox()
        Me.TextBox6 = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Button13 = New System.Windows.Forms.Button()
        Me.Button18 = New System.Windows.Forms.Button()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.ComboBox7 = New System.Windows.Forms.ComboBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Button20 = New System.Windows.Forms.Button()
        Me.Button19 = New System.Windows.Forms.Button()
        Me.Button21 = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.MessageLabel = New System.Windows.Forms.Label()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.Panel23 = New System.Windows.Forms.Panel()
        Me.OptionCheckedListBox = New System.Windows.Forms.CheckedListBox()
        Me.Panel24 = New System.Windows.Forms.Panel()
        Me.Panel25 = New System.Windows.Forms.Panel()
        Me.Panel8 = New System.Windows.Forms.Panel()
        Me.Panel15 = New System.Windows.Forms.Panel()
        Me.Panel20 = New System.Windows.Forms.Panel()
        Me.Panel22 = New System.Windows.Forms.Panel()
        Me.Panel21 = New System.Windows.Forms.Panel()
        Me.Panel19 = New System.Windows.Forms.Panel()
        Me.Panel9 = New System.Windows.Forms.Panel()
        Me.Panel26 = New System.Windows.Forms.Panel()
        Me.Panel30 = New System.Windows.Forms.Panel()
        Me.Panel29 = New System.Windows.Forms.Panel()
        Me.Panel7 = New System.Windows.Forms.Panel()
        Me.Panel13 = New System.Windows.Forms.Panel()
        Me.Panel16 = New System.Windows.Forms.Panel()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Button23 = New System.Windows.Forms.Button()
        Me.RadioButton6 = New System.Windows.Forms.RadioButton()
        Me.RadioButton5 = New System.Windows.Forms.RadioButton()
        Me.RadioButton4 = New System.Windows.Forms.RadioButton()
        Me.Button11 = New System.Windows.Forms.Button()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.videoPanel = New System.Windows.Forms.Panel()
        Me.VideoView1 = New LibVLCSharp.WinForms.VideoView()
        Me.Panel36 = New System.Windows.Forms.Panel()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.ComboBox5 = New System.Windows.Forms.ComboBox()
        Me.ProgressBar2 = New System.Windows.Forms.ProgressBar()
        Me.DeviceComboBox = New System.Windows.Forms.ComboBox()
        Me.Button24 = New System.Windows.Forms.Button()
        Me.TrackBar1 = New System.Windows.Forms.TrackBar()
        Me.Button25 = New System.Windows.Forms.Button()
        Me.WebView21 = New Microsoft.Web.WebView2.WinForms.WebView2()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.Panel12 = New System.Windows.Forms.Panel()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Panel14 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.SearchTextBox = New System.Windows.Forms.TextBox()
        Me.Panel28 = New System.Windows.Forms.Panel()
        Me.Panel27 = New System.Windows.Forms.Panel()
        Me.Panel17 = New System.Windows.Forms.Panel()
        Me.Panel31 = New System.Windows.Forms.Panel()
        Me.Panel34 = New System.Windows.Forms.Panel()
        Me.Button16 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button17 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Panel35 = New System.Windows.Forms.Panel()
        Me.PathComboBox = New System.Windows.Forms.ComboBox()
        Me.TargetComboBox = New System.Windows.Forms.ComboBox()
        Me.Panel33 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel11 = New System.Windows.Forms.Panel()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.nextButton = New System.Windows.Forms.Button()
        Me.previousButton = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.RadioButton11 = New System.Windows.Forms.RadioButton()
        Me.RadioButton13 = New System.Windows.Forms.RadioButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RadioButton15 = New System.Windows.Forms.RadioButton()
        Me.RadioButton14 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton3 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.Panel18 = New System.Windows.Forms.Panel()
        Me.ContextMenuStrip4 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.GroupBox2.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.Panel6.SuspendLayout()
        Me.Panel23.SuspendLayout()
        Me.Panel24.SuspendLayout()
        Me.Panel25.SuspendLayout()
        Me.Panel8.SuspendLayout()
        Me.Panel15.SuspendLayout()
        Me.Panel20.SuspendLayout()
        Me.Panel22.SuspendLayout()
        Me.Panel21.SuspendLayout()
        Me.Panel19.SuspendLayout()
        Me.Panel9.SuspendLayout()
        Me.Panel26.SuspendLayout()
        Me.Panel30.SuspendLayout()
        Me.Panel29.SuspendLayout()
        Me.Panel7.SuspendLayout()
        Me.Panel13.SuspendLayout()
        Me.Panel16.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.videoPanel.SuspendLayout()
        CType(Me.VideoView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel36.SuspendLayout()
        CType(Me.TrackBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.WebView21, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel12.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.Panel14.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel28.SuspendLayout()
        Me.Panel27.SuspendLayout()
        Me.Panel17.SuspendLayout()
        Me.Panel31.SuspendLayout()
        Me.Panel34.SuspendLayout()
        Me.Panel35.SuspendLayout()
        Me.Panel33.SuspendLayout()
        Me.Panel11.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Panel18.SuspendLayout()
        Me.SuspendLayout()
        '
        'ExtList
        '
        Me.ExtList.BackColor = System.Drawing.SystemColors.WindowFrame
        resources.ApplyResources(Me.ExtList, "ExtList")
        Me.ExtList.ForeColor = System.Drawing.SystemColors.Window
        Me.ExtList.FormattingEnabled = True
        Me.ExtList.Name = "ExtList"
        '
        'FileCollection
        '
        Me.FileCollection.AllowDrop = True
        Me.FileCollection.BackColor = System.Drawing.SystemColors.GrayText
        Me.FileCollection.ContextMenuStrip = Me.ContextMenuStrip1
        resources.ApplyResources(Me.FileCollection, "FileCollection")
        Me.FileCollection.ForeColor = System.Drawing.SystemColors.Window
        Me.FileCollection.FormattingEnabled = True
        Me.FileCollection.Name = "FileCollection"
        Me.FileCollection.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        resources.ApplyResources(Me.ContextMenuStrip1, "ContextMenuStrip1")
        '
        'ContextMenuStrip2
        '
        Me.ContextMenuStrip2.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.ContextMenuStrip2.Name = "ContextMenuStrip2"
        resources.ApplyResources(Me.ContextMenuStrip2, "ContextMenuStrip2")
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.SpecialButton)
        Me.GroupBox2.Controls.Add(Me.DeleteButton)
        Me.GroupBox2.Controls.Add(Me.CopyButton)
        Me.GroupBox2.Controls.Add(Me.MoveButton)
        Me.GroupBox2.Controls.Add(Me.OpenButton)
        resources.ApplyResources(Me.GroupBox2, "GroupBox2")
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.TabStop = False
        '
        'SpecialButton
        '
        resources.ApplyResources(Me.SpecialButton, "SpecialButton")
        Me.SpecialButton.Name = "SpecialButton"
        Me.SpecialButton.TabStop = True
        Me.SpecialButton.UseVisualStyleBackColor = True
        '
        'DeleteButton
        '
        resources.ApplyResources(Me.DeleteButton, "DeleteButton")
        Me.DeleteButton.Name = "DeleteButton"
        Me.DeleteButton.TabStop = True
        Me.DeleteButton.UseVisualStyleBackColor = True
        '
        'CopyButton
        '
        resources.ApplyResources(Me.CopyButton, "CopyButton")
        Me.CopyButton.Name = "CopyButton"
        Me.CopyButton.TabStop = True
        Me.CopyButton.UseVisualStyleBackColor = True
        '
        'MoveButton
        '
        resources.ApplyResources(Me.MoveButton, "MoveButton")
        Me.MoveButton.Name = "MoveButton"
        Me.MoveButton.TabStop = True
        Me.MoveButton.UseVisualStyleBackColor = True
        '
        'OpenButton
        '
        Me.OpenButton.Checked = True
        resources.ApplyResources(Me.OpenButton, "OpenButton")
        Me.OpenButton.Name = "OpenButton"
        Me.OpenButton.TabStop = True
        Me.OpenButton.UseVisualStyleBackColor = True
        '
        'SubFileCollection
        '
        Me.SubFileCollection.AllowDrop = True
        Me.SubFileCollection.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.SubFileCollection.ContextMenuStrip = Me.ContextMenuStrip3
        resources.ApplyResources(Me.SubFileCollection, "SubFileCollection")
        Me.SubFileCollection.ForeColor = System.Drawing.SystemColors.Window
        Me.SubFileCollection.FormattingEnabled = True
        Me.SubFileCollection.Name = "SubFileCollection"
        Me.SubFileCollection.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        '
        'ContextMenuStrip3
        '
        Me.ContextMenuStrip3.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.ContextMenuStrip3.Name = "ContextMenuStrip3"
        resources.ApplyResources(Me.ContextMenuStrip3, "ContextMenuStrip3")
        '
        'Button4
        '
        resources.ApplyResources(Me.Button4, "Button4")
        Me.Button4.Name = "Button4"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Label3
        '
        resources.ApplyResources(Me.Label3, "Label3")
        Me.Label3.Name = "Label3"
        '
        'Button5
        '
        resources.ApplyResources(Me.Button5, "Button5")
        Me.Button5.Name = "Button5"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button7
        '
        resources.ApplyResources(Me.Button7, "Button7")
        Me.Button7.Name = "Button7"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'ComboBox2
        '
        resources.ApplyResources(Me.ComboBox2, "ComboBox2")
        Me.ComboBox2.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Name = "ComboBox2"
        '
        'Label4
        '
        resources.ApplyResources(Me.Label4, "Label4")
        Me.Label4.Name = "Label4"
        '
        'Label5
        '
        resources.ApplyResources(Me.Label5, "Label5")
        Me.Label5.Name = "Label5"
        '
        'Label6
        '
        resources.ApplyResources(Me.Label6, "Label6")
        Me.Label6.Name = "Label6"
        '
        'Label7
        '
        resources.ApplyResources(Me.Label7, "Label7")
        Me.Label7.Name = "Label7"
        '
        'PasswordBox
        '
        Me.PasswordBox.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.PasswordBox.ForeColor = System.Drawing.SystemColors.Window
        Me.PasswordBox.FormattingEnabled = True
        resources.ApplyResources(Me.PasswordBox, "PasswordBox")
        Me.PasswordBox.Name = "PasswordBox"
        '
        'Button6
        '
        resources.ApplyResources(Me.Button6, "Button6")
        Me.Button6.Name = "Button6"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Button8
        '
        resources.ApplyResources(Me.Button8, "Button8")
        Me.Button8.Name = "Button8"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'Button9
        '
        resources.ApplyResources(Me.Button9, "Button9")
        Me.Button9.Name = "Button9"
        Me.Button9.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        resources.ApplyResources(Me.ProgressBar1, "ProgressBar1")
        Me.ProgressBar1.Name = "ProgressBar1"
        '
        'Label8
        '
        resources.ApplyResources(Me.Label8, "Label8")
        Me.Label8.Name = "Label8"
        '
        'TextBox3
        '
        resources.ApplyResources(Me.TextBox3, "TextBox3")
        Me.TextBox3.Name = "TextBox3"
        '
        'TextBox4
        '
        resources.ApplyResources(Me.TextBox4, "TextBox4")
        Me.TextBox4.Name = "TextBox4"
        '
        'CheckBox5
        '
        resources.ApplyResources(Me.CheckBox5, "CheckBox5")
        Me.CheckBox5.Name = "CheckBox5"
        Me.CheckBox5.UseVisualStyleBackColor = True
        '
        'CheckBox6
        '
        resources.ApplyResources(Me.CheckBox6, "CheckBox6")
        Me.CheckBox6.Name = "CheckBox6"
        Me.CheckBox6.UseVisualStyleBackColor = True
        '
        'orderComboBox
        '
        resources.ApplyResources(Me.orderComboBox, "orderComboBox")
        Me.orderComboBox.FormattingEnabled = True
        Me.orderComboBox.Name = "orderComboBox"
        '
        'Label9
        '
        resources.ApplyResources(Me.Label9, "Label9")
        Me.Label9.Name = "Label9"
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        resources.ApplyResources(Me.DataGridView1, "DataGridView1")
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Height = 27
        '
        'CheckBox10
        '
        resources.ApplyResources(Me.CheckBox10, "CheckBox10")
        Me.CheckBox10.Name = "CheckBox10"
        Me.CheckBox10.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        Me.OpenFileDialog1.Multiselect = True
        '
        'CheckBox11
        '
        resources.ApplyResources(Me.CheckBox11, "CheckBox11")
        Me.CheckBox11.Name = "CheckBox11"
        Me.CheckBox11.UseVisualStyleBackColor = True
        '
        'Button12
        '
        resources.ApplyResources(Me.Button12, "Button12")
        Me.Button12.Name = "Button12"
        Me.Button12.UseVisualStyleBackColor = True
        '
        'CheckBox9
        '
        resources.ApplyResources(Me.CheckBox9, "CheckBox9")
        Me.CheckBox9.Name = "CheckBox9"
        Me.CheckBox9.UseVisualStyleBackColor = True
        '
        'TextBox6
        '
        resources.ApplyResources(Me.TextBox6, "TextBox6")
        Me.TextBox6.Name = "TextBox6"
        '
        'Label12
        '
        resources.ApplyResources(Me.Label12, "Label12")
        Me.Label12.Name = "Label12"
        '
        'Label13
        '
        resources.ApplyResources(Me.Label13, "Label13")
        Me.Label13.ForeColor = System.Drawing.SystemColors.Window
        Me.Label13.Name = "Label13"
        '
        'Button13
        '
        resources.ApplyResources(Me.Button13, "Button13")
        Me.Button13.Name = "Button13"
        Me.Button13.UseVisualStyleBackColor = True
        '
        'Button18
        '
        resources.ApplyResources(Me.Button18, "Button18")
        Me.Button18.Name = "Button18"
        Me.Button18.UseVisualStyleBackColor = True
        '
        'Label16
        '
        resources.ApplyResources(Me.Label16, "Label16")
        Me.Label16.Name = "Label16"
        '
        'ComboBox7
        '
        Me.ComboBox7.FormattingEnabled = True
        resources.ApplyResources(Me.ComboBox7, "ComboBox7")
        Me.ComboBox7.Name = "ComboBox7"
        '
        'Label17
        '
        resources.ApplyResources(Me.Label17, "Label17")
        Me.Label17.ForeColor = System.Drawing.SystemColors.Window
        Me.Label17.Name = "Label17"
        '
        'Button20
        '
        resources.ApplyResources(Me.Button20, "Button20")
        Me.Button20.Name = "Button20"
        Me.Button20.UseVisualStyleBackColor = True
        '
        'Button19
        '
        resources.ApplyResources(Me.Button19, "Button19")
        Me.Button19.Name = "Button19"
        Me.Button19.UseVisualStyleBackColor = True
        '
        'Button21
        '
        resources.ApplyResources(Me.Button21, "Button21")
        Me.Button21.Name = "Button21"
        Me.Button21.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Button13)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.Label13)
        resources.ApplyResources(Me.Panel2, "Panel2")
        Me.Panel2.Name = "Panel2"
        '
        'Panel5
        '
        Me.Panel5.Controls.Add(Me.MessageLabel)
        Me.Panel5.Controls.Add(Me.orderComboBox)
        Me.Panel5.Controls.Add(Me.Button18)
        Me.Panel5.Controls.Add(Me.TextBox3)
        Me.Panel5.Controls.Add(Me.ComboBox2)
        Me.Panel5.Controls.Add(Me.Label3)
        Me.Panel5.Controls.Add(Me.Label9)
        resources.ApplyResources(Me.Panel5, "Panel5")
        Me.Panel5.Name = "Panel5"
        '
        'MessageLabel
        '
        resources.ApplyResources(Me.MessageLabel, "MessageLabel")
        Me.MessageLabel.Name = "MessageLabel"
        '
        'Panel6
        '
        Me.Panel6.Controls.Add(Me.Panel23)
        Me.Panel6.Controls.Add(Me.Panel24)
        Me.Panel6.Controls.Add(Me.Panel25)
        resources.ApplyResources(Me.Panel6, "Panel6")
        Me.Panel6.Name = "Panel6"
        '
        'Panel23
        '
        Me.Panel23.Controls.Add(Me.OptionCheckedListBox)
        Me.Panel23.Controls.Add(Me.Label12)
        resources.ApplyResources(Me.Panel23, "Panel23")
        Me.Panel23.Name = "Panel23"
        '
        'OptionCheckedListBox
        '
        Me.OptionCheckedListBox.BackColor = System.Drawing.SystemColors.WindowFrame
        resources.ApplyResources(Me.OptionCheckedListBox, "OptionCheckedListBox")
        Me.OptionCheckedListBox.ForeColor = System.Drawing.SystemColors.Window
        Me.OptionCheckedListBox.FormattingEnabled = True
        Me.OptionCheckedListBox.Name = "OptionCheckedListBox"
        '
        'Panel24
        '
        Me.Panel24.Controls.Add(Me.Button4)
        Me.Panel24.Controls.Add(Me.Button9)
        resources.ApplyResources(Me.Panel24, "Panel24")
        Me.Panel24.Name = "Panel24"
        '
        'Panel25
        '
        Me.Panel25.Controls.Add(Me.GroupBox2)
        Me.Panel25.Controls.Add(Me.ComboBox7)
        Me.Panel25.Controls.Add(Me.Label16)
        resources.ApplyResources(Me.Panel25, "Panel25")
        Me.Panel25.Name = "Panel25"
        '
        'Panel8
        '
        Me.Panel8.Controls.Add(Me.Panel15)
        Me.Panel8.Controls.Add(Me.Panel19)
        resources.ApplyResources(Me.Panel8, "Panel8")
        Me.Panel8.Name = "Panel8"
        '
        'Panel15
        '
        Me.Panel15.Controls.Add(Me.Panel20)
        Me.Panel15.Controls.Add(Me.Panel22)
        Me.Panel15.Controls.Add(Me.Panel21)
        resources.ApplyResources(Me.Panel15, "Panel15")
        Me.Panel15.Name = "Panel15"
        '
        'Panel20
        '
        Me.Panel20.Controls.Add(Me.ExtList)
        resources.ApplyResources(Me.Panel20, "Panel20")
        Me.Panel20.Name = "Panel20"
        '
        'Panel22
        '
        Me.Panel22.Controls.Add(Me.CheckBox11)
        Me.Panel22.Controls.Add(Me.CheckBox10)
        resources.ApplyResources(Me.Panel22, "Panel22")
        Me.Panel22.Name = "Panel22"
        '
        'Panel21
        '
        Me.Panel21.Controls.Add(Me.Label6)
        Me.Panel21.Controls.Add(Me.CheckBox6)
        resources.ApplyResources(Me.Panel21, "Panel21")
        Me.Panel21.Name = "Panel21"
        '
        'Panel19
        '
        Me.Panel19.Controls.Add(Me.Button21)
        Me.Panel19.Controls.Add(Me.Button5)
        Me.Panel19.Controls.Add(Me.Button7)
        Me.Panel19.Controls.Add(Me.TextBox4)
        Me.Panel19.Controls.Add(Me.PasswordBox)
        Me.Panel19.Controls.Add(Me.Label4)
        resources.ApplyResources(Me.Panel19, "Panel19")
        Me.Panel19.Name = "Panel19"
        '
        'Panel9
        '
        Me.Panel9.Controls.Add(Me.Panel26)
        Me.Panel9.Controls.Add(Me.Panel30)
        Me.Panel9.Controls.Add(Me.Panel29)
        resources.ApplyResources(Me.Panel9, "Panel9")
        Me.Panel9.Name = "Panel9"
        '
        'Panel26
        '
        Me.Panel26.Controls.Add(Me.SubFileCollection)
        resources.ApplyResources(Me.Panel26, "Panel26")
        Me.Panel26.Name = "Panel26"
        '
        'Panel30
        '
        Me.Panel30.Controls.Add(Me.Button19)
        Me.Panel30.Controls.Add(Me.Label5)
        Me.Panel30.Controls.Add(Me.Label17)
        resources.ApplyResources(Me.Panel30, "Panel30")
        Me.Panel30.Name = "Panel30"
        '
        'Panel29
        '
        Me.Panel29.Controls.Add(Me.CheckBox9)
        Me.Panel29.Controls.Add(Me.TextBox6)
        Me.Panel29.Controls.Add(Me.Label8)
        Me.Panel29.Controls.Add(Me.Button20)
        Me.Panel29.Controls.Add(Me.Button12)
        Me.Panel29.Controls.Add(Me.ProgressBar1)
        resources.ApplyResources(Me.Panel29, "Panel29")
        Me.Panel29.Name = "Panel29"
        '
        'Panel7
        '
        Me.Panel7.Controls.Add(Me.Panel13)
        Me.Panel7.Controls.Add(Me.Panel6)
        resources.ApplyResources(Me.Panel7, "Panel7")
        Me.Panel7.Name = "Panel7"
        '
        'Panel13
        '
        Me.Panel13.Controls.Add(Me.Panel16)
        Me.Panel13.Controls.Add(Me.Panel5)
        resources.ApplyResources(Me.Panel13, "Panel13")
        Me.Panel13.Name = "Panel13"
        '
        'Panel16
        '
        Me.Panel16.Controls.Add(Me.GroupBox3)
        Me.Panel16.Controls.Add(Me.Button11)
        Me.Panel16.Controls.Add(Me.Label15)
        Me.Panel16.Controls.Add(Me.videoPanel)
        Me.Panel16.Controls.Add(Me.DataGridView1)
        Me.Panel16.Controls.Add(Me.WebView21)
        Me.Panel16.Controls.Add(Me.PictureBox1)
        Me.Panel16.Controls.Add(Me.RichTextBox1)
        resources.ApplyResources(Me.Panel16, "Panel16")
        Me.Panel16.Name = "Panel16"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Button23)
        Me.GroupBox3.Controls.Add(Me.RadioButton6)
        Me.GroupBox3.Controls.Add(Me.RadioButton5)
        Me.GroupBox3.Controls.Add(Me.RadioButton4)
        resources.ApplyResources(Me.GroupBox3, "GroupBox3")
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.TabStop = False
        '
        'Button23
        '
        resources.ApplyResources(Me.Button23, "Button23")
        Me.Button23.Name = "Button23"
        Me.Button23.UseVisualStyleBackColor = True
        '
        'RadioButton6
        '
        resources.ApplyResources(Me.RadioButton6, "RadioButton6")
        Me.RadioButton6.Name = "RadioButton6"
        Me.RadioButton6.TabStop = True
        Me.RadioButton6.UseVisualStyleBackColor = True
        '
        'RadioButton5
        '
        resources.ApplyResources(Me.RadioButton5, "RadioButton5")
        Me.RadioButton5.Name = "RadioButton5"
        Me.RadioButton5.TabStop = True
        Me.RadioButton5.UseVisualStyleBackColor = True
        '
        'RadioButton4
        '
        resources.ApplyResources(Me.RadioButton4, "RadioButton4")
        Me.RadioButton4.Name = "RadioButton4"
        Me.RadioButton4.TabStop = True
        Me.RadioButton4.UseVisualStyleBackColor = True
        '
        'Button11
        '
        resources.ApplyResources(Me.Button11, "Button11")
        Me.Button11.Name = "Button11"
        Me.Button11.UseVisualStyleBackColor = True
        '
        'Label15
        '
        resources.ApplyResources(Me.Label15, "Label15")
        Me.Label15.Name = "Label15"
        '
        'videoPanel
        '
        Me.videoPanel.Controls.Add(Me.VideoView1)
        Me.videoPanel.Controls.Add(Me.Panel36)
        resources.ApplyResources(Me.videoPanel, "videoPanel")
        Me.videoPanel.Name = "videoPanel"
        '
        'VideoView1
        '
        Me.VideoView1.BackColor = System.Drawing.Color.Black
        resources.ApplyResources(Me.VideoView1, "VideoView1")
        Me.VideoView1.MediaPlayer = Nothing
        Me.VideoView1.Name = "VideoView1"
        '
        'Panel36
        '
        Me.Panel36.Controls.Add(Me.Label11)
        Me.Panel36.Controls.Add(Me.ComboBox5)
        Me.Panel36.Controls.Add(Me.ProgressBar2)
        Me.Panel36.Controls.Add(Me.DeviceComboBox)
        Me.Panel36.Controls.Add(Me.Button24)
        Me.Panel36.Controls.Add(Me.TrackBar1)
        Me.Panel36.Controls.Add(Me.Button25)
        resources.ApplyResources(Me.Panel36, "Panel36")
        Me.Panel36.Name = "Panel36"
        '
        'Label11
        '
        resources.ApplyResources(Me.Label11, "Label11")
        Me.Label11.Name = "Label11"
        '
        'ComboBox5
        '
        resources.ApplyResources(Me.ComboBox5, "ComboBox5")
        Me.ComboBox5.FormattingEnabled = True
        Me.ComboBox5.Name = "ComboBox5"
        '
        'ProgressBar2
        '
        resources.ApplyResources(Me.ProgressBar2, "ProgressBar2")
        Me.ProgressBar2.Maximum = 1000
        Me.ProgressBar2.Name = "ProgressBar2"
        '
        'DeviceComboBox
        '
        resources.ApplyResources(Me.DeviceComboBox, "DeviceComboBox")
        Me.DeviceComboBox.FormattingEnabled = True
        Me.DeviceComboBox.Name = "DeviceComboBox"
        '
        'Button24
        '
        resources.ApplyResources(Me.Button24, "Button24")
        Me.Button24.Name = "Button24"
        Me.Button24.UseVisualStyleBackColor = True
        '
        'TrackBar1
        '
        resources.ApplyResources(Me.TrackBar1, "TrackBar1")
        Me.TrackBar1.Maximum = 100
        Me.TrackBar1.Name = "TrackBar1"
        Me.TrackBar1.Value = 80
        '
        'Button25
        '
        resources.ApplyResources(Me.Button25, "Button25")
        Me.Button25.Name = "Button25"
        Me.Button25.UseVisualStyleBackColor = True
        '
        'WebView21
        '
        Me.WebView21.AllowExternalDrop = True
        Me.WebView21.CreationProperties = Nothing
        Me.WebView21.DefaultBackgroundColor = System.Drawing.Color.White
        resources.ApplyResources(Me.WebView21, "WebView21")
        Me.WebView21.Name = "WebView21"
        Me.WebView21.ZoomFactor = 1.0R
        '
        'PictureBox1
        '
        Me.PictureBox1.ContextMenuStrip = Me.ContextMenuStrip2
        resources.ApplyResources(Me.PictureBox1, "PictureBox1")
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.TabStop = False
        '
        'RichTextBox1
        '
        Me.RichTextBox1.BackColor = System.Drawing.SystemColors.WindowFrame
        resources.ApplyResources(Me.RichTextBox1, "RichTextBox1")
        Me.RichTextBox1.ForeColor = System.Drawing.SystemColors.Info
        Me.RichTextBox1.Name = "RichTextBox1"
        '
        'Panel12
        '
        Me.Panel12.Controls.Add(Me.Panel4)
        Me.Panel12.Controls.Add(Me.Panel8)
        resources.ApplyResources(Me.Panel12, "Panel12")
        Me.Panel12.Name = "Panel12"
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.Panel14)
        Me.Panel4.Controls.Add(Me.Panel9)
        resources.ApplyResources(Me.Panel4, "Panel4")
        Me.Panel4.Name = "Panel4"
        '
        'Panel14
        '
        Me.Panel14.Controls.Add(Me.Panel3)
        Me.Panel14.Controls.Add(Me.Panel1)
        Me.Panel14.Controls.Add(Me.Panel2)
        resources.ApplyResources(Me.Panel14, "Panel14")
        Me.Panel14.Name = "Panel14"
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.Button3)
        Me.Panel3.Controls.Add(Me.FileCollection)
        resources.ApplyResources(Me.Panel3, "Panel3")
        Me.Panel3.Name = "Panel3"
        '
        'Button3
        '
        resources.ApplyResources(Me.Button3, "Button3")
        Me.Button3.Name = "Button3"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.SearchTextBox)
        Me.Panel1.Controls.Add(Me.Panel28)
        Me.Panel1.Controls.Add(Me.Panel27)
        resources.ApplyResources(Me.Panel1, "Panel1")
        Me.Panel1.Name = "Panel1"
        '
        'SearchTextBox
        '
        Me.SearchTextBox.BackColor = System.Drawing.SystemColors.WindowFrame
        resources.ApplyResources(Me.SearchTextBox, "SearchTextBox")
        Me.SearchTextBox.Name = "SearchTextBox"
        '
        'Panel28
        '
        Me.Panel28.Controls.Add(Me.Button8)
        Me.Panel28.Controls.Add(Me.Button6)
        resources.ApplyResources(Me.Panel28, "Panel28")
        Me.Panel28.Name = "Panel28"
        '
        'Panel27
        '
        Me.Panel27.Controls.Add(Me.CheckBox5)
        resources.ApplyResources(Me.Panel27, "Panel27")
        Me.Panel27.Name = "Panel27"
        '
        'Panel17
        '
        Me.Panel17.Controls.Add(Me.Panel31)
        Me.Panel17.Controls.Add(Me.Panel11)
        resources.ApplyResources(Me.Panel17, "Panel17")
        Me.Panel17.Name = "Panel17"
        '
        'Panel31
        '
        Me.Panel31.Controls.Add(Me.Panel34)
        Me.Panel31.Controls.Add(Me.Panel35)
        Me.Panel31.Controls.Add(Me.Panel33)
        resources.ApplyResources(Me.Panel31, "Panel31")
        Me.Panel31.Name = "Panel31"
        '
        'Panel34
        '
        Me.Panel34.Controls.Add(Me.Button16)
        Me.Panel34.Controls.Add(Me.Button2)
        Me.Panel34.Controls.Add(Me.Button17)
        Me.Panel34.Controls.Add(Me.Button1)
        resources.ApplyResources(Me.Panel34, "Panel34")
        Me.Panel34.Name = "Panel34"
        '
        'Button16
        '
        resources.ApplyResources(Me.Button16, "Button16")
        Me.Button16.Name = "Button16"
        Me.Button16.UseVisualStyleBackColor = True
        '
        'Button2
        '
        resources.ApplyResources(Me.Button2, "Button2")
        Me.Button2.Name = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button17
        '
        resources.ApplyResources(Me.Button17, "Button17")
        Me.Button17.Name = "Button17"
        Me.Button17.UseVisualStyleBackColor = True
        '
        'Button1
        '
        resources.ApplyResources(Me.Button1, "Button1")
        Me.Button1.Name = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Panel35
        '
        Me.Panel35.Controls.Add(Me.PathComboBox)
        Me.Panel35.Controls.Add(Me.TargetComboBox)
        resources.ApplyResources(Me.Panel35, "Panel35")
        Me.Panel35.Name = "Panel35"
        '
        'PathComboBox
        '
        Me.PathComboBox.AllowDrop = True
        resources.ApplyResources(Me.PathComboBox, "PathComboBox")
        Me.PathComboBox.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.PathComboBox.ForeColor = System.Drawing.SystemColors.Window
        Me.PathComboBox.FormattingEnabled = True
        Me.PathComboBox.Name = "PathComboBox"
        '
        'TargetComboBox
        '
        Me.TargetComboBox.AllowDrop = True
        resources.ApplyResources(Me.TargetComboBox, "TargetComboBox")
        Me.TargetComboBox.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.TargetComboBox.ForeColor = System.Drawing.SystemColors.Window
        Me.TargetComboBox.FormattingEnabled = True
        Me.TargetComboBox.Name = "TargetComboBox"
        '
        'Panel33
        '
        Me.Panel33.Controls.Add(Me.Label2)
        Me.Panel33.Controls.Add(Me.Label1)
        resources.ApplyResources(Me.Panel33, "Panel33")
        Me.Panel33.Name = "Panel33"
        '
        'Label2
        '
        resources.ApplyResources(Me.Label2, "Label2")
        Me.Label2.Name = "Label2"
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'Panel11
        '
        Me.Panel11.Controls.Add(Me.Label10)
        Me.Panel11.Controls.Add(Me.CheckBox3)
        Me.Panel11.Controls.Add(Me.nextButton)
        Me.Panel11.Controls.Add(Me.previousButton)
        Me.Panel11.Controls.Add(Me.GroupBox4)
        Me.Panel11.Controls.Add(Me.GroupBox1)
        resources.ApplyResources(Me.Panel11, "Panel11")
        Me.Panel11.Name = "Panel11"
        '
        'Label10
        '
        resources.ApplyResources(Me.Label10, "Label10")
        Me.Label10.Name = "Label10"
        '
        'CheckBox3
        '
        resources.ApplyResources(Me.CheckBox3, "CheckBox3")
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'nextButton
        '
        resources.ApplyResources(Me.nextButton, "nextButton")
        Me.nextButton.Name = "nextButton"
        Me.nextButton.UseVisualStyleBackColor = True
        '
        'previousButton
        '
        resources.ApplyResources(Me.previousButton, "previousButton")
        Me.previousButton.Name = "previousButton"
        Me.previousButton.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.CheckBox2)
        Me.GroupBox4.Controls.Add(Me.RadioButton11)
        Me.GroupBox4.Controls.Add(Me.RadioButton13)
        resources.ApplyResources(Me.GroupBox4, "GroupBox4")
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.TabStop = False
        '
        'CheckBox2
        '
        resources.ApplyResources(Me.CheckBox2, "CheckBox2")
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'RadioButton11
        '
        resources.ApplyResources(Me.RadioButton11, "RadioButton11")
        Me.RadioButton11.Name = "RadioButton11"
        Me.RadioButton11.UseVisualStyleBackColor = True
        '
        'RadioButton13
        '
        resources.ApplyResources(Me.RadioButton13, "RadioButton13")
        Me.RadioButton13.Name = "RadioButton13"
        Me.RadioButton13.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RadioButton15)
        Me.GroupBox1.Controls.Add(Me.RadioButton14)
        Me.GroupBox1.Controls.Add(Me.RadioButton2)
        Me.GroupBox1.Controls.Add(Me.RadioButton3)
        Me.GroupBox1.Controls.Add(Me.RadioButton1)
        resources.ApplyResources(Me.GroupBox1, "GroupBox1")
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.TabStop = False
        '
        'RadioButton15
        '
        resources.ApplyResources(Me.RadioButton15, "RadioButton15")
        Me.RadioButton15.Name = "RadioButton15"
        Me.RadioButton15.UseVisualStyleBackColor = True
        '
        'RadioButton14
        '
        resources.ApplyResources(Me.RadioButton14, "RadioButton14")
        Me.RadioButton14.Name = "RadioButton14"
        Me.RadioButton14.TabStop = True
        Me.RadioButton14.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        resources.ApplyResources(Me.RadioButton2, "RadioButton2")
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton3
        '
        resources.ApplyResources(Me.RadioButton3, "RadioButton3")
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.TabStop = True
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        resources.ApplyResources(Me.RadioButton1, "RadioButton1")
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'Panel18
        '
        Me.Panel18.Controls.Add(Me.Panel7)
        Me.Panel18.Controls.Add(Me.Panel12)
        resources.ApplyResources(Me.Panel18, "Panel18")
        Me.Panel18.Name = "Panel18"
        '
        'ContextMenuStrip4
        '
        Me.ContextMenuStrip4.Name = "ContextMenuStrip4"
        resources.ApplyResources(Me.ContextMenuStrip4, "ContextMenuStrip4")
        '
        'Form1
        '
        Me.AllowDrop = True
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.Controls.Add(Me.Panel18)
        Me.Controls.Add(Me.Panel17)
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.KeyPreview = True
        Me.Name = "Form1"
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel5.ResumeLayout(False)
        Me.Panel5.PerformLayout()
        Me.Panel6.ResumeLayout(False)
        Me.Panel23.ResumeLayout(False)
        Me.Panel23.PerformLayout()
        Me.Panel24.ResumeLayout(False)
        Me.Panel25.ResumeLayout(False)
        Me.Panel25.PerformLayout()
        Me.Panel8.ResumeLayout(False)
        Me.Panel15.ResumeLayout(False)
        Me.Panel20.ResumeLayout(False)
        Me.Panel22.ResumeLayout(False)
        Me.Panel22.PerformLayout()
        Me.Panel21.ResumeLayout(False)
        Me.Panel21.PerformLayout()
        Me.Panel19.ResumeLayout(False)
        Me.Panel19.PerformLayout()
        Me.Panel9.ResumeLayout(False)
        Me.Panel26.ResumeLayout(False)
        Me.Panel30.ResumeLayout(False)
        Me.Panel30.PerformLayout()
        Me.Panel29.ResumeLayout(False)
        Me.Panel29.PerformLayout()
        Me.Panel7.ResumeLayout(False)
        Me.Panel13.ResumeLayout(False)
        Me.Panel16.ResumeLayout(False)
        Me.Panel16.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.videoPanel.ResumeLayout(False)
        CType(Me.VideoView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel36.ResumeLayout(False)
        Me.Panel36.PerformLayout()
        CType(Me.TrackBar1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.WebView21, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel12.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.Panel14.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel28.ResumeLayout(False)
        Me.Panel27.ResumeLayout(False)
        Me.Panel27.PerformLayout()
        Me.Panel17.ResumeLayout(False)
        Me.Panel31.ResumeLayout(False)
        Me.Panel34.ResumeLayout(False)
        Me.Panel35.ResumeLayout(False)
        Me.Panel33.ResumeLayout(False)
        Me.Panel33.PerformLayout()
        Me.Panel11.ResumeLayout(False)
        Me.Panel11.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.Panel18.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ExtList As CheckedListBox
    Friend WithEvents FileCollection As ListBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents DeleteButton As RadioButton
    Friend WithEvents CopyButton As RadioButton
    Friend WithEvents MoveButton As RadioButton
    Friend WithEvents OpenButton As RadioButton
    Friend WithEvents SubFileCollection As ListBox
    Friend WithEvents Button4 As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents Button5 As Button
    Friend WithEvents Button7 As Button
    Friend WithEvents ComboBox2 As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents PasswordBox As ListBox
    Friend WithEvents Button6 As Button
    Friend WithEvents Button8 As Button
    Friend WithEvents Button9 As Button
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents Label8 As Label
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents TextBox4 As TextBox
    Friend WithEvents CheckBox5 As CheckBox
    Friend WithEvents CheckBox6 As CheckBox

    Private Sub FileCollection_MouseMove(sender As Object, e As MouseEventArgs) Handles fileCollection.MouseMove, subFileCollection.MouseMove, TargetComboBox.MouseMove, PathComboBox.MouseMove
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
    Select(Function(x) remove_label(x)).ToArray()
            '' 判斷滑鼠是否在控制項範圍內
            If sender Is FileCollection Then file_Preview(True)
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
    Private Sub Form1_Index_DragDrop(sender As ListBox, e As DragEventArgs) Handles fileCollection.DragDrop, subFileCollection.DragDrop
        Dim lb As ListBox = DirectCast(sender, ListBox) ' 接收拖曳的 ListBox
        ' 指定正確的目標
        targetList = lb
        Form4.form4_load()
        ' 支援拖曳檔案
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            Form4.setdata(files)
            ' 支援拖曳 URL
        ElseIf e.Data.GetDataPresent(DataFormats.Text) Then
            Dim url As String = CType(e.Data.GetData(DataFormats.Text), String)

            Form4.setdata(url)
            ' 只接受 http(s) 開頭的連結
            'If url.StartsWith("http://") OrElse url.StartsWith("https://") Then

            'End If
        End If
        Form4.Button1_Click(sender, e)
    End Sub

    Private dragfrom = Nothing

    Private Sub Form1_DragEnter(sender As Object, e As DragEventArgs) Handles fileCollection.DragEnter, subFileCollection.DragEnter, PathComboBox.DragEnter, TargetComboBox.DragEnter
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


