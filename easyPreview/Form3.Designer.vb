<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form3
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form3))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.FileNameTextBox = New System.Windows.Forms.TextBox()
        Me.InputRadioButton = New System.Windows.Forms.RadioButton()
        Me.CommonUsedRadioButton = New System.Windows.Forms.RadioButton()
        Me.RadioButton3 = New System.Windows.Forms.RadioButton()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.DesktopRadioButton = New System.Windows.Forms.RadioButton()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.DateRadioButton = New System.Windows.Forms.RadioButton()
        Me.NewNameRadioButton = New System.Windows.Forms.RadioButton()
        Me.OriginalNameRadioButton = New System.Windows.Forms.RadioButton()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.DirectoryNameTextBox = New System.Windows.Forms.TextBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.AppendRadioButton = New System.Windows.Forms.RadioButton()
        Me.OverWriteRadioButton = New System.Windows.Forms.RadioButton()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.RadioButton15 = New System.Windows.Forms.RadioButton()
        Me.TrfRadioButton = New System.Windows.Forms.RadioButton()
        Me.RadioButton10 = New System.Windows.Forms.RadioButton()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(21, 60)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(32, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "名稱:"
        '
        'FileNameTextBox
        '
        Me.FileNameTextBox.Location = New System.Drawing.Point(57, 57)
        Me.FileNameTextBox.Margin = New System.Windows.Forms.Padding(2)
        Me.FileNameTextBox.Name = "FileNameTextBox"
        Me.FileNameTextBox.Size = New System.Drawing.Size(273, 22)
        Me.FileNameTextBox.TabIndex = 1
        '
        'InputRadioButton
        '
        Me.InputRadioButton.AutoSize = True
        Me.InputRadioButton.Checked = True
        Me.InputRadioButton.Location = New System.Drawing.Point(4, 19)
        Me.InputRadioButton.Margin = New System.Windows.Forms.Padding(2)
        Me.InputRadioButton.Name = "InputRadioButton"
        Me.InputRadioButton.Size = New System.Drawing.Size(47, 16)
        Me.InputRadioButton.TabIndex = 2
        Me.InputRadioButton.TabStop = True
        Me.InputRadioButton.Text = "輸入"
        Me.InputRadioButton.UseVisualStyleBackColor = True
        '
        'CommonUsedRadioButton
        '
        Me.CommonUsedRadioButton.AutoSize = True
        Me.CommonUsedRadioButton.Location = New System.Drawing.Point(55, 19)
        Me.CommonUsedRadioButton.Margin = New System.Windows.Forms.Padding(2)
        Me.CommonUsedRadioButton.Name = "CommonUsedRadioButton"
        Me.CommonUsedRadioButton.Size = New System.Drawing.Size(47, 16)
        Me.CommonUsedRadioButton.TabIndex = 3
        Me.CommonUsedRadioButton.Text = "常用"
        Me.CommonUsedRadioButton.UseVisualStyleBackColor = True
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Location = New System.Drawing.Point(106, 19)
        Me.RadioButton3.Margin = New System.Windows.Forms.Padding(2)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(47, 16)
        Me.RadioButton3.TabIndex = 4
        Me.RadioButton3.Text = "來源"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(272, 197)
        Me.Button1.Margin = New System.Windows.Forms.Padding(2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(56, 29)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "儲存"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.DesktopRadioButton)
        Me.GroupBox1.Controls.Add(Me.InputRadioButton)
        Me.GroupBox1.Controls.Add(Me.CommonUsedRadioButton)
        Me.GroupBox1.Controls.Add(Me.RadioButton3)
        Me.GroupBox1.Location = New System.Drawing.Point(23, 93)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Size = New System.Drawing.Size(244, 43)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "路徑"
        '
        'DesktopRadioButton
        '
        Me.DesktopRadioButton.AutoSize = True
        Me.DesktopRadioButton.Location = New System.Drawing.Point(161, 19)
        Me.DesktopRadioButton.Name = "DesktopRadioButton"
        Me.DesktopRadioButton.Size = New System.Drawing.Size(47, 16)
        Me.DesktopRadioButton.TabIndex = 5
        Me.DesktopRadioButton.TabStop = True
        Me.DesktopRadioButton.Text = "桌面"
        Me.DesktopRadioButton.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.DateRadioButton)
        Me.GroupBox2.Controls.Add(Me.NewNameRadioButton)
        Me.GroupBox2.Controls.Add(Me.OriginalNameRadioButton)
        Me.GroupBox2.Location = New System.Drawing.Point(23, 140)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Size = New System.Drawing.Size(188, 38)
        Me.GroupBox2.TabIndex = 7
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "命名"
        '
        'DateRadioButton
        '
        Me.DateRadioButton.AutoSize = True
        Me.DateRadioButton.Location = New System.Drawing.Point(130, 18)
        Me.DateRadioButton.Margin = New System.Windows.Forms.Padding(2)
        Me.DateRadioButton.Name = "DateRadioButton"
        Me.DateRadioButton.Size = New System.Drawing.Size(47, 16)
        Me.DateRadioButton.TabIndex = 7
        Me.DateRadioButton.Text = "日期"
        Me.DateRadioButton.UseVisualStyleBackColor = True
        '
        'NewNameRadioButton
        '
        Me.NewNameRadioButton.AutoSize = True
        Me.NewNameRadioButton.Location = New System.Drawing.Point(67, 18)
        Me.NewNameRadioButton.Margin = New System.Windows.Forms.Padding(2)
        Me.NewNameRadioButton.Name = "NewNameRadioButton"
        Me.NewNameRadioButton.Size = New System.Drawing.Size(59, 16)
        Me.NewNameRadioButton.TabIndex = 6
        Me.NewNameRadioButton.Text = "新名稱"
        Me.NewNameRadioButton.UseVisualStyleBackColor = True
        '
        'OriginalNameRadioButton
        '
        Me.OriginalNameRadioButton.AutoSize = True
        Me.OriginalNameRadioButton.Checked = True
        Me.OriginalNameRadioButton.Location = New System.Drawing.Point(4, 18)
        Me.OriginalNameRadioButton.Margin = New System.Windows.Forms.Padding(2)
        Me.OriginalNameRadioButton.Name = "OriginalNameRadioButton"
        Me.OriginalNameRadioButton.Size = New System.Drawing.Size(59, 16)
        Me.OriginalNameRadioButton.TabIndex = 5
        Me.OriginalNameRadioButton.TabStop = True
        Me.OriginalNameRadioButton.Text = "原名稱"
        Me.OriginalNameRadioButton.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(21, 26)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(32, 12)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "路徑:"
        '
        'DirectoryNameTextBox
        '
        Me.DirectoryNameTextBox.Location = New System.Drawing.Point(57, 23)
        Me.DirectoryNameTextBox.Margin = New System.Windows.Forms.Padding(2)
        Me.DirectoryNameTextBox.Name = "DirectoryNameTextBox"
        Me.DirectoryNameTextBox.Size = New System.Drawing.Size(273, 22)
        Me.DirectoryNameTextBox.TabIndex = 9
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.AppendRadioButton)
        Me.GroupBox3.Controls.Add(Me.OverWriteRadioButton)
        Me.GroupBox3.Location = New System.Drawing.Point(215, 135)
        Me.GroupBox3.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox3.Size = New System.Drawing.Size(118, 43)
        Me.GroupBox3.TabIndex = 10
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "寫入方式"
        '
        'AppendRadioButton
        '
        Me.AppendRadioButton.AutoSize = True
        Me.AppendRadioButton.Location = New System.Drawing.Point(56, 20)
        Me.AppendRadioButton.Margin = New System.Windows.Forms.Padding(2)
        Me.AppendRadioButton.Name = "AppendRadioButton"
        Me.AppendRadioButton.Size = New System.Drawing.Size(59, 16)
        Me.AppendRadioButton.TabIndex = 1
        Me.AppendRadioButton.Text = "不覆寫"
        Me.AppendRadioButton.UseVisualStyleBackColor = True
        '
        'OverWriteRadioButton
        '
        Me.OverWriteRadioButton.AutoSize = True
        Me.OverWriteRadioButton.Checked = True
        Me.OverWriteRadioButton.Location = New System.Drawing.Point(5, 20)
        Me.OverWriteRadioButton.Margin = New System.Windows.Forms.Padding(2)
        Me.OverWriteRadioButton.Name = "OverWriteRadioButton"
        Me.OverWriteRadioButton.Size = New System.Drawing.Size(47, 16)
        Me.OverWriteRadioButton.TabIndex = 0
        Me.OverWriteRadioButton.TabStop = True
        Me.OverWriteRadioButton.Text = "覆寫"
        Me.OverWriteRadioButton.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.RadioButton15)
        Me.GroupBox4.Controls.Add(Me.TrfRadioButton)
        Me.GroupBox4.Controls.Add(Me.RadioButton10)
        Me.GroupBox4.Location = New System.Drawing.Point(23, 182)
        Me.GroupBox4.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox4.Size = New System.Drawing.Size(144, 44)
        Me.GroupBox4.TabIndex = 11
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "寫入格式"
        '
        'RadioButton15
        '
        Me.RadioButton15.AutoSize = True
        Me.RadioButton15.Location = New System.Drawing.Point(47, 20)
        Me.RadioButton15.Margin = New System.Windows.Forms.Padding(2)
        Me.RadioButton15.Name = "RadioButton15"
        Me.RadioButton15.Size = New System.Drawing.Size(48, 16)
        Me.RadioButton15.TabIndex = 2
        Me.RadioButton15.Text = ".wmc"
        Me.RadioButton15.UseVisualStyleBackColor = True
        '
        'TrfRadioButton
        '
        Me.TrfRadioButton.AutoSize = True
        Me.TrfRadioButton.Checked = True
        Me.TrfRadioButton.Location = New System.Drawing.Point(99, 20)
        Me.TrfRadioButton.Margin = New System.Windows.Forms.Padding(2)
        Me.TrfRadioButton.Name = "TrfRadioButton"
        Me.TrfRadioButton.Size = New System.Drawing.Size(37, 16)
        Me.TrfRadioButton.TabIndex = 1
        Me.TrfRadioButton.TabStop = True
        Me.TrfRadioButton.Text = ".trf"
        Me.TrfRadioButton.UseVisualStyleBackColor = True
        '
        'RadioButton10
        '
        Me.RadioButton10.AutoSize = True
        Me.RadioButton10.Location = New System.Drawing.Point(5, 20)
        Me.RadioButton10.Margin = New System.Windows.Forms.Padding(2)
        Me.RadioButton10.Name = "RadioButton10"
        Me.RadioButton10.Size = New System.Drawing.Size(38, 16)
        Me.RadioButton10.TabIndex = 0
        Me.RadioButton10.Text = ".txt"
        Me.RadioButton10.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(184, 200)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(57, 23)
        Me.Button2.TabIndex = 13
        Me.Button2.Text = "網頁"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Form3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(344, 246)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.DirectoryNameTextBox)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.FileNameTextBox)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "Form3"
        Me.Text = "Form3"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents FileNameTextBox As TextBox
    Friend WithEvents InputRadioButton As RadioButton
    Friend WithEvents CommonUsedRadioButton As RadioButton
    Friend WithEvents RadioButton3 As RadioButton
    Friend WithEvents Button1 As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents OriginalNameRadioButton As RadioButton
    Friend WithEvents DateRadioButton As RadioButton
    Friend WithEvents NewNameRadioButton As RadioButton
    Friend WithEvents Label2 As Label
    Friend WithEvents DirectoryNameTextBox As TextBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents AppendRadioButton As RadioButton
    Friend WithEvents OverWriteRadioButton As RadioButton
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents TrfRadioButton As RadioButton
    Friend WithEvents RadioButton10 As RadioButton
    Friend WithEvents RadioButton15 As RadioButton
    Friend WithEvents DesktopRadioButton As RadioButton
    Friend WithEvents Button2 As Button
End Class
