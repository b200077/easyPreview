<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form2
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form2))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.videoPanel = New System.Windows.Forms.Panel()
        Me.videoControlPanel = New System.Windows.Forms.Panel()
        Me.ComboBox5 = New System.Windows.Forms.ComboBox()
        Me.ProgressBar2 = New System.Windows.Forms.ProgressBar()
        Me.ComboBox8 = New System.Windows.Forms.ComboBox()
        Me.Button24 = New System.Windows.Forms.Button()
        Me.TrackBar1 = New System.Windows.Forms.TrackBar()
        Me.Button25 = New System.Windows.Forms.Button()
        Me.VideoView1 = New LibVLCSharp.WinForms.VideoView()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.WebView21 = New Microsoft.Web.WebView2.WinForms.WebView2()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Button7 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.Button10 = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Button8 = New System.Windows.Forms.Button()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.videoPanel.SuspendLayout()
        Me.videoControlPanel.SuspendLayout()
        CType(Me.TrackBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.VideoView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.WebView21, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.FlowLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.PictureBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PictureBox1.Location = New System.Drawing.Point(0, 0)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(904, 749)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(61, 4)
        '
        'Button5
        '
        Me.Button5.AutoSize = True
        Me.Button5.Location = New System.Drawing.Point(1850, 993)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(54, 39)
        Me.Button5.TabIndex = 7
        Me.Button5.Text = "下一頁"
        Me.Button5.UseVisualStyleBackColor = True
        Me.Button5.Visible = False
        '
        'Button6
        '
        Me.Button6.AutoSize = True
        Me.Button6.Location = New System.Drawing.Point(1850, 952)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(54, 39)
        Me.Button6.TabIndex = 8
        Me.Button6.Text = "上一頁"
        Me.Button6.UseVisualStyleBackColor = True
        Me.Button6.Visible = False
        '
        'TextBox3
        '
        Me.TextBox3.ContextMenuStrip = Me.ContextMenuStrip2
        Me.TextBox3.Dock = System.Windows.Forms.DockStyle.Top
        Me.TextBox3.Font = New System.Drawing.Font("新細明體", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.TextBox3.Location = New System.Drawing.Point(0, 0)
        Me.TextBox3.Margin = New System.Windows.Forms.Padding(2)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(904, 30)
        Me.TextBox3.TabIndex = 14
        '
        'ContextMenuStrip2
        '
        Me.ContextMenuStrip2.Name = "ContextMenuStrip2"
        Me.ContextMenuStrip2.Size = New System.Drawing.Size(61, 4)
        '
        'Panel1
        '
        Me.Panel1.AllowDrop = True
        Me.Panel1.AutoSize = True
        Me.Panel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.Panel1.Controls.Add(Me.videoPanel)
        Me.Panel1.Controls.Add(Me.Label15)
        Me.Panel1.Controls.Add(Me.WebView21)
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 30)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(904, 749)
        Me.Panel1.TabIndex = 17
        '
        'videoPanel
        '
        Me.videoPanel.Controls.Add(Me.Button8)
        Me.videoPanel.Controls.Add(Me.videoControlPanel)
        Me.videoPanel.Controls.Add(Me.VideoView1)
        Me.videoPanel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.videoPanel.Location = New System.Drawing.Point(0, 0)
        Me.videoPanel.Name = "videoPanel"
        Me.videoPanel.Size = New System.Drawing.Size(904, 749)
        Me.videoPanel.TabIndex = 86
        Me.videoPanel.Visible = False
        '
        'videoControlPanel
        '
        Me.videoControlPanel.Controls.Add(Me.ComboBox5)
        Me.videoControlPanel.Controls.Add(Me.ProgressBar2)
        Me.videoControlPanel.Controls.Add(Me.ComboBox8)
        Me.videoControlPanel.Controls.Add(Me.Button24)
        Me.videoControlPanel.Controls.Add(Me.TrackBar1)
        Me.videoControlPanel.Controls.Add(Me.Button25)
        Me.videoControlPanel.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.videoControlPanel.Location = New System.Drawing.Point(0, 692)
        Me.videoControlPanel.Name = "videoControlPanel"
        Me.videoControlPanel.Size = New System.Drawing.Size(904, 57)
        Me.videoControlPanel.TabIndex = 81
        '
        'ComboBox5
        '
        Me.ComboBox5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ComboBox5.FormattingEnabled = True
        Me.ComboBox5.Location = New System.Drawing.Point(762, 27)
        Me.ComboBox5.Name = "ComboBox5"
        Me.ComboBox5.Size = New System.Drawing.Size(130, 20)
        Me.ComboBox5.TabIndex = 76
        Me.ComboBox5.Visible = False
        '
        'ProgressBar2
        '
        Me.ProgressBar2.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar2.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ProgressBar2.Location = New System.Drawing.Point(3, 6)
        Me.ProgressBar2.Maximum = 1000
        Me.ProgressBar2.Name = "ProgressBar2"
        Me.ProgressBar2.Size = New System.Drawing.Size(895, 12)
        Me.ProgressBar2.TabIndex = 82
        '
        'ComboBox8
        '
        Me.ComboBox8.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ComboBox8.FormattingEnabled = True
        Me.ComboBox8.Location = New System.Drawing.Point(502, 27)
        Me.ComboBox8.Name = "ComboBox8"
        Me.ComboBox8.Size = New System.Drawing.Size(201, 20)
        Me.ComboBox8.TabIndex = 77
        Me.ComboBox8.Visible = False
        '
        'Button24
        '
        Me.Button24.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Button24.Location = New System.Drawing.Point(3, 24)
        Me.Button24.Name = "Button24"
        Me.Button24.Size = New System.Drawing.Size(54, 23)
        Me.Button24.TabIndex = 79
        Me.Button24.Text = "播放"
        Me.Button24.UseVisualStyleBackColor = True
        '
        'TrackBar1
        '
        Me.TrackBar1.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.TrackBar1.Location = New System.Drawing.Point(119, 24)
        Me.TrackBar1.Maximum = 100
        Me.TrackBar1.Name = "TrackBar1"
        Me.TrackBar1.Size = New System.Drawing.Size(100, 45)
        Me.TrackBar1.TabIndex = 81
        '
        'Button25
        '
        Me.Button25.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Button25.Location = New System.Drawing.Point(63, 24)
        Me.Button25.Name = "Button25"
        Me.Button25.Size = New System.Drawing.Size(53, 23)
        Me.Button25.TabIndex = 80
        Me.Button25.Text = "停止"
        Me.Button25.UseVisualStyleBackColor = True
        '
        'VideoView1
        '
        Me.VideoView1.BackColor = System.Drawing.Color.Black
        Me.VideoView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.VideoView1.Location = New System.Drawing.Point(0, 0)
        Me.VideoView1.MediaPlayer = Nothing
        Me.VideoView1.Name = "VideoView1"
        Me.VideoView1.Size = New System.Drawing.Size(904, 749)
        Me.VideoView1.TabIndex = 83
        Me.VideoView1.Text = "VideoView1"
        '
        'Label15
        '
        Me.Label15.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(821, 724)
        Me.Label15.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(80, 12)
        Me.Label15.TabIndex = 85
        Me.Label15.Text = "0000px 0000px"
        '
        'WebView21
        '
        Me.WebView21.AllowExternalDrop = True
        Me.WebView21.CreationProperties = Nothing
        Me.WebView21.DefaultBackgroundColor = System.Drawing.Color.White
        Me.WebView21.Dock = System.Windows.Forms.DockStyle.Fill
        Me.WebView21.Location = New System.Drawing.Point(0, 0)
        Me.WebView21.Name = "WebView21"
        Me.WebView21.Size = New System.Drawing.Size(904, 749)
        Me.WebView21.TabIndex = 11
        Me.WebView21.ZoomFactor = 1.0R
        '
        'Button7
        '
        Me.Button7.AutoSize = True
        Me.Button7.Location = New System.Drawing.Point(351, 3)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(54, 22)
        Me.Button7.TabIndex = 13
        Me.Button7.Text = "清單"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.AutoSize = True
        Me.Button4.Location = New System.Drawing.Point(410, 2)
        Me.Button4.Margin = New System.Windows.Forms.Padding(2)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(52, 22)
        Me.Button4.TabIndex = 6
        Me.Button4.Text = "<<"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'TextBox2
        '
        Me.TextBox2.Font = New System.Drawing.Font("新細明體", 12.0!)
        Me.TextBox2.Location = New System.Drawing.Point(467, 3)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(48, 27)
        Me.TextBox2.TabIndex = 11
        Me.TextBox2.Text = "0"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("新細明體", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label2.Location = New System.Drawing.Point(521, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(27, 24)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "/0"
        '
        'Button3
        '
        Me.Button3.AutoSize = True
        Me.Button3.Location = New System.Drawing.Point(553, 2)
        Me.Button3.Margin = New System.Windows.Forms.Padding(2)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(56, 22)
        Me.Button3.TabIndex = 5
        Me.Button3.Text = ">>"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Dock = System.Windows.Forms.DockStyle.Right
        Me.CheckBox1.Location = New System.Drawing.Point(614, 3)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(72, 27)
        Me.CheckBox1.TabIndex = 14
        Me.CheckBox1.Text = "固定上層"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.AutoSize = True
        Me.Button1.Location = New System.Drawing.Point(692, 3)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(27, 22)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "-"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("新細明體", 12.0!)
        Me.TextBox1.Location = New System.Drawing.Point(725, 3)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(33, 27)
        Me.TextBox1.TabIndex = 1
        Me.TextBox1.Text = "100"
        Me.TextBox1.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("新細明體", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label1.Location = New System.Drawing.Point(764, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(29, 24)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "%"
        Me.Label1.Visible = False
        '
        'Button2
        '
        Me.Button2.AutoSize = True
        Me.Button2.Location = New System.Drawing.Point(799, 3)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(29, 22)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "+"
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.Controls.Add(Me.Button10)
        Me.FlowLayoutPanel1.Controls.Add(Me.Button2)
        Me.FlowLayoutPanel1.Controls.Add(Me.Label1)
        Me.FlowLayoutPanel1.Controls.Add(Me.TextBox1)
        Me.FlowLayoutPanel1.Controls.Add(Me.Button1)
        Me.FlowLayoutPanel1.Controls.Add(Me.CheckBox1)
        Me.FlowLayoutPanel1.Controls.Add(Me.Button3)
        Me.FlowLayoutPanel1.Controls.Add(Me.Label2)
        Me.FlowLayoutPanel1.Controls.Add(Me.TextBox2)
        Me.FlowLayoutPanel1.Controls.Add(Me.Button4)
        Me.FlowLayoutPanel1.Controls.Add(Me.Button7)
        Me.FlowLayoutPanel1.Controls.Add(Me.Label3)
        Me.FlowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.FlowLayoutPanel1.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(0, 779)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(904, 32)
        Me.FlowLayoutPanel1.TabIndex = 15
        '
        'Button10
        '
        Me.Button10.Location = New System.Drawing.Point(834, 3)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(67, 23)
        Me.Button10.TabIndex = 16
        Me.Button10.Text = "遠端按鈕"
        Me.Button10.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label3.Location = New System.Drawing.Point(308, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(37, 33)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Label3"
        '
        'Button8
        '
        Me.Button8.Location = New System.Drawing.Point(817, 663)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(75, 23)
        Me.Button8.TabIndex = 84
        Me.Button8.Text = "remove"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(904, 811)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.FlowLayoutPanel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "Form2"
        Me.Text = "Form2"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.videoPanel.ResumeLayout(False)
        Me.videoControlPanel.ResumeLayout(False)
        Me.videoControlPanel.PerformLayout()
        CType(Me.TrackBar1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.VideoView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.WebView21, System.ComponentModel.ISupportInitialize).EndInit()
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.FlowLayoutPanel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents Button5 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents ContextMenuStrip2 As ContextMenuStrip
    Friend WithEvents WebView21 As Microsoft.Web.WebView2.WinForms.WebView2
    Friend WithEvents Timer1 As Timer
    Friend WithEvents Label15 As Label
    Friend WithEvents Button7 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Button3 As Button
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents Button1 As Button
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Button2 As Button
    Friend WithEvents FlowLayoutPanel1 As FlowLayoutPanel
    Friend WithEvents Label3 As Label
    Friend WithEvents Button10 As Button
    Friend WithEvents videoPanel As Panel
    Friend WithEvents VideoView1 As LibVLCSharp.WinForms.VideoView
    Friend WithEvents videoControlPanel As Panel
    Friend WithEvents ComboBox5 As ComboBox
    Friend WithEvents ProgressBar2 As ProgressBar
    Friend WithEvents ComboBox8 As ComboBox
    Friend WithEvents Button24 As Button
    Friend WithEvents TrackBar1 As TrackBar
    Friend WithEvents Button25 As Button
    Friend WithEvents Button8 As Button
End Class
