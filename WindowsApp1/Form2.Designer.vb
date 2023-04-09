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
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.WebView21 = New Microsoft.Web.WebView2.WinForms.WebView2()
        Me.ChromiumWebBrowser1 = New CefSharp.WinForms.ChromiumWebBrowser()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.FlowLayoutPanel1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.WebView21, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.PictureBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PictureBox1.Location = New System.Drawing.Point(0, 0)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(1904, 955)
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(61, 4)
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("PMingLiU", 12.0!)
        Me.TextBox1.Location = New System.Drawing.Point(311, 3)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(33, 27)
        Me.TextBox1.TabIndex = 1
        Me.TextBox1.Text = "100"
        Me.TextBox1.Visible = False
        '
        'Button1
        '
        Me.Button1.AutoSize = True
        Me.Button1.Location = New System.Drawing.Point(266, 3)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(39, 39)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "-"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'Button2
        '
        Me.Button2.AutoSize = True
        Me.Button2.Location = New System.Drawing.Point(385, 3)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(38, 39)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "+"
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("PMingLiU", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label1.Location = New System.Drawing.Point(350, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(29, 24)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "%"
        Me.Label1.Visible = False
        '
        'Button3
        '
        Me.Button3.AutoSize = True
        Me.Button3.Location = New System.Drawing.Point(205, 2)
        Me.Button3.Margin = New System.Windows.Forms.Padding(2)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(56, 39)
        Me.Button3.TabIndex = 5
        Me.Button3.Text = ">>"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.AutoSize = True
        Me.Button4.Location = New System.Drawing.Point(62, 2)
        Me.Button4.Margin = New System.Windows.Forms.Padding(2)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(52, 39)
        Me.Button4.TabIndex = 6
        Me.Button4.Text = "<<"
        Me.Button4.UseVisualStyleBackColor = True
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
        'TextBox2
        '
        Me.TextBox2.Font = New System.Drawing.Font("PMingLiU", 12.0!)
        Me.TextBox2.Location = New System.Drawing.Point(119, 3)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(48, 27)
        Me.TextBox2.TabIndex = 11
        Me.TextBox2.Text = "0"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("PMingLiU", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label2.Location = New System.Drawing.Point(173, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(27, 24)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "/0"
        '
        'Button7
        '
        Me.Button7.AutoSize = True
        Me.Button7.Location = New System.Drawing.Point(3, 3)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(54, 39)
        Me.Button7.TabIndex = 13
        Me.Button7.Text = "清單"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'TextBox3
        '
        Me.TextBox3.ContextMenuStrip = Me.ContextMenuStrip2
        Me.TextBox3.Dock = System.Windows.Forms.DockStyle.Top
        Me.TextBox3.Font = New System.Drawing.Font("PMingLiU", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.TextBox3.Location = New System.Drawing.Point(0, 0)
        Me.TextBox3.Margin = New System.Windows.Forms.Padding(2)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.ShortcutsEnabled = False
        Me.TextBox3.Size = New System.Drawing.Size(1904, 30)
        Me.TextBox3.TabIndex = 14
        '
        'ContextMenuStrip2
        '
        Me.ContextMenuStrip2.Name = "ContextMenuStrip2"
        Me.ContextMenuStrip2.Size = New System.Drawing.Size(61, 4)
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.Controls.Add(Me.Button7)
        Me.FlowLayoutPanel1.Controls.Add(Me.Button4)
        Me.FlowLayoutPanel1.Controls.Add(Me.TextBox2)
        Me.FlowLayoutPanel1.Controls.Add(Me.Label2)
        Me.FlowLayoutPanel1.Controls.Add(Me.Button3)
        Me.FlowLayoutPanel1.Controls.Add(Me.Button1)
        Me.FlowLayoutPanel1.Controls.Add(Me.TextBox1)
        Me.FlowLayoutPanel1.Controls.Add(Me.Label1)
        Me.FlowLayoutPanel1.Controls.Add(Me.Button2)
        Me.FlowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(0, 993)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(1904, 43)
        Me.FlowLayoutPanel1.TabIndex = 15
        '
        'Panel1
        '
        Me.Panel1.AllowDrop = True
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Controls.Add(Me.WebView21)
        Me.Panel1.Controls.Add(Me.ChromiumWebBrowser1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 38)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1904, 955)
        Me.Panel1.TabIndex = 17
        '
        'WebView21
        '
        Me.WebView21.AllowExternalDrop = True
        Me.WebView21.CreationProperties = Nothing
        Me.WebView21.DefaultBackgroundColor = System.Drawing.Color.White
        Me.WebView21.Location = New System.Drawing.Point(0, 0)
        Me.WebView21.Name = "WebView21"
        Me.WebView21.Size = New System.Drawing.Size(1904, 955)
        Me.WebView21.TabIndex = 11
        Me.WebView21.Visible = False
        Me.WebView21.ZoomFactor = 1.0R
        '
        'ChromiumWebBrowser1
        '
        Me.ChromiumWebBrowser1.ActivateBrowserOnCreation = False
        Me.ChromiumWebBrowser1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ChromiumWebBrowser1.Location = New System.Drawing.Point(0, 0)
        Me.ChromiumWebBrowser1.Margin = New System.Windows.Forms.Padding(2)
        Me.ChromiumWebBrowser1.Name = "ChromiumWebBrowser1"
        Me.ChromiumWebBrowser1.Size = New System.Drawing.Size(1904, 955)
        Me.ChromiumWebBrowser1.TabIndex = 10
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1904, 1036)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.FlowLayoutPanel1)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button5)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "Form2"
        Me.Text = "Form2"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.FlowLayoutPanel1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        CType(Me.WebView21, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Button7 As Button
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents FlowLayoutPanel1 As FlowLayoutPanel
    Friend WithEvents Panel1 As Panel
    Friend WithEvents ChromiumWebBrowser1 As CefSharp.WinForms.ChromiumWebBrowser
    Friend WithEvents ContextMenuStrip2 As ContextMenuStrip
    Friend WithEvents WebView21 As Microsoft.Web.WebView2.WinForms.WebView2
End Class
