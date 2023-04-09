<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form4
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。
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

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請勿使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.CheckedListBox1 = New System.Windows.Forms.CheckedListBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(13, 141)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(265, 22)
        Me.TextBox1.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(156, 3)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(54, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "確認"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(216, 3)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(55, 23)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "取消"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'CheckedListBox1
        '
        Me.CheckedListBox1.FormattingEnabled = True
        Me.CheckedListBox1.Location = New System.Drawing.Point(13, 13)
        Me.CheckedListBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.CheckedListBox1.Name = "CheckedListBox1"
        Me.CheckedListBox1.ScrollAlwaysVisible = True
        Me.CheckedListBox1.Size = New System.Drawing.Size(266, 123)
        Me.CheckedListBox1.TabIndex = 3
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(2, 7)
        Me.CheckBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(120, 16)
        Me.CheckBox1.TabIndex = 4
        Me.CheckBox1.Text = "插入選取的下一項"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.CheckBox1)
        Me.Panel1.Controls.Add(Me.Button2)
        Me.Panel1.Location = New System.Drawing.Point(13, 234)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(274, 31)
        Me.Panel1.TabIndex = 5
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(15, 169)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(54, 23)
        Me.Button3.TabIndex = 5
        Me.Button3.Text = "google"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(88, 198)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(67, 23)
        Me.Button4.TabIndex = 6
        Me.Button4.Text = "常用書籤"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(15, 198)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(67, 23)
        Me.Button5.TabIndex = 7
        Me.Button5.Text = "ACG書籤"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(227, 169)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(48, 16)
        Me.CheckBox2.TabIndex = 8
        Me.CheckBox2.Text = "全選"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'Form4
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(287, 277)
        Me.Controls.Add(Me.CheckBox2)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.CheckedListBox1)
        Me.Controls.Add(Me.TextBox1)
        Me.Name = "Form4"
        Me.Text = "Form4"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents CheckedListBox1 As CheckedListBox
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents CheckBox2 As CheckBox
End Class
