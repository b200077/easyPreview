Imports System.IO

Public Class Form3
    Private Sub Form3_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        RadioButton1_CheckedChanged(sender, e)
        RadioButton4_CheckedChanged(sender, e)
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim desk As String = TextBox2.Text
        If String.IsNullOrEmpty(TextBox2.Text) Then
            desk = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString
        End If
        Dim ext_format As String = ""
        Dim record As String = ""
        If RadioButton10.Checked = True Then
            ext_format = ".txt"
        Else
            ext_format = ".wpl"
        End If
        record = desk & "\" & TextBox1.Text & ext_format
        If RadioButton7.Checked = True Then
            File.Delete(record)
        End If
        If Not File.Exists(record) Then
            Dim filestream = File.Create(record)
            filestream.Close()
        End If
        Dim ori_content As String = ""
        If RadioButton8.Checked = True Then
            'Using sr As StreamReader = New StreamReader(record)
            ' ori_content = sr.ReadToEnd()
            ' End Using
            ' Using sr As StreamWriter = New StreamWriter(record)
            'sr.Write(ori_content)
            ' End Using
        End If
        If ext_format = ".txt" Then
            Using sr As StreamWriter = New StreamWriter(record, True)
                For Each foundfile In Form1.目錄檔案集合.Items
                    sr.WriteLine(foundfile)
                Next
            End Using
        Else
            Using sr As StreamWriter = New StreamWriter(record, True)
                sr.Write("<?wpl version=""1.0""?>
<smil>
    <head>
        <meta name=""Generator"" content=""Microsoft Windows Media Player -- 12.0.19041.1387""/>
        <meta name = ""ItemCount"" content=""23""/>
        <author/>
        <title>102</title>
    </head>
    <body>
        <seq>")
                For Each foundfile In Form1.目錄檔案集合.Items
                    sr.WriteLine($"<media src=""{foundfile}""/>")
                Next
                sr.Write("        </seq>
    </body>
</smil>
")
            End Using
        End If
        MsgBox("輸出完成")
        Me.Close()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        TextBox2.Text = Path.GetDirectoryName(Form1.OpenFileDialog1.FileName)
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        TextBox2.Text = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        TextBox2.Text = Form1.ComboBox4.Text
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        TextBox1.Text = Path.GetFileNameWithoutExtension(Form1.OpenFileDialog1.FileName)
    End Sub

    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged
        TextBox1.Text = DateTime.Now.ToString("yyyy_MM_dd")
    End Sub

    Private Sub RadioButton9_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton9.CheckedChanged
        TextBox1.Text = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile).ToString & "\Music\Playlists"
    End Sub
End Class