Imports System.IO
Imports System.Text


Public Class Form3
    Public keywords As String() = New String() {"なし", "音声のみ", "反転", "無"}
    Public ext_in_turn As String() = New String() {".wav", ".flac", ".mp3"}
    Public Sub Form3_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        RadioButton1_CheckedChanged(sender, e)
        RadioButton4_CheckedChanged(sender, e)
        TextBox1.Text = Path.GetFileNameWithoutExtension(Form1.OpenFileDialog1.FileName)
        If Not String.IsNullOrEmpty(Form1.TextBox5.Text) Then
            RadioButton8.Checked = True
            RadioButton5.Checked = True
            RadioButton5_CheckedChanged(sender, e)
        End If
        If Form1.ComboBox5.Text.ToString = "Combine trf files" Then
            RadioButton8.Checked = True
        End If
        For Each RadiobuttonInGroup As RadioButton In GroupBox5.Controls
            AddHandler RadiobuttonInGroup.CheckedChanged, AddressOf Group_CheckedChanged
        Next
    End Sub
    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '防過載
        If Form1.目錄檔案集合.Items.Count > 10000 Then Return
        Dim desk As String = TextBox2.Text
        If String.IsNullOrEmpty(TextBox2.Text) Then
            desk = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString
        End If
        Dim ext_format As String = GroupBox4.Controls.OfType(Of RadioButton)().FirstOrDefault(Function(n) n.Checked).Text
        Dim record As String = desk & "\" & TextBox1.Text & ext_format
        Try
            Using sr As New StreamWriter(record, Not RadioButton7.Checked)
                If Not Form1.OpenFileDialog1.FileName.Contains(".wmc") Then
                    Dim files = Form1.目錄檔案集合.Items.Cast(Of String).ToArray
                    For Each foundfile In files
                        If Form1.ComboBox5.Text = "Combine trf files" And Path.GetExtension(foundfile) = ".trf" Then
                            Form1.trf_backup(foundfile)
                            Form1.壓縮檔案集合.Items.Clear()
                            Form1.read_text_data(foundfile, Form1.壓縮檔案集合)
                            For Each trfFiles In Form1.壓縮檔案集合.Items
                                sr.WriteLine(trfFiles)
                            Next
                        Else
                            sr.WriteLine(foundfile)

                        End If
                    Next
                    'write selected item
                    sr.WriteLine("[selected]" & Form1.目錄檔案集合.SelectedIndex.ToString)
                    'sr.WriteLine("[opened]{" & String.Join(", ", Form6.ListBox1.Items.Cast(Of String).ToList) & "}")
                    'write play progress bar
                    If Form1.AxWindowsMediaPlayer1.Visible Then
                        sr.WriteLine("[playTime]" & Form1.AxWindowsMediaPlayer1.Ctlcontrols.currentPosition.ToString)
                    End If
                    If Form2.Visible And Form2.PictureBox1.Visible Then
                        sr.WriteLine("[Form2_open]")
                    End If
                Else
                    For Each foundfile In Form1.壓縮檔案集合.Items
                        sr.WriteLine(foundfile)
                    Next
                End If

            End Using
        Catch
        End Try
        '----------導致崩潰----------
        ' Form1.Form1_Close(sender, e) 
        If Me.Visible Then MsgBox(record & "輸出完成")
        If String.IsNullOrEmpty(Form1.TextBox5.Text) Then
            Me.Visible = False
            Form1.OpenFileDialog1.FileName = record
            Form1.ComboBox3.Text = record

        ElseIf Me.Visible = True Then
            Form1.CheckBox5.Checked = True
            Form1.Button11_Click_1(sender, e)
            Form1.TextBox5.Text = ""
        End If
        Close()
    End Sub
    Public Function read_item_from_trf(ListPath As String)
        Dim list As String() = New String() {}
        Dim trfAddress As New StringBuilder("")
        trfAddress.Append(ListPath & Path.DirectorySeparatorChar & "playlist.trf")
        Using sr As New StreamReader(trfAddress.ToString)
            While (sr.Peek() >= 0)
                '去掉網址以外字元
                Dim item As String = sr.ReadLine()
                If String.IsNullOrEmpty(item) Then Continue While
                Dim invalidOrders As String() = {"[selected]", "[playTime]", "[opened]", "[Form2_open]"}
                For Each order In invalidOrders
                    If item.Contains(order) Then Continue While
                Next
                list.Append(item.Trim)
            End While

        End Using
        Return list
    End Function
    Public Function get_media_list(ListItems As Object)
        Dim list As String() = New String() {}
        Dim compare As String() = New String() {}
        Dim content As String() = New String() {}
        For Each foundfile In ListItems
            '情況1 檔案在第一層
            Dim type As String = check_file_ext(foundfile)
            If type <> "no_compare" Then
                compare.Append(type)
                compare.Append(foundfile)
            Else
                get_media_list(Directory.GetDirectories(foundfile))
            End If
        Next
        compare = select_media_form(compare)
        If compare.Length <> 0 Then
            content = Array.FindAll(Directory.GetFiles(compare(1)), Function(s) Path.GetExtension(s) = compare(0))
            For Each i In content
                list.Append(i)
            Next
        End If
        Return list
    End Function
    Private Function select_media_form(compare As String())
        For Each ext In ext_in_turn
            For i = 0 To compare.Length - 1
                If compare(i) = ext Then
                    Return New String() {compare(i), compare(i + 1)}
                End If
            Next
        Next
        Return New String() {}
    End Function
    Private Function select_keyword(file() As String)
        For Each keys In keywords
            file = Array.FindAll(file, Function(s) Path.GetFileName(s) <> keys)
        Next
        Return file
    End Function
    Private Function check_file_ext(directory_name As String)
        For Each ext In ext_in_turn
            Dim content() As String = Array.FindAll(Directory.GetFiles(directory_name), Function(s) Path.GetExtension(s) = ext)
            content = select_keyword(content)
            If content.Length <> 0 Then
                Return ext
            End If
        Next
        Return "no_compare"
    End Function
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        TextBox2.Text = Path.GetDirectoryName(Form1.OpenFileDialog1.FileName)
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        TextBox2.Text = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString & "\常用"
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        TextBox2.Text = Form1.ComboBox3.Text
    End Sub



    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.Click
        If RadioButton6.Checked Then TextBox1.Text = DateTime.Now.ToString("yyyy_MM_dd")
    End Sub
    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.Click
        If RadioButton5.Checked Then TextBox1.Text = Form1.TextBox5.Text
    End Sub
    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.Click
        If RadioButton4.Checked Then TextBox1.Text = Path.GetFileNameWithoutExtension(Form1.OpenFileDialog1.FileName)
    End Sub
    Private Sub Group_CheckedChanged(sender As Object, e As EventArgs)
        RadioButton8.Checked = True
        If TextBox1.Text.Contains("網頁") Then
            TextBox1.Text = TextBox1.Text.Split("[網頁]")(0)
        End If
        Dim text As String = CType(sender, RadioButton).Text
        If text = "網頁" Then
            TextBox1.Text += "[網頁]"
            Return
        End If
        TextBox1.Text += "[網頁][" & text & "]"
    End Sub









    ' Private Sub RadioButton9_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton9.CheckedChanged
    '      TextBox2.Text = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile).ToString & "\Music\Playlists"
    '  End Sub
End Class