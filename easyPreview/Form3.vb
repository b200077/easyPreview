Imports System.ComponentModel
Imports System.IO


Public Class Form3

    Public Sub Form3_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        For Each RadiobuttonInGroup As RadioButton In GroupBox2.Controls
            AddHandler RadiobuttonInGroup.CheckedChanged, AddressOf Group2_CheckedChanged
        Next
        For Each RadiobuttonInGroup As RadioButton In GroupBox1.Controls
            AddHandler RadiobuttonInGroup.CheckedChanged, AddressOf Group1_CheckedChanged
        Next
        Group1_CheckedChanged(InputRadioButton, e)
        Group2_CheckedChanged(OriginalNameRadioButton, e)
        FileNameTextBox.Text = Path.GetFileNameWithoutExtension(Form1.PathComboBox.Text)
        If Not String.IsNullOrEmpty(Form1.SearchTextBox.Text) Then
            CommonUsedRadioButton.Checked = True
            AppendRadioButton.Checked = True
            NewNameRadioButton.Checked = True
            Group1_CheckedChanged(CommonUsedRadioButton, e)
            Group2_CheckedChanged(NewNameRadioButton, e)
            DirectoryNameTextBox.Text = Path.Combine(DirectoryNameTextBox.Text, "個別網站")
        End If
        If Form1.CheckedListBox1_isCheck("Combine trf files") Then
            AppendRadioButton.Checked = True
            CommonUsedRadioButton.Checked = True
            DateRadioButton.Checked = True
            Group1_CheckedChanged(CommonUsedRadioButton, e)
            Group2_CheckedChanged(DateRadioButton, e)
            FileNameTextBox.Text += "網頁整合"
        End If
    End Sub
    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ext_format As String = GroupBox4.Controls.OfType(Of RadioButton)().FirstOrDefault(Function(n) n.Checked).Text
        Form1.saveTrf(DirectoryNameTextBox.Text, FileNameTextBox.Text, Not OverWriteRadioButton.Checked, ext_format)
    End Sub
    Public Sub AutoSort(collections As List(Of String), tarlist As ListBox)
        If collections Is Nothing Or tarlist Is Nothing Then Return
        Dim collectionsNumber As Integer = collections.Count
        tarlist.BeginUpdate()
        For Each webType As String In Constants.GetValue("webTypes")
            Dim savename As String = webType
            Dim keyword As String = webType
            If webType.Contains(":") Then
                Dim tempSort As String() = webType.Split(":")
                savename = tempSort(0)
                keyword = tempSort(1)
            End If
            Dim record As String = Path.Combine(Form1.commonUsed, "個別網站", savename & ".trf")
            Dim lines As New List(Of String)
            For i As Integer = collections.Count - 1 To 0 Step -1
                Dim foundfile = collections(i)
                If Not foundfile.Contains(keyword) Then Continue For
                lines.Add(foundfile)
                Form1.SubFileCollection.Items.Add($"{foundfile}寫入{savename}")
                collections.RemoveAt(i)
                If tarlist IsNot Nothing Then tarlist.Items.Remove(foundfile)
            Next
            ' 如果有符合的檔案，寫入檔案
            If lines.Count > 0 Then
                lines.Reverse()
                Using sw As New StreamWriter(record, True)
                    sw.WriteLine(String.Join(Environment.NewLine, lines))
                End Using
            End If
            collectionsNumber = collections.Count
            If collectionsNumber = 0 Then
                Exit For
            End If
        Next
        If tarlist Is Nothing And collectionsNumber > 0 Then
            Form1.targetList.Items.AddRange(collections.ToArray)
        End If
        tarlist.EndUpdate()
    End Sub
    Public Sub AutoSort(collection As String, tarlist As ListBox)
        If tarlist Is Nothing Or collection Is Nothing Then Return

        For Each webType As String In Constants.GetValue("webTypes")
            Dim savename As String = webType
            Dim keyword As String = webType
            If webType.Contains(":") Then
                Dim tempSort As String() = webType.Split(":")
                savename = tempSort(0)
                keyword = tempSort(1)
            End If

            If Not collection.Contains(keyword) Then Continue For
            Form1.SubFileCollection.Items.Add($"{collection}寫入{savename}")
            If tarlist IsNot Nothing Then tarlist.Items.Remove(collection)
            ' 如果有符合的檔案，寫入檔案
            Dim record As String = Path.Combine(Form1.commonUsed, "個別網站", savename & ".trf")
            Using sw As New StreamWriter(record, True)
                sw.WriteLine(collection)
            End Using
            Return
        Next
        tarlist.Items.Add(collection)
    End Sub
    Private Sub Group2_CheckedChanged(sender As Object, e As EventArgs)
        Dim name As String = CType(sender, RadioButton).Name
        Select Case name
            Case "OriginalNameRadioButton"
                If OriginalNameRadioButton.Checked Then FileNameTextBox.Text = Path.GetFileNameWithoutExtension(Form1.PathComboBox.Text)
            Case "NewNameRadioButton"
                If NewNameRadioButton.Checked Then FileNameTextBox.Text = Form1.SearchTextBox.Text
            Case "DateRadioButton"
                If DateRadioButton.Checked Then FileNameTextBox.Text = DateTime.Now.ToString("yyyy_MM_dd")
        End Select
    End Sub
    Private Sub Group1_CheckedChanged(sender As Object, e As EventArgs)
        Dim name As String = CType(sender, RadioButton).Name
        Select Case name
            Case "InputRadioButton"
                If Not File.Exists(Form1.PathComboBox.Text) Then Return
                DirectoryNameTextBox.Text = Path.GetDirectoryName(Form1.PathComboBox.Text)
            Case "CommonUsedRadioButton"
                DirectoryNameTextBox.Text = Form1.commonUsed
            Case "RadioButton3"
                DirectoryNameTextBox.Text = Form1.PathComboBox.Text
            Case "DesktopRadioButton"
                DirectoryNameTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        End Select
    End Sub
    Private Sub WebRadioButton_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        AppendRadioButton.Checked = True
        If FileNameTextBox.Text.Contains("網頁") Then
            FileNameTextBox.Text = FileNameTextBox.Text.Split("[網頁]")(0)
        End If
        CommonUsedRadioButton.Checked = True
        DateRadioButton.Checked = True
        Group1_CheckedChanged(CommonUsedRadioButton, e)
        Group2_CheckedChanged(DateRadioButton, e)
        FileNameTextBox.Text += "[網頁]"
    End Sub
End Class