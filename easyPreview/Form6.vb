Imports System.Text.RegularExpressions

Public Class Form6
    Dim source As List(Of String)
    Public Sub Form6_setup(s As List(Of String))
        If s Is Nothing Then Return
        Label1.Text = "共有 " & s.Count & " 項資源，請輸入要下載的數量："
        source = s
        StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Shown
        RadioButton1.Checked = True
        RadioButton_CheckedChanged(sender, e)
        ActiveControl = TextBox1
    End Sub
    Public Sub Form6_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e Is Nothing Then Return
        Select Case e.KeyCode
            Case Keys.Enter
                Button1_Click(sender, e)
            Case Keys.Escape
                Form_Hide(sender, e)
        End Select
    End Sub

    Private Async Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim parts = TextBox1.Text.Split(","c)
        Dim selectedIndexes As New HashSet(Of Integer)
        For Each part In parts
            Dim token = part.Trim()
            Dim pattern As String = "(\d+)~(\d+)"
            Dim m As Match = Regex.Match(token, pattern)
            If m.Success Then
                Dim startIdx = Integer.Parse(m.Groups(1).Value)
                Dim endIdx = Integer.Parse(m.Groups(2).Value)
                If startIdx > endIdx Then
                    Dim tmp = startIdx
                    startIdx = endIdx
                    endIdx = tmp
                End If
                For i = startIdx To endIdx
                    Dim idx = i - 1
                    If idx >= 0 AndAlso idx < source.Count Then
                        selectedIndexes.Add(idx)
                    End If
                Next
                Continue For
            End If
            Dim n As Integer
            If Integer.TryParse(token, n) Then
                Dim idx = n - 1
                If idx >= 0 AndAlso idx < source.Count Then
                    selectedIndexes.Add(idx)
                End If
            End If

        Next
        If selectedIndexes.Count = 0 Then
            MessageBox.Show("輸入無效，請輸入數字。")
            Return
        End If
        For Each i In selectedIndexes
            Dim uri = source(i)
            Console.WriteLine(uri)
            Await Form2.DownloadOriginalImage(uri)
        Next
    End Sub

    Private Sub RadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.Click, RadioButton2.Click
        If RadioButton1.Checked Then
            TextBox1.Text = "1"
        ElseIf RadioButton2.Checked Then
            TextBox1.Text = "1~" & source.Count.ToString()
        End If
        ActiveControl = TextBox1
    End Sub

    Private Sub Form_Hide(sender As Object, e As EventArgs) Handles Button2.Click, Button1.Click
        Hide()
    End Sub

End Class