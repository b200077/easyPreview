
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Windows
Public Class Form4
    Public isEditngListbox As Boolean = False
    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        isEditngListbox = True
        Dim targets As String() = CheckedListBox1.CheckedItems.Cast(Of String).ToArray()
        Form1.AddItem(targets)
        isEditngListbox = False
        Form1.refresh_backup()
        Close()
    End Sub
    'Private Sub DeleteItem()
    '    For Each item In CheckedListBox1.CheckedItems
    '        tarForm.Items.Remove(item)
    '    Next
    'End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Close()
    End Sub
    Public Sub Setdata(data As String)
        Setdata({data})
    End Sub
    Public Sub Setdata(data As String())
        If data Is Nothing Then Return
        Dim index As Integer = 0
        For Each files As String In data
            CheckedListBox1.Items.Add(files)
            CheckedListBox1.SetItemChecked(index, True)
            index += 1
        Next
    End Sub
    Public Sub Setdata_from_clipboard()
        Dim items As New List(Of String)

        If Clipboard.ContainsFileDropList Then
            Dim files As Specialized.StringCollection = Clipboard.GetFileDropList
            For Each clipfile As String In files
                items.Add(clipfile)
            Next

        ElseIf Clipboard.GetText.Contains(vbCrLf) Then
            Dim text As String = Clipboard.GetText
            For Each line In text.Split({vbCrLf}, StringSplitOptions.None)
                Dim trimmed As String = line.Trim()
                If String.IsNullOrEmpty(trimmed) Then Continue For
                items.Add(trimmed)
            Next

        Else
            Dim trimmed As String = Clipboard.GetText.Trim
            If Not String.IsNullOrEmpty(trimmed) Then
                items.Add(trimmed)
            End If
        End If

        AddData_toCheckListBox(items)
    End Sub

    Private Sub AddData_toCheckListBox(data As List(Of String))
        If data.Count = 0 Then Return
        Dim startIndex As Integer = CheckedListBox1.Items.Count
        CheckedListBox1.Items.AddRange(data.ToArray())
        For i = startIndex To CheckedListBox1.Items.Count - 1
            CheckedListBox1.SetItemChecked(i, True)
        Next
    End Sub

    Public Sub Form4_Load() Handles MyBase.Load
        Dim f1l As Drawing.Point = Form1.Location
        Location = New Drawing.Point(46 + f1l.X, 35 + f1l.Y)
        CheckedListBox1.Items.Clear()
    End Sub

End Class