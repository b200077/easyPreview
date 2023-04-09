Imports AxWMPLib
Imports ImageMagick
Imports System.IO
Imports System.IO.Compression
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports VisioForge.Libs.AForge.Imaging

Public Class Form4
    Public tarForm As ListBox
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If CheckedListBox1.Items.Count > 0 Then
            If CheckBox1.Checked = False Then
                For Each item In CheckedListBox1.CheckedItems
                    tarForm.Items.Add(item)
                Next
            Else
                Dim selectedindex As Integer = tarForm.SelectedIndex + 1
                Dim oriForm As Array = tarForm.Items.Cast(Of String).ToArray
                Dim count As Integer = CheckedListBox1.CheckedItems.Count
                tarForm.Items.Clear()
                For i = 0 To oriForm.Length - 1 + count

                    If i < selectedindex Then
                        tarForm.Items.Add(oriForm(i))
                    ElseIf i < selectedindex + count Then
                        tarForm.Items.Add(CheckedListBox1.CheckedItems(i - selectedindex))
                    Else
                        tarForm.Items.Add(oriForm(i - count))
                    End If
                Next
                tarForm.SetSelected(selectedindex, True)
            End If
        Else
            If CheckBox1.Checked = False Then
                tarForm.Items.Add(TextBox1.Text)
            Else
                Dim selectedindex As Integer = tarForm.SelectedIndex + 1
                Dim oriForm As Array = tarForm.Items.Cast(Of String).ToArray

                tarForm.Items.Clear()
                For i = 0 To oriForm.Length
                    If i < selectedindex Then
                        tarForm.Items.Add(oriForm(i))
                    ElseIf i = selectedindex Then
                        tarForm.Items.Add(TextBox1.Text)
                    Else
                        tarForm.Items.Add(oriForm(i - 1))
                    End If
                Next
                tarForm.SetSelected(selectedindex, True)
            End If
        End If
        Form1.delete_repeat_item()
        Form1.refresh_backup()
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
    Public Sub form4_load() Handles MyBase.Load

        If Clipboard.ContainsFileDropList Then
            CheckedListBox1.Items.Clear()
            Dim files As Specialized.StringCollection = Clipboard.GetFileDropList
            Dim index As Integer = 0
            For Each clipfile As String In files
                CheckedListBox1.Items.Add(clipfile)
                CheckedListBox1.SetItemChecked(index, True)
                index += 1
            Next

        ElseIf Clipboard.GetText.Contains(vbCrLf) Then
            Dim text As String = Clipboard.GetText
            Dim text_subs = text.Split(vbCrLf)
            Dim index As Integer = 0
            Console.WriteLine(text_subs.Length)
            For Each line In text_subs

                If String.IsNullOrEmpty(line.Trim()) Then Continue For
                CheckedListBox1.Items.Add(line.Trim())
                CheckedListBox1.SetItemChecked(index, True)
                index += 1
                Console.WriteLine(line.Trim())
            Next
        Else
            TextBox1.Text = Clipboard.GetText
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If CheckedListBox1.Items.Contains("https://www.google.com/") Then Return
        CheckedListBox1.Items.Add("https://www.google.com/")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim labelPath = "F:\desktop\常用\常用書籤.trf"
        add_item_from_trf(labelPath)
    End Sub
    Private Sub add_item_from_trf(labelPath As String)
        Using sr As New StreamReader(labelPath)
            While (sr.Peek() >= 0)
                '去掉網址以外字元
                Dim item As String = sr.ReadLine()
                If String.IsNullOrEmpty(item) Then Continue While
                Dim invalidOrders As String() = {"[selected]", "[playTime]", "[opened]", "[Form2_open]"}
                For Each order In invalidOrders
                    If item.Contains(order) Then Continue While
                Next
                CheckedListBox1.Items.Add(item.Trim)
            End While

        End Using
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim labelPath = "F:\desktop\常用\ACG書籤.trf"
        add_item_from_trf(labelPath)
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged

        For sel = 0 To CheckedListBox1.Items.Count - 1
            If CheckBox2.Checked Then
                CheckedListBox1.SetItemChecked(sel, True)
            Else
                CheckedListBox1.SetItemChecked(sel, False)
            End If
        Next
    End Sub
End Class