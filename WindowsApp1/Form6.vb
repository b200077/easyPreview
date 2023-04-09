Public Class Form6
    Public undoList As ArrayList
    Public undoLabel As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        undoList.Add(ListBox1.Items.Cast(Of String).ToArray)
        ListBox1.Items.Remove(ListBox1.SelectedItem)
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Form2.Index = CInt(ListBox1.SelectedItem)
        Form2.Button3_Click(sender, e)
    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim newForm As New Form1
        For Each item In ListBox1.Items
            Dim realName As String
            If Form1.壓縮檔案集合.Items.Count = 0 And Form1.壓縮檔案集合.SelectedItem = Nothing Then
                realName = Form1.目錄檔案集合.Items(item)
            Else
                realName = Form1.壓縮檔案集合.Items(item)
            End If
            newForm.目錄檔案集合.Items.Add(realName)
        Next
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click, Button2.Click
        Dim buttonName As String = CType(sender, Button).Text
        If buttonName = "redo" Then
            If undoLabel >= undoList.Count Then Return
            undoLabel += 1
        ElseIf buttonName = "undo" Then
            If undoLabel = 0 Then Return
            undoLabel -= 1
        End If
        refresh_list()
    End Sub
    Private Sub refresh_list()
        ListBox1.Items.Clear()
        For Each label In undoList(undoLabel)
            ListBox1.Items.Add(label)
        Next
    End Sub
End Class