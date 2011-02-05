Option Explicit On

Imports Microsoft.VisualBasic
Imports System.IO

Public Class Form1
    Dim file1byte As Integer
    Dim file2byte As Integer
    Dim fs1 As FileStream
    Dim fs2 As FileStream
    Private Sub OpenNewInputFile()

        fs1 = New IO.FileStream(TextBox3.Text, IO.FileMode.Open)
        fs1.Close()
    End Sub
    Private Sub OpenInputFile()

        fs1 = New IO.FileStream(TextBox2.Text, IO.FileMode.Open)
        fs1.Close()
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        With dlg
            .InitialDirectory = Application.StartupPath
            .ShowDialog()
            If .FileName <> "" Then
                TextBox2.Text = .FileName
                Call OpenInputFile()
            End If
        End With
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        With dlg
            .InitialDirectory = Application.StartupPath
            .ShowDialog()
            If .FileName <> "" Then
                TextBox3.Text = .FileName
                Call OpenNewInputFile()
            End If
        End With
    End Sub
    Private Function FileCompare(ByVal file1 As String, ByVal file2 As String) As Boolean
        Dim file1byte As Integer
        Dim file2byte As Integer
        Dim fs1 As FileStream
        Dim fs2 As FileStream

        If TextBox2.Text = "" Then
            Exit Function
        ElseIf TextBox3.Text = "" Then
            Exit Function
        ElseIf TextBox2.Text = TextBox3.Text Then
            Exit Function
            fs1.Close()
            fs1 = Nothing
            fs2.Close()
            fs2 = Nothing
        End If
        ' Open the two files.
        fs1 = New FileStream(file1, FileMode.Open)
        fs2 = New FileStream(file2, FileMode.Open)

        ' Check the file sizes.
        If (fs1.Length <> fs2.Length) Then
            ' Close the file
            fs1.Close()
            fs1 = Nothing
            fs2.Close()
            fs2 = Nothing

            ' Return a False value.
            Return False
        End If

        ' Read and compare a byte from each file 
        Do
            file1byte = fs1.ReadByte()
            file2byte = fs2.ReadByte()
        Loop While ((file1byte = file2byte) And (file1byte <> -1))

        ' Close the files.
        fs1.Close()
        fs2.Close()

        ' Return the success of the comparison. 
        Return ((file1byte - file2byte) = 0)
    End Function


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ' Compare the two files 
        If (FileCompare(Me.TextBox2.Text, Me.TextBox3.Text)) Then
            MessageBox.Show("Files are identical")
        ElseIf TextBox2.Text = "" Then
            MsgBox("Please select a file to comapare", MsgBoxStyle.Information + MsgBoxStyle.OkOnly)
        ElseIf TextBox3.Text = "" Then
            MsgBox("Please select a file to compare", MsgBoxStyle.Information + MsgBoxStyle.OkOnly)
        ElseIf TextBox2.Text = TextBox3.Text Then
            MsgBox("Files are identical", MsgBoxStyle.Information + MsgBoxStyle.OkOnly)
        Else
            MessageBox.Show("Files do not match")
        End If

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class

