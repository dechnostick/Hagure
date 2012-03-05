Imports System.Diagnostics
Imports System.IO
Imports System.Text

Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Label1.Text = "�N�����Ԃ��v�肽���v���O������" _
            & vbNewLine & "�h���b�v���Ă��������B" _
            & vbNewLine & "(�t�@�C���ł��\�ł�)"
    End Sub

    Private Sub Form1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragEnter

        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub Form1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragDrop

        Dim f As String = CType(e.Data.GetData(DataFormats.FileDrop, False), String())(0)
        Dim p As Process = Process.Start(f)
        Try
            p.WaitForInputIdle()
        Catch ex As Exception

            Me.Label1.Text = Path.GetFileName(f) & vbNewLine & "�v���ł��܂���ł����B"
            Return
        End Try

        Dim t As String = (Math.Floor((DateTime.Now - p.StartTime).TotalMilliseconds) / 1000).ToString()
        Me.Label1.Text = Path.GetFileName(f) & vbNewLine & t & " �b" & vbNewLine

        Dim l As String = Process.GetCurrentProcess().ProcessName & ".csv"

        Using sw As StreamWriter = New StreamWriter(l, True, Encoding.GetEncoding(932))
            sw.WriteLine("""" & DateTime.Now.ToString() & """,""" & f & """,""" & t & """")
            sw.Close()
        End Using
    End Sub
End Class
