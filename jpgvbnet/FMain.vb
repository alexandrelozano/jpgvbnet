Imports System.Threading

Public Class FMain

    Private jpg As jpg
    Private t As Thread

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click

        OpenFileDialog1.Filter = "JPG files (*.jpg)|*.jpg"
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then

            Cursor = Cursors.WaitCursor

            lblJPGFilename.Text = My.Computer.FileSystem.GetName(OpenFileDialog1.FileName)
            lblJPGSize.Text = GetFileSize(OpenFileDialog1.FileName)
            lblJPGMilliseconds.Text = ""

            PictureBox1.Image = Nothing
            PictureBox1.Refresh()

            MenuStrip1.Enabled = False
            Timer1.Enabled = True

            Dim stopWatch = New Stopwatch()
            jpg = New jpg
            Dim bmp As Bitmap = Nothing
            t = New Thread(Sub()
                               stopWatch.Start()
                               bmp = jpg.DoJPG(OpenFileDialog1.FileName)
                               stopWatch.Stop()
                           End Sub)
            t.Start()

            While t.ThreadState() = ThreadState.Running
                Application.DoEvents()
            End While

            lblJPGMilliseconds.Text = stopWatch.ElapsedMilliseconds & " ms"

            PictureBox1.Image = bmp
            PictureBox1.Refresh()

            pgbJPGDecode.Value = pgbJPGDecode.Maximum

            MenuStrip1.Enabled = True
            Timer1.Enabled = False
            Cursor = Cursors.Default

        End If

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click

        If t IsNot Nothing AndAlso t.IsAlive Then
            t.Abort()
        End If

        End

    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click

        MsgBox("JPG Viewer is a port in VB.NET " + vbCrLf + "by Alexandre Lozano Vilanova of" + vbCrLf _
        + "JPGView in VB6 by Dmitry Brant" + vbCrLf _
        + vbCrLf + "http://www.dmitrybrant.com", vbInformation, "About JPGView")

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        pgbJPGDecode.Maximum = jpg.flen
        pgbJPGDecode.Value = jpg.findex

    End Sub

    Public Function GetFileSize(ByVal TheFile As String) As String

        Dim DoubleBytes As Double

        If TheFile.Length = 0 Then Return ""
        If Not System.IO.File.Exists(TheFile) Then Return ""

        Dim TheSize As ULong = My.Computer.FileSystem.GetFileInfo(TheFile).Length
        Dim SizeType As String = ""

        Try
            Select Case TheSize
                Case Is >= 1099511627776
                    DoubleBytes = CDbl(TheSize / 1099511627776) 'TB
                    Return FormatNumber(DoubleBytes, 2) & " TB"
                Case 1073741824 To 1099511627775
                    DoubleBytes = CDbl(TheSize / 1073741824) 'GB
                    Return FormatNumber(DoubleBytes, 2) & " GB"
                Case 1048576 To 1073741823
                    DoubleBytes = CDbl(TheSize / 1048576) 'MB
                    Return FormatNumber(DoubleBytes, 2) & " MB"
                Case 1024 To 1048575
                    DoubleBytes = CDbl(TheSize / 1024) 'KB
                    Return FormatNumber(DoubleBytes, 2) & " KB"
                Case 0 To 1023
                    DoubleBytes = TheSize ' bytes
                    Return FormatNumber(DoubleBytes, 2) & " bytes"
                Case Else
                    Return ""
            End Select
        Catch
            Return ""
        End Try

    End Function

End Class
