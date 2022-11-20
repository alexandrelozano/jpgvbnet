Imports System.Threading

Public Class FMain

    Private jpg As jpg
    Private t As Thread

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click

        OpenFileDialog1.Filter = "JPG files (*.jpg)|*.jpg"
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then

            Cursor = Cursors.WaitCursor

            PictureBox1.Image = Nothing
            PictureBox1.Refresh()

            MenuStrip1.Enabled = False
            ProgressBar1.Visible = True
            Timer1.Enabled = True

            jpg = New jpg
            Dim bmp As Bitmap = Nothing
            t = New Thread(Sub()
                               bmp = jpg.DoJPG(OpenFileDialog1.FileName)
                           End Sub)
            t.Start()

            While t.ThreadState() = ThreadState.Running
                Application.DoEvents()
            End While

            PictureBox1.Image = bmp
            PictureBox1.Refresh()

            MenuStrip1.Enabled = True
            ProgressBar1.Visible = False
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

        t.Suspend()

        ProgressBar1.Maximum = jpg.flen
        ProgressBar1.Value = jpg.findex
        ProgressBar1.Refresh()

        PictureBox1.Image = jpg.bmp.Clone

        t.Resume()

    End Sub

End Class
