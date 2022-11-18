Public Class FMain
    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click

        OpenFileDialog1.Filter = "JPG files (*.jpg)|*.jpg"
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then

            Dim jpg = New jpg
            Dim bmp = jpg.DoJPG(OpenFileDialog1.FileName)

            PictureBox1.Width = bmp.Width
            PictureBox1.Height = bmp.Height
            PictureBox1.Image = bmp
            PictureBox1.Refresh()

        End If

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click

        End

    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click

        MsgBox("JPG Viewer is a port in VB.NET " + vbCrLf + "by Alexandre Lozano Vilanova of" + vbCrLf _
        + "JPGView in VB6 by Dmitry Brant" + vbCrLf _
        + vbCrLf + "http://www.dmitrybrant.com", vbInformation, "About JPGView")

    End Sub

End Class
