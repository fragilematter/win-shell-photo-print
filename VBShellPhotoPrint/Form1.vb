Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Using openFileDialog1 As Windows.Forms.OpenFileDialog = New Windows.Forms.OpenFileDialog()
            openFileDialog1.Filter = "JPEG Files (*.jpg)|*.jpg|GIF Files (*.gif)|*.gif|All Files (*.*)|*.*|Bitmap Files (*.bmp)|*.bmp|PNG Files (*.png)|*.png|TIFF Files (*.tif, *.tiff)|*.tif;*.tiff|Icon Files (*.ico)|*.ico"
            openFileDialog1.FilterIndex = 1
            openFileDialog1.Multiselect = True
            Dim result As DialogResult = openFileDialog1.ShowDialog()
            If result = DialogResult.OK Then
                PhotoPrinter.Instance.PrintEm(openFileDialog1.FileNames)
            End If
        End Using
    End Sub
End Class
