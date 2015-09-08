Imports System.IO
Imports System.Drawing.Drawing2D

Public Class MainForm
    Private ax As Integer = 0
    Private bx As Integer = 0
    Private ay As Integer = 0
    Private by As Integer = 0
    Private TileW As Integer = 64
    Private z As Integer = 8

    Private Delegate Sub DelegPrintoPictureBox(ByVal pctb As PictureBox, ByVal image1 As Image)
    Private Delegate Sub DelegPrintoLabel(ByVal lbl As Label, ByVal txt As String)
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'todo check if nothing and use textbox value
        TextBox1.Text = Me.z
        TextBox2.Text = Me.ax
        TextBox3.Text = Me.bx
        TextBox4.Text = Me.ay
        TextBox5.Text = Me.by

        TextBox6.Text = Me.UpdateTxtbxX()
        TextBox7.Text = Me.UpdateTxtbxY()

        TextBox8.Text = Me.TileW
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox6.Text = UpdateTxtbxX().ToString
        End If
    End Sub
    Private Function UpdateTxtbxX() As Integer
        Dim d As Integer = Me.bx - Me.ax
        Dim e As Integer = d * Me.TileW
        Return e
    End Function
    Private Function UpdateTxtbxY() As Integer
        Dim d As Integer = Me.by - Me.ay
        Dim e As Integer = d * Me.TileW
        Return e
    End Function

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged, TextBox3.TextChanged
        If CheckBox1.Checked = True Then
            Me.ax = TextBox2.Text
            Me.bx = TextBox3.Text
            TextBox6.Text = Me.UpdateTxtbxX().ToString
        End If
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged, TextBox4.TextChanged
        If CheckBox1.Checked = True Then
            Me.ay = TextBox4.Text
            Me.by = TextBox5.Text
            TextBox7.Text = UpdateTxtbxY().ToString
        End If
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        PictureBox1.Width = TextBox6.Text
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        PictureBox1.Height = TextBox7.Text
    End Sub

    Private Sub Main()
        Dim theThread _
            As New Threading.Thread(
                AddressOf CombineTile)
        theThread.Start()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Main()
    End Sub
    Public Shared Function ResizeImage(ByVal image As Image, _
  ByVal size As Size, Optional ByVal preserveAspectRatio As Boolean = True) As Image
        Dim newWidth As Integer
        Dim newHeight As Integer
        If preserveAspectRatio Then
            Dim originalWidth As Integer = image.Width
            Dim originalHeight As Integer = image.Height
            Dim percentWidth As Single = CSng(size.Width) / CSng(originalWidth)
            Dim percentHeight As Single = CSng(size.Height) / CSng(originalHeight)
            Dim percent As Single = If(percentHeight < percentWidth,
                    percentHeight, percentWidth)
            newWidth = CInt(originalWidth * percent)
            newHeight = CInt(originalHeight * percent)
        Else
            newWidth = size.Width
            newHeight = size.Height
        End If
        Dim newImage As Image = New Bitmap(newWidth, newHeight)
        Using graphicsHandle As Graphics = Graphics.FromImage(newImage)
            graphicsHandle.InterpolationMode = InterpolationMode.NearestNeighbor
            graphicsHandle.DrawImage(image, 0, 0, newWidth, newHeight)
        End Using
        Return newImage
    End Function
    Private Sub CombineTile()

        Dim stopWatch As New Stopwatch()
        Me.PrintoPictureBox(PictureBox1, Nothing)
        Dim axmi As Integer = CInt(TextBox2.Text)
        Dim axmx As Integer = CInt(TextBox3.Text)
        Dim aymi As Integer = CInt(TextBox4.Text)
        Dim aymx As Integer = CInt(TextBox5.Text)
        Dim TW As Integer = TileW 'CInt(TextBox8.Text)
        Dim path As String = TextBox9.Text
        Dim z As Integer = CInt(TextBox1.Text)
        Dim PosCursX As Integer = 0
        Dim PosCursY As Integer = 0

        Dim folders() As String = Directory.GetDirectories(path & "\" & z)
        Dim xmin As Integer = 0
        Dim xmax As Integer = 0
        Dim ymin As Integer = 0
        Dim ymax As Integer = 0
        For Each _folder As String In folders

            Dim files() As String = Directory.GetFiles(_folder, "*.png")
            _folder = New DirectoryInfo(_folder).Name
            For Each _file As String In files
                _file = IO.Path.GetFileNameWithoutExtension(_file)
                If CInt(_file) < xmin Then
                    ymin = CInt(_file)
                ElseIf CInt(_file) > ymax Then
                    ymax = CInt(_file)
                End If
            Next
            If CInt(_folder) < xmin Then
                xmin = CInt(_folder)
            ElseIf CInt(_folder) > xmax Then
                xmax = CInt(_folder)
            End If
        Next
        Me.Invoke(Sub()
                      TextBox2.Text = xmin
                      TextBox3.Text = xmax
                      TextBox4.Text = ymin
                      TextBox5.Text = ymax
                  End Sub)
        Dim ss As Integer = (xmax - xmin + 1) * TileW
        Dim ss2 As Integer = (ymax - ymin + 1) * TileW
        Dim image2 As New Bitmap(ss, ss2)
        stopWatch.Start()
        For x As Integer = xmin To xmax
            ' Me.PrintoLabel(Label12, x)
            For y As Integer = ymin To ymax
                PosCursX = (x - xmin) * TW
                PosCursY = (y - ymin) * TW

                '   Me.PrintoLabel(Label13, y)
                Dim fullpath As String = path & "\" & z & "\" & x & "\" & y & ".png"
                Dim g As Graphics = Graphics.FromImage(image2)
                Try
                    If System.IO.File.Exists(fullpath) Then
                        Try
                            Dim imageL As New Bitmap(Image.FromFile(fullpath))
                            'PictureBox2.Image = imageL
                            imageL = ResizeImage(imageL, New Size(TileW, TileW), True)
                            g.DrawImage(imageL, New Point(CInt(PosCursX), CInt(PosCursY)))
                        Catch ex As Exception
                            g.Dispose()
                        End Try

                    Else
                        Try
                            Dim Rect1 As Rectangle
                            Rect1 = New Rectangle(New Point(CInt(PosCursX), CInt(PosCursY)), New Size(TileW, TileW))
                            Dim labrush As New SolidBrush(Color.Black)
                            g.FillRectangle(labrush, Rect1)
                            g.DrawRectangle(Pens.Black, Rect1)
                            '  Console.WriteLine(fullpath)
                            labrush.Dispose()
                        Catch ex As Exception
                            g.Dispose()
                            g = Nothing
                        End Try

                    End If
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                End Try
                g.Dispose()
                g = Nothing
            Next
        Next
        ExportMyMap2(image2)

        stopWatch.Stop()
        Dim ts As TimeSpan = stopWatch.Elapsed
        Dim elapsedTime As String = [String].Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10)
        Me.PrintoLabel(Label12, elapsedTime)
        image2.Dispose()
    End Sub

    Private Sub PrintoPictureBox(ByVal pctb As PictureBox, ByVal image1 As Image)
        If pctb.InvokeRequired Then
            pctb.Invoke(New DelegPrintoPictureBox(AddressOf PrintoPictureBox), New Object() {pctb, image1})
        Else
            pctb.Image = image1
        End If
    End Sub
    Private Sub PrintoLabel(ByVal lbl As Label, ByVal txt As String)
        If lbl.InvokeRequired Then
            lbl.Invoke(New DelegPrintoLabel(AddressOf PrintoLabel), New Object() {lbl, txt})
        Else
            lbl.Text = txt
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With FolderBrowserDialog1
            .Description = "TODO"
            If TextBox9.Text <> "" Then .SelectedPath = TextBox9.Text
            If .ShowDialog = System.Windows.Forms.DialogResult.Cancel Then
                Exit Sub
            End If
            TextBox9.Text = .SelectedPath
        End With
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.ExportMyMap("png")
    End Sub
    Public Function ExportMyMap2(ByVal img As Bitmap) As Boolean
        Me.Invoke(Sub()
                      Dim fily As String
                      Dim SaveFileDialog1 As New SaveFileDialog
                      If TextBox10.Text.Length > 7 Then
                          fily = TextBox10.Text
                      Else
                          With SaveFileDialog1
                              .Filter = ("Image Files|*.jpg;*.gif;*.bmp;*.png;*.jpeg|All Files|*.*")
                              .DefaultExt = "png"
                              .FileName = "*." & "png"
                              .Title = "Export My Map"
                              If .ShowDialog = System.Windows.Forms.DialogResult.Cancel Then
                                  '  Return False
                                  Exit Sub
                              End If
                              fily = .FileName
                              TextBox10.Text = fily
                              If .FileName.Length <= 0 Or .FileName = "" Then
                                  '  Return False
                                  Exit Sub
                              End If
                          End With
                      End If
                      img.Save(fily)
                      Try
                          Me.PrintoPictureBox(PictureBox1, Image.FromFile(fily))
                      Catch ex As Exception

                      End Try

                  End Sub)

    End Function
    Public Function ExportMyMap(ByVal imgFormat As String) As Boolean
        Dim fily As String
        Dim SaveFileDialog1 As New SaveFileDialog
        If TextBox10.Text.Length > 7 Then
            fily = TextBox10.Text
        Else
            With SaveFileDialog1
                .Filter = ("Image Files|*.jpg;*.gif;*.bmp;*.png;*.jpeg|All Files|*.*")
                .DefaultExt = imgFormat
                .FileName = "*." & imgFormat
                .Title = "Export My Map"
                If .ShowDialog = System.Windows.Forms.DialogResult.Cancel Then
                    Return False
                    Exit Function
                End If
                fily = .FileName
                TextBox10.Text = fily
                If .FileName.Length <= 0 Or .FileName = "" Then
                    Return False
                    Exit Function
                End If
            End With
        End If
        Try
            Select Case imgFormat
                Case Is = "png"
                    PictureBox1.Image.Save(Path.GetFullPath(fily), Imaging.ImageFormat.Png)
                Case Is = "jpg"
                    PictureBox1.Image.Save(Path.GetFullPath(fily), Imaging.ImageFormat.Jpeg)
                Case Is = "bmp"
                    PictureBox1.Image.Save(Path.GetFullPath(fily), Imaging.ImageFormat.Bmp)
            End Select
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
End Class
