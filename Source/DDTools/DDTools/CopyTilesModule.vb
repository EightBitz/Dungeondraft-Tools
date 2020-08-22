Imports System.IO
Imports ImageProcessor
Imports ImageProcessor.Imaging.Formats
Module CopyTilesModule
    Public Function ImageCallBack() As Boolean
        Return True
    End Function
    Public Function LoadThumbnail(ByVal Filename As String)
        'Dim RowIndex As Integer = 0
        Dim imgcallback As Image.GetThumbnailImageAbort = New Image.GetThumbnailImageAbort(AddressOf ImageCallBack)
        'Dim Filename As String
        Dim FileExtension As String
        Dim image As Image

        'For Each AssetDataGridRow In AssetDataGridView.Rows

        'Filename = AssetDataGridRow.Cells("FilePath").Value
        If File.Exists(Filename) Then
            FileExtension = Path.GetExtension(Filename)

            Select Case FileExtension.ToLower
                Case ".bmp"
                    image = New Bitmap(Filename)
                Case ".dds"
                    image = Nothing
                Case ".exr"
                    image = Nothing
                Case ".hdr"
                    image = Nothing
                Case ".jpg"
                    image = New Bitmap(Filename)
                Case ".jpeg"
                    image = New Bitmap(Filename)
                Case ".png"
                    image = New Bitmap(Filename)
                Case ".svg"
                    image = Nothing
                Case ".svgz"
                    image = Nothing
                Case ".tga"
                    image = Nothing
                Case ".webp"
                    image = GetWebPImage(Filename)
                Case Else
                    image = Nothing
            End Select
        Else
            image = Nothing
        End If

        '
        ' Aspect Ratio
        'https://eikhart.com/blog/aspect-ratio-calculator#:~:text=There%20is%20a%20simple%20formula,%3D%20(%20newHeight%20*%20aspectRatio%20)%20.
        '
        Dim imgThumbnail As Bitmap
        If Not image Is Nothing Then
            Dim aspectRatio As Double = image.Width / image.Height
            Dim thumbHeight As Double = 60 / aspectRatio
            Dim thumbWidth As Double = thumbHeight * aspectRatio
            imgThumbnail = New Bitmap(image.GetThumbnailImage(CInt(thumbWidth), CInt(thumbHeight), imgcallback, New IntPtr))

            'AssetDataGridRow.Cells("Thumbnail").Value = imgThumbnail
        Else
            imgThumbnail = Nothing
        End If

        Return imgThumbnail
        'Next
    End Function

    Private Function GetWebPImage(Filename As String)
        Dim photoBytes As Byte() = File.ReadAllBytes(Filename)
        ' Format is automatically detected though can be changed.
        Dim WebPImage As Image

        Dim format As ISupportedImageFormat = New PngFormat With {
                .Quality = 70
            }
        Dim size As Size = New Size(150, 0)

        Using inStream As MemoryStream = New MemoryStream(photoBytes)

            Using outStream As MemoryStream = New MemoryStream()

                ' Initialize the ImageFactory using the overload to preserve EXIF metadata.
                Using imageFactory As ImageFactory = New ImageFactory(preserveExifData:=True)
                    ' Load, resize, set the format and quality and save an image.
                    imageFactory.Load(inStream).Resize(size).Format(format).Save(outStream)
                End Using
                ' Do something with the stream.
                WebPImage = Image.FromStream(outStream)
            End Using
        End Using
        Return WebPImage
    End Function
    Public Sub LoadTilesSub(Source As String, CreateLog As Boolean, LogFileName As String, SelectTile As Boolean, Indent As String)
        Dim FileTypes() As String = {".bmp", ".dds", ".exr", ".hdr", ".jpg", ".jpeg", ".png", ".tga", ".svg", ".svgz", ".webp"}
        Dim Message As String

        Dim VHNames As New ArrayList
        Dim HINames As New ArrayList
        Dim Tiles As New ArrayList
        Dim BaseName

        'Get the list of all source files for supported image types where the file name does not end with "_LO.*" or "_VL.*"
        'This tool is geared toward Campaign Cartographer 3+. With CC3+'s naming convention,
        'files ending with "_LO.*" are of low quality and files ending with "_VL.*" are of very low quality, so we want to skip those.

        Message = Indent & "Getting all source files." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        Dim SourceFiles = From File In (Directory.GetFiles(Source, "*.*", SearchOption.AllDirectories))
                          Where FileTypes.Contains(My.Computer.FileSystem.GetFileInfo(File).Extension.ToLower()) _
                              And Not (Path.GetFileNameWithoutExtension(File)).EndsWith("_LO") _
                              And Not (Path.GetFileNameWithoutExtension(File)).EndsWith("_VL")

        'Assuming CC3+ again, after filtering out "_LO.*" and "_VL." files, we're left with "_VH.*" and "_HI.*"
        '(very high quality and high quality) and standard quality. If we copy all the files, we'll end up with
        'duplicate images. Ultimately, we just want one copy of each image, and preferably the highest version
        'of each image. To help with this, We're setting up one array to hold the stripped-down file names of
        'very-high quality images and another array to hold the stripped-down file names of high-quality images.
        'We'll use these arrays for later comparison when determining which files to copy.
        'If we're not copying from CC3+ folders, then this part won't be of any consequence.

        Message = Indent & "Creating lists of high-quality and very-high-quality files for later comparison to avoid duplicates." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        For Each file In SourceFiles
            BaseName = Path.GetFileNameWithoutExtension(file)
            If BaseName.EndsWith("_VH") Then
                VHNames.Add(BaseName.Replace("_VH", ""))
            ElseIf BaseName.EndsWith("_HI") Then
                HINames.Add(BaseName.Replace("_HI", ""))
            End If
        Next

        'Here, we're creating a collection of objects that, according to CC3+'s naming conventions, are of very-high quality.
        Message = Indent & "Getting all objects of very high quality." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        Dim VHTiles = From File In (Directory.GetFiles(Source, "*.*", SearchOption.AllDirectories))
                      Where FileTypes.Contains(My.Computer.FileSystem.GetFileInfo(File).Extension.ToLower()) _
                                       And (Path.GetFileNameWithoutExtension(File)).EndsWith("_VH")

        'Here, we're creating a collection of objects that, according to CC3+'s naming conventions, are of high quality.
        Message = Indent & "Getting all objects of high quality." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        Dim HITiles = From File In (Directory.GetFiles(Source, "*.*", SearchOption.AllDirectories))
                      Where FileTypes.Contains(My.Computer.FileSystem.GetFileInfo(File).Extension.ToLower()) _
                                       And (Path.GetFileNameWithoutExtension(File)).EndsWith("_HI")

        'Here, we're creating a collection of objects that, according to CC3+'s naming conventions, are of standard quality.
        Message = Indent & "Getting all objects of standard quality." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        Dim STDTiles = From File In (Directory.GetFiles(Source, "*.*", SearchOption.AllDirectories))
                       Where FileTypes.Contains(My.Computer.FileSystem.GetFileInfo(File).Extension.ToLower()) _
                                       And Not (Path.GetFileNameWithoutExtension(File)).EndsWith("_VH") _
                                       And Not (Path.GetFileNameWithoutExtension(File)).EndsWith("_HI") _
                                       And Not (Path.GetFileNameWithoutExtension(File)).EndsWith("_LO") _
                                       And Not (Path.GetFileNameWithoutExtension(File)).EndsWith("_VL")

        'Very-high quality is the default quality we're after, so any such objects we've found
        'will be added, unconditionally, to our final collection of objects to be copied.
        Message = Indent & "Adding all very-high-quality objects to the copy list." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        For Each File As String In VHTiles
            Tiles.Add(Path.GetFileName(File))
        Next

        'We want to grab any high-quality objects that do not have a very-high-quality counterpart.
        'Before adding high-quality objects to our final collection of objects to be copied, we will
        'compare the stripped-down name against the names we collected in the VHNames list.
        Message = Indent & "Adding high-quality objects to the copy list that are not duplicates." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        For Each file As String In HITiles
            BaseName = (Path.GetFileNameWithoutExtension(file)).Replace("_HI", "")
            If Not VHNames.Contains(BaseName) Then
                Tiles.Add(Path.GetFileName(file))
            End If
        Next

        'We want to grab any standard-quality objects that have neither a very-high-quality counterpart nor a high-quality counterpart.
        'Before adding standard-quality objects to our final collection of objects to be copied, we will
        'compare the stripped-down name against the names we collected in the VHNames list as well as
        'the names we collected in the HINames list.
        Message = Indent & "Adding standard-quality objects to the copy list that are not duplicates." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        For Each file As String In STDTiles
            BaseName = Path.GetFileNameWithoutExtension(file)
            If Not VHNames.Contains(BaseName) And Not HINames.Contains(BaseName) Then
                Tiles.Add(Path.GetFileName(file))
            End If
        Next

        'Center our column headers in the data grid view.
        DDToolsForm.CopyTilesDataGridView.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        'Configure our columns for the data grid view.
        Dim SelectRow As New DataGridViewCheckBoxColumn With {
            .HeaderText = "De/Select Row",
            .Name = "SelectRow"
        }
        DDToolsForm.CopyTilesDataGridView.Columns.Add(SelectRow)

        Dim ThumbnailColumn As New DataGridViewImageColumn With {
            .HeaderText = "Thumbnail",
            .Name = "Thumbnail",
            .ReadOnly = True
        }
        DDToolsForm.CopyTilesDataGridView.Columns.Add(ThumbnailColumn)

        Dim NameColumn As New DataGridViewTextBoxColumn With {
            .HeaderText = "Tile Name",
            .Name = "TileName",
            .ReadOnly = True
        }
        DDToolsForm.CopyTilesDataGridView.Columns.Add(NameColumn)
        DDToolsForm.CopyTilesDataGridView.Columns("TileName").Width = 182

        Dim PatternColumn As New DataGridViewCheckBoxColumn With {
            .HeaderText = "Copy to Patterns",
            .Name = "Patterns"
        }
        DDToolsForm.CopyTilesDataGridView.Columns.Add(PatternColumn)

        Dim TerrainColumn As New DataGridViewCheckBoxColumn With {
            .HeaderText = "Copy to Terrain",
            .Name = "Terrain"
        }
        DDToolsForm.CopyTilesDataGridView.Columns.Add(TerrainColumn)

        Dim SimpleTile As New DataGridViewCheckBoxColumn With {
            .HeaderText = "Copy to Tilesets",
            .Name = "SimpleTile"
        }
        DDToolsForm.CopyTilesDataGridView.Columns.Add(SimpleTile)

        'Add a top row where each checkboxes will be used to select all or none of the items
        'in its respective column.
        Dim NewRow As Object() = New Object() {SelectTile, Nothing, "De/Select Column", SelectTile, SelectTile, SelectTile}
        DDToolsForm.CopyTilesDataGridView.Rows.Add(NewRow)

        Dim Image As Image = Nothing
        Dim RowIndex As Integer = 1
        For Each TileName As String In Tiles

            'Get a thumbnail image of the current tile to be used in the thumbnail column of the current row.
            If Not Source.EndsWith("\") Then Source &= "\"
            Dim Filename As String = Source & TileName
            Dim Thumbnail = LoadThumbnail(Filename)

            'Image = Image.FromFile(Source & TileName)
            'Dim Thumbnail = Image.GetThumbnailImage(50, 50, Nothing, New IntPtr())
            'Dim ImgByte() As Byte
            'Dim ms As New MemoryStream
            'Thumbnail.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            'ImgByte = ms.GetBuffer()

            'Add a new row for the current tile.
            NewRow = {SelectTile, Thumbnail, TileName, SelectTile, SelectTile, SelectTile}
            DDToolsForm.CopyTilesDataGridView.Rows.Add(NewRow)

            'Set the height of the row to accomodate the thumbnail image.
            DDToolsForm.CopyTilesDataGridView.Rows(RowIndex).Height = 60
            RowIndex += 1
        Next

        'Add a handler for clicking checkboxes.
        AddHandler DDToolsForm.CopyTilesDataGridView.CellContentClick, AddressOf DDToolsForm.CopyTilesDataGridView_CellClick
    End Sub
End Module
