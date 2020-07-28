Imports System.IO
Module CopyTilesModule
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
            Image = Image.FromFile(Source & TileName)
            Dim Thumbnail = Image.GetThumbnailImage(50, 50, Nothing, New IntPtr())
            Dim ImgByte() As Byte
            Dim ms As New MemoryStream
            Thumbnail.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            ImgByte = ms.GetBuffer()

            'Add a new row for the current tile.
            NewRow = {SelectTile, ImgByte, TileName, SelectTile, SelectTile, SelectTile}
            DDToolsForm.CopyTilesDataGridView.Rows.Add(NewRow)

            'Set the height of the row to accomodate the thumbnail image.
            DDToolsForm.CopyTilesDataGridView.Rows(RowIndex).Height = 60
            RowIndex += 1
        Next

        'Add a handler for clicking checkboxes.
        AddHandler DDToolsForm.CopyTilesDataGridView.CellContentClick, AddressOf DDToolsForm.CopyTilesDataGridView_CellClick
    End Sub
End Module
