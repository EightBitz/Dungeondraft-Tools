Imports System.IO
Module CopyAssetsModule
    Public Sub CopyAssetsSub(Source As String, Destination As String, CreateTagFile As Boolean, DefaultTag As String, Portals As Boolean, CreateLog As Boolean, LogFileName As String, Indent As String)

        Dim ParentSource As String = Source
        Dim ParentDestination As String = Destination
        Dim VHNames As New ArrayList
        Dim HINames As New ArrayList
        Dim DestinationObjects As New ArrayList
        Dim DestinationPortals As New ArrayList
        Dim ObjectFolders As New ArrayList
        Dim PortalFolders As New ArrayList
        Dim ReplaceSource As String
        Dim FolderDestination As String
        Dim CopySource As String
        Dim CopyDestination As String
        Dim PortalSource As String
        Dim Message As String
        Dim BaseName As String
        Dim FileTypes() As String = {".bmp", ".dds", ".exr", ".hdr", ".jpg", ".jpeg", ".png", ".tga", ".svg", ".svgz", ".webp"}

        Dim TexturePath As String = "\textures"
        Dim ObjectPath As String = TexturePath & "\objects"
        Dim PortalPath As String = TexturePath & "\portals"

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

        'CC3+ names its doors starting with "Door " or "Doors ", and it names its windows starting with "Window " or "Windows ".
        'In case the option was selected to separate portals, we want to keep standard objects separate from portals.

        'Here, we're creating a collection of objects that, according to CC3+'s naming conventions,
        'are of very-high quality and are not portals (doors or windows).
        Message = Indent & "Getting all objects of very high quality." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        Dim VHDestinationObjects = From File In (Directory.GetFiles(Source, "*.*", SearchOption.AllDirectories))
                                   Where FileTypes.Contains((Path.GetExtension(File)).ToLower()) _
                                       And (Path.GetFileNameWithoutExtension(File)).EndsWith("_VH") _
                                       And Not (Path.GetFileName(File)).StartsWith("Door ") _
                                       And Not (Path.GetFileName(File)).StartsWith("Doors ") _
                                       And Not (Path.GetFileName(File)).StartsWith("Window ") _
                                       And Not (Path.GetFileName(File)).StartsWith("Windows ")

        'Here, we're creating a collection of objects that, according to CC3+'s naming conventions,
        'are of high quality and are not portals (doors or windows).
        Message = Indent & "Getting all objects of high quality." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        Dim HIDestinationObjects = From File In (Directory.GetFiles(Source, "*.*", SearchOption.AllDirectories))
                                   Where FileTypes.Contains((Path.GetExtension(File)).ToLower()) _
                                       And (Path.GetFileNameWithoutExtension(File)).EndsWith("_HI") _
                                       And Not (Path.GetFileName(File)).StartsWith("Door ") _
                                       And Not (Path.GetFileName(File)).StartsWith("Doors ") _
                                       And Not (Path.GetFileName(File)).StartsWith("Window ") _
                                       And Not (Path.GetFileName(File)).StartsWith("Windows ")

        'Here, we're creating a collection of objects that, according to CC3+'s naming conventions,
        'are of standard quality and are not portals (doors or windows).
        Message = Indent & "Getting all objects of standard quality." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        Dim STDDestinationObjects = From File In (Directory.GetFiles(Source, "*.*", SearchOption.AllDirectories))
                                    Where FileTypes.Contains((Path.GetExtension(File)).ToLower()) _
                                       And Not (Path.GetFileNameWithoutExtension(File)).EndsWith("_VH") _
                                       And Not (Path.GetFileNameWithoutExtension(File)).EndsWith("_HI") _
                                       And Not (Path.GetFileNameWithoutExtension(File)).EndsWith("_LO") _
                                       And Not (Path.GetFileNameWithoutExtension(File)).EndsWith("_VL") _
                                       And Not (Path.GetFileName(File)).StartsWith("Door ") _
                                       And Not (Path.GetFileName(File)).StartsWith("Doors ") _
                                       And Not (Path.GetFileName(File)).StartsWith("Window ") _
                                       And Not (Path.GetFileName(File)).StartsWith("Windows ")

        'Very-high quality is the default quality we're after, so any such objects we've found
        'will be added, unconditionally, to our final collection of objects to be copied.
        Message = Indent & "Adding all very-high-quality objects to the copy list." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        For Each File As String In VHDestinationObjects
            DestinationObjects.Add(File)
        Next

        'We want to grab any high-quality objects that do not have a very-high-quality counterpart.
        'Before adding high-quality objects to our final collection of objects to be copied, we will
        'compare the stripped-down name against the names we collected in the VHNames list.
        Message = Indent & "Adding high-quality objects to the copy list that are not duplicates." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        For Each file As String In HIDestinationObjects
            BaseName = Path.GetFileNameWithoutExtension(file).Replace("_HI", "")
            If Not VHNames.Contains(BaseName) Then
                DestinationObjects.Add(file)
            End If
        Next

        'We want to grab any standard-quality objects that have neither a very-high-quality counterpart
        'nor a high-quality counterpart.
        'Before adding standard-quality objects to our final collection of objects to be copied, we will
        'compare the stripped-down name against the names we collected in the VHNames list as well as
        'the names we collected in the HINames list.
        Message = Indent & "Adding standard-quality objects to the copy list that are not duplicates." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        For Each file As String In STDDestinationObjects
            BaseName = Path.GetFileNameWithoutExtension(file)
            If Not VHNames.Contains(BaseName) And Not HINames.Contains(BaseName) Then
                DestinationObjects.Add(file)
            End If
        Next

        'Here, we're creating a collection of objects that, according to CC3+'s naming conventions,
        'are of very-high quality and are portals (doors or windows).
        Message = Indent & "Getting all portals of very high quality." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        Dim VHDestinationPortals = From File In (Directory.GetFiles(Source, "*.*", SearchOption.AllDirectories))
                                   Where FileTypes.Contains((Path.GetExtension(File)).ToLower()) _
                                       And (Path.GetFileNameWithoutExtension(File)).EndsWith("_VH") _
                                       And ((Path.GetFileName(File)).StartsWith("Door ") _
                                       Or (Path.GetFileName(File)).StartsWith("Doors ") _
                                       Or (Path.GetFileName(File)).StartsWith("Window ") _
                                       Or (Path.GetFileName(File)).StartsWith("Windows "))

        'Here, we're creating a collection of objects that, according to CC3+'s naming conventions,
        'are of high-quality and are portals (doors or windows).
        Message = Indent & "Getting all portals of high quality." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        Dim HIDestinationPortals = From File In (Directory.GetFiles(Source, "*.*", SearchOption.AllDirectories))
                                   Where FileTypes.Contains((Path.GetExtension(File)).ToLower()) _
                                       And (Path.GetFileNameWithoutExtension(File)).EndsWith("_HI") _
                                       And ((Path.GetFileName(File)).StartsWith("Door ") _
                                       Or (Path.GetFileName(File)).StartsWith("Doors ") _
                                       Or (Path.GetFileName(File)).StartsWith("Window ") _
                                       Or (Path.GetFileName(File)).StartsWith("Windows "))

        'Here, we're creating a collection of objects that, according to CC3+'s naming conventions,
        'are of standard-quality and are portals (doors or windows).
        Message = Indent & "Getting all portals of standard quality." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        Dim STDDestinationPortals = From File In (Directory.GetFiles(Source, "*.*", SearchOption.AllDirectories))
                                    Where FileTypes.Contains((Path.GetExtension(File)).ToLower()) _
                                       And Not (Path.GetFileNameWithoutExtension(File)).EndsWith("_VH") _
                                       And Not (Path.GetFileNameWithoutExtension(File)).EndsWith("_HI") _
                                       And Not (Path.GetFileNameWithoutExtension(File)).EndsWith("_LO") _
                                       And Not (Path.GetFileNameWithoutExtension(File)).EndsWith("_VL") _
                                       And ((Path.GetFileName(File)).StartsWith("Door ") _
                                       Or (Path.GetFileName(File)).StartsWith("Doors ") _
                                       Or (Path.GetFileName(File)).StartsWith("Window ") _
                                       Or (Path.GetFileName(File)).StartsWith("Windows "))

        'Very-high quality is the default quality we're after, so any such portals we've found
        'will be added, unconditionally, to our final collection of portals to be copied.
        Message = Indent & "Adding all very-high-quality portals to the copy list." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        For Each File As String In VHDestinationPortals
            DestinationPortals.Add(File)
        Next

        'We want to grab any high-quality portals that do not have a very-high-quality counterpart.
        'Before adding high-quality portals to our final collection of portals to be copied, we will
        'compare the stripped-down name against the names we collected in the VHNames list.
        Message = Indent & "Adding all high-quality portals to the copy list that are not duplicates." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        For Each file As String In HIDestinationPortals
            BaseName = (Path.GetFileNameWithoutExtension(file)).Replace("_HI", "")
            If Not VHNames.Contains(BaseName) Then
                DestinationPortals.Add(file)
            End If
        Next

        'We want to grab any standard-quality portals that have neither a very-high-quality counterpart
        'nor a high-quality counterpart.
        'Before adding standard-quality portals to our final collection of portals to be copied, we will
        'compare the stripped-down name against the names we collected in the VHNames list as well as
        'the names we collected in the HINames list.
        Message = Indent & "Adding all standard-quality portals to the copy list that are not duplicates." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        For Each file As String In STDDestinationPortals
            BaseName = Path.GetFileNameWithoutExtension(file)
            If Not VHNames.Contains(BaseName) And Not HINames.Contains(BaseName) Then
                DestinationPortals.Add(file)
            End If
        Next

        If DestinationObjects.Count >= 1 Then
            'Get the list of folders that contain objects.
            Message = Indent & "Determining source and destination folders for objects." & vbCrLf
            OutputForm.OutputTextBox.AppendText(Message)
            If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
            For Each file As String In DestinationObjects
                Dim folder As String = (My.Computer.FileSystem.GetFileInfo(file)).DirectoryName
                If Not ObjectFolders.Contains(folder) Then
                    ObjectFolders.Add(folder)
                End If
            Next

            For Each folder As String In ObjectFolders
                'Ultimately, we want the objects copied to a "\textures\objects" subfolder.
                'But we want to determine if that path already exists before we add it again.
                If folder.Contains(TexturePath) Then
                    ReplaceSource = Destination
                Else
                    ReplaceSource = Destination & ObjectPath
                End If
                FolderDestination = folder.Replace(Source, ReplaceSource)

                'If the destination folder does not exist, create it.
                If Not My.Computer.FileSystem.DirectoryExists(FolderDestination) Then
                    Message = Indent & "Creating folder: " & FolderDestination & vbCrLf
                    OutputForm.OutputTextBox.AppendText(Message)
                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                    My.Computer.FileSystem.CreateDirectory(FolderDestination)
                Else
                    Message = Indent & "Folder exists: " & FolderDestination & vbCrLf
                    OutputForm.OutputTextBox.AppendText(Message)
                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                End If
            Next

            'Copy each object file.
            For Each file As String In DestinationObjects
                CopySource = My.Computer.FileSystem.GetFileInfo(file).FullName
                CopyDestination = CopySource.Replace(Source, ReplaceSource)

                If My.Computer.FileSystem.FileExists(CopyDestination) Then
                    Message = Indent & "Destination file already exists: " & CopyDestination & vbCrLf
                    OutputForm.OutputTextBox.AppendText(Message)
                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                Else
                    Message = Indent & "Copying from: " & CopySource & vbCrLf
                    Message = Indent & "          to: " & CopyDestination & vbCrLf
                    OutputForm.OutputTextBox.AppendText(Message)
                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                    My.Computer.FileSystem.CopyFile(CopySource, CopyDestination)
                End If
            Next
        End If

        If DestinationPortals.Count >= 1 Then
            'Get the list of folders that contain portals.
            Message = Indent & "Determining source and destination folders for portals." & vbCrLf
            OutputForm.OutputTextBox.AppendText(Message)
            If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
            For Each file As String In DestinationPortals
                Dim folder As String = (My.Computer.FileSystem.GetFileInfo(file)).DirectoryName
                If Not PortalFolders.Contains(folder) Then
                    PortalFolders.Add(folder)
                End If
            Next

            For Each folder In PortalFolders
                'Ultimately, we want the portals copied to a "\textures\objects" subfolder,
                'or to a "\textures\portals" subfolder, depending on whether or not
                '"Separate Portals" is checked.
                If Portals Then
                    'If the current option is to separate portals, then...
                    If folder.Contains(PortalPath) Then
                        '...if the portals are in the "textures\portals" folder,
                        'Set the source folder as the user-specified source folder.
                        'Set the destination folder as the user-specified destination folder.
                        PortalSource = Source
                        ReplaceSource = Destination
                    ElseIf folder.Contains(ObjectPath) Then
                        '...if the portals the are in the "textures\objects" folder,
                        'Append "\textures\objects" to the user-specified source folder.
                        'Append "\textures\portals" to the user-specified destination folder. 
                        PortalSource = Source & ObjectPath
                        ReplaceSource = Destination & PortalPath
                    Else
                        '...if the portals are in neither folder,
                        'Set the source folder as the user-specified source folder.
                        'Append "\textures\portals" to the user-specified destination folder.
                        PortalSource = Source
                        ReplaceSource = Destination & PortalPath
                    End If
                ElseIf Not Portals Then
                    'If the current option is to not separate portals, then...
                    If folder.Contains(PortalPath) Then
                        '...if the portals are in the "textures\portals" folder,
                        'Append "\textures\portals" to the user-specified source folder.
                        'Append "\textures\objects" to the user-specified destination folder.
                        PortalSource = Source & PortalPath
                        ReplaceSource = Destination & ObjectPath
                    ElseIf folder.Contains(ObjectPath) Then
                        '...if the portals the are in the "textures\objects" folder,
                        'Set the source folder as the user-specified source folder.
                        'Set the destination folder as that user-specified destination folder.
                        PortalSource = Source
                        ReplaceSource = Destination
                    Else
                        '...if the portals are in neither folder,
                        'Set the source folder as the user-specified source folder.
                        'Append "\textures\objects" to the user-specified destination folder.
                        PortalSource = Source
                        ReplaceSource = Destination & ObjectPath
                    End If
                End If
                FolderDestination = folder.Replace(PortalSource, ReplaceSource)

                'If the destination folder does not exist, create it.
                If Not My.Computer.FileSystem.DirectoryExists(FolderDestination) Then
                    Message = Indent & "Creating folder: " & FolderDestination & vbCrLf
                    OutputForm.OutputTextBox.AppendText(Message)
                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                    My.Computer.FileSystem.CreateDirectory(FolderDestination)
                Else
                    Message = Indent & "Folder exists: " & FolderDestination & vbCrLf
                    OutputForm.OutputTextBox.AppendText(Message)
                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                End If
            Next

            'Copy each portal file.
            For Each file As String In DestinationPortals
                CopySource = My.Computer.FileSystem.GetFileInfo(file).FullName
                CopyDestination = CopySource.Replace(PortalSource, ReplaceSource)
                If My.Computer.FileSystem.FileExists(CopyDestination) Then
                    Message = Indent & "Destination file already exists: " & CopyDestination & vbCrLf
                    OutputForm.OutputTextBox.AppendText(Message)
                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                Else
                    Message = Indent & "Copying from: " & CopySource & vbCrLf
                    Message &= Indent & "          to: " & CopyDestination & vbCrLf
                    OutputForm.OutputTextBox.AppendText(Message)
                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                    My.Computer.FileSystem.CopyFile(CopySource, CopyDestination)
                End If
            Next
        End If

        If CreateTagFile And My.Computer.FileSystem.DirectoryExists(Destination) Then
            Dim AssetFolder As String = My.Computer.FileSystem.GetDirectoryInfo(Destination).Name
            TagAssetsSub(Destination, AssetFolder, DefaultTag, CreateLog, LogFileName, Indent)
        End If
    End Sub
End Module
