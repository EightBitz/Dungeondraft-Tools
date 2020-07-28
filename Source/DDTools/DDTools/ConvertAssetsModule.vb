Imports System.IO
Module ConvertAssetsModule
    Public Sub ConvertAssetsSub(Source As String, Destination As String, CreateLog As Boolean, LogFileName As String, Indent As String, ImageMagickEXE As String)
        Dim Message As String
        Dim FileTypes() As String = {".bmp", ".dds", ".exr", ".hdr", ".jpg", ".jpeg", ".png", ".tga", ".svg", ".svgz", ".webp"}
        Dim FolderList As String() = Directory.GetDirectories(Source, "*.*", SearchOption.AllDirectories)
        Dim FileList = Directory.GetFiles(Source, "*.*", SearchOption.AllDirectories)
        Dim FileContent As String
        Dim ConvertFiles = Nothing
        Dim CopyFiles = Nothing
        Dim OtherFiles = Nothing
        Dim DestinationFile As String

        'Check if the parent destination folder exists, and if it doesn't, create it.
        If My.Computer.FileSystem.DirectoryExists(Destination) Then
            Message = Indent & "Destination folder exists: " & Destination & vbCrLf
            OutputForm.OutputTextBox.AppendText(Message)
            If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        Else
            Message = Indent & "Creating destination folder: " & Destination & vbCrLf
            OutputForm.OutputTextBox.AppendText(Message)
            If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
            My.Computer.FileSystem.CreateDirectory(Destination)
        End If

        'Check if all the subfolders exist, and if they don't, create them.
        Message = Indent & "Replicating folder structure." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        Dim DestinationInfo = My.Computer.FileSystem.GetDirectoryInfo(Destination)

        For Each folder As String In FolderList
            Dim destinationfolder = folder.Replace(Source, Destination)
            If My.Computer.FileSystem.DirectoryExists(destinationfolder) Then
                Message = Indent & "Destination subfolder exists: " & destinationfolder & vbCrLf
                OutputForm.OutputTextBox.AppendText(Message)
                If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
            Else
                Message = Indent & "Creating destination subfolder: " & destinationfolder & vbCrLf
                OutputForm.OutputTextBox.AppendText(Message)
                If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                My.Computer.FileSystem.CreateDirectory(destinationfolder)
            End If
        Next
        Message = Indent & "Finished replicating folder structure." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)

        'Get the list of all folders that contain a "textures" subfolder.
        'This will determine if we have structured asset folders, and we will proceed differently based on that result.
        Dim TextureFolders = From Folder In Directory.GetDirectories(Source, "*.*", SearchOption.AllDirectories)
                             Where My.Computer.FileSystem.GetDirectoryInfo(Folder).FullName.Contains("\textures\")

        If TextureFolders.Count >= 1 Then
            Message = Indent & """textures"" folders were found. Assuming structured asset folders." & vbCrLf
            OutputForm.OutputTextBox.AppendText(Message)
            If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)

            'If we have structured asset folders, we only want to convert files in the textures\objects folders.
            'Converting other types of assets will cause Dungeondraft to crash.
            'We also want to avoid reconverting files that are already in .webp format.
            Message = Indent & "Creating list of files to convert." & vbCrLf
            OutputForm.OutputTextBox.AppendText(Message)
            If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
            ConvertFiles = From File In Directory.GetFiles(Source, "*.*", SearchOption.AllDirectories)
                           Where My.Computer.FileSystem.GetDirectoryInfo(File).FullName.Contains("\textures\objects\") _
                           And FileTypes.Contains(My.Computer.FileSystem.GetFileInfo(File).Extension.ToLower()) _
                                   And My.Computer.FileSystem.GetFileInfo(File).Extension.ToLower() <> ".webp"

            'If we have structured asset folders, we only want to copy files that are not in the textures\objects folders.
            'Also, if any files are already in .webp format, we just want to copy those as well.
            Message = Indent & "Creating list of files to copy without converting." & vbCrLf
            OutputForm.OutputTextBox.AppendText(Message)
            If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
            CopyFiles = From File In Directory.GetFiles(Source, "*.*", SearchOption.AllDirectories)
                        Where FileTypes.Contains(My.Computer.FileSystem.GetFileInfo(File).Extension.ToLower()) _
                        And (Not My.Computer.FileSystem.GetFileInfo(File).FullName.Contains("\textures\objects\") _
                                Or My.Computer.FileSystem.GetFileInfo(File).Extension.ToLower() = ".webp")

            'If we have structured asset folders, get the list of all files that are not supported image files and
            'are not "default.dungeondraft_tags" files. This list of files will be copied rather than converted.
            '(pack.json files, data files for tilesets and walls, etc.)
            OtherFiles = From File In Directory.GetFiles(Source, "*.*", SearchOption.AllDirectories)
                         Where Not FileTypes.Contains(My.Computer.FileSystem.GetFileInfo(File).Extension.ToLower()) _
                             And My.Computer.FileSystem.GetFileInfo(File).Name <> "default.dungeondraft_tags"

            'Get the list of all "default.dungeondraft_tags" files.
            'We will have to update the content of these files to change extensions of referenced files to .webp.
            Dim TagFiles = Directory.GetFiles(Source, "default.dungeondraft_tags", SearchOption.AllDirectories)

            'Update the content of each default.dungeondraft_tags file so that each referenced file has a .webp extension.
            Message = Indent & "Updating all instances of default.dungeondraft_tags to reference converted objects with .webp extensions." & vbCrLf
            OutputForm.OutputTextBox.AppendText(Message)
            If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
            For Each file In TagFiles
                DestinationFile = (My.Computer.FileSystem.GetFileInfo(file).FullName).Replace(Source, Destination)
                FileContent = My.Computer.FileSystem.ReadAllText(file)

                For Each Extension In FileTypes
                    FileContent = Microsoft.VisualBasic.Strings.Replace(FileContent, Extension, ".webp", 1, -1, Constants.vbTextCompare)
                Next
                My.Computer.FileSystem.WriteAllText(DestinationFile, FileContent, False)
            Next
        Else
            Message = Indent & "No ""textures"" folder(s) found. Assuming conversion of all supported image files." & vbCrLf
            OutputForm.OutputTextBox.AppendText(Message)
            If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
            'If we do not have structured asset folders, simply convert all supported images.
            ConvertFiles = From File In Directory.GetFiles(Source, "*.*", SearchOption.AllDirectories)
                           Where FileTypes.Contains(My.Computer.FileSystem.GetFileInfo(File).Extension.ToLower())
        End If

        'Use ImageMagick to convert specified image files to .webp format.
        If Not ConvertFiles Is Nothing Then
            Message = Indent & "Converting image files for which conversion is supported." & vbCrLf
            OutputForm.OutputTextBox.AppendText(Message)
            If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
            For Each SourceFile As String In ConvertFiles
                Dim FileExt = My.Computer.FileSystem.GetFileInfo(SourceFile).Extension
                DestinationFile = SourceFile.Replace(FileExt, ".webp")
                DestinationFile = DestinationFile.Replace(Source, Destination)
                If My.Computer.FileSystem.FileExists(DestinationFile) Then
                    Message = Indent & "Destination file already exists: " & DestinationFile & vbCrLf
                    OutputForm.OutputTextBox.AppendText(Message)
                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                Else
                    Message = Indent & "Converting from: " & SourceFile & vbCrLf
                    OutputForm.OutputTextBox.AppendText(Message)
                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                    Message = Indent & "             to: " & DestinationFile & vbCrLf
                    OutputForm.OutputTextBox.AppendText(Message)
                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)

                    Dim MagickArguments As String = """" & SourceFile & """ """ & DestinationFile & """"
                    Dim MagickStartInfo As New ProcessStartInfo With {
                        .FileName = ImageMagickEXE,
                        .Arguments = MagickArguments,
                        .RedirectStandardError = True,
                        .RedirectStandardOutput = True,
                        .UseShellExecute = False,
                        .CreateNoWindow = True
                    }
                    Dim Magick As Process = Process.Start(MagickStartInfo)
                    Magick.WaitForExit()
                End If
            Next
            Message = vbCrLf
            OutputForm.OutputTextBox.AppendText(Message)
            If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        Else
            Message = Indent & "No images were found to convert." & vbCrLf & vbCrLf
            OutputForm.OutputTextBox.AppendText(Message)
            If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        End If

        'Copy specified files that are either already converted or are not to be converted.
        If Not CopyFiles Is Nothing Then
            Message = Indent & "Copying Image files for which conversion is not supported." & vbCrLf
            OutputForm.OutputTextBox.AppendText(Message)
            If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
            For Each SourceFile As String In CopyFiles
                DestinationFile = SourceFile.Replace(Source, Destination)

                If My.Computer.FileSystem.FileExists(DestinationFile) Then
                    Message = Indent & "Destination file already exists: " & DestinationFile & vbCrLf
                    OutputForm.OutputTextBox.AppendText(Message)
                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                Else
                    Message = Indent & "Copying from: " & SourceFile & vbCrLf
                    OutputForm.OutputTextBox.AppendText(Message)
                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                    Message = Indent & "          to: " & DestinationFile & vbCrLf
                    OutputForm.OutputTextBox.AppendText(Message)
                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                    My.Computer.FileSystem.CopyFile(SourceFile, DestinationFile)
                End If
            Next
        Else
            Message = Indent & "No images were found to copy." & vbCrLf & vbCrLf
            OutputForm.OutputTextBox.AppendText(Message)
            If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        End If

        'Copy any other relevant files (pack.json files, data files for tilesets and walls, etc.)
        If Not OtherFiles Is Nothing Then
            Message = Indent & "Copying data and metadata files." & vbCrLf
            OutputForm.OutputTextBox.AppendText(Message)
            If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
            For Each SourceFile As String In OtherFiles
                DestinationFile = SourceFile.Replace(Source, Destination)
                If My.Computer.FileSystem.FileExists(DestinationFile) Then
                    Message = Indent & "Destination file already exists: " & DestinationFile & vbCrLf
                    OutputForm.OutputTextBox.AppendText(Message)
                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                Else
                    Message = Indent & "Copying from: " & SourceFile & vbCrLf
                    OutputForm.OutputTextBox.AppendText(Message)
                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                    Message = Indent & "          to: " & DestinationFile & vbCrLf
                    OutputForm.OutputTextBox.AppendText(Message)
                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                    My.Computer.FileSystem.CopyFile(SourceFile, DestinationFile)
                End If
            Next
        Else
            Message = Indent & "No data or metadata files were found to copy." & vbCrLf & vbCrLf
            OutputForm.OutputTextBox.AppendText(Message)
            If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        End If
    End Sub
End Module
