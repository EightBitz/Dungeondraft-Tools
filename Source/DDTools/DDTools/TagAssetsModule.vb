Imports System.IO
Imports Newtonsoft.Json
Module TagAssetsModule
    Public Sub TagAssetsSub(Source As String, AssetFolder As String, DefaultTag As String, CreateLog As Boolean, LogFileName As String, Indent As String)
        Dim TagsFilePath As String = Source & "\data"
        Dim TagsFile As String = TagsFilePath & "\default.dungeondraft_tags"
        Dim TexturePath As String
        Dim ObjectPath As String
        Dim ColorablePath As String
        Dim Message As String

        'Get the exact, case-sensitive path name for the textures folder.
        Dim TexturePathArray = From subfolder In My.Computer.FileSystem.GetDirectories(Source)
                               Where (My.Computer.FileSystem.GetDirectoryInfo(subfolder)).Name.ToLower() = "textures"
        For Each path As String In TexturePathArray
            TexturePath = path
        Next

        'Get the exact, case-sensitive path name for the textures\objects folder.
        Dim ObjectPathArray = From subfolder In My.Computer.FileSystem.GetDirectories(TexturePath)
                              Where (My.Computer.FileSystem.GetDirectoryInfo(subfolder)).Name.ToLower() = "objects"
        For Each path As String In ObjectPathArray
            ObjectPath = path
        Next

        'Get the exact, case-sensitive path name for the textures\objects\colorable folder.
        Dim ColorablePathArray = From subfolder In My.Computer.FileSystem.GetDirectories(ObjectPath.ToString())
                                 Where (My.Computer.FileSystem.GetDirectoryInfo(subfolder)).Name.ToLower() = "colorable"
        For Each path As String In ColorablePathArray
            ColorablePath = path
        Next

        'Set the values for the main keys that will appear in the JSON file.
        Dim tags As String = "tags"
        Dim colorable As String = "Colorable"
        Dim sets As String = "sets"

        'Set up ordered dictionaries that we can later serialize for later covnersion to JSON.
        Dim TagObject As New System.Collections.Specialized.OrderedDictionary
        Dim FolderObject As New System.Collections.Specialized.OrderedDictionary
        Dim ColorableObject As New System.Collections.Specialized.OrderedDictionary
        Dim SetObject As New System.Collections.Specialized.OrderedDictionary
        Dim TagSetMembers() As String

        'Get the list of subfolders in textures\objects.
        Message = Indent & "Getting subfolder names to use as tags." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        Dim SubFolders As String() = Directory.GetDirectories(ObjectPath)

        'For each subfolder, truncate path from the string value so we're just left with the folder name.
        For i = 0 To SubFolders.Count - 1
            SubFolders(i) = SubFolders(i).Replace(ObjectPath & "\", "")
        Next

        'Get the list files from the root of textures\objects.
        Message = Indent & "Getting root objects." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        Dim RootFiles As String() = Directory.GetFiles(ObjectPath)

        'For each file, truncate the path up to "textures" so we're left with "textures\objects\<filename>".
        'Then replace the backslashes with forward slashes so we're left with "textures/objects/<filename>".
        For i = 0 To RootFiles.Count - 1
            RootFiles(i) = RootFiles(i).Replace(Source & "\", "")
            RootFiles(i) = RootFiles(i).Replace("\", "/")
        Next

        'If a textures\objects\colorable folder exists, get the list of files from textures\objects\colorable.
        Message = Indent & "Getting root colorable objects." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        Dim RootColorableFiles As String()
        If My.Computer.FileSystem.DirectoryExists(ColorablePath) Then
            RootColorableFiles = Directory.GetFiles(ColorablePath)

            'For each file, truncate the path up to "textures" so we're left with "textures\objects\colorable\<filename>".
            'Then replace the backslashes with forward slashes so we're left with "textures/objects/colorable/<filename>".
            For i = 0 To RootColorableFiles.Count - 1
                RootColorableFiles(i) = RootColorableFiles(i).Replace(Source & "\", "")
                RootColorableFiles(i) = RootColorableFiles(i).Replace("\", "/")
            Next
        End If

        Dim RootCount As Integer
        Dim RootObjects() As String
        RootCount = 0
        If RootFiles IsNot Nothing And RootColorableFiles IsNot Nothing Then
            'If there were files in both the root of textures\objects, and in textures\objects\colorable,
            'then combine both lists of files into the RootObjects array.
            RootCount = RootFiles.Count + RootColorableFiles.Count - 1
            ReDim RootObjects(RootCount)
            RootFiles.CopyTo(RootObjects, 0)
            RootColorableFiles.CopyTo(RootObjects, RootFiles.Count)
        ElseIf RootFiles IsNot Nothing Then
            'If there were files in the root of textures\objects, then copy that list into the RootObjects array.
            RootCount = RootFiles.Count
            ReDim RootObjects(RootCount - 1)
            RootFiles.CopyTo(RootObjects, 0)
        ElseIf RootColorableFiles IsNot Nothing Then
            'If there were files in textures\objects\colorable, then copy that list into the RootObjects array.
            RootCount = RootColorableFiles.Count
            ReDim RootObjects(RootCount)
            RootColorableFiles.CopyTo(RootObjects, 0)
        End If

        'Get a recursive list of all files in all subfolders under textures\objects.
        Message = Indent & "Getting all objects from subfolders." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        Dim AllObjects As String() = Directory.GetFiles(ObjectPath, "*.*", SearchOption.AllDirectories)

        'For each file, truncate the path up to "textures" so we're left with "textures\objects\<path>\<filename>".
        'Then replace the backslashes with forward slashes so we're left with "textures/objects/<path>/<filename>".
        For i = 0 To AllObjects.Count - 1
            AllObjects(i) = AllObjects(i).Replace(Source & "\", "")
            AllObjects(i) = AllObjects(i).Replace("\", "/")
        Next

        Message = Indent & "Adding the default tag, if there is one, and adding subfolders to the tag set." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        If DefaultTag = "" Or RootObjects.Count = 0 Then
            'If the default tag value is empty, or if there are no files in the root of textures\objects,
            'then add the list of subfolders to the SetObject ordered dictionary.

            SetObject.Add(AssetFolder, SubFolders)
        ElseIf DefaultTag <> "" And RootObjects.Count >= 1 And Not SubFolders.Contains(DefaultTag) Then
            'If the default tag value is not empty, and if there are files in the root of textures\objects,
            'and if the default tag value does not conflict with the name of a subfolder,
            'then add the default tag to the FolderObject ordered dictionary with its
            'associated files from textures\objects and textures\objects\colorable.

            FolderObject.Add(DefaultTag, RootObjects)

            'Then copy the default tag into the first element of the TagSetMembers array,
            'then append the list of subfolders to to the TagSetMembers array,
            'then add the TagSetMembers array to the SetObject ordered dictionary.

            ReDim TagSetMembers(SubFolders.Count)
            TagSetMembers(0) = DefaultTag
            SubFolders.CopyTo(TagSetMembers, 1)
            SetObject.Add(AssetFolder, TagSetMembers)
        ElseIf DefaultTag <> "" And RootObjects.Count >= 1 And SubFolders.Contains(DefaultTag) Then
            'If the default tag value is not empty, and if there are files in the root of textures\objects,
            'and if the default tag value does conflict with the name of a subfolder,
            'then add the list of subfolders to the SetObject ordered dictionary.
            'We will resolve the conflict later by combining both the files in the root of textures objects,
            'and the files from the identically named subfolder into the same array.

            Message = Indent & "Default Tag value conflicts with a folder name. Will attempt to merge assets." & vbCrLf
            OutputForm.OutputTextBox.AppendText(Message)
            If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
            SetObject.Add(AssetFolder, SubFolders)
        End If

        'For each subfolder, add the folder and its associated files to the FolderObject ordered dictionary.
        'Each subfolder name is what will be considered a Tag, and each file contained within that subfolder 
        'will be associated with its parent folder name as a tag.
        Dim SetCount As Integer
        Dim subset As String()
        Message = Indent & "Adding the tags for the subfolders." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        For Each folder As String In SubFolders
            If folder.ToLower() <> "colorable" Then
                'We don't want to tag any files in textures\objects\colorable.
                'Those files, if they exist, have already been tagged with the default tag.
                Dim folderset As String() = Array.FindAll(AllObjects, Function(s) s.Contains("/" & folder & "/"))
                If DefaultTag = folder And RootObjects.Count >= 1 Then
                    'If the default tag value conflicts with this folder name,
                    'and if there were files in the root of textures\objects or in textures\objects\colorable,
                    'then combine those files into the same array as the files from this folder.

                    SetCount = folderset.Count + RootObjects.Count - 1
                    ReDim subset(SetCount)
                    RootObjects.CopyTo(subset, 0)
                    folderset.CopyTo(subset, RootObjects.Count)

                    'Then add the combined list of files to the FolderObject ordered dictionary,
                    'using the folder name as the tag.

                    FolderObject.Add(folder, subset)
                Else
                    'If there is no conflict between the default tag value and this folder name,
                    'then add the list of files from this subfolder to the FolderObject ordered dictionary,
                    'using the folder name as the tag.

                    SetCount = folderset.Count
                    ReDim subset(SetCount - 1)
                    folderset.CopyTo(subset, 0)
                    FolderObject.Add(folder, subset)
                End If
            End If
        Next

        Message = Indent & "Adding the tag for all colorable objects from all subfolders." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        'Get a recursive list of all files in all "colorable" folders.
        'If any are found, add them to the FolderObject ordered dictionary under a "Colorable" tag.
        Dim AllColorableObjects As String() = Array.FindAll(AllObjects, Function(s) s.ToLower().Contains("/colorable/"))
        If AllColorableObjects.Count >= 1 Then
            FolderObject.Add(colorable, AllColorableObjects)
        End If

        Message = Indent & "Creating the tag object." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        'Add the FolderObject ordered dictionary to the TagObject ordered dictionary.
        TagObject.Add(tags, FolderObject)

        'Add the SetObject ordered dictionary to the TagObject ordered dictionary.
        TagObject.Add(sets, SetObject)

        Message = Indent & "Converting the tag object to JSON." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        'Convert the TagObject ordered dictionary to a JSON-formatted string.
        Dim JSONString As String = JsonConvert.SerializeObject(TagObject, Formatting.Indented)

        Message = Indent & "Looking for the data folder, and creating it if it doesn't exist." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        'Check for a data folder, and if it doesn't exist, create it.
        If Not My.Computer.FileSystem.DirectoryExists(TagsFilePath) Then
            My.Computer.FileSystem.CreateDirectory(TagsFilePath)
        End If

        Message = Indent & "Writing the tag file to " & TagsFile & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
        'Write the JSON string as a default.dungeondraft_tags file to the data folder.
        My.Computer.FileSystem.WriteAllText(TagsFile, JSONString, False)
    End Sub
End Module
