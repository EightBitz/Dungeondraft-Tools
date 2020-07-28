Imports System.IO
Imports Newtonsoft.Json
Module PackAssetsModule
    Public Sub LoadAssetFolders(Source As String, CreateLog As Boolean, LogFileName As String, SelectFolder As Boolean, Indent As String)

        'Initialize the columns for DDToolsForm.PackAssetsDataGridView
        Dim SelectColumn As New DataGridViewCheckBoxColumn With {
            .HeaderText = "Select Folder",
            .Name = "SelectFolder"
        }
        DDToolsForm.PackAssetsDataGridView.Columns.Add(SelectColumn)

        Dim FolderColumn As New DataGridViewTextBoxColumn With {
            .HeaderText = "Folder Name",
            .Name = "FolderName",
            .Width = 199,
            .ReadOnly = True
        }
        DDToolsForm.PackAssetsDataGridView.Columns.Add(FolderColumn)

        Dim NameColumn As New DataGridViewTextBoxColumn With {
            .HeaderText = "Pack Name",
            .Name = "PackName",
            .Width = 199
        }
        DDToolsForm.PackAssetsDataGridView.Columns.Add(NameColumn)

        Dim IDColumn As New DataGridViewTextBoxColumn With {
            .HeaderText = "Pack ID",
            .Name = "PackID",
            .ReadOnly = True
        }
        DDToolsForm.PackAssetsDataGridView.Columns.Add(IDColumn)

        Dim VersionColumn As New DataGridViewTextBoxColumn With {
            .HeaderText = "Version",
            .Name = "Version",
            .Width = 70
        }
        DDToolsForm.PackAssetsDataGridView.Columns.Add(VersionColumn)

        Dim AuthorColumn As New DataGridViewTextBoxColumn With {
            .HeaderText = "Author",
            .Name = "Author",
            .Width = 130
        }
        DDToolsForm.PackAssetsDataGridView.Columns.Add(AuthorColumn)

        'Get the list of folders in the Source folder, then iterate through them.
        Dim Folders = My.Computer.FileSystem.GetDirectories(Source)
        Dim PackInfo
        Dim RowIndex As Integer = 0
        For Each AssetFolder As String In Folders
            Dim FolderName As String = (My.Computer.FileSystem.GetDirectoryInfo(AssetFolder)).Name
            'If the current folder has a "textures" subfolder, then assume it's a pack folder.
            'Get the pack.json info if it's there, but whether it is or not, add the folder to the DataGridView.
            'If there is no pack.json file, then set the folder name as the pack name, and set the version as "1".
            If My.Computer.FileSystem.DirectoryExists(AssetFolder & "\textures") Then
                If My.Computer.FileSystem.FileExists(AssetFolder & "\pack.json") Then
                    Dim rawJson = File.ReadAllText(AssetFolder & "\pack.json")
                    PackInfo = JsonConvert.DeserializeObject(rawJson)
                    Dim NewRow As Object() = New Object() {SelectFolder, FolderName, PackInfo("name"), PackInfo("id"), PackInfo("version"), PackInfo("author")}
                    DDToolsForm.PackAssetsDataGridView.Rows.Add(NewRow)
                Else
                    Dim NewRow As Object() = New Object() {SelectFolder, FolderName, FolderName, "", "1", ""}
                    DDToolsForm.PackAssetsDataGridView.Rows.Add(NewRow)
                End If
                RowIndex += 1
            End If
        Next
    End Sub
    Public Sub RunPackEXE(PackEXE As String, PackArguments As String)
        'Call dungeondraft-pack.exe with the specified arguments, and wait for it to exit.
        Dim PackStartInfo As New ProcessStartInfo With {
                .FileName = PackEXE,
                .Arguments = PackArguments,
                .RedirectStandardError = True,
                .RedirectStandardOutput = True,
                .UseShellExecute = False,
                .CreateNoWindow = True
            }
        Dim Pack As Process = Process.Start(PackStartInfo)
        Pack.WaitForExit()
    End Sub
    Public Sub PackAssetsSub(Source As String, Destination As String, RowIndex As Integer, CreateLog As Boolean, LogFileName As String, Indent As String, PackEXE As String, Overwrite As Boolean)

        'Initialize some variables.
        Dim PackJSON As String = Source & "\pack.json"
        Dim JSONExists As Boolean = False
        Dim PackInfo
        Dim NewPackInfo

        Dim OriginalPackName As String
        Dim OriginalPackID As String
        Dim OriginalPackVersion As String
        Dim OriginalPackAuthor As String

        Dim NewPackName As String
        Dim NewPackID As String
        Dim NewPackVersion As String
        Dim NewPackAuthor As String

        Dim ArgSource As String
        Dim ArgDestination As String
        Dim ArgPackName As String
        Dim ArgPackID As String
        Dim ArgPackVersion As String
        Dim ArgPackAuthor As String

        Dim BaseArguments As String
        Dim PackArguments As String

        'Get the current, user-editable values for creating the pack.
        NewPackName = DDToolsForm.PackAssetsDataGridView.Rows(RowIndex).Cells(2).Value
        NewPackID = DDToolsForm.PackAssetsDataGridView.Rows(RowIndex).Cells(3).Value
        NewPackVersion = DDToolsForm.PackAssetsDataGridView.Rows(RowIndex).Cells(4).Value
        NewPackAuthor = DDToolsForm.PackAssetsDataGridView.Rows(RowIndex).Cells(5).Value

        'Define a set of variables enclosed in quotes to use for command line arguments.
        ArgSource = """" & Source & """"
        ArgDestination = """" & Destination.TrimEnd("\") & """"
        ArgPackName = """" & NewPackName & """"
        ArgPackID = """" & NewPackID & """"
        ArgPackVersion = """" & NewPackVersion & """"
        ArgPackAuthor = """" & NewPackAuthor & """"

        'Set up the minimum number of arguments we need to package a folder.
        BaseArguments = ArgSource & " " & ArgDestination

        'If the user selected to overwrite existing files, then add that to the minimum arguments.
        If Overwrite Then BaseArguments = "-overwrite " & BaseArguments

        If My.Computer.FileSystem.FileExists(PackJSON) Then
            'If there's an existing pack.json file, we want to proceed differently than if there isn't.
            JSONExists = True

            'Load the pack.json file into PackInfo
            Dim rawJson = File.ReadAllText(PackJSON)
            PackInfo = JsonConvert.DeserializeObject(rawJson)

            'Store the values from the pack.json file so we can later determine if the user has edited values.
            OriginalPackName = PackInfo("name")
            OriginalPackID = PackInfo("id")
            OriginalPackVersion = PackInfo("version")
            OriginalPackAuthor = PackInfo("author")

            If NewPackName = OriginalPackName And NewPackVersion = OriginalPackVersion And NewPackAuthor = OriginalPackAuthor Then
                'If no values have been edited, then create the pack with the minimum arguments.
                RunPackEXE(PackEXE, BaseArguments)
            ElseIf NewPackName = OriginalPackName And (NewPackVersion <> OriginalPackVersion Or NewPackAuthor <> OriginalPackAuthor) Then
                'If other values besides the name have been edited, then we want to add arguments to update values in the existing pack.json file.
                'This will generate a new pack ID, but since the name hasn't changed, we want to preserve the original pack ID.
                'So we're not going to create the pack yet. We just want to update the pack.json file.
                PackArguments = "-name " & ArgPackName & " -version " & ArgPackVersion & " -author " & ArgPackAuthor & " -editpack -genpack " & BaseArguments
                RunPackEXE(PackEXE, PackArguments)

                'Read in the updated pack.json file, restore the original ID, then save the file.
                Dim newJson = File.ReadAllText(PackJSON)
                NewPackInfo = JsonConvert.DeserializeObject(newJson)
                NewPackInfo("id") = OriginalPackID
                Dim NewJSONString As String = JsonConvert.SerializeObject(NewPackInfo, Formatting.Indented)
                My.Computer.FileSystem.WriteAllText(PackJSON, NewJSONString, False, System.Text.Encoding.ASCII)

                'Create the pack with the the updated info, but the same name and ID.
                RunPackEXE(PackEXE, BaseArguments)
            ElseIf NewPackName <> OriginalPackName Then
                'If the name has been edited, regardless of whether or not other values have been edited, we will treat this as the creation of a new pack.
                PackArguments = "-name " & ArgPackName & " -version " & ArgPackVersion & " -author " & ArgPackAuthor & " -editpack " & BaseArguments
                RunPackEXE(PackEXE, PackArguments)
            End If
        Else
            'If there is no existing pack.json file, we will treat this as the creation of a new pack.
            PackArguments = "-name " & ArgPackName & " -version " & ArgPackVersion & " -author " & ArgPackAuthor & " " & BaseArguments
            RunPackEXE(PackEXE, PackArguments)
        End If
    End Sub
End Module
