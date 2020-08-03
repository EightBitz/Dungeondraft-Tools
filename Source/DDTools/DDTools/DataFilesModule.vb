Imports System.IO
Imports Newtonsoft.Json

Public Class AssetObject
    Public Property AssetType As String
    Public Property FilePath As String
    Public Property Path As String
    Public Property Name As String
    Public Property Type As String
    Public Property Color As String
    Public Property DataFolderName As String
    Public Property DataFileName As String
End Class

Module DataFilesModule
    Public Function GetFileInfo(Source As String)
        Dim TileRoot As String = Source & "\textures\tilesets"
        Dim AssetList As New List(Of AssetObject)
        Dim Subfolders() As String = {TileRoot & "\simple", TileRoot & "\smart", TileRoot & "\smart_double", Source & "\textures\walls"}
        For Each AssetFolder In Subfolders
            If Directory.Exists(AssetFolder) Then
                Dim AssetFileList = Directory.GetFiles(AssetFolder)
                If AssetFileList.Count >= 1 Then
                    For Each AssetFile In AssetFileList
                        If Not Path.GetFileNameWithoutExtension(AssetFile).EndsWith("_end") Then
                            Dim Asset As New AssetObject

                            If AssetFolder.EndsWith("\simple") Then
                                Asset.AssetType = "Tile: Simple"
                                Asset.DataFileName = Path.GetFileNameWithoutExtension(AssetFile) & ".dungeondraft_tileset"
                            ElseIf AssetFolder.EndsWith("\smart") Then
                                Asset.AssetType = "Tile: Smart"
                                Asset.DataFileName = Path.GetFileNameWithoutExtension(AssetFile) & ".dungeondraft_tileset"
                            ElseIf AssetFolder.EndsWith("\smart_double") Then
                                Asset.AssetType = "Tile: Smart_Double"
                                Asset.DataFileName = Path.GetFileNameWithoutExtension(AssetFile) & ".dungeondraft_tileset"
                            ElseIf AssetFolder.EndsWith("\walls") Then
                                Asset.AssetType = "Wall"
                                Asset.DataFileName = Path.GetFileNameWithoutExtension(AssetFile) & ".dungeondraft_wall"
                            End If

                            Asset.FilePath = AssetFile
                            Asset.Name = Path.GetFileNameWithoutExtension(AssetFile)
                            If Asset.AssetType = "Wall" Then
                                Asset.Type = ""
                                Asset.DataFolderName = Source & "\data\walls"
                            Else
                                Asset.Name = Path.GetFileNameWithoutExtension(AssetFile)
                                Asset.Type = "normal"
                                Asset.DataFolderName = Source & "\data\tilesets"
                            End If

                            Asset.Path = Microsoft.VisualBasic.Strings.Replace(AssetFile, Source & "\", "", 1, -1, Constants.vbTextCompare).Replace("\", "/")
                            Asset.Color = "ffffff"

                            AssetList.Add(Asset)
                        End If
                    Next
                End If
            End If
        Next
        Return AssetList
    End Function

    Public Function GetDataInfo(Source As String, AssetList As List(Of AssetObject))
        Dim DataFolderList() As String = {Source & "\data\tilesets", Source & "\data\walls"}
        Dim DataFileList
        Dim rawJson
        Dim JsonAsset

        For Each DataFolder In DataFolderList
            If Directory.Exists(DataFolder) Then
                DataFileList = Directory.GetFiles(DataFolder)
                For Each DataFile In DataFileList
                    rawJson = File.ReadAllText(DataFile)
                    JsonAsset = JsonConvert.DeserializeObject(rawJson)
                    For Each Asset In AssetList
                        If Asset.Path = JsonAsset("path") Then
                            Asset.Color = JsonAsset("color")
                            If Asset.AssetType <> "Wall" Then
                                Asset.Name = JsonAsset("name")
                                Asset.Type = JsonAsset("type")
                            End If
                            Asset.DataFolderName = DataFolder
                            Asset.DataFileName = Path.GetFileName(DataFile)
                        End If
                    Next
                Next
            End If
        Next
        Return AssetList
    End Function

    Public Sub GetTilesAndWalls(Source As String)
        'Center our column headers in the data grid view.
        DDToolsForm.DataFilesDataGridView.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        'Configure our columns for the data grid view.
        Dim AssetTypeColumn As New DataGridViewTextBoxColumn With {
            .HeaderText = "Asset Type",
            .Name = "AssetType",
            .Width = 130,
            .ReadOnly = True
        }
        DDToolsForm.DataFilesDataGridView.Columns.Add(AssetTypeColumn)

        Dim NameColumn As New DataGridViewTextBoxColumn With {
            .HeaderText = "Name",
            .Name = "Name",
            .Width = 182
        }
        DDToolsForm.DataFilesDataGridView.Columns.Add(NameColumn)

        Dim ColorTypeColumn As New DataGridViewComboBoxColumn With {
            .HeaderText = "Color Type",
            .Name = "ColorType",
            .Width = 110,
            .DataSource = {"", "custom_color", "normal"}
        }
        DDToolsForm.DataFilesDataGridView.Columns.Add(ColorTypeColumn)

        Dim ColorColumn As New DataGridViewTextBoxColumn With {
            .HeaderText = "Color",
            .Name = "Color"
        }
        DDToolsForm.DataFilesDataGridView.Columns.Add(ColorColumn)
        DDToolsForm.DataFilesDataGridView.Columns(3).DefaultCellStyle.Font = New Font("Courier New", 10, FontStyle.Regular)

        Dim DataFileColumn As New DataGridViewTextBoxColumn With {
            .HeaderText = "Data File",
            .Name = "DataFile",
            .Width = 379
        }
        DDToolsForm.DataFilesDataGridView.Columns.Add(DataFileColumn)

        Dim FolderColumn As New DataGridViewTextBoxColumn With {
            .HeaderText = "Folder",
            .Name = "Folder",
            .Visible = False,
            .ReadOnly = True
        }
        DDToolsForm.DataFilesDataGridView.Columns.Add(FolderColumn)

        Dim PathColumn As New DataGridViewTextBoxColumn With {
            .HeaderText = "Path",
            .Name = "Path",
            .Visible = False,
            .ReadOnly = True
        }
        DDToolsForm.DataFilesDataGridView.Columns.Add(PathColumn)

        Dim AssetList As List(Of AssetObject)
        AssetList = GetFileInfo(Source)
        AssetList = GetDataInfo(Source, AssetList)

        Dim RowIndex As Integer = 0
        For Each Asset In AssetList
            Dim NewRow As Object() = New Object() {Asset.AssetType, Asset.Name, Asset.Type, Asset.Color, Asset.DataFileName, Asset.DataFolderName, Asset.Path}
            DDToolsForm.DataFilesDataGridView.Rows.Add(NewRow)
            If Asset.AssetType = "Wall" Then
                DDToolsForm.DataFilesDataGridView.Rows(RowIndex).Cells(1).ReadOnly = True
                DDToolsForm.DataFilesDataGridView.Rows(RowIndex).Cells(2).ReadOnly = True
            End If
            RowIndex += 1
        Next

    End Sub
End Module
