Imports Newtonsoft.Json
Imports System.IO

Module MapDetailsModule
    Public Class TextureClass
        Public Property PackName As String
        Public Property PackVersion As String
        Public Property PackAuthor As String
        Public Property AssetName As String
        Public Property AssetType As String
        Public Property LevelName As String
        Public Property Instances As Integer
    End Class
    Public Function ParseTexture(Texture, AssetManifest, AssetType, LevelName)
        Dim PackID As String
        Dim PackName As String = ""
        Dim PackVersion As String = "<no version number>"
        Dim PackAuthor As String = "<no author name>"
        Dim AssetName As String

        'Find the pack ID in the texture string.
        Dim SplitTexture = Texture.Split("/")
        PackID = SplitTexture(3)
        AssetName = SplitTexture(SplitTexture.Length - 1)

        'Iterate through the asset manifest until we find a matching pack ID,
        'then get the pack name of the maatching ID.
        For Each AssetPack In AssetManifest
            If PackID = AssetPack("id") Then
                PackName = AssetPack("name")
                If AssetPack("version") <> "" Then PackVersion = AssetPack("version")
                If AssetPack("author") <> "" Then PackAuthor = AssetPack("author")
                Exit For
            End If
        Next

        'If we didn't find the pack ID in the asset manifest, then put that info in the PackName field.
        If PackName = "" Then PackName = "Pack ID " & PackID & " not found in Asset Manifest"

        'Create a UsedAsset record with all the info we found.
        Dim UsedAsset As New TextureClass With {
            .PackName = PackName,
            .PackVersion = PackVersion,
            .PackAuthor = PackAuthor,
            .AssetName = AssetName,
            .AssetType = AssetType,
            .LevelName = LevelName
        }

        'Return the UsedAsset record as the result of this function.
        Return UsedAsset
    End Function

    Public Sub MapDetailsSub(MapSource As String, MapFile As String, CreateLog As Boolean, LogFileName As String, Indent As String)
        'Initialize some variables
        Dim Message As String
        Dim rawJson = File.ReadAllText(MapSource)
        Dim MapInfo = JsonConvert.DeserializeObject(rawJson)
        Dim Header = MapInfo("header")
        Dim AssetManifest = Header("asset_manifest")
        Dim World = MapInfo("world")
        Dim TraceObject

        'Get some basic map info.
        Message = Indent & "Created with version: " & Header("creation_build") & vbCrLf
        Message &= Indent & "Created on (yyyy/m/d): " & Header("creation_date")("year") & "/" & Header("creation_date")("month") & "/" & Header("creation_date")("day") & vbCrLf
        Message &= Indent & "Grid size (W x H): " & World("width") & " x " & World("height") & vbCrLf

        'If the map includes a trace image, get that info.
        TraceObject = Header("editor_state")("trace_image")
        If TraceObject Is Nothing Then
            Message &= Indent & "Trace Image: none" & vbCrLf & vbCrLf
        ElseIf TraceObject.HasValues Then
            Message &= Indent & "Trace Image: " & Header("editor_state")("trace_image")("image") & vbCrLf & vbCrLf
        Else
            Message &= Indent & "Trace Image: none" & vbCrLf & vbCrLf
        End If

        'Get the list of custom asset packs that were selected when this map was saved.
        Message &= Indent & "Asset Manifest (Pack Name, Version #, Author Name):" & vbCrLf
        If AssetManifest.Count > 0 Then
            For Each AssetPack In AssetManifest
                Message &= Indent & Indent & AssetPack("name") & ", Version " & AssetPack("version") & ", " & AssetPack("author") & vbCrLf
            Next
        Else
            Message &= Indent & Indent & "No custom asset packs found in manifest." & vbCrLf
        End If
        Message &= vbCrLf

        'Initialize some more variables.
        Dim AssetTypeArray As String() = {"lights", "materials", "objects", "paths", "patterns", "portals", "terrain", "walls"}
        Dim TerrainArray As String() = {"texture_1", "texture_2", "texture_3", "texture_4"}
        Dim levelnumber As String
        Dim levelname As String
        Dim texture As String
        Dim layernumber As String
        Dim UsedAssetList As New List(Of TextureClass)
        Dim SearchIndex As Integer

        'Iterate through each level of the map.
        For Each Level In World("levels")
            'Get the number and name assigned to the current level.
            levelnumber = Level.Name.ToString
            levelname = World("levels")(levelnumber)("label")

            'Iterate through each asset type in the AssetTypeArray, and find all assets of that type.
            For Each AssetType In AssetTypeArray
                Select Case AssetType
                    Case "materials"
                        'Iterate through all assets where the asset type is "materials".
                        If Not World("levels")(levelnumber)(AssetType) Is Nothing Then
                            For Each layer In World("levels")(levelnumber)(AssetType)
                                'Materials are separated by layers (instead of having the layer stored as a property of the material),
                                'so we have to iterate through each layer under "materials" to find all materials used.
                                layernumber = layer.name.ToString
                                For Each asset In World("levels")(levelnumber)("materials")(layernumber)
                                    'Get the "texture" value for the current asset.
                                    texture = asset("texture")
                                    If texture.StartsWith("res://packs/") Then
                                        'If this is from a custom asset pack, then create a new record for it.
                                        Dim UsedAsset As New TextureClass

                                        'Call the "ParseTexture" function to get the info for the custom asset and
                                        'for the custom asset pack it's from.
                                        UsedAsset = ParseTexture(texture, AssetManifest, AssetType, levelname)

                                        'If this asset has already been added to our UsedAssetList, then increment the Instances property
                                        'of the existing record instead of adding the new one as a duplicate.
                                        'Otherwise, add the new record to the list And set the Instances property to 1.
                                        SearchIndex = UsedAssetList.FindIndex(Function(p) p.AssetName = UsedAsset.AssetName)
                                        If SearchIndex >= 0 Then
                                            UsedAssetList(SearchIndex).Instances += 1
                                        Else
                                            UsedAsset.Instances = 1
                                            UsedAssetList.Add(UsedAsset)
                                        End If

                                    End If
                                Next
                            Next
                        End If
                    Case "terrain"
                        'Iterate through all assets where the asset type is "terrain".
                        For Each asset In TerrainArray
                            If Not World("levels")(levelnumber)(AssetType)(asset) Is Nothing Then
                                'Get the "texture" value for the current asset.
                                texture = World("levels")(levelnumber)(AssetType)(asset)
                                If texture.StartsWith("res://packs/") Then
                                    'If this is from a custom asset pack, then create a new record for it.
                                    Dim UsedAsset As New TextureClass

                                    'Call the "ParseTexture" function to get the info for the custom asset and
                                    'for the custom asset pack it's from.
                                    UsedAsset = ParseTexture(texture, AssetManifest, AssetType, levelname)

                                    'If this asset has already been added to our UsedAssetList, then increment the Instances property
                                    'of the existing record instead of adding the new one as a duplicate.
                                    'Otherwise, add the new record to the list And set the Instances property to 1.
                                    SearchIndex = UsedAssetList.FindIndex(Function(p) p.AssetName = UsedAsset.AssetName)
                                    If SearchIndex >= 0 Then
                                        UsedAssetList(SearchIndex).Instances += 1
                                    Else
                                        UsedAsset.Instances = 1
                                        UsedAssetList.Add(UsedAsset)
                                    End If

                                End If
                            End If
                        Next
                    Case "walls"
                        'Iterate through all assets where the asset type is "walls".
                        If Not World("levels")(levelnumber)(AssetType) Is Nothing Then
                            For Each asset In World("levels")(levelnumber)(AssetType)
                                'Get the "texture" value for the current asset.
                                texture = asset("texture")
                                If texture.StartsWith("res://packs/") Then
                                    'If this is from a custom asset pack, then create a new record for it.
                                    Dim UsedAsset As New TextureClass

                                    'Call the "ParseTexture" function to get the info for the custom asset and
                                    'for the custom asset pack it's from.
                                    UsedAsset = ParseTexture(texture, AssetManifest, AssetType, levelname)

                                    'If this asset has already been added to our UsedAssetList, then increment the Instances property
                                    'of the existing record instead of adding the new one as a duplicate.
                                    'Otherwise, add the new record to the list And set the Instances property to 1.
                                    SearchIndex = UsedAssetList.FindIndex(Function(p) p.AssetName = UsedAsset.AssetName)
                                    If SearchIndex >= 0 Then
                                        UsedAssetList(SearchIndex).Instances += 1
                                    Else
                                        UsedAsset.Instances = 1
                                        UsedAssetList.Add(UsedAsset)
                                    End If

                                End If

                                'Iterate through all portal assets anchored to the current wall.
                                For Each portal In asset("portals")
                                    'Get the "texture" value for the current asset.
                                    texture = portal("texture")
                                    If texture.StartsWith("res://packs/") Then
                                        'If this is from a custom asset pack, then create a new record for it.
                                        Dim UsedAsset As New TextureClass

                                        'Call the "ParseTexture" function to get the info for the custom asset and
                                        'for the custom asset pack it's from.
                                        UsedAsset = ParseTexture(texture, AssetManifest, "portals", levelname)

                                        'If this asset has already been added to our UsedAssetList, then increment the Instances property
                                        'of the existing record instead of adding the new one as a duplicate.
                                        'Otherwise, add the new record to the list And set the Instances property to 1.
                                        SearchIndex = UsedAssetList.FindIndex(Function(p) p.AssetName = UsedAsset.AssetName)
                                        If SearchIndex >= 0 Then
                                            UsedAssetList(SearchIndex).Instances += 1
                                        Else
                                            UsedAsset.Instances = 1
                                            UsedAssetList.Add(UsedAsset)
                                        End If

                                    End If
                                Next
                            Next
                        End If
                    Case Else
                        'Iterate through all assets of the current type where the type is not defined above.
                        'i.e. "lights", "objects", "paths", "patterns", and freestanding "portals".
                        If Not World("levels")(levelnumber)(AssetType) Is Nothing Then
                            For Each asset In World("levels")(levelnumber)(AssetType)
                                'Get the "texture" value for the current asset.
                                texture = asset("texture")
                                If Not texture Is Nothing Then
                                    If texture.StartsWith("res://packs/") Then
                                        'If this is from a custom asset pack, then create a new record for it.
                                        Dim UsedAsset As New TextureClass

                                        'Call the "ParseTexture" function to get the info for the custom asset and
                                        'for the custom asset pack it's from.
                                        UsedAsset = ParseTexture(texture, AssetManifest, AssetType, levelname)

                                        'If this asset has already been added to our UsedAssetList, then increment the Instances property
                                        'of the existing record instead of adding the new one as a duplicate.
                                        'Otherwise, add the new record to the list And set the Instances property to 1.
                                        SearchIndex = UsedAssetList.FindIndex(Function(p) p.AssetName = UsedAsset.AssetName)
                                        If SearchIndex >= 0 Then
                                            UsedAssetList(SearchIndex).Instances += 1
                                        Else
                                            UsedAsset.Instances = 1
                                            UsedAssetList.Add(UsedAsset)
                                        End If
                                    End If
                                End If
                            Next
                        End If
                End Select
            Next
        Next

        Message &= Indent & "Custom assets in use:" & vbCrLf
        If UsedAssetList.Count > 0 Then
            'Our UsedAssetList has at least one record, then do all the things.
            Dim PackName As String
            Dim PackVersion As String
            Dim PackAuthor As String
            Dim Instances As String

            'Sort the UsedAssetList by the PackName property.
            UsedAssetList.Sort(Function(x, y) x.PackName.CompareTo(y.PackName))

            'List all the custom assets that have been used, grouping them by their respective pack names.


            'Message &= Indent & Indent & AssetPack("name") & ", Version " & AssetPack("version") & ", " & AssetPack("author") & vbCrLf
            PackName = UsedAssetList(0).PackName
            PackVersion = UsedAssetList(0).PackVersion
            PackAuthor = UsedAssetList(0).PackAuthor
            Message &= Indent & Indent & "From " & PackName & ", Version " & PackVersion & ", by " & PackAuthor & ":" & vbCrLf
            For AssetIndex = 0 To UsedAssetList.Count - 1
                If PackName <> UsedAssetList(AssetIndex).PackName Then
                    PackName = UsedAssetList(AssetIndex).PackName
                    PackVersion = UsedAssetList(AssetIndex).PackVersion
                    PackAuthor = UsedAssetList(AssetIndex).PackAuthor
                    Message &= vbCrLf & Indent & Indent & "From " & PackName & ", Version " & PackVersion & ", by " & PackAuthor & ":" & vbCrLf
                End If
                If UsedAssetList(AssetIndex).Instances = 1 Then Instances = "instance" Else Instances = "instances"
                Message &= Indent & Indent & Indent & UsedAssetList(AssetIndex).AssetName & ", " & UsedAssetList(AssetIndex).AssetType & ", " & UsedAssetList(AssetIndex).Instances & " " & Instances & " found." & vbCrLf & vbCrLf
            Next
        Else
            'If our UsedAssetList has no records, then indicate that no custom assets have been used.
            Message &= Indent & Indent & "No custom assets are in use." & vbCrLf & vbCrLf
        End If

        Message &= Indent & "Embedded assets:" & vbCrLf
        If Not World("embedded") Is Nothing Then
            For Each Asset In World("embedded")
                Message &= Indent & Indent & Asset.Name & vbCrLf
            Next
        Else
            Message &= Indent & Indent & "No embedded assets were found." & vbCrLf
        End If


        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)

    End Sub
End Module
