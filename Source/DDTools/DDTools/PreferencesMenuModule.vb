Imports Newtonsoft.Json
Imports System.IO
Module PreferencesMenuModule
    Public Function BuildConfigObject()
        Dim TagAssetsConfig As New System.Collections.Specialized.OrderedDictionary
        Dim ConvertAssetsConfig As New System.Collections.Specialized.OrderedDictionary
        Dim ConvertPacksConfig As New System.Collections.Specialized.OrderedDictionary
        Dim CopyAssetsConfig As New System.Collections.Specialized.OrderedDictionary
        Dim CopyTilesConfig As New System.Collections.Specialized.OrderedDictionary
        Dim MapDetailsConfig As New System.Collections.Specialized.OrderedDictionary
        Dim PackAssetsConfig As New System.Collections.Specialized.OrderedDictionary
        Dim UnpackAssetsConfig As New System.Collections.Specialized.OrderedDictionary

        TagAssetsConfig.Add("source", "")
        TagAssetsConfig.Add("default_tag", "")
        TagAssetsConfig.Add("create_log", False)
        TagAssetsConfig.Add("select_all", False)

        ConvertAssetsConfig.Add("source", "")
        ConvertAssetsConfig.Add("destination", "")
        ConvertAssetsConfig.Add("create_log", False)
        ConvertAssetsConfig.Add("select_all", False)

        ConvertPacksConfig.Add("source", "")
        ConvertPacksConfig.Add("destination", "")
        ConvertPacksConfig.Add("cleanup", True)
        ConvertPacksConfig.Add("create_log", False)
        ConvertPacksConfig.Add("select_all", False)

        CopyAssetsConfig.Add("source", "")
        CopyAssetsConfig.Add("destination", "")
        CopyAssetsConfig.Add("create_tags", False)
        CopyAssetsConfig.Add("separate_portals", True)
        CopyAssetsConfig.Add("create_log", False)
        CopyAssetsConfig.Add("select_all", False)

        CopyTilesConfig.Add("source", "")
        CopyTilesConfig.Add("destination", "")
        CopyTilesConfig.Add("create_log", False)
        CopyTilesConfig.Add("select_all", False)

        MapDetailsConfig.Add("source", "")
        MapDetailsConfig.Add("create_log", False)
        MapDetailsConfig.Add("select_all", False)

        PackAssetsConfig.Add("source", "")
        PackAssetsConfig.Add("destination", "")
        PackAssetsConfig.Add("overwrite", False)
        PackAssetsConfig.Add("create_log", False)
        PackAssetsConfig.Add("select_all", False)

        UnpackAssetsConfig.Add("source", "")
        UnpackAssetsConfig.Add("destination", "")
        UnpackAssetsConfig.Add("create_log", False)
        UnpackAssetsConfig.Add("select_all", False)

        Dim ConfigObject As New System.Collections.Specialized.OrderedDictionary
        ConfigObject.Add("tag_assets", TagAssetsConfig)
        ConfigObject.Add("convert_assets", ConvertAssetsConfig)
        ConfigObject.Add("convert_packs", ConvertPacksConfig)
        ConfigObject.Add("copy_assets", CopyAssetsConfig)
        ConfigObject.Add("copy_tiles", CopyTilesConfig)
        ConfigObject.Add("map_details", MapDetailsConfig)
        ConfigObject.Add("pack_assets", PackAssetsConfig)
        ConfigObject.Add("unpack_assets", UnpackAssetsConfig)

        Dim SerialConfig As String = JsonConvert.SerializeObject(ConfigObject, Formatting.Indented)
        Dim JSONConfig = JsonConvert.DeserializeObject(SerialConfig)
        Return JSONConfig
    End Function

    Public Function GetSavedConfig(ConfigFile As String)
        Dim rawJson = File.ReadAllText(ConfigFile)
        Dim ConfigObject = JsonConvert.DeserializeObject(rawJson)
        Return ConfigObject
    End Function

    Public Sub SaveNewConfig(ConfigObject As Object, ConfigFileName As String)
        Dim NewConfig As String = JsonConvert.SerializeObject(ConfigObject, Formatting.Indented)
        My.Computer.FileSystem.WriteAllText(ConfigFileName, NewConfig, False, System.Text.Encoding.ASCII)
    End Sub
End Module
