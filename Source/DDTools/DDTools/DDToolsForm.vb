Imports System.IO
Imports Newtonsoft.Json
Imports System.Environment
Public Class DDToolsForm
    Private Sub DDToolsForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PreferencesToolStripMenu.Visible = False
        'TitlePanel.Hide()
        TagAssetsGroupBox.Hide()
        ConvertAssetsGroupBox.Hide()
        ConvertPacksGroupBox.Hide()
        CopyAssetsGroupBox.Hide()
        CopyTilesGroupBox.Hide()
        DataFilesGroupBox.Hide()
        MapDetailsGroupBox.Hide()
        PackAssetsGroupBox.Hide()
        UnpackAssetsGroupBox.Hide()

        Me.Text = "EightBitz's DDTools - Version " + My.Application.Info.Version.ToString

        VersionLabel.Text = "Version " & My.Application.Info.Version.ToString
        GitHubLinkLabel.Text = "The latest version of this program can be found in its GitHub repository."
        GitHubLinkLabel.Links.Add(55, 17, "https://github.com/EightBitz/Dungeondraft-Custom-Tools")
        CreativeCommonsLinkLabel.Text = "This work is licensed under a Creative Commons Attribution-NonCommercial 4.0 International License."
        CreativeCommonsLinkLabel.Links.Add(30, 68, "https://creativecommons.org/licenses/by-nc/4.0/legalcode")
        EmailLinkLabel.Text = "Email: eightbitz73@outlook.com"
        EmailLinkLabel.Links.Add(7, 23, "mailto:eightbitz73@outlook.com")

        Me.Size = New Size(1000, 453)
        TitlePanel.BringToFront()
        TitlePanel.Show()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        AboutBox1.Show()
    End Sub

    Private Sub GitHubLinkLabel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles GitHubLinkLabel.LinkClicked
        System.Diagnostics.Process.Start(e.Link.LinkData)
    End Sub

    Private Sub CreativeCommonsLinkLabel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles CreativeCommonsLinkLabel.LinkClicked
        System.Diagnostics.Process.Start(e.Link.LinkData)
    End Sub

    Private Sub EmailLinkLabel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles EmailLinkLabel.LinkClicked
        System.Diagnostics.Process.Start(e.Link.LinkData)
    End Sub

    Private Sub DocumentationMenuItem_Click(sender As Object, e As EventArgs) Handles DocumentationMenuItem.Click
        Dim DocFile As String = My.Application.Info.DirectoryPath & "\Documentation\EightBitz's Dungeondraft Tools Documentation.pdf"
        System.Diagnostics.Process.Start(DocFile)
    End Sub

    Private Sub LicenseMenuItem_Click(sender As Object, e As EventArgs) Handles LicenseMenuItem.Click
        Dim LicenseFile As String = My.Application.Info.DirectoryPath & "\Documentation\LICENSE.html"
        System.Diagnostics.Process.Start(LicenseFile)
    End Sub

    Private Sub READMEMenuItem_Click(sender As Object, e As EventArgs) Handles READMEMenuItem.Click
        Dim READMEFile As String = My.Application.Info.DirectoryPath & "\Documentation\README.html"
        System.Diagnostics.Process.Start(READMEFile)
    End Sub

    '###### Main Menu Items ######
    Private Sub TagAssetsMenuItem_Click(sender As Object, e As EventArgs) Handles TagAssetsMenuItem.Click
        TitlePanel.Hide()
        'TagAssetsGroupBox.Hide()
        ConvertAssetsGroupBox.Hide()
        ConvertPacksGroupBox.Hide()
        CopyAssetsGroupBox.Hide()
        CopyTilesGroupBox.Hide()
        DataFilesGroupBox.Hide()
        MapDetailsGroupBox.Hide()
        PackAssetsGroupBox.Hide()
        UnpackAssetsGroupBox.Hide()

        Dim ConfigFileName As String = GlobalVariables.ConfigFolderName & "\" & GlobalVariables.ConfigFileName
        If File.Exists(ConfigFileName) Then
            Dim ConfigObject = GetSavedConfig(ConfigFileName)
            Dim ActiveTool As String = "tag_assets"
            GetSavedConfig(ConfigFileName)

            TagAssetsSourceTextBox.Text = ConfigObject(ActiveTool)("source")
            TagAssetsDefaultTagTextBox.Text = ConfigObject(ActiveTool)("default_tag")
            TagAssetsLogCheckBox.Checked = ConfigObject(ActiveTool)("create_log")
            TagAssetsSelectAllCheckBox.Checked = ConfigObject(ActiveTool)("select_all")

            TagAssetsSourceTextBox_LostFocus(sender, e)
        End If

        Me.Size = New Size(802, 653)
        PreferencesToolStripMenu.Visible = True
        TagAssetsGroupBox.BringToFront()
        TagAssetsGroupBox.Show()
    End Sub

    Private Sub ConvertAssetssMenuItem_Click(sender As Object, e As EventArgs) Handles ConvertAssetsMenuItem.Click
        TitlePanel.Hide()
        TagAssetsGroupBox.Hide()
        'ConvertAssetsGroupBox.Hide()
        ConvertPacksGroupBox.Hide()
        CopyAssetsGroupBox.Hide()
        CopyTilesGroupBox.Hide()
        DataFilesGroupBox.Hide()
        MapDetailsGroupBox.Hide()
        PackAssetsGroupBox.Hide()
        UnpackAssetsGroupBox.Hide()

        Dim ConfigFileName As String = GlobalVariables.ConfigFolderName & "\" & GlobalVariables.ConfigFileName
        If File.Exists(ConfigFileName) Then
            Dim ConfigObject = GetSavedConfig(ConfigFileName)
            Dim ActiveTool As String = "convert_assets"
            GetSavedConfig(ConfigFileName)

            ConvertAssetsSourceTextBox.Text = ConfigObject(ActiveTool)("source")
            ConvertAssetsDestinationTextBox.Text = ConfigObject(ActiveTool)("destination")
            ConvertAssetsLogCheckBox.Checked = ConfigObject(ActiveTool)("create_log")
            ConvertAssetsSelectAllCheckBox.Checked = ConfigObject(ActiveTool)("select_all")

            ConvertAssetsSourceTextBox_LostFocus(sender, e)
        End If

        Me.Size = New Size(802, 653)
        PreferencesToolStripMenu.Visible = True
        ConvertAssetsGroupBox.BringToFront()
        ConvertAssetsGroupBox.Show()
    End Sub

    Private Sub ConvertPacksMenuItem_Click(sender As Object, e As EventArgs) Handles ConvertPacksMenuItem.Click
        TitlePanel.Hide()
        TagAssetsGroupBox.Hide()
        ConvertAssetsGroupBox.Hide()
        'ConvertPacksGroupBox.Hide()
        CopyAssetsGroupBox.Hide()
        CopyTilesGroupBox.Hide()
        DataFilesGroupBox.Hide()
        MapDetailsGroupBox.Hide()
        PackAssetsGroupBox.Hide()
        UnpackAssetsGroupBox.Hide()

        Dim ConfigFileName As String = GlobalVariables.ConfigFolderName & "\" & GlobalVariables.ConfigFileName
        If File.Exists(ConfigFileName) Then
            Dim ConfigObject = GetSavedConfig(ConfigFileName)
            Dim ActiveTool As String = "convert_packs"
            GetSavedConfig(ConfigFileName)

            ConvertPacksSourceTextBox.Text = ConfigObject(ActiveTool)("source")
            ConvertPacksDestinationTextBox.Text = ConfigObject(ActiveTool)("destination")
            ConvertPacksCleanUpCheckBox.Checked = ConfigObject(ActiveTool)("cleanup")
            ConvertPacksLogCheckBox.Checked = ConfigObject(ActiveTool)("create_log")
            ConvertPacksSelectAllCheckBox.Checked = ConfigObject(ActiveTool)("select_all")

            ConvertPacksSourceTextBox_LostFocus(sender, e)
        End If

        Me.Size = New Size(802, 653)
        PreferencesToolStripMenu.Visible = True
        ConvertPacksGroupBox.BringToFront()
        ConvertPacksGroupBox.Show()
    End Sub

    Private Sub CopyAssetsMenuItem_Click(sender As Object, e As EventArgs) Handles CopyAssetsMenuItem.Click
        TitlePanel.Hide()
        TagAssetsGroupBox.Hide()
        ConvertAssetsGroupBox.Hide()
        ConvertPacksGroupBox.Hide()
        'CopyAssetsGroupBox.Hide()
        CopyTilesGroupBox.Hide()
        DataFilesGroupBox.Hide()
        MapDetailsGroupBox.Hide()
        PackAssetsGroupBox.Hide()
        UnpackAssetsGroupBox.Hide()

        Dim ConfigFileName As String = GlobalVariables.ConfigFolderName & "\" & GlobalVariables.ConfigFileName
        If File.Exists(ConfigFileName) Then
            Dim ConfigObject = GetSavedConfig(ConfigFileName)
            Dim ActiveTool As String = "copy_assets"
            GetSavedConfig(ConfigFileName)

            CopyAssetsSourceTextBox.Text = ConfigObject(ActiveTool)("source")
            CopyAssetsDestinationTextBox.Text = ConfigObject(ActiveTool)("destination")
            CopyAssetsCreateTagsCheckBox.Checked = ConfigObject(ActiveTool)("create_tags")
            CopyAssetsPortalsCheckBox.Checked = ConfigObject(ActiveTool)("separate_portals")
            CopyAssetsLogCheckBox.Checked = ConfigObject(ActiveTool)("create_log")
            CopyAssetsSelectAllCheckBox.Checked = ConfigObject(ActiveTool)("select_all")

            CopyAssetsSourceTextBox_LostFocus(sender, e)
        End If

        Me.Size = New Size(802, 653)
        PreferencesToolStripMenu.Visible = True
        CopyAssetsGroupBox.BringToFront()
        CopyAssetsGroupBox.Show()
    End Sub
    Private Sub CopyTilesMenuItem_Click(sender As Object, e As EventArgs) Handles CopyTilesMenuItem.Click
        TitlePanel.Hide()
        TagAssetsGroupBox.Hide()
        ConvertAssetsGroupBox.Hide()
        ConvertPacksGroupBox.Hide()
        CopyAssetsGroupBox.Hide()
        'CopyTilesGroupBox.Hide()
        DataFilesGroupBox.Hide()
        MapDetailsGroupBox.Hide()
        PackAssetsGroupBox.Hide()
        UnpackAssetsGroupBox.Hide()

        Dim ConfigFileName As String = GlobalVariables.ConfigFolderName & "\" & GlobalVariables.ConfigFileName
        If File.Exists(ConfigFileName) Then
            Dim ConfigObject = GetSavedConfig(ConfigFileName)
            Dim ActiveTool As String = "copy_tiles"
            GetSavedConfig(ConfigFileName)

            CopyTilesSourceTextBox.Text = ConfigObject(ActiveTool)("source")
            CopyTilesDestinationTextBox.Text = ConfigObject(ActiveTool)("destination")
            CopyTilesLogCheckBox.Checked = ConfigObject(ActiveTool)("create_log")
            CopyTilesSelectAllCheckBox.Checked = ConfigObject(ActiveTool)("select_all")

            CopyTilesSourceTextBox_LostFocus(sender, e)
        End If

        Me.Size = New Size(802, 653)
        PreferencesToolStripMenu.Visible = True
        CopyTilesGroupBox.BringToFront()
        CopyTilesGroupBox.Show()
    End Sub

    Private Sub DataFilesMenuItem_Click(sender As Object, e As EventArgs) Handles DataFilesMenuItem.Click
        TitlePanel.Hide()
        TagAssetsGroupBox.Hide()
        ConvertAssetsGroupBox.Hide()
        ConvertPacksGroupBox.Hide()
        CopyAssetsGroupBox.Hide()
        CopyTilesGroupBox.Hide()
        'DataFilesGroupBox.Hide()
        MapDetailsGroupBox.Hide()
        PackAssetsGroupBox.Hide()
        UnpackAssetsGroupBox.Hide()

        Dim ConfigFileName As String = GlobalVariables.ConfigFolderName & "\" & GlobalVariables.ConfigFileName
        If File.Exists(ConfigFileName) Then
            Dim ConfigObject = GetSavedConfig(ConfigFileName)
            Dim ActiveTool As String = "data_files"
            GetSavedConfig(ConfigFileName)

            If Not ConfigObject(ActiveTool) Is Nothing Then
                DataFilesSourceTextBox.Text = ConfigObject(ActiveTool)("source")
                DataFilesLogCheckBox.Checked = ConfigObject(ActiveTool)("create_log")
                DataFilesSourceTextBox_LostFocus(sender, e)
            End If
        End If

        Me.Size = New Size(1000, 653)
        PreferencesToolStripMenu.Visible = True
        DataFilesGroupBox.BringToFront()
        DataFilesGroupBox.Show()
    End Sub

    Private Sub MapDetailsMenuItem_Click(sender As Object, e As EventArgs) Handles MapDetailsMenuItem.Click
        TitlePanel.Hide()
        TagAssetsGroupBox.Hide()
        ConvertAssetsGroupBox.Hide()
        ConvertPacksGroupBox.Hide()
        CopyAssetsGroupBox.Hide()
        CopyTilesGroupBox.Hide()
        DataFilesGroupBox.Hide()
        'MapDetailsGroupBox.Hide()
        PackAssetsGroupBox.Hide()
        UnpackAssetsGroupBox.Hide()

        Dim ConfigFileName As String = GlobalVariables.ConfigFolderName & "\" & GlobalVariables.ConfigFileName
        If File.Exists(ConfigFileName) Then
            Dim ConfigObject = GetSavedConfig(ConfigFileName)
            Dim ActiveTool As String = "map_details"
            GetSavedConfig(ConfigFileName)

            MapDetailsSourceTextBox.Text = ConfigObject(ActiveTool)("source")
            MapDetailsLogCheckBox.Checked = ConfigObject(ActiveTool)("create_log")
            MapDetailsSelectAllCheckBox.Checked = ConfigObject(ActiveTool)("select_all")

            MapDetailsSourceTextBox_LostFocus(sender, e)
        End If

        Me.Size = New Size(802, 653)
        PreferencesToolStripMenu.Visible = True
        MapDetailsGroupBox.BringToFront()
        MapDetailsGroupBox.Show()
    End Sub

    Private Sub PackAssetsMenuItem_Click(sender As Object, e As EventArgs) Handles PackAssetsMenuItem.Click
        TitlePanel.Hide()
        TagAssetsGroupBox.Hide()
        ConvertAssetsGroupBox.Hide()
        ConvertPacksGroupBox.Hide()
        CopyAssetsGroupBox.Hide()
        CopyTilesGroupBox.Hide()
        DataFilesGroupBox.Hide()
        MapDetailsGroupBox.Hide()
        'PackAssetsGroupBox.Hide()
        UnpackAssetsGroupBox.Hide()

        Dim ConfigFileName As String = GlobalVariables.ConfigFolderName & "\" & GlobalVariables.ConfigFileName
        If File.Exists(ConfigFileName) Then
            Dim ConfigObject = GetSavedConfig(ConfigFileName)
            Dim ActiveTool As String = "pack_assets"
            GetSavedConfig(ConfigFileName)

            PackAssetsSourceTextBox.Text = ConfigObject(ActiveTool)("source")
            PackAssetsDestinationTextBox.Text = ConfigObject(ActiveTool)("destination")
            PackAssetsOverwriteCheckBox.Checked = ConfigObject(ActiveTool)("overwrite")
            PackAssetsLogCheckBox.Checked = ConfigObject(ActiveTool)("create_log")
            PackAssetsSelectAllCheckBox.Checked = ConfigObject(ActiveTool)("select_all")

            PackAssetsSourceTextBox_LostFocus(sender, e)
        End If

        Me.Size = New Size(1000, 653)
        PreferencesToolStripMenu.Visible = True
        PackAssetsGroupBox.BringToFront()
        PackAssetsGroupBox.Show()
    End Sub

    Private Sub UnpackAssetsMenuItem_Click(sender As Object, e As EventArgs) Handles UnpackAssetsMenuItem.Click
        TitlePanel.Hide()
        TagAssetsGroupBox.Hide()
        ConvertAssetsGroupBox.Hide()
        ConvertPacksGroupBox.Hide()
        CopyAssetsGroupBox.Hide()
        CopyTilesGroupBox.Hide()
        DataFilesGroupBox.Hide()
        MapDetailsGroupBox.Hide()
        PackAssetsGroupBox.Hide()
        'UnpackAssetsGroupBox.Hide()

        Dim ConfigFileName As String = GlobalVariables.ConfigFolderName & "\" & GlobalVariables.ConfigFileName
        If File.Exists(ConfigFileName) Then
            Dim ConfigObject = GetSavedConfig(ConfigFileName)
            Dim ActiveTool As String = "unpack_assets"
            GetSavedConfig(ConfigFileName)

            UnpackAssetsSourceTextBox.Text = ConfigObject(ActiveTool)("source")
            UnpackAssetsDestinationTextBox.Text = ConfigObject(ActiveTool)("destination")
            UnpackAssetsLogCheckBox.Checked = ConfigObject(ActiveTool)("create_log")
            PackAssetsSelectAllCheckBox.Checked = ConfigObject(ActiveTool)("select_all")

            UnpackAssetsSourceTextBox_LostFocus(sender, e)
        End If

        Me.Size = New Size(802, 653)
        PreferencesToolStripMenu.Visible = True
        UnpackAssetsGroupBox.BringToFront()
        UnpackAssetsGroupBox.Show()
    End Sub

    '###### Preference Menu Items ######
    Private Sub SavePrefsMenuItem_Click(sender As Object, e As EventArgs) Handles SavePrefsMenuItem.Click
        Dim ConfigFileName As String = GlobalVariables.ConfigFileName
        Dim ConfigFolderName As String = GlobalVariables.ConfigFolderName
        Dim ConfigObject
        Dim ActiveTool As String
        If My.Computer.FileSystem.FileExists(ConfigFileName) Then
            ConfigObject = GetSavedConfig(ConfigFileName)
        Else
            ConfigObject = BuildConfigObject()
        End If

        If TagAssetsGroupBox.Visible = True Then
            ActiveTool = "tag_assets"
            If ConfigObject(ActiveTool) Is Nothing Then
                Dim NewPrefs As New Newtonsoft.Json.Linq.JObject
                ConfigObject.Add(ActiveTool, NewPrefs)
            End If
            If ConfigObject(ActiveTool)("source") Is Nothing Then ConfigObject(ActiveTool).Add("source", "")
            ConfigObject(ActiveTool)("source") = TagAssetsSourceTextBox.Text
            If ConfigObject(ActiveTool)("default_tag") Is Nothing Then ConfigObject(ActiveTool).Add("default_tag", "")
            ConfigObject(ActiveTool)("default_tag") = TagAssetsDefaultTagTextBox.Text
            If ConfigObject(ActiveTool)("create_log") Is Nothing Then ConfigObject(ActiveTool).Add("create_log", "")
            ConfigObject(ActiveTool)("create_log") = TagAssetsLogCheckBox.Checked
            If ConfigObject(ActiveTool)("select_all") Is Nothing Then ConfigObject(ActiveTool).Add("select_all", "")
            ConfigObject(ActiveTool)("select_all") = TagAssetsSelectAllCheckBox.Checked
        ElseIf ConvertAssetsGroupBox.Visible = True Then
            ActiveTool = "convert_assets"
            If ConfigObject(ActiveTool) Is Nothing Then
                Dim NewPrefs As New Newtonsoft.Json.Linq.JObject
                ConfigObject.Add(ActiveTool, NewPrefs)
            End If
            If ConfigObject(ActiveTool)("source") Is Nothing Then ConfigObject(ActiveTool).Add("source", "")
            ConfigObject(ActiveTool)("source") = ConvertAssetsSourceTextBox.Text
            If ConfigObject(ActiveTool)("destination") Is Nothing Then ConfigObject(ActiveTool).Add("destination", "")
            ConfigObject(ActiveTool)("destination") = ConvertAssetsDestinationTextBox.Text
            If ConfigObject(ActiveTool)("create_log") Is Nothing Then ConfigObject(ActiveTool).Add("create_log", "")
            ConfigObject(ActiveTool)("create_log") = ConvertAssetsLogCheckBox.Checked
            If ConfigObject(ActiveTool)("select_all") Is Nothing Then ConfigObject(ActiveTool).Add("select_all", "")
            ConfigObject(ActiveTool)("select_all") = ConvertAssetsSelectAllCheckBox.Checked
        ElseIf ConvertPacksGroupBox.Visible = True Then
            ActiveTool = "convert_packs"
            If ConfigObject(ActiveTool) Is Nothing Then
                Dim NewPrefs As New Newtonsoft.Json.Linq.JObject
                ConfigObject.Add(ActiveTool, NewPrefs)
            End If
            If ConfigObject(ActiveTool)("source") Is Nothing Then ConfigObject(ActiveTool).Add("source", "")
            ConfigObject(ActiveTool)("source") = ConvertPacksSourceTextBox.Text
            If ConfigObject(ActiveTool)("destination") Is Nothing Then ConfigObject(ActiveTool).Add("destination", "")
            ConfigObject(ActiveTool)("destination") = ConvertPacksDestinationTextBox.Text
            If ConfigObject(ActiveTool)("cleanup") Is Nothing Then ConfigObject(ActiveTool).Add("cleanup", "")
            ConfigObject(ActiveTool)("cleanup") = ConvertPacksCleanUpCheckBox.Checked
            If ConfigObject(ActiveTool)("create_log") Is Nothing Then ConfigObject(ActiveTool).Add("create_log", "")
            ConfigObject(ActiveTool)("create_log") = ConvertPacksLogCheckBox.Checked
            If ConfigObject(ActiveTool)("select_all") Is Nothing Then ConfigObject(ActiveTool).Add("select_all", "")
            ConfigObject(ActiveTool)("select_all") = ConvertPacksSelectAllCheckBox.Checked
        ElseIf CopyAssetsGroupBox.Visible = True Then
            ActiveTool = "copy_assets"
            If ConfigObject(ActiveTool) Is Nothing Then
                Dim NewPrefs As New Newtonsoft.Json.Linq.JObject
                ConfigObject.Add(ActiveTool, NewPrefs)
            End If
            If ConfigObject(ActiveTool)("source") Is Nothing Then ConfigObject(ActiveTool).Add("source", "")
            ConfigObject(ActiveTool)("source") = CopyAssetsSourceTextBox.Text
            If ConfigObject(ActiveTool)("destination") Is Nothing Then ConfigObject(ActiveTool).Add("destination", "")
            ConfigObject(ActiveTool)("destination") = CopyAssetsDestinationTextBox.Text
            If ConfigObject(ActiveTool)("create_tags") Is Nothing Then ConfigObject(ActiveTool).Add("create_tags", "")
            ConfigObject(ActiveTool)("create_tags") = CopyAssetsCreateTagsCheckBox.Checked
            If ConfigObject(ActiveTool)("separate_portals") Is Nothing Then ConfigObject(ActiveTool).Add("separate_portals", "")
            ConfigObject(ActiveTool)("separate_portals") = CopyAssetsPortalsCheckBox.Checked
            If ConfigObject(ActiveTool)("create_log") Is Nothing Then ConfigObject(ActiveTool).Add("create_log", "")
            ConfigObject(ActiveTool)("create_log") = CopyAssetsLogCheckBox.Checked
            If ConfigObject(ActiveTool)("select_all") Is Nothing Then ConfigObject(ActiveTool).Add("select_all", "")
            ConfigObject(ActiveTool)("select_all") = CopyAssetsSelectAllCheckBox.Checked
        ElseIf CopyTilesGroupBox.Visible = True Then
            ActiveTool = "copy_tiles"
            If ConfigObject(ActiveTool) Is Nothing Then
                Dim NewPrefs As New Newtonsoft.Json.Linq.JObject
                ConfigObject.Add(ActiveTool, NewPrefs)
            End If
            If ConfigObject(ActiveTool)("source") Is Nothing Then ConfigObject(ActiveTool).Add("source", "")
            ConfigObject(ActiveTool)("source") = CopyTilesSourceTextBox.Text
            If ConfigObject(ActiveTool)("destination") Is Nothing Then ConfigObject(ActiveTool).Add("destination", "")
            ConfigObject(ActiveTool)("destination") = CopyTilesDestinationTextBox.Text
            If ConfigObject(ActiveTool)("create_log") Is Nothing Then ConfigObject(ActiveTool).Add("create_log", "")
            ConfigObject(ActiveTool)("create_log") = CopyTilesLogCheckBox.Checked
            If ConfigObject(ActiveTool)("select_all") Is Nothing Then ConfigObject(ActiveTool).Add("select_all", "")
            ConfigObject(ActiveTool)("select_all") = CopyTilesSelectAllCheckBox.Checked
        ElseIf DataFilesGroupBox.Visible = True Then
            ActiveTool = "data_files"
            If ConfigObject(ActiveTool) Is Nothing Then
                Dim NewPrefs As New Newtonsoft.Json.Linq.JObject
                ConfigObject.Add(ActiveTool, NewPrefs)
            End If
            If ConfigObject(ActiveTool)("source") Is Nothing Then ConfigObject(ActiveTool).Add("source", "")
            ConfigObject(ActiveTool)("source") = DataFilesSourceTextBox.Text
            If ConfigObject(ActiveTool)("create_log") Is Nothing Then ConfigObject(ActiveTool).Add("create_log", "")
            ConfigObject(ActiveTool)("create_log") = DataFilesLogCheckBox.Checked
        ElseIf MapDetailsGroupBox.Visible = True Then
            ActiveTool = "map_details"
            If ConfigObject(ActiveTool) Is Nothing Then
                Dim NewPrefs As New Newtonsoft.Json.Linq.JObject
                ConfigObject.Add(ActiveTool, NewPrefs)
            End If
            If ConfigObject(ActiveTool)("source") Is Nothing Then ConfigObject(ActiveTool).Add("source", "")
            ConfigObject(ActiveTool)("source") = MapDetailsSourceTextBox.Text
            If ConfigObject(ActiveTool)("create_log") Is Nothing Then ConfigObject(ActiveTool).Add("create_log", "")
            ConfigObject(ActiveTool)("create_log") = MapDetailsLogCheckBox.Checked
            If ConfigObject(ActiveTool)("select_all") Is Nothing Then ConfigObject(ActiveTool).Add("select_all", "")
            ConfigObject(ActiveTool)("select_all") = MapDetailsSelectAllCheckBox.Checked
        ElseIf PackAssetsGroupBox.Visible = True Then
            ActiveTool = "pack_assets"
            If ConfigObject(ActiveTool) Is Nothing Then
                Dim NewPrefs As New Newtonsoft.Json.Linq.JObject
                ConfigObject.Add(ActiveTool, NewPrefs)
            End If
            If ConfigObject(ActiveTool)("source") Is Nothing Then ConfigObject(ActiveTool).Add("source", "")
            ConfigObject(ActiveTool)("source") = PackAssetsSourceTextBox.Text
            If ConfigObject(ActiveTool)("destination") Is Nothing Then ConfigObject(ActiveTool).Add("destination", "")
            ConfigObject(ActiveTool)("destination") = PackAssetsDestinationTextBox.Text
            If ConfigObject(ActiveTool)("overwrite") Is Nothing Then ConfigObject(ActiveTool).Add("overwrite", "")
            ConfigObject(ActiveTool)("overwrite") = PackAssetsOverwriteCheckBox.Checked
            If ConfigObject(ActiveTool)("create_log") Is Nothing Then ConfigObject(ActiveTool).Add("create_log", "")
            ConfigObject(ActiveTool)("create_log") = PackAssetsLogCheckBox.Checked
            If ConfigObject(ActiveTool)("select_all") Is Nothing Then ConfigObject(ActiveTool).Add("select_all", "")
            ConfigObject(ActiveTool)("select_all") = PackAssetsSelectAllCheckBox.Checked
        ElseIf UnpackAssetsGroupBox.Visible = True Then
            ActiveTool = "unpack_assets"
            If ConfigObject(ActiveTool) Is Nothing Then
                Dim NewPrefs As New Newtonsoft.Json.Linq.JObject
                ConfigObject.Add(ActiveTool, NewPrefs)
            End If
            If ConfigObject(ActiveTool)("source") Is Nothing Then ConfigObject(ActiveTool).Add("source", "")
            ConfigObject(ActiveTool)("source") = UnpackAssetsSourceTextBox.Text
            If ConfigObject(ActiveTool)("destination") Is Nothing Then ConfigObject(ActiveTool).Add("destination", "")
            ConfigObject(ActiveTool)("destination") = UnpackAssetsDestinationTextBox.Text
            If ConfigObject(ActiveTool)("create_log") Is Nothing Then ConfigObject(ActiveTool).Add("create_log", "")
            ConfigObject(ActiveTool)("create_log") = UnpackAssetsLogCheckBox.Checked
            If ConfigObject(ActiveTool)("select_all") Is Nothing Then ConfigObject(ActiveTool).Add("select_all", "")
            ConfigObject(ActiveTool)("select_all") = UnpackAssetsSelectAllCheckBox.Checked
        End If

        SaveNewConfig(ConfigObject, ConfigFolderName, ConfigFileName)
    End Sub

    Private Sub LoadPrefsMenuItem_Click(sender As Object, e As EventArgs) Handles LoadPrefsMenuItem.Click
        Dim ConfigFileName As String = GlobalVariables.ConfigFolderName & "\" & GlobalVariables.ConfigFileName
        Dim ActiveTool As String

        If My.Computer.FileSystem.FileExists(ConfigFileName) Then
            Dim ConfigObject = GetSavedConfig(ConfigFileName)

            If TagAssetsGroupBox.Visible = True Then
                ActiveTool = "tag_assets"
                TagAssetsSourceTextBox.Text = ConfigObject(ActiveTool)("source")
                TagAssetsDefaultTagTextBox.Text = ConfigObject(ActiveTool)("default_tag")
                TagAssetsLogCheckBox.Checked = ConfigObject(ActiveTool)("create_log")
                TagAssetsSelectAllCheckBox.Checked = ConfigObject(ActiveTool)("select_all")
                TagAssetsSourceTextBox_LostFocus(sender, e)
            ElseIf ConvertAssetsGroupBox.Visible = True Then
                ActiveTool = "convert_assets"
                ConvertAssetsSourceTextBox.Text = ConfigObject(ActiveTool)("source")
                ConvertAssetsDestinationTextBox.Text = ConfigObject(ActiveTool)("destination")
                ConvertAssetsLogCheckBox.Checked = ConfigObject(ActiveTool)("create_log")
                ConvertAssetsSelectAllCheckBox.Checked = ConfigObject(ActiveTool)("select_all")
                ConvertAssetsSourceTextBox_LostFocus(sender, e)
            ElseIf ConvertPacksGroupBox.Visible = True Then
                ActiveTool = "convert_packs"
                ConvertPacksSourceTextBox.Text = ConfigObject(ActiveTool)("source")
                ConvertPacksDestinationTextBox.Text = ConfigObject(ActiveTool)("destination")
                ConvertPacksCleanUpCheckBox.Checked = ConfigObject(ActiveTool)("cleanup")
                ConvertPacksLogCheckBox.Checked = ConfigObject(ActiveTool)("create_log")
                ConvertPacksSelectAllCheckBox.Checked = ConfigObject("convertpacks")("select_all")
                ConvertPacksSourceTextBox_LostFocus(sender, e)
            ElseIf CopyAssetsGroupBox.Visible = True Then
                ActiveTool = "copy_assets"
                CopyAssetsSourceTextBox.Text = ConfigObject(ActiveTool)("source")
                CopyAssetsDestinationTextBox.Text = ConfigObject(ActiveTool)("destination")
                CopyAssetsCreateTagsCheckBox.Checked = ConfigObject(ActiveTool)("create_tags")
                CopyAssetsPortalsCheckBox.Checked = ConfigObject(ActiveTool)("separate_portals")
                CopyAssetsLogCheckBox.Checked = ConfigObject(ActiveTool)("create_log")
                CopyAssetsSelectAllCheckBox.Checked = ConfigObject(ActiveTool)("select_all")
                CopyAssetsSourceTextBox_LostFocus(sender, e)
            ElseIf CopyTilesGroupBox.Visible = True Then
                ActiveTool = "copy_tiles"
                CopyTilesSourceTextBox.Text = ConfigObject(ActiveTool)("source")
                CopyTilesDestinationTextBox.Text = ConfigObject(ActiveTool)("destination")
                CopyTilesLogCheckBox.Checked = ConfigObject(ActiveTool)("create_log")
                CopyTilesSelectAllCheckBox.Checked = ConfigObject(ActiveTool)("select_all")
                CopyTilesSourceTextBox_LostFocus(sender, e)
            ElseIf DataFilesGroupBox.Visible = True Then
                ActiveTool = "data_files"
                DataFilesSourceTextBox.Text = ConfigObject(ActiveTool)("source")
                DataFilesLogCheckBox.Checked = ConfigObject(ActiveTool)("create_log")
                DataFilesSourceTextBox_LostFocus(sender, e)
            ElseIf MapDetailsGroupBox.Visible = True Then
                ActiveTool = "map_details"
                MapDetailsSourceTextBox.Text = ConfigObject(ActiveTool)("source")
                MapDetailsLogCheckBox.Checked = ConfigObject(ActiveTool)("create_log")
                MapDetailsSelectAllCheckBox.Checked = ConfigObject(ActiveTool)("select_all")
                MapDetailsSourceTextBox_LostFocus(sender, e)
            ElseIf PackAssetsGroupBox.Visible = True Then
                ActiveTool = "pack_assets"
                PackAssetsSourceTextBox.Text = ConfigObject(ActiveTool)("source")
                PackAssetsDestinationTextBox.Text = ConfigObject(ActiveTool)("destination")
                PackAssetsOverwriteCheckBox.Checked = ConfigObject(ActiveTool)("overwrite")
                PackAssetsLogCheckBox.Checked = ConfigObject(ActiveTool)("create_log")
                PackAssetsSelectAllCheckBox.Checked = ConfigObject(ActiveTool)("select_all")
                PackAssetsSourceTextBox_LostFocus(sender, e)
            ElseIf UnpackAssetsGroupBox.Visible = True Then
                ActiveTool = "unpack_assets"
                UnpackAssetsSourceTextBox.Text = ConfigObject(ActiveTool)("source")
                UnpackAssetsDestinationTextBox.Text = ConfigObject(ActiveTool)("destination")
                UnpackAssetsLogCheckBox.Checked = ConfigObject(ActiveTool)("create_log")
                UnpackAssetsSelectAllCheckBox.Checked = ConfigObject(ActiveTool)("select_all")
                UnpackAssetsSourceTextBox_LostFocus(sender, e)
            End If
        End If
    End Sub

    '###### Tag Assets Group Box ######
    Private Sub TagAssetsSourceTextBox_LostFocus(sender As Object, e As EventArgs) Handles TagAssetsSourceTextBox.LostFocus
        TagAssetsCheckedListBox.Items.Clear()
        Dim SourceFolderName As String
        Dim IsFolderNameValid As Boolean
        Dim DoesFolderExist As Boolean
        SourceFolderName = TagAssetsSourceTextBox.Text
        IsFolderNameValid = IsValidPathName(SourceFolderName)
        DoesFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If IsFolderNameValid And DoesFolderExist Then
            If My.Computer.FileSystem.DirectoryExists(SourceFolderName & "\textures\objects") Then
                Dim SourceFolderInfo = My.Computer.FileSystem.GetDirectoryInfo(SourceFolderName)
                TagAssetsSourceTextBox.Text = SourceFolderInfo.Parent.FullName
                TagAssetsCheckedListBox.Items.Add(SourceFolderInfo.Name)
                TagAssetsCheckedListBox.SetItemChecked(TagAssetsCheckedListBox.Items.Count - 1, True)
            Else
                For Each TagFolder As String In My.Computer.FileSystem.GetDirectories(SourceFolderName)
                    If My.Computer.FileSystem.DirectoryExists(TagFolder & "\textures\objects") Then
                        Dim FolderName As New System.IO.DirectoryInfo(TagFolder)
                        TagAssetsCheckedListBox.Items.Add(FolderName.Name)
                        If TagAssetsSelectAllCheckBox.Checked Then TagAssetsCheckedListBox.SetItemChecked(TagAssetsCheckedListBox.Items.Count - 1, True)
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub TagAssetsBrowseButton_Click(sender As Object, e As EventArgs) Handles TagAssetsBrowseButton.Click
        TagAssetsCheckedListBox.Items.Clear()
        TagAssetsSourceBrowserDialog.ShowDialog()
        TagAssetsSourceTextBox.Text = TagAssetsSourceBrowserDialog.SelectedPath

        Dim SourceFolderName As String
        Dim IsFolderNameValid As Boolean
        Dim DoesFolderExist As Boolean
        SourceFolderName = TagAssetsSourceTextBox.Text
        IsFolderNameValid = IsValidPathName(SourceFolderName)
        DoesFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If IsFolderNameValid And DoesFolderExist Then
            If My.Computer.FileSystem.DirectoryExists(SourceFolderName & "\textures\objects") Then
                Dim SourceFolderInfo = My.Computer.FileSystem.GetDirectoryInfo(SourceFolderName)
                TagAssetsSourceTextBox.Text = SourceFolderInfo.Parent.FullName
                TagAssetsCheckedListBox.Items.Add(SourceFolderInfo.Name)
                TagAssetsCheckedListBox.SetItemChecked(TagAssetsCheckedListBox.Items.Count - 1, True)
            Else
                For Each TagFolder As String In My.Computer.FileSystem.GetDirectories(SourceFolderName)
                    If My.Computer.FileSystem.DirectoryExists(TagFolder & "\textures\objects") Then
                        Dim FolderName As New System.IO.DirectoryInfo(TagFolder)
                        TagAssetsCheckedListBox.Items.Add(FolderName.Name)
                        If TagAssetsSelectAllCheckBox.Checked Then TagAssetsCheckedListBox.SetItemChecked(TagAssetsCheckedListBox.Items.Count - 1, True)
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub TagAssetsSelectAllCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles TagAssetsSelectAllCheckBox.CheckedChanged
        If TagAssetsSelectAllCheckBox.Checked Then
            TagAssetsSelectAllButton_Click(sender, e)
        Else
            TagAssetsSelectNoneButton_Click(sender, e)
        End If
    End Sub

    Private Sub TagAssetsSelectAllButton_Click(sender As Object, e As EventArgs) Handles TagAssetsSelectAllButton.Click
        Dim Count As Integer
        For Count = 0 To TagAssetsCheckedListBox.Items.Count - 1
            TagAssetsCheckedListBox.SetItemChecked(Count, True)
        Next
    End Sub

    Private Sub TagAssetsSelectNoneButton_Click(sender As Object, e As EventArgs) Handles TagAssetsSelectNoneButton.Click
        Dim Count As Integer
        For Count = 0 To TagAssetsCheckedListBox.Items.Count - 1
            TagAssetsCheckedListBox.SetItemChecked(Count, False)
        Next
    End Sub

    Private Sub TagAssetsStartButton_Click(sender As Object, e As EventArgs) Handles TagAssetsStartButton.Click
        Dim SourceFolderName As String = TagAssetsSourceTextBox.Text
        Dim IsSourceFolderNameValid As String
        Dim DoesSourceFolderExist As Boolean
        Dim DefaultTag As String = TagAssetsDefaultTagTextBox.Text
        Dim CreateLog As Boolean = TagAssetsLogCheckBox.Checked
        Dim LogFileName As String = "TagAssets.log"
        Dim TagSource As String
        Dim Indent As String = "    " '4 spaces
        Dim SubIndent As String = "        " '8 spaces
        Dim Message As String

        IsSourceFolderNameValid = IsValidPathName(SourceFolderName)
        DoesSourceFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If SourceFolderName <> "" And IsSourceFolderNameValid And DoesSourceFolderExist Then
            If TagAssetsCheckedListBox.CheckedItems.Count >= 1 Then
                OutputForm.OutputTextBox.Text = ""
                OutputForm.Show()
                OutputForm.BringToFront()
                Message = "### Starting selected folders at " & DateTime.Now & "." & vbCrLf & vbCrLf
                OutputForm.OutputTextBox.AppendText(Message)
                If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, False)
                Dim SelectedFolders = TagAssetsCheckedListBox.CheckedItems
                For Each AssetFolder In SelectedFolders
                    If SourceFolderName.EndsWith("\") Then
                        TagSource = SourceFolderName & AssetFolder
                    Else
                        TagSource = SourceFolderName & "\" & AssetFolder
                    End If
                    Message = Indent & "Starting " & AssetFolder & " at " & DateTime.Now & vbCrLf
                    OutputForm.OutputTextBox.AppendText(Message)
                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                    TagAssetsSub(TagSource, AssetFolder, DefaultTag, CreateLog, LogFileName, SubIndent)
                    Message = Indent & "Finished " & AssetFolder & " at " & DateTime.Now & vbCrLf & vbCrLf
                    OutputForm.OutputTextBox.AppendText(Message)
                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                Next
                Message = "### Finished selected folders at " & DateTime.Now & "." & vbCrLf
                OutputForm.OutputTextBox.AppendText(Message)
                If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
            Else
                MsgBox("Nothing was selected.")
            End If
        Else
            MsgBox("Source folder name is either invalid or does not exist.")
        End If

    End Sub

    '###### Convert Assets Group Box ######
    Private Sub ConvertAssetsSourceTextBox_LostFocus(sender As Object, e As EventArgs) Handles ConvertAssetsSourceTextBox.LostFocus
        ConvertAssetsCheckedListBox.Items.Clear()
        Dim UserFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim SourceFolderName As String
        Dim IsFolderNameValid As Boolean
        Dim DoesFolderExist As Boolean
        Dim SourceFolder As System.IO.DirectoryInfo
        Dim DestinationFolderName As String

        SourceFolderName = ConvertAssetsSourceTextBox.Text
        IsFolderNameValid = IsValidPathName(SourceFolderName)
        DoesFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If IsFolderNameValid And DoesFolderExist Then
            SourceFolder = New System.IO.DirectoryInfo(SourceFolderName)
            DestinationFolderName = UserFolder & "\Dungeondraft\Converted Assets\" & SourceFolder.Name
            If ConvertAssetsDestinationTextBox.Text = "" Then ConvertAssetsDestinationTextBox.Text = DestinationFolderName
            For Each AssetFolder As String In My.Computer.FileSystem.GetDirectories(SourceFolderName)
                Dim FolderName As New System.IO.DirectoryInfo(AssetFolder)
                ConvertAssetsCheckedListBox.Items.Add(FolderName.Name)
                If ConvertAssetsSelectAllCheckBox.Checked Then ConvertAssetsCheckedListBox.SetItemChecked(ConvertAssetsCheckedListBox.Items.Count - 1, True)
            Next
        End If
    End Sub

    Private Sub ConvertAssetsSourceBrowseButton_Click(sender As Object, e As EventArgs) Handles ConvertAssetsSourceBrowseButton.Click
        ConvertAssetsCheckedListBox.Items.Clear()
        ConvertAssetsSourceBrowserDialog.ShowDialog()
        ConvertAssetsSourceTextBox.Text = ConvertAssetsSourceBrowserDialog.SelectedPath

        Dim UserFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim SourceFolderName As String
        Dim IsFolderNameValid As Boolean
        Dim DoesFolderExist As Boolean
        Dim SourceFolder As System.IO.DirectoryInfo
        Dim DestinationFolderName As String

        SourceFolderName = ConvertAssetsSourceTextBox.Text
        IsFolderNameValid = IsValidPathName(SourceFolderName)
        DoesFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If IsFolderNameValid And DoesFolderExist Then
            SourceFolder = New System.IO.DirectoryInfo(SourceFolderName)
            DestinationFolderName = UserFolder & "\Dungeondraft\Converted Assets\" & SourceFolder.Name
            If ConvertAssetsDestinationTextBox.Text = "" Then ConvertAssetsDestinationTextBox.Text = DestinationFolderName
            For Each AssetFolder As String In My.Computer.FileSystem.GetDirectories(SourceFolderName)
                Dim FolderName As New System.IO.DirectoryInfo(AssetFolder)
                ConvertAssetsCheckedListBox.Items.Add(FolderName.Name)
                If ConvertAssetsSelectAllCheckBox.Checked Then ConvertAssetsCheckedListBox.SetItemChecked(ConvertAssetsCheckedListBox.Items.Count - 1, True)
            Next
        End If
    End Sub

    Private Sub ConvertAssetsDestinationBrowseButton_Click(sender As Object, e As EventArgs) Handles ConvertAssetsDestinationBrowseButton.Click
        ConvertAssetsDestinationBrowserDialog.ShowDialog()
        ConvertAssetsDestinationTextBox.Text = ConvertAssetsDestinationBrowserDialog.SelectedPath
    End Sub

    Private Sub ConvertAssetsSelectAllCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles ConvertAssetsSelectAllCheckBox.CheckedChanged
        If ConvertAssetsSelectAllCheckBox.Checked Then
            ConvertAssetsSelectAllButton_Click(sender, e)
        Else
            ConvertAssetsSelectNoneButton_Click(sender, e)
        End If
    End Sub

    Private Sub ConvertAssetsSelectAllButton_Click(sender As Object, e As EventArgs) Handles ConvertAssetsSelectAllButton.Click
        Dim Count As Integer
        For Count = 0 To ConvertAssetsCheckedListBox.Items.Count - 1
            ConvertAssetsCheckedListBox.SetItemChecked(Count, True)
        Next
    End Sub
    Private Sub ConvertAssetsSelectNoneButton_Click(sender As Object, e As EventArgs) Handles ConvertAssetsSelectNoneButton.Click
        Dim Count As Integer
        For Count = 0 To ConvertAssetsCheckedListBox.Items.Count - 1
            ConvertAssetsCheckedListBox.SetItemChecked(Count, False)
        Next
    End Sub
    Private Sub ConvertAssetsStartButton_Click(sender As Object, e As EventArgs) Handles ConvertAssetsStartButton.Click

        Dim SourceFolderName As String = ConvertAssetsSourceTextBox.Text
        Dim DestinationFolderName As String = ConvertAssetsDestinationTextBox.Text
        Dim IsSourceFolderNameValid As String
        Dim DoesSourceFolderExist As Boolean
        Dim IsDestinationFolderNameValid As String
        Dim LogFileName As String = "ConvertAssets.log"
        Dim CreateLog = ConvertAssetsLogCheckBox.Checked
        Dim CopySource As String
        Dim CopyDestination As String
        Dim Indent As String = "    " '4 spaces
        Dim SubIndent As String = "        " '8 spaces
        Dim Message As String
        Dim ProgramFiles As String = Environment.GetEnvironmentVariable("ProgramFiles")
        Dim Magick As Boolean
        Dim ImageMagickEXE As String

        Magick = False
        Dim ImageMagickFolder = Directory.GetDirectories(ProgramFiles, "ImageMagick*")
        For Each folder As String In ImageMagickFolder
            ImageMagickEXE = folder & "\magick.exe"
            If My.Computer.FileSystem.FileExists(ImageMagickEXE) Then Magick = True
        Next

        IsSourceFolderNameValid = IsValidPathName(SourceFolderName)
        DoesSourceFolderExist = System.IO.Directory.Exists(SourceFolderName)

        IsDestinationFolderNameValid = IsValidPathName(DestinationFolderName)

        If Magick Then
            If SourceFolderName <> "" And IsSourceFolderNameValid And DoesSourceFolderExist Then
                If DestinationFolderName <> "" And IsDestinationFolderNameValid Then
                    If ConvertAssetsCheckedListBox.CheckedItems.Count >= 1 Then
                        OutputForm.OutputTextBox.Text = ""
                        OutputForm.Show()
                        OutputForm.BringToFront()
                        Message = "### Starting selected folders at " & DateTime.Now & "." & vbCrLf & vbCrLf
                        OutputForm.OutputTextBox.AppendText(Message)
                        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, False)
                        Dim SelectedFolders = ConvertAssetsCheckedListBox.CheckedItems
                        For Each AssetFolder In SelectedFolders
                            If SourceFolderName.EndsWith("\") Then
                                CopySource = SourceFolderName & AssetFolder
                            Else
                                CopySource = SourceFolderName & "\" & AssetFolder
                            End If

                            If DestinationFolderName.EndsWith("\") Then
                                CopyDestination = DestinationFolderName & AssetFolder
                            Else
                                CopyDestination = DestinationFolderName & "\" & AssetFolder
                            End If

                            Message = Indent & "Starting " & AssetFolder & " at " & DateTime.Now & vbCrLf
                            OutputForm.OutputTextBox.AppendText(Message)
                            If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                            ConvertAssetsSub(CopySource, CopyDestination, CreateLog, LogFileName, SubIndent, ImageMagickEXE)
                            Message = Indent & "Finished " & AssetFolder & " at " & DateTime.Now & vbCrLf & vbCrLf
                            OutputForm.OutputTextBox.AppendText(Message)
                            If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                        Next
                        Message = "### Finished selected folders at " & DateTime.Now & "." & vbCrLf
                        OutputForm.OutputTextBox.AppendText(Message)
                        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                    End If
                Else
                    MsgBox("Destination folder name is invalid.")
                End If
            Else
                MsgBox("Source folder name is either invalid or does not exist.")
            End If
        Else
            MsgBox("Cannot find ImageMagick installation in " & ProgramFiles)
        End If
    End Sub

    '###### Convert Packs Group Box ######
    Private Sub ConvertPacksSourceTextBox_LostFocus(sender As Object, e As EventArgs) Handles ConvertPacksSourceTextBox.LostFocus
        ConvertPacksCheckedListBox.Items.Clear()

        Dim UserFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim SourceFolderName As String
        Dim IsFolderNameValid As Boolean
        Dim DoesFolderExist As Boolean
        Dim SourceFolder As System.IO.DirectoryInfo
        Dim DestinationFolderName As String

        SourceFolderName = ConvertPacksSourceTextBox.Text
        IsFolderNameValid = IsValidPathName(SourceFolderName)
        DoesFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If IsFolderNameValid And DoesFolderExist Then
            SourceFolder = New System.IO.DirectoryInfo(SourceFolderName)
            DestinationFolderName = UserFolder & "\Dungeondraft\" & SourceFolder.Name
            If ConvertPacksDestinationTextBox.Text = "" Then ConvertPacksDestinationTextBox.Text = DestinationFolderName
            For Each PackFile As String In My.Computer.FileSystem.GetFiles(SourceFolderName)
                Dim PackName As New System.IO.DirectoryInfo(PackFile)
                If PackName.Extension = ".dungeondraft_pack" Then
                    ConvertPacksCheckedListBox.Items.Add(PackName.Name)
                    If ConvertPacksSelectAllCheckBox.Checked Then ConvertPacksCheckedListBox.SetItemChecked(ConvertPacksCheckedListBox.Items.Count - 1, True)
                End If
            Next
        End If
    End Sub

    Private Sub ConvertPacksSourceBrowseButton_Click(sender As Object, e As EventArgs) Handles ConvertPacksSourceBrowseButton.Click
        ConvertPacksCheckedListBox.Items.Clear()
        ConvertPacksSourceBrowserDialog.ShowDialog()
        ConvertPacksSourceTextBox.Text = ConvertPacksSourceBrowserDialog.SelectedPath

        Dim UserFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim SourceFolderName As String
        Dim IsFolderNameValid As Boolean
        Dim DoesFolderExist As Boolean
        Dim SourceFolder As System.IO.DirectoryInfo
        Dim DestinationFolderName As String

        SourceFolderName = ConvertPacksSourceTextBox.Text
        IsFolderNameValid = IsValidPathName(SourceFolderName)
        DoesFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If IsFolderNameValid And DoesFolderExist Then
            SourceFolder = New System.IO.DirectoryInfo(SourceFolderName)
            DestinationFolderName = UserFolder & "\Dungeondraft\" & SourceFolder.Name
            If ConvertPacksDestinationTextBox.Text = "" Then ConvertPacksDestinationTextBox.Text = DestinationFolderName
            For Each PackFile As String In My.Computer.FileSystem.GetFiles(SourceFolderName)
                Dim PackName As New System.IO.DirectoryInfo(PackFile)
                If PackName.Extension = ".dungeondraft_pack" Then
                    ConvertPacksCheckedListBox.Items.Add(PackName.Name)
                    If ConvertPacksSelectAllCheckBox.Checked Then ConvertPacksCheckedListBox.SetItemChecked(ConvertPacksCheckedListBox.Items.Count - 1, True)
                End If
            Next
        End If
    End Sub

    Private Sub ConvertPacksDestinationBrowseButton_Click(sender As Object, e As EventArgs) Handles ConvertPacksDestinationBrowseButton.Click
        ConvertPacksDestinationBrowserDialog.ShowDialog()
        ConvertPacksDestinationTextBox.Text = ConvertPacksDestinationBrowserDialog.SelectedPath
    End Sub

    Private Sub ConvertPacksSelectAllCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles ConvertPacksSelectAllCheckBox.CheckedChanged
        If ConvertPacksSelectAllCheckBox.Checked Then
            ConvertPacksSelectAllButton_Click(sender, e)
        Else
            ConvertPacksSelectNoneButton_Click(sender, e)
        End If
    End Sub

    Private Sub ConvertPacksSelectAllButton_Click(sender As Object, e As EventArgs) Handles ConvertPacksSelectAllButton.Click
        Dim Count As Integer
        For Count = 0 To ConvertPacksCheckedListBox.Items.Count - 1
            ConvertPacksCheckedListBox.SetItemChecked(Count, True)
        Next
    End Sub

    Private Sub ConvertPacksSelectNoneButton_Click(sender As Object, e As EventArgs) Handles ConvertPacksSelectNoneButton.Click
        Dim Count As Integer
        For Count = 0 To ConvertPacksCheckedListBox.Items.Count - 1
            ConvertPacksCheckedListBox.SetItemChecked(Count, False)
        Next
    End Sub

    Private Sub ConvertPacksStartButton_Click(sender As Object, e As EventArgs) Handles ConvertPacksStartButton.Click
        Dim SourceFolderName As String = ConvertPacksSourceTextBox.Text
        Dim DestinationFolderName As String = ConvertPacksDestinationTextBox.Text
        Dim IsSourceFolderNameValid As String
        Dim DoesSourceFolderExist As Boolean
        Dim IsDestinationFolderNameValid As String
        Dim CleanUp As Boolean = ConvertPacksCleanUpCheckBox.Checked
        Dim LogFileName As String = "ConvertPacks.log"
        Dim UnpackEXE As String = "dungeondraft-unpack.exe"
        Dim PackEXE As String = "dungeondraft-pack.exe"
        Dim UnpackEXEexists As Boolean
        Dim PackEXEexists As Boolean
        Dim MagickExists As Boolean
        Dim ProgramFiles As String = Environment.GetEnvironmentVariable("ProgramFiles")
        Dim ImageMagickEXE As String
        Dim Message As String
        Dim CreateLog = ConvertPacksLogCheckBox.Checked
        Dim Indent As String = "    " '4 spaces
        Dim SubIndent As String = "        " '8 spaces

        MagickExists = False
        UnpackEXEexists = False
        PackEXEexists = False

        If My.Computer.FileSystem.FileExists(UnpackEXE) Then UnpackEXEexists = True
        If My.Computer.FileSystem.FileExists(PackEXE) Then PackEXEexists = True
        Dim ImageMagickFolder = Directory.GetDirectories(ProgramFiles, "ImageMagick*")
        For Each folder As String In ImageMagickFolder
            ImageMagickEXE = folder & "\magick.exe"
            If My.Computer.FileSystem.FileExists(ImageMagickEXE) Then MagickExists = True
        Next

        IsSourceFolderNameValid = IsValidPathName(SourceFolderName)
        DoesSourceFolderExist = System.IO.Directory.Exists(SourceFolderName)
        IsDestinationFolderNameValid = IsValidPathName(DestinationFolderName)

        If UnpackEXEexists And PackEXEexists And MagickExists Then
            If SourceFolderName <> "" And IsSourceFolderNameValid And DoesSourceFolderExist Then
                If DestinationFolderName <> "" And IsDestinationFolderNameValid Then
                    If ConvertPacksCheckedListBox.CheckedItems.Count >= 1 Then
                        OutputForm.OutputTextBox.Text = ""
                        OutputForm.Show()
                        OutputForm.BringToFront()
                        Message = "### Starting selected packs at " & DateTime.Now & "." & vbCrLf & vbCrLf
                        OutputForm.OutputTextBox.AppendText(Message)
                        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, False)
                        Dim SelectedPacks = ConvertPacksCheckedListBox.CheckedItems
                        For Each PackFile In SelectedPacks
                            If SourceFolderName.EndsWith("\") Then SourceFolderName = SourceFolderName.TrimEnd("\")

                            If DestinationFolderName.EndsWith("\") Then DestinationFolderName = DestinationFolderName.TrimEnd("\")

                            Message = Indent & "Starting " & PackFile & " at " & DateTime.Now & vbCrLf
                            OutputForm.OutputTextBox.AppendText(Message)
                            If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                            ConvertPackSub(SourceFolderName, DestinationFolderName, PackFile, CreateLog, LogFileName, SubIndent, ImageMagickEXE, UnpackEXE, PackEXE, CleanUp)
                            Message = Indent & "Finished " & PackFile & " at " & DateTime.Now & vbCrLf & vbCrLf
                            OutputForm.OutputTextBox.AppendText(Message)
                            If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                        Next
                        Message = "### Finished selected packs at " & DateTime.Now & "." & vbCrLf
                        OutputForm.OutputTextBox.AppendText(Message)
                        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                    Else
                        MsgBox("Nothing was selected.")
                    End If

                Else
                    MsgBox("Destination folder name Is invalid.")
                End If
            Else
                MsgBox("Source folder name Is either invalid Or does Not exist.")
            End If
        Else
            Message = "One or more dependencies are missing:" & vbCrLf
            If Not UnpackEXEexists Then Message &= UnpackEXE & " not found." & vbCrLf
            If Not MagickExists Then Message &= "ImageMagick not found." & vbCrLf
            If Not PackEXEexists Then Message &= PackEXE & " not found."
            MsgBox(Message)
        End If
    End Sub

    '###### Copy Assets Group Box ######
    Private Sub CopyAssetsSourceTextBox_LostFocus(sender As Object, e As EventArgs) Handles CopyAssetsSourceTextBox.LostFocus
        CopyAssetsCheckedListBox.Items.Clear()
        Dim UserFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim SourceFolderName As String
        Dim IsFolderNameValid As Boolean
        Dim DoesFolderExist As Boolean
        Dim SourceFolder As System.IO.DirectoryInfo
        Dim DestinationFolderName As String

        SourceFolderName = CopyAssetsSourceTextBox.Text
        IsFolderNameValid = IsValidPathName(SourceFolderName)
        DoesFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If IsFolderNameValid And DoesFolderExist Then
            SourceFolder = New System.IO.DirectoryInfo(SourceFolderName)
            DestinationFolderName = UserFolder & "\Dungeondraft\Copied Assets\" & SourceFolder.Name
            If CopyAssetsDestinationTextBox.Text = "" Then CopyAssetsDestinationTextBox.Text = DestinationFolderName
            For Each AssetFolder As String In My.Computer.FileSystem.GetDirectories(SourceFolderName)
                Dim FolderName As New System.IO.DirectoryInfo(AssetFolder)
                CopyAssetsCheckedListBox.Items.Add(FolderName.Name)
                If CopyAssetsSelectAllCheckBox.Checked Then CopyAssetsCheckedListBox.SetItemChecked(CopyAssetsCheckedListBox.Items.Count - 1, True)
            Next
        End If
    End Sub

    Private Sub CopyAssetsSourceBrowseButton_Click(sender As Object, e As EventArgs) Handles CopyAssetsSourceBrowseButton.Click
        CopyAssetsCheckedListBox.Items.Clear()
        CopyAssetsSourceBrowserDialog.ShowDialog()
        CopyAssetsSourceTextBox.Text = CopyAssetsSourceBrowserDialog.SelectedPath

        Dim UserFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim SourceFolderName As String
        Dim IsFolderNameValid As Boolean
        Dim DoesFolderExist As Boolean
        Dim SourceFolder As System.IO.DirectoryInfo
        Dim DestinationFolderName As String

        SourceFolderName = CopyAssetsSourceTextBox.Text
        IsFolderNameValid = IsValidPathName(SourceFolderName)
        DoesFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If IsFolderNameValid And DoesFolderExist Then
            SourceFolder = New System.IO.DirectoryInfo(SourceFolderName)
            DestinationFolderName = UserFolder & "\Dungeondraft\Copied Assets\" & SourceFolder.Name
            If CopyAssetsDestinationTextBox.Text = "" Then CopyAssetsDestinationTextBox.Text = DestinationFolderName
            For Each AssetFolder As String In My.Computer.FileSystem.GetDirectories(SourceFolderName)
                Dim FolderName As New System.IO.DirectoryInfo(AssetFolder)
                CopyAssetsCheckedListBox.Items.Add(FolderName.Name)
                If CopyAssetsSelectAllCheckBox.Checked Then CopyAssetsCheckedListBox.SetItemChecked(CopyAssetsCheckedListBox.Items.Count - 1, True)
            Next
        End If
    End Sub

    Private Sub CopyAssetsDestinationBrowseButton_Click(sender As Object, e As EventArgs) Handles CopyAssetsDestinationBrowseButton.Click
        CopyAssetsDestinationBrowserDialog.ShowDialog()
        CopyAssetsDestinationTextBox.Text = CopyAssetsDestinationBrowserDialog.SelectedPath
    End Sub

    Private Sub CopyAssetsSelectAllCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles CopyAssetsSelectAllCheckBox.CheckedChanged
        If CopyAssetsSelectAllCheckBox.Checked Then
            CopyAssetsSelectAllButton_Click(sender, e)
        Else
            CopyAssetsSelectNoneButton_Click(sender, e)
        End If
    End Sub

    Private Sub CopyAssetsSelectAllButton_Click(sender As Object, e As EventArgs) Handles CopyAssetsSelectAllButton.Click
        Dim Count As Integer
        For Count = 0 To CopyAssetsCheckedListBox.Items.Count - 1
            CopyAssetsCheckedListBox.SetItemChecked(Count, True)
        Next
    End Sub
    Private Sub CopyAssetsSelectNoneButton_Click(sender As Object, e As EventArgs) Handles CopyAssetsSelectNoneButton.Click
        Dim Count As Integer
        For Count = 0 To CopyAssetsCheckedListBox.Items.Count - 1
            CopyAssetsCheckedListBox.SetItemChecked(Count, False)
        Next
    End Sub
    Private Sub CopyAssetsStartButton_Click(sender As Object, e As EventArgs) Handles CopyAssetsStartButton.Click
        Dim PowerShellPath As String = Environment.SystemDirectory & "\WindowsPowerShell\v1.0\powershell.exe"
        Dim ScriptName As String = "'.\DDCopyAssets.ps1'"
        Dim SourceFolderName As String = CopyAssetsSourceTextBox.Text
        Dim DestinationFolderName As String = CopyAssetsDestinationTextBox.Text
        Dim CreateTagFile As Boolean = CopyAssetsCreateTagsCheckBox.Checked
        Dim Portals As Boolean = CopyAssetsPortalsCheckBox.Checked
        Dim IsSourceFolderNameValid As String
        Dim DoesSourceFolderExist As Boolean
        Dim IsDestinationFolderNameValid As String
        Dim CreateLog = CopyAssetsLogCheckBox.Checked
        Dim LogFileName As String = "CopyAssets.log"
        Dim CopySource As String
        Dim CopyDestination As String
        Dim Indent As String = "    " '4 spaces
        Dim SubIndent As String = "        " '8 spaces
        Dim Message As String
        Dim DefaultTag As String = ""

        If CopyAssetsCreateTagsCheckBox.Checked Then
            Dim ConfigObject = GetSavedConfig(GlobalVariables.ConfigFileName)
            DefaultTag = ConfigObject("tag_assets")("default_tag")
        End If

        IsSourceFolderNameValid = IsValidPathName(SourceFolderName)
        DoesSourceFolderExist = System.IO.Directory.Exists(SourceFolderName)

        IsDestinationFolderNameValid = IsValidPathName(DestinationFolderName)

        If SourceFolderName <> "" And IsSourceFolderNameValid And DoesSourceFolderExist Then
            If DestinationFolderName <> "" And IsDestinationFolderNameValid Then
                If CopyAssetsCheckedListBox.CheckedItems.Count >= 1 Then
                    OutputForm.OutputTextBox.Text = ""
                    OutputForm.Show()
                    OutputForm.BringToFront()
                    Message = "### Starting selected folders at " & DateTime.Now & "." & vbCrLf & vbCrLf
                    OutputForm.OutputTextBox.AppendText(Message)
                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, False)
                    Dim SelectedFolders = CopyAssetsCheckedListBox.CheckedItems
                    For Each AssetFolder In SelectedFolders
                        If SourceFolderName.EndsWith("\") Then
                            CopySource = SourceFolderName & AssetFolder
                        Else
                            CopySource = SourceFolderName & "\" & AssetFolder
                        End If

                        If DestinationFolderName.EndsWith("\") Then
                            CopyDestination = DestinationFolderName & AssetFolder
                        Else
                            CopyDestination = DestinationFolderName & "\" & AssetFolder
                        End If

                        Message = Indent & "Starting " & AssetFolder & " at " & DateTime.Now & vbCrLf
                        OutputForm.OutputTextBox.AppendText(Message)
                        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                        CopyAssetsSub(CopySource, CopyDestination, CreateTagFile, DefaultTag, Portals, CreateLog, LogFileName, SubIndent)
                        Message = Indent & "Finished " & AssetFolder & " at " & DateTime.Now & vbCrLf & vbCrLf
                        OutputForm.OutputTextBox.AppendText(Message)
                        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                    Next
                    Message = "### Finished selected folders at " & DateTime.Now & "." & vbCrLf
                    OutputForm.OutputTextBox.AppendText(Message)
                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                Else
                    MsgBox("Nothing was selected.")
                End If
            Else
                MsgBox("Destination folder name is invalid.")
            End If
        Else
            MsgBox("Source folder name is either invalid or does not exist.")
        End If
    End Sub

    '###### Copy Tiles Group Box ######
    Private Sub CopyTilesSourceTextBox_LostFocus(sender As Object, e As EventArgs) Handles CopyTilesSourceTextBox.LostFocus
        CopyTilesDataGridView.Rows.Clear()
        CopyTilesDataGridView.Columns.Clear()
        Dim UserFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim SourceFolderName As String
        Dim IsFolderNameValid As Boolean
        Dim DoesFolderExist As Boolean
        Dim SourceFolder As System.IO.DirectoryInfo
        Dim DestinationFolderName As String
        Dim CreateLog As Boolean = TagAssetsLogCheckBox.Checked
        Dim LogFileName As String = "CopyTiles.log"
        Dim SubIndent As String = "        "
        Dim SelectAll As Boolean = CopyTilesSelectAllCheckBox.Checked

        SourceFolderName = CopyTilesSourceTextBox.Text
        IsFolderNameValid = IsValidPathName(SourceFolderName)
        DoesFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If IsFolderNameValid And DoesFolderExist Then
            SourceFolder = New System.IO.DirectoryInfo(SourceFolderName)
            DestinationFolderName = UserFolder & "\Dungeondraft\Copied Assets\" & SourceFolder.Name
            If CopyTilesDestinationTextBox.Text = "" Then CopyTilesDestinationTextBox.Text = DestinationFolderName
            LoadTilesSub(SourceFolderName, CreateLog, LogFileName, SelectAll, SubIndent)
        End If
    End Sub

    Private Sub CopyTilesSourceBrowseButton_Click(sender As Object, e As EventArgs) Handles CopyTilesSourceBrowseButton.Click
        CopyTilesDataGridView.Rows.Clear()
        CopyTilesDataGridView.Columns.Clear()
        CopyTilesSourceBrowserDialog.ShowDialog()
        CopyTilesSourceTextBox.Text = CopyTilesSourceBrowserDialog.SelectedPath

        Dim UserFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim SourceFolderName As String
        Dim IsFolderNameValid As Boolean
        Dim DoesFolderExist As Boolean
        Dim SourceFolder As System.IO.DirectoryInfo
        Dim DestinationFolderName As String
        Dim CreateLog As Boolean = TagAssetsLogCheckBox.Checked
        Dim LogFileName As String = "CopyTiles.log"
        Dim SubIndent As String = "        "
        Dim SelectAll As Boolean = CopyTilesSelectAllCheckBox.Checked

        SourceFolderName = CopyTilesSourceTextBox.Text
        IsFolderNameValid = IsValidPathName(SourceFolderName)
        DoesFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If IsFolderNameValid And DoesFolderExist Then
            SourceFolder = New System.IO.DirectoryInfo(SourceFolderName)
            DestinationFolderName = UserFolder & "\Dungeondraft\Copied Assets\" & SourceFolder.Name
            If CopyTilesDestinationTextBox.Text = "" Then CopyTilesDestinationTextBox.Text = DestinationFolderName
            LoadTilesSub(SourceFolderName, CreateLog, LogFileName, SelectAll, SubIndent)
        End If
    End Sub

    Private Sub CopyTileSelectAllCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles CopyTilesSelectAllCheckBox.CheckedChanged
        For RowIndex As Integer = 0 To CopyTilesDataGridView.Rows.Count - 1
            CopyTilesDataGridView.Rows(RowIndex).Cells(0).Value = CopyTilesSelectAllCheckBox.Checked
            CopyTilesDataGridView.Rows(RowIndex).Cells(3).Value = CopyTilesSelectAllCheckBox.Checked
            CopyTilesDataGridView.Rows(RowIndex).Cells(4).Value = CopyTilesSelectAllCheckBox.Checked
            CopyTilesDataGridView.Rows(RowIndex).Cells(5).Value = CopyTilesSelectAllCheckBox.Checked
        Next
    End Sub

    Private Sub CopyTilesStartButton_Click(sender As Object, e As EventArgs) Handles CopyTilesStartButton.Click
        Dim SourceFolderName As String = CopyTilesSourceTextBox.Text
        Dim DestinationFolderName As String = CopyTilesDestinationTextBox.Text
        Dim CopySource As String
        Dim CopyDestination As String
        Dim FileName As String
        Dim BaseName As String
        Dim Extension As String
        Dim TileName As String
        Dim Message As String
        Dim CreateLog As Boolean = CopyTilesLogCheckBox.Checked
        Dim LogFileName As String = "CopyTiles.log"
        Dim Indent As String = "    "

        If Not SourceFolderName.EndsWith("\") Then SourceFolderName &= "\"
        If Not DestinationFolderName.EndsWith("\") Then DestinationFolderName &= "\"

        Dim PatternsPath As String = DestinationFolderName & "textures\patterns\normal\"
        Dim TerrainPath As String = DestinationFolderName & "textures\terrain\"
        Dim SimpleTilePath As String = DestinationFolderName & "textures\tilesets\simple\"
        Dim SimpleTileDataPath As String = DestinationFolderName & "data\tilesets\"
        Dim SimpleTileBaseName As String
        Dim SimpleTileFileName As String
        Dim SimpleTileSetDataFile As String


        Dim Row As DataGridViewRow

        OutputForm.OutputTextBox.Text = ""
        OutputForm.Show()
        OutputForm.BringToFront()
        Message = "### Starting selected tiles at " & DateTime.Now & "." & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, False)

        For RowIndex = 1 To CopyTilesDataGridView.Rows.Count - 1
            Row = CopyTilesDataGridView.Rows(RowIndex)
            TileName = Row.Cells(2).Value
            If TileName <> "" Then
                CopySource = SourceFolderName & TileName

                For ColumnIndex = 3 To 5
                    Row.Cells(ColumnIndex).Value = Convert.ToBoolean(Row.Cells(ColumnIndex).EditedFormattedValue)
                Next

                FileName = Path.GetFileName(CopySource)
                BaseName = Path.GetFileNameWithoutExtension(CopySource)
                Extension = Path.GetExtension(CopySource)

                If Row.Cells("Patterns").Value Then
                    CopyDestination = PatternsPath & FileName
                    If My.Computer.FileSystem.FileExists(CopyDestination) Then
                        Message = Indent & "Destination file already exists: " & CopyDestination & vbCrLf
                        OutputForm.OutputTextBox.AppendText(Message)
                        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                    Else
                        If Not My.Computer.FileSystem.DirectoryExists(PatternsPath) Then My.Computer.FileSystem.CreateDirectory(PatternsPath)
                        Message = Indent & "Copying from: " & CopySource & vbCrLf
                        Message &= Indent & "          to: " & CopyDestination & vbCrLf
                        OutputForm.OutputTextBox.AppendText(Message)
                        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                        My.Computer.FileSystem.CopyFile(CopySource, CopyDestination)
                    End If
                End If

                If Row.Cells("Terrain").Value Then
                    CopyDestination = TerrainPath & BaseName & "_terrain" & Extension
                    If My.Computer.FileSystem.FileExists(CopyDestination) Then
                        Message = Indent & "Destination file already exists: " & CopyDestination & vbCrLf
                        OutputForm.OutputTextBox.AppendText(Message)
                        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                    Else
                        If Not My.Computer.FileSystem.DirectoryExists(TerrainPath) Then My.Computer.FileSystem.CreateDirectory(TerrainPath)
                        Message = Indent & "Copying from: " & CopySource & vbCrLf
                        Message &= Indent & "          to: " & CopyDestination & vbCrLf
                        OutputForm.OutputTextBox.AppendText(Message)
                        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                        My.Computer.FileSystem.CopyFile(CopySource, CopyDestination)
                    End If
                End If

                If Row.Cells("SimpleTile").Value Then
                    SimpleTileBaseName = BaseName & "_simple"
                    SimpleTileFileName = BaseName & Extension
                    CopyDestination = SimpleTilePath & BaseName & "_simple" & Extension

                    If My.Computer.FileSystem.FileExists(CopyDestination) Then
                        Message = Indent & "Destination file already exists: " & CopyDestination & vbCrLf
                        OutputForm.OutputTextBox.AppendText(Message)
                        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                    Else
                        If Not My.Computer.FileSystem.DirectoryExists(SimpleTilePath) Then My.Computer.FileSystem.CreateDirectory(SimpleTilePath)
                        Message = Indent & "Copying from: " & CopySource & vbCrLf
                        Message &= Indent & "          to: " & CopyDestination & vbCrLf
                        OutputForm.OutputTextBox.AppendText(Message)
                        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                        My.Computer.FileSystem.CopyFile(CopySource, CopyDestination)

                        Dim SimpleTileSetData As New System.Collections.Specialized.OrderedDictionary
                        SimpleTileSetData.Add("path", "textures/tilesets/simple/" & BaseName & "_simple" & Extension)
                        SimpleTileSetData.Add("name", BaseName)
                        SimpleTileSetData.Add("type", "normal")
                        SimpleTileSetData.Add("color", "ffffff")

                        If Not My.Computer.FileSystem.DirectoryExists(SimpleTileDataPath) Then My.Computer.FileSystem.CreateDirectory(SimpleTileDataPath)
                        SimpleTileSetDataFile = SimpleTileDataPath & SimpleTileBaseName & ".dungeondraft_tileset"
                        Dim JSONString As String = JsonConvert.SerializeObject(SimpleTileSetData, Formatting.Indented)
                        My.Computer.FileSystem.WriteAllText(SimpleTileSetDataFile, JSONString, False, System.Text.Encoding.ASCII)
                    End If
                End If
                End If
        Next
        Message = "### Finished selected tiles at " & DateTime.Now & "." & vbCrLf & vbCrLf
        OutputForm.OutputTextBox.AppendText(Message)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
    End Sub

    Public Sub CopyTilesDataGridView_CellClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs)
        'Dim opposite As Boolean
        'Check to ensure that the row CheckBox is clicked.
        'If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 2 Then
        If e.ColumnIndex <> 1 And e.ColumnIndex <> 2 Then

            'Reference the GridView Row.
            Dim row As DataGridViewRow = CopyTilesDataGridView.Rows(e.RowIndex)
            Dim toprow As DataGridViewRow = CopyTilesDataGridView.Rows(0)
            Dim checkrow As DataGridViewRow
            Dim checkcolumn As DataGridViewColumn

            'Set the CheckBox selection.
            'For column As Integer = 1 To 6
            row.Cells(e.ColumnIndex).Value = Convert.ToBoolean(row.Cells(e.ColumnIndex).EditedFormattedValue)
            If e.RowIndex = 0 And e.ColumnIndex = 0 Then
                For eachrow As Integer = 0 To CopyTilesDataGridView.Rows.Count - 1
                    checkrow = CopyTilesDataGridView.Rows(eachrow)
                    If checkrow.Cells(2).Value <> "" Then
                        checkrow.Cells(e.ColumnIndex).Value = row.Cells(e.ColumnIndex).Value
                        For eachcolumn As Integer = 3 To 5
                            checkcolumn = CopyTilesDataGridView.Columns(eachcolumn)
                            If row.Cells(2).Value <> "" Then checkrow.Cells(eachcolumn).Value = row.Cells(e.ColumnIndex).Value
                        Next

                    End If
                Next
            ElseIf e.RowIndex = 0 Then
                For eachrow As Integer = 1 To CopyTilesDataGridView.Rows.Count - 1
                    checkrow = CopyTilesDataGridView.Rows(eachrow)
                    If checkrow.Cells(2).Value <> "" Then checkrow.Cells(e.ColumnIndex).Value = row.Cells(e.ColumnIndex).Value
                Next
            ElseIf e.ColumnIndex = 0 Then
                For eachcolumn As Integer = 3 To 5
                    checkcolumn = CopyTilesDataGridView.Columns(eachcolumn)
                    If row.Cells(2).Value <> "" Then row.Cells(eachcolumn).Value = row.Cells(e.ColumnIndex).Value
                Next
            Else
                If Not row.Cells(e.ColumnIndex).Value Then
                    toprow.Cells(e.ColumnIndex).Value = False
                    row.Cells(0).Value = False
                End If
            End If

        End If
    End Sub

    '###### Data Files Group Box ######
    Private Sub DataFilesSourceTextBox_LostFocus(sender As Object, e As EventArgs) Handles DataFilesSourceTextBox.LostFocus
        DataFilesDataGridView.Rows.Clear()
        DataFilesDataGridView.Columns.Clear()

        Dim SourceFolderName As String
        Dim IsFolderNameValid As Boolean
        Dim DoesFolderExist As Boolean
        Dim SourceFolder As System.IO.DirectoryInfo

        SourceFolderName = DataFilesSourceTextBox.Text
        IsFolderNameValid = IsValidPathName(SourceFolderName)
        DoesFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If IsFolderNameValid And DoesFolderExist Then
            SourceFolder = New System.IO.DirectoryInfo(SourceFolderName)
            SourceFolderName.TrimEnd("\")
            GetTilesAndWalls(SourceFolderName)
        Else
            MsgBox("Source folder name is either invalid or does not exist.")
        End If
    End Sub

    Private Sub DataFilesSourceBrowseButton_Click(sender As Object, e As EventArgs) Handles DataFilesSourceBrowseButton.Click
        DataFilesDataGridView.Rows.Clear()
        DataFilesDataGridView.Columns.Clear()
        DataFilesSourceBrowserDialog.ShowDialog()
        DataFilesSourceTextBox.Text = DataFilesSourceBrowserDialog.SelectedPath

        Dim SourceFolderName As String
        Dim IsFolderNameValid As Boolean
        Dim DoesFolderExist As Boolean
        Dim SourceFolder As System.IO.DirectoryInfo

        SourceFolderName = DataFilesSourceTextBox.Text
        IsFolderNameValid = IsValidPathName(SourceFolderName)
        DoesFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If IsFolderNameValid And DoesFolderExist Then
            SourceFolder = New System.IO.DirectoryInfo(SourceFolderName)
            SourceFolderName.TrimEnd("\")
            GetTilesAndWalls(SourceFolderName)
        Else
            MsgBox("Source folder name is either invalid or does not exist.")
        End If
    End Sub

    Public Sub DataFilesDataGridView_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataFilesDataGridView.CellValueChanged
        If e.ColumnIndex = 2 Then
            Dim row As DataGridViewRow
            row = DataFilesDataGridView.Rows(e.RowIndex)
            If row.Cells(2).Value = "" Then row.Cells(2).Value = "normal"
        End If
    End Sub

    Public Sub DataFilesDataGridView_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataFilesDataGridView.CellDoubleClick
        If e.ColumnIndex = 3 Then
            If DataFilesColorDialog.ShowDialog() = DialogResult.OK Then
                Dim NewColor = DataFilesColorDialog.Color
                Dim R = Hex(NewColor.R)
                Dim G = Hex(NewColor.G)
                Dim B = Hex(NewColor.B)
                If R.Length = 1 Then R = "0" & R
                If G.Length = 1 Then G = "0" & G
                If B.Length = 1 Then B = "0" & B
                Dim HexColor = R & G & B
                Dim row As DataGridViewRow
                row = DataFilesDataGridView.Rows(e.RowIndex)
                row.Cells(3).Value = HexColor.ToLower
            End If
        End If
    End Sub

    Private Sub DataFilesStartButton_Click(sender As Object, e As EventArgs) Handles DataFilesStartButton.Click
        Dim DataFolder As String
        Dim DataFile As String
        Dim Indent As String = "    "
        Dim OutputMessage As String
        Dim LogFileName As String = "DataFiles.log"
        Dim CreateLog As Boolean = DataFilesLogCheckBox.Checked

        OutputForm.OutputTextBox.Text = ""
        OutputForm.Show()
        OutputMessage = "### Starting at " & DateTime.Now & vbCrLf
        OutputForm.OutputTextBox.AppendText(OutputMessage)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, OutputMessage, False)

        OutputMessage = Indent & "Deleting existing tileset and wall data files to eliminate strays." & vbCrLf
        OutputForm.OutputTextBox.AppendText(OutputMessage)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, OutputMessage, True)

        Dim DataFolderList() As String = {DataFilesSourceTextBox.Text & "\data\tilesets", DataFilesSourceTextBox.Text & "\data\walls"}
        For Each SubFolder In DataFolderList
            If Directory.Exists(SubFolder) Then
                For Each DataFile In Directory.GetFiles(SubFolder)
                    File.Delete(DataFile)
                Next
            End If
        Next

        For Each SubFolder In DataFolderList
            If Directory.Exists(SubFolder) Then
                OutputMessage = Indent & "Directory exists: " & SubFolder & vbCrLf
                OutputForm.OutputTextBox.AppendText(OutputMessage)
                If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, OutputMessage, True)
            Else
                OutputMessage = Indent & "Creating directory: " & SubFolder & vbCrLf
                OutputForm.OutputTextBox.AppendText(OutputMessage)
                If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, OutputMessage, True)
                Directory.CreateDirectory(SubFolder)
            End If
        Next

        For Each row As DataGridViewRow In DataFilesDataGridView.Rows
            Dim AssetInfo As New System.Collections.Specialized.OrderedDictionary
            If row.Cells(0).Value = "Wall" Then
                AssetInfo.Add("path", row.Cells(6).Value)
                AssetInfo.Add("color", row.Cells(3).Value)
            Else
                AssetInfo.Add("path", row.Cells(6).Value)
                AssetInfo.Add("name", row.Cells(1).Value)
                AssetInfo.Add("type", row.Cells(2).Value)
                AssetInfo.Add("color", row.Cells(3).Value)
            End If
            DataFolder = row.Cells(5).Value

            DataFile = DataFolder & "\" & row.Cells(4).Value

            OutputMessage = Indent & "Writing " & DataFile & vbCrLf
            OutputForm.OutputTextBox.AppendText(OutputMessage)
            If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, OutputMessage, True)
            Dim DataContent As String = JsonConvert.SerializeObject(AssetInfo, Formatting.Indented)
            My.Computer.FileSystem.WriteAllText(DataFile, DataContent, False, System.Text.Encoding.ASCII)
        Next
        OutputMessage = "### Finished at " & DateTime.Now
        OutputForm.OutputTextBox.AppendText(OutputMessage)
        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, OutputMessage, True)
    End Sub

    '###### Map Details Group Box ######
    Private Sub MapDetailsSourceTextBox_LostFocus(sender As Object, e As EventArgs) Handles MapDetailsSourceTextBox.LostFocus
        MapDetailsCheckedListBox.Items.Clear()
        Dim SourceFolderName As String
        Dim IsFolderNameValid As Boolean
        Dim DoesFolderExist As Boolean
        SourceFolderName = MapDetailsSourceTextBox.Text
        IsFolderNameValid = IsValidPathName(SourceFolderName)
        DoesFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If IsFolderNameValid And DoesFolderExist Then
            Dim SourceFiles = Directory.GetFiles(SourceFolderName, "*.dungeondraft_map")
            'Where Path.GetExtension(File).ToLower() = ".dungeondraft_map"

            For Each File In SourceFiles
                Dim FileName As New System.IO.FileInfo(File)
                MapDetailsCheckedListBox.Items.Add(FileName.Name)
                If MapDetailsSelectAllCheckBox.Checked Then MapDetailsCheckedListBox.SetItemChecked(MapDetailsCheckedListBox.Items.Count - 1, True)
            Next
        End If
    End Sub

    Private Sub MapDetailsBrowseButton_Click(sender As Object, e As EventArgs) Handles MapDetailsBrowseButton.Click
        MapDetailsCheckedListBox.Items.Clear()
        MapDetailsSourceBrowserDialog.ShowDialog()
        MapDetailsSourceTextBox.Text = MapDetailsSourceBrowserDialog.SelectedPath

        Dim SourceFolderName As String
        Dim IsFolderNameValid As Boolean
        Dim DoesFolderExist As Boolean
        SourceFolderName = MapDetailsSourceTextBox.Text
        IsFolderNameValid = IsValidPathName(SourceFolderName)
        DoesFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If IsFolderNameValid And DoesFolderExist Then
            Dim SourceFiles = Directory.GetFiles(SourceFolderName, "*.dungeondraft_map")
            For Each File In SourceFiles
                Dim FileName As New System.IO.FileInfo(File)
                MapDetailsCheckedListBox.Items.Add(FileName.Name)
                If MapDetailsSelectAllCheckBox.Checked Then MapDetailsCheckedListBox.SetItemChecked(MapDetailsCheckedListBox.Items.Count - 1, True)
            Next
        End If
    End Sub

    Private Sub MapDetailsSelectAllCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles MapDetailsSelectAllCheckBox.CheckedChanged
        If MapDetailsSelectAllCheckBox.Checked Then
            MapDetailsSelectAllButton_Click(sender, e)
        Else
            MapDetailsSelectNoneButton_Click(sender, e)
        End If
    End Sub

    Private Sub MapDetailsSelectAllButton_Click(sender As Object, e As EventArgs) Handles MapDetailsSelectAllButton.Click
        Dim Count As Integer
        For Count = 0 To MapDetailsCheckedListBox.Items.Count - 1
            MapDetailsCheckedListBox.SetItemChecked(Count, True)
        Next
    End Sub

    Private Sub MapDetailsSelectNoneButton_Click(sender As Object, e As EventArgs) Handles MapDetailsSelectNoneButton.Click
        Dim Count As Integer
        For Count = 0 To MapDetailsCheckedListBox.Items.Count - 1
            MapDetailsCheckedListBox.SetItemChecked(Count, False)
        Next
    End Sub

    Private Sub MapDetailsStartButton_Click(sender As Object, e As EventArgs) Handles MapDetailsStartButton.Click
        Dim SourceFolderName As String = MapDetailsSourceTextBox.Text
        Dim IsSourceFolderNameValid As String
        Dim DoesSourceFolderExist As Boolean
        Dim CreateLog As Boolean = MapDetailsLogCheckBox.Checked
        Dim LogFileName As String = "MapDetails.log"
        Dim MapSource As String
        Dim Indent As String = "    " '4 spaces
        Dim SubIndent As String = "        " '8 spaces
        Dim Message As String

        IsSourceFolderNameValid = IsValidPathName(SourceFolderName)
        DoesSourceFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If SourceFolderName <> "" And IsSourceFolderNameValid And DoesSourceFolderExist Then
            If MapDetailsCheckedListBox.CheckedItems.Count >= 1 Then
                OutputForm.OutputTextBox.Text = ""
                OutputForm.Show()
                OutputForm.BringToFront()
                Message = "### Starting selected map files at " & DateTime.Now & "." & vbCrLf & vbCrLf
                OutputForm.OutputTextBox.AppendText(Message)
                If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, False)
                Dim SelectedFiles = MapDetailsCheckedListBox.CheckedItems
                For Each MapFile In SelectedFiles
                    If SourceFolderName.EndsWith("\") Then
                        MapSource = SourceFolderName & MapFile
                    Else
                        MapSource = SourceFolderName & "\" & MapFile
                    End If
                    Message = Indent & "Starting " & MapFile & " at " & DateTime.Now & vbCrLf
                    OutputForm.OutputTextBox.AppendText(Message)
                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                    MapDetailsSub(MapSource, MapFile, CreateLog, LogFileName, SubIndent)
                    Message = Indent & "Finished " & MapFile & " at " & DateTime.Now & vbCrLf & vbCrLf
                    OutputForm.OutputTextBox.AppendText(Message)
                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                Next
                Message = "### Finished selected folders at " & DateTime.Now & "." & vbCrLf
                OutputForm.OutputTextBox.AppendText(Message)
                If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
            Else
                MsgBox("Nothing was selected.")
            End If
        Else
            MsgBox("Source folder name is either invalid or does not exist.")
        End If
    End Sub

    '###### Pack Assets Group Box ######
    Private Sub PackAssetsSourceTextBox_LostFocus(sender As Object, e As EventArgs) Handles PackAssetsSourceTextBox.LostFocus
        PackAssetsDataGridView.Rows.Clear()
        PackAssetsDataGridView.Columns.Clear()
        Dim UserFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim SourceFolderName As String
        Dim IsFolderNameValid As Boolean
        Dim DoesFolderExist As Boolean
        Dim SourceFolder As System.IO.DirectoryInfo
        Dim DestinationFolderName As String
        Dim CreateLog As Boolean = PackAssetsLogCheckBox.Checked
        Dim LogFileName As String = "PackAssets.log"
        Dim SubIndent As String = "        "
        Dim SelectAll As Boolean = PackAssetsSelectAllCheckBox.Checked

        SourceFolderName = PackAssetsSourceTextBox.Text
        IsFolderNameValid = IsValidPathName(SourceFolderName)
        DoesFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If IsFolderNameValid And DoesFolderExist Then
            SourceFolder = New System.IO.DirectoryInfo(SourceFolderName)
            DestinationFolderName = UserFolder & "\Dungeondraft"
            If PackAssetsDestinationTextBox.Text = "" Then PackAssetsDestinationTextBox.Text = DestinationFolderName
            LoadAssetFolders(SourceFolderName, CreateLog, LogFileName, SelectAll, SubIndent)
        End If
    End Sub

    Private Sub PackAssetsSourceBrowseButton_Click(sender As Object, e As EventArgs) Handles PackAssetsSourceBrowseButton.Click
        PackAssetsDataGridView.Rows.Clear()
        PackAssetsDataGridView.Columns.Clear()
        PackAssetsSourceBrowserDialog.ShowDialog()
        PackAssetsSourceTextBox.Text = PackAssetsSourceBrowserDialog.SelectedPath
        Dim UserFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim SourceFolderName As String
        Dim IsFolderNameValid As Boolean
        Dim DoesFolderExist As Boolean
        Dim SourceFolder As System.IO.DirectoryInfo
        Dim DestinationFolderName As String
        Dim CreateLog As Boolean = PackAssetsLogCheckBox.Checked
        Dim LogFileName As String = "PackAssets.log"
        Dim SubIndent As String = "        "
        Dim SelectAll As Boolean = PackAssetsSelectAllCheckBox.Checked

        SourceFolderName = PackAssetsSourceTextBox.Text
        IsFolderNameValid = IsValidPathName(SourceFolderName)
        DoesFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If IsFolderNameValid And DoesFolderExist Then
            SourceFolder = New System.IO.DirectoryInfo(SourceFolderName)
            DestinationFolderName = UserFolder & "\Dungeondraft\Repacked\" & SourceFolder.Name
            If PackAssetsDestinationTextBox.Text = "" Then PackAssetsDestinationTextBox.Text = DestinationFolderName
            LoadAssetFolders(SourceFolderName, CreateLog, LogFileName, SelectAll, SubIndent)
        End If
    End Sub

    Private Sub PackAssetsDestinationBrowseButton_Click(sender As Object, e As EventArgs) Handles PackAssetsDestinationBrowseButton.Click
        PackAssetsDestinationBrowserDialog.ShowDialog()
        PackAssetsDestinationTextBox.Text = PackAssetsDestinationBrowserDialog.SelectedPath
    End Sub

    Private Sub PackAssetsSelectAllCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles PackAssetsSelectAllCheckBox.CheckedChanged
        If PackAssetsSelectAllCheckBox.Checked Then
            PackAssetsSelectAllButton_Click(sender, e)
        Else
            PackAssetsSelectNoneButton_Click(sender, e)
        End If
    End Sub

    Private Sub PackAssetsRefreshButton_Click(sender As Object, e As EventArgs) Handles PackAssetsRefreshButton.Click
        PackAssetsDataGridView.Rows.Clear()
        PackAssetsDataGridView.Columns.Clear()
        Dim SourceFolderName As String
        Dim IsFolderNameValid As Boolean
        Dim DoesFolderExist As Boolean
        Dim CreateLog As Boolean = PackAssetsLogCheckBox.Checked
        Dim LogFileName As String = "PackAssets.log"
        Dim SubIndent As String = "        "
        Dim SelectAll As Boolean = PackAssetsSelectAllCheckBox.Checked

        SourceFolderName = PackAssetsSourceTextBox.Text
        IsFolderNameValid = IsValidPathName(SourceFolderName)
        DoesFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If IsFolderNameValid And DoesFolderExist Then
            LoadAssetFolders(SourceFolderName, CreateLog, LogFileName, SelectAll, SubIndent)
        End If
    End Sub
    Private Sub PackAssetsSelectAllButton_Click(sender As Object, e As EventArgs) Handles PackAssetsSelectAllButton.Click
        Dim RowIndex As Integer
        For RowIndex = 0 To PackAssetsDataGridView.Rows.Count - 1
            If PackAssetsDataGridView.Rows(RowIndex).Cells(1).Value <> "" Then
                PackAssetsDataGridView.Rows(RowIndex).Cells(0).Value = True
            End If
        Next
    End Sub

    Private Sub PackAssetsSelectNoneButton_Click(sender As Object, e As EventArgs) Handles PackAssetsSelectNoneButton.Click
        Dim RowIndex As Integer
        For RowIndex = 0 To PackAssetsDataGridView.Rows.Count - 1
            If PackAssetsDataGridView.Rows(RowIndex).Cells(1).Value <> "" Then
                PackAssetsDataGridView.Rows(RowIndex).Cells(0).Value = True
            End If
        Next
    End Sub

    Private Sub PackAssetsStartButton_Click(sender As Object, e As EventArgs) Handles PackAssetsStartButton.Click
        Dim PackEXE As String = "dungeondraft-pack.exe"
        Dim SourceFolderName As String = PackAssetsSourceTextBox.Text
        Dim DestinationFolderName As String = PackAssetsDestinationTextBox.Text
        Dim IsSourceFolderNameValid As String
        Dim DoesSourceFolderExist As Boolean
        Dim IsDestinationFolderNameValid As String
        Dim PackEXEexists As Boolean
        Dim Message As String
        Dim Indent As String = "    " '4 spaces
        Dim SubIndent As String = "        " '8 spaces
        Dim RowIndex As Integer
        Dim CreateLog As Boolean = PackAssetsLogCheckBox.Checked
        Dim LogFileName As String = "PackAssets.log"
        Dim Overwrite As Boolean = PackAssetsOverwriteCheckBox.Checked
        Dim SelectAll As Boolean = PackAssetsSelectAllCheckBox.Checked

        Dim PackSource As String
        Dim PackDestination As String

        Dim FolderName As String
        Dim PackName As String
        Dim Version As String
        Dim Author As String


        IsSourceFolderNameValid = IsValidPathName(SourceFolderName)
        DoesSourceFolderExist = System.IO.Directory.Exists(SourceFolderName)
        IsDestinationFolderNameValid = IsValidPathName(DestinationFolderName)

        PackEXEexists = My.Computer.FileSystem.FileExists(PackEXE)

        If PackEXEexists Then
            If SourceFolderName <> "" And IsSourceFolderNameValid And DoesSourceFolderExist Then
                If DestinationFolderName <> "" And IsDestinationFolderNameValid Then

                    If PackAssetsDataGridView.Rows.Count >= 1 Then
                        If Not SourceFolderName.EndsWith("\") Then SourceFolderName &= "\"
                        If Not DestinationFolderName.EndsWith("\") Then DestinationFolderName &= "\"
                        'If Not DestinationFolderName.EndsWith("\") Then DestinationFolderName &= "\"

                        OutputForm.OutputTextBox.Text = ""
                        OutputForm.Show()
                        OutputForm.BringToFront()
                        Message = "### Starting selected asset folders at " & DateTime.Now & "." & vbCrLf
                        If PackAssetsOverwriteCheckBox.Checked Then
                            Message &= Indent & "(Overwrite is enabled.)" & vbCrLf & vbCrLf
                        Else
                            Message &= Indent & "(Overwrite is disabled.)" & vbCrLf & vbCrLf
                        End If
                        OutputForm.OutputTextBox.AppendText(Message)
                        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, False)

                        If Not My.Computer.FileSystem.DirectoryExists(DestinationFolderName) Then My.Computer.FileSystem.CreateDirectory(DestinationFolderName)
                        For RowIndex = 0 To PackAssetsDataGridView.Rows.Count - 1
                            If PackAssetsDataGridView.Rows(RowIndex).Cells(0).Value And PackAssetsDataGridView.Rows(RowIndex).Cells(1).Value <> "" Then

                                FolderName = PackAssetsDataGridView.Rows(RowIndex).Cells(1).Value
                                PackName = PackAssetsDataGridView.Rows(RowIndex).Cells(2).Value & ".dungeondraft_pack"
                                Version = PackAssetsDataGridView.Rows(RowIndex).Cells(4).Value
                                Author = PackAssetsDataGridView.Rows(RowIndex).Cells(5).Value
                                If Author = "" Then Author = "<no author name>"

                                PackSource = SourceFolderName & FolderName
                                PackDestination = DestinationFolderName & PackName

                                If My.Computer.FileSystem.FileExists(PackDestination) And Not Overwrite Then
                                    Message = Indent & "Destination file already exists: " & PackAssetsDataGridView.Rows(RowIndex).Cells(1).Value & vbCrLf
                                    OutputForm.OutputTextBox.AppendText(Message)
                                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                                Else
                                    Message = Indent & "Packaging: " & FolderName & ", Version " & Version & ", by " & Author & vbCrLf
                                    Message &= Indent & "       to: " & PackDestination & vbCrLf
                                    OutputForm.OutputTextBox.AppendText(Message)
                                    If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                                    PackAssetsSub(PackSource, DestinationFolderName, RowIndex, CreateLog, LogFileName, SubIndent, PackEXE, Overwrite)
                                End If
                            End If
                        Next

                        Message = "### Finished selected asset folders at " & DateTime.Now & "." & vbCrLf
                        OutputForm.OutputTextBox.AppendText(Message)
                        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                        PackAssetsDataGridView.Rows.Clear()
                        PackAssetsDataGridView.Columns.Clear()
                        SourceFolderName = PackAssetsSourceTextBox.Text
                        LoadAssetFolders(SourceFolderName, CreateLog, LogFileName, SelectAll, SubIndent)
                    Else
                        MsgBox("Asset folder list is empty.")
                    End If

                Else
                    MsgBox("Destination folder name is invalid.")
                End If
            Else
                MsgBox("Source folder name is either invalid or does not exist.")
            End If
        Else
            MsgBox(PackEXE & " not found.")
        End If
    End Sub

    '###### Unpack Assets Group Box ######
    Private Sub UnpackAssetsSourceTextBox_LostFocus(sender As Object, e As EventArgs) Handles UnpackAssetsSourceTextBox.LostFocus
        UnpackAssetsCheckedListBox.Items.Clear()

        Dim UserFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim SourceFolderName As String
        Dim IsFolderNameValid As Boolean
        Dim DoesFolderExist As Boolean
        Dim SourceFolder As System.IO.DirectoryInfo
        Dim DestinationFolderName As String

        SourceFolderName = UnpackAssetsSourceTextBox.Text
        IsFolderNameValid = IsValidPathName(SourceFolderName)
        DoesFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If IsFolderNameValid And DoesFolderExist Then
            SourceFolder = New System.IO.DirectoryInfo(SourceFolderName)
            DestinationFolderName = UserFolder & "\Dungeondraft"
            If UnpackAssetsDestinationTextBox.Text = "" Then UnpackAssetsDestinationTextBox.Text = DestinationFolderName
            For Each PackFile As String In My.Computer.FileSystem.GetFiles(SourceFolderName)
                Dim PackName As New System.IO.DirectoryInfo(PackFile)
                If PackName.Extension = ".dungeondraft_pack" Then
                    UnpackAssetsCheckedListBox.Items.Add(PackName.Name)
                    If UnpackAssetsSelectAllCheckBox.Checked Then UnpackAssetsCheckedListBox.SetItemChecked(UnpackAssetsCheckedListBox.Items.Count - 1, True)
                End If
            Next
        End If
    End Sub

    Private Sub UnpackAssetsSourceBrowseButton_Click(sender As Object, e As EventArgs) Handles UnpackAssetsSourceBrowseButton.Click
        UnpackAssetsCheckedListBox.Items.Clear()
        UnpackAssetsSourceBrowserDialog.ShowDialog()
        UnpackAssetsSourceTextBox.Text = UnpackAssetsSourceBrowserDialog.SelectedPath

        Dim UserFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim SourceFolderName As String
        Dim IsFolderNameValid As Boolean
        Dim DoesFolderExist As Boolean
        Dim SourceFolder As System.IO.DirectoryInfo
        Dim DestinationFolderName As String

        SourceFolderName = UnpackAssetsSourceTextBox.Text
        IsFolderNameValid = IsValidPathName(SourceFolderName)
        DoesFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If IsFolderNameValid And DoesFolderExist Then
            SourceFolder = New System.IO.DirectoryInfo(SourceFolderName)
            DestinationFolderName = UserFolder & "\Dungeondraft\Unpacked\" & SourceFolder.Name
            If UnpackAssetsDestinationTextBox.Text = "" Then UnpackAssetsDestinationTextBox.Text = DestinationFolderName
            For Each PackFile As String In My.Computer.FileSystem.GetFiles(SourceFolderName)
                Dim PackName As New System.IO.DirectoryInfo(PackFile)
                If PackName.Extension = ".dungeondraft_pack" Then
                    UnpackAssetsCheckedListBox.Items.Add(PackName.Name)
                    If UnpackAssetsSelectAllCheckBox.Checked Then UnpackAssetsCheckedListBox.SetItemChecked(UnpackAssetsCheckedListBox.Items.Count - 1, True)
                End If
            Next
        Else
            MsgBox("Source folder name is either invalid or does not exist.")
        End If
    End Sub

    Private Sub UnpackAssetsDestinationBrowseButton_Click(sender As Object, e As EventArgs) Handles UnpackAssetsDestinationBrowseButton.Click
        UnpackAssetsDestinationBrowserDialog.ShowDialog()
        UnpackAssetsDestinationTextBox.Text = UnpackAssetsDestinationBrowserDialog.SelectedPath
    End Sub

    Private Sub UnpackAssetsSelectAllCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles UnpackAssetsSelectAllCheckBox.CheckedChanged
        If UnpackAssetsSelectAllCheckBox.Checked Then
            UnpackAssetsSelectAllButton_Click(sender, e)
        Else
            UnpackAssetsSelectNoneButton_Click(sender, e)
        End If
    End Sub

    Private Sub UnpackAssetsSelectAllButton_Click(sender As Object, e As EventArgs) Handles UnpackAssetsSelectAllButton.Click
        Dim Count As Integer
        For Count = 0 To UnpackAssetsCheckedListBox.Items.Count - 1
            UnpackAssetsCheckedListBox.SetItemChecked(Count, True)
        Next
    End Sub

    Private Sub UnpackAssetsSelectNoneButton_Click(sender As Object, e As EventArgs) Handles UnpackAssetsSelectNoneButton.Click
        Dim Count As Integer
        For Count = 0 To UnpackAssetsCheckedListBox.Items.Count - 1
            UnpackAssetsCheckedListBox.SetItemChecked(Count, False)
        Next
    End Sub

    Private Sub UnpackAssetsStartButton_Click(sender As Object, e As EventArgs) Handles UnpackAssetsStartButton.Click
        Dim UnpackEXE As String = "dungeondraft-unpack.exe"
        Dim SourceFolderName As String = UnpackAssetsSourceTextBox.Text
        Dim DestinationFolderName As String = UnpackAssetsDestinationTextBox.Text
        Dim IsSourceFolderNameValid As String
        Dim DoesSourceFolderExist As Boolean
        Dim IsDestinationFolderNameValid As String
        Dim UnpackEXEexists As Boolean
        Dim Message As String
        Dim Indent As String = "    " '4 spaces
        Dim PackBaseName As String
        Dim LogFileName As String = "UnpackAssets.log"
        Dim CreateLog As Boolean = UnpackAssetsLogCheckBox.Checked

        IsSourceFolderNameValid = IsValidPathName(SourceFolderName)
        DoesSourceFolderExist = System.IO.Directory.Exists(SourceFolderName)
        IsDestinationFolderNameValid = IsValidPathName(DestinationFolderName)

        UnpackEXEexists = My.Computer.FileSystem.FileExists(UnpackEXE)

        If UnpackEXEexists Then
            If SourceFolderName <> "" And IsSourceFolderNameValid And DoesSourceFolderExist Then
                If DestinationFolderName <> "" And IsDestinationFolderNameValid Then
                    If UnpackAssetsCheckedListBox.CheckedItems.Count >= 1 Then
                        OutputForm.OutputTextBox.Text = ""
                        OutputForm.Show()
                        OutputForm.BringToFront()
                        Message = "### Starting selected pack files at " & DateTime.Now & "." & vbCrLf
                        OutputForm.OutputTextBox.AppendText(Message)
                        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, False)
                        For Each PackFile As String In UnpackAssetsCheckedListBox.CheckedItems
                            PackBaseName = Path.GetFileNameWithoutExtension(SourceFolderName & "\" & PackFile)
                            Message = Indent & "Unpacking: " & SourceFolderName & "\" & PackFile & vbCrLf
                            Message &= Indent & "       to: " & DestinationFolderName & "\" & PackBaseName & vbCrLf
                            OutputForm.OutputTextBox.AppendText(Message)
                            If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                            Dim UnpackArguments As String = """" & SourceFolderName & "\" & PackFile & """ " & """" & DestinationFolderName & """"
                            Dim UnpackStartInfo As New ProcessStartInfo With {
                                .FileName = UnpackEXE,
                                .Arguments = UnpackArguments,
                                .RedirectStandardError = True,
                                .RedirectStandardOutput = True,
                                .UseShellExecute = False,
                                .CreateNoWindow = True
                            }
                            Dim Unpack As Process = Process.Start(UnpackStartInfo)
                            Unpack.WaitForExit()
                        Next
                        Message = "### Finished selected pack files at " & DateTime.Now & "." & vbCrLf
                        OutputForm.OutputTextBox.AppendText(Message)
                        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                    Else
                        MsgBox("Nothing was selected.")
                    End If

                Else
                    MsgBox("Destination folder name is invalid.")
                End If
            Else
                MsgBox("Source folder name is either invalid or does not exist.")
            End If
        Else
            MsgBox(UnpackEXE & " not found.")
        End If
    End Sub

    Public Class GlobalVariables
        Public Shared Property ConfigFileName As String = "DDTools.config"
        Public Shared Property ConfigFolderName As String = GetFolderPath(SpecialFolder.ApplicationData) & "\EightBitz\DDTools"
    End Class
End Class
