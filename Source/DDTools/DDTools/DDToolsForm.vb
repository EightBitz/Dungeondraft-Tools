Imports System.IO
Public Class DDToolsForm

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ConvertAssetsGroupBox.Hide()
        ConvertPacksGroupBox.Hide()
        CopyAssetsGroupBox.Hide()
        TagAssetsGroupBox.Hide()
        UnpackGroupBox.Hide()
        PackGroupBox.Hide()

        Me.Size = New Size(798, 245)
        TitlePanel.BringToFront()
        TitlePanel.Show()
    End Sub

    '###### Menu Items ######
    Private Sub ConvertAssetssMenuItem_Click(sender As Object, e As EventArgs) Handles ConvertAssetsMenuItem.Click
        TitlePanel.Hide()
        'DDConvertAssetsGroupBox.Hide()
        ConvertPacksGroupBox.Hide()
        CopyAssetsGroupBox.Hide()
        TagAssetsGroupBox.Hide()
        UnpackGroupBox.Hide()
        PackGroupBox.Hide()

        Me.Size = New Size(798, 649)
        ConvertAssetsGroupBox.BringToFront()
        ConvertAssetsGroupBox.Show()
    End Sub

    Private Sub ConvertPacksMenuItem_Click(sender As Object, e As EventArgs) Handles ConvertPacksMenuItem.Click
        TitlePanel.Hide()
        ConvertAssetsGroupBox.Hide()
        'DDConvertPacksGroupBox.Hide()
        CopyAssetsGroupBox.Hide()
        TagAssetsGroupBox.Hide()
        UnpackGroupBox.Hide()
        PackGroupBox.Hide()

        Me.Size = New Size(798, 649)
        ConvertPacksGroupBox.BringToFront()
        ConvertPacksGroupBox.Show()
    End Sub

    Private Sub CopyAssetsMenuItem_Click(sender As Object, e As EventArgs) Handles CopyAssetsMenuItem.Click
        TitlePanel.Hide()
        ConvertAssetsGroupBox.Hide()
        ConvertPacksGroupBox.Hide()
        'DDCopyAssetsGroupBox.Hide()
        TagAssetsGroupBox.Hide()
        UnpackGroupBox.Hide()
        PackGroupBox.Hide()

        Me.Size = New Size(798, 649)
        CopyAssetsGroupBox.BringToFront()
        CopyAssetsGroupBox.Show()
    End Sub

    Private Sub TagAssetsMenuItem_Click(sender As Object, e As EventArgs) Handles TagAssetsMenuItem.Click
        TitlePanel.Hide()
        ConvertAssetsGroupBox.Hide()
        ConvertPacksGroupBox.Hide()
        CopyAssetsGroupBox.Hide()
        'TagAssetsGroupBox.Hide()
        UnpackGroupBox.Hide()
        PackGroupBox.Hide()

        Me.Size = New Size(798, 649)
        TagAssetsGroupBox.BringToFront()
        TagAssetsGroupBox.Show()
    End Sub

    Private Sub DungeondraftunpackexeMenuItem_Click(sender As Object, e As EventArgs) Handles DungeondraftunpackexeMenuItem.Click
        TitlePanel.Hide()
        ConvertAssetsGroupBox.Hide()
        ConvertPacksGroupBox.Hide()
        CopyAssetsGroupBox.Hide()
        TagAssetsGroupBox.Hide()
        'DDUnpackGroupBox.Hide()
        PackGroupBox.Hide()

        Me.Size = New Size(798, 649)
        UnpackGroupBox.BringToFront()
        UnpackGroupBox.Show()
    End Sub

    Private Sub DungeondraftpackexeMenuItem_Click(sender As Object, e As EventArgs) Handles DungeondraftpackexeMenuItem.Click
        TitlePanel.Hide()
        ConvertAssetsGroupBox.Hide()
        ConvertPacksGroupBox.Hide()
        CopyAssetsGroupBox.Hide()
        TagAssetsGroupBox.Hide()
        UnpackGroupBox.Hide()
        'DDPackGroupBox.Hide()

        Me.Size = New Size(798, 649)
        PackGroupBox.BringToFront()
        PackGroupBox.Show()
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
            ConvertAssetsDestinationTextBox.Text = DestinationFolderName
            For Each AssetFolder As String In My.Computer.FileSystem.GetDirectories(SourceFolderName)
                Dim FolderName As New System.IO.DirectoryInfo(AssetFolder)
                ConvertAssetsCheckedListBox.Items.Add(FolderName.Name)
                ConvertAssetsCheckedListBox.SetItemChecked(ConvertAssetsCheckedListBox.Items.Count - 1, True)
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
            ConvertAssetsDestinationTextBox.Text = DestinationFolderName
            For Each AssetFolder As String In My.Computer.FileSystem.GetDirectories(SourceFolderName)
                Dim FolderName As New System.IO.DirectoryInfo(AssetFolder)
                ConvertAssetsCheckedListBox.Items.Add(FolderName.Name)
                ConvertAssetsCheckedListBox.SetItemChecked(ConvertAssetsCheckedListBox.Items.Count - 1, True)
            Next
        End If
    End Sub

    Private Sub ConvertAssetsDestinationBrowseButton_Click(sender As Object, e As EventArgs) Handles ConvertAssetsDestinationBrowseButton.Click
        ConvertAssetsDestinationBrowserDialog.ShowDialog()
        ConvertAssetsDestinationTextBox.Text = ConvertAssetsDestinationBrowserDialog.SelectedPath
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
        Dim CreateLog = CopyAssetsLogCheckBox.Checked
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
                        Message = "### Finished selected folders " & DateTime.Now & "." & vbCrLf
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
            DestinationFolderName = UserFolder & "\Dungeondraft"
            ConvertPacksDestinationTextBox.Text = DestinationFolderName
            For Each PackFile As String In My.Computer.FileSystem.GetFiles(SourceFolderName)
                Dim PackName As New System.IO.DirectoryInfo(PackFile)
                If PackName.Extension = ".dungeondraft_pack" Then
                    ConvertPacksCheckedListBox.Items.Add(PackName.Name)
                    ConvertPacksCheckedListBox.SetItemChecked(ConvertPacksCheckedListBox.Items.Count - 1, True)
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
            DestinationFolderName = UserFolder & "\Dungeondraft" & SourceFolder.Name
            ConvertPacksDestinationTextBox.Text = DestinationFolderName
            For Each PackFile As String In My.Computer.FileSystem.GetFiles(SourceFolderName)
                Dim PackName As New System.IO.DirectoryInfo(PackFile)
                If PackName.Extension = ".dungeondraft_pack" Then
                    ConvertPacksCheckedListBox.Items.Add(PackName.Name)
                    ConvertPacksCheckedListBox.SetItemChecked(ConvertPacksCheckedListBox.Items.Count - 1, True)
                End If
            Next
        End If
    End Sub

    Private Sub ConvertPacksDestinationBrowseButton_Click(sender As Object, e As EventArgs) Handles ConvertPacksDestinationBrowseButton.Click
        ConvertPacksDestinationBrowserDialog.ShowDialog()
        ConvertPacksDestinationTextBox.Text = ConvertPacksDestinationBrowserDialog.SelectedPath
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
        Dim CreateLog = CopyAssetsLogCheckBox.Checked
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

        If Not UnpackEXEexists And Not PackEXEexists And MagickExists Then
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
                        Message = "### Finished selected packs " & DateTime.Now & "." & vbCrLf
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
            CopyAssetsDestinationTextBox.Text = DestinationFolderName
            For Each AssetFolder As String In My.Computer.FileSystem.GetDirectories(SourceFolderName)
                Dim FolderName As New System.IO.DirectoryInfo(AssetFolder)
                CopyAssetsCheckedListBox.Items.Add(FolderName.Name)
                CopyAssetsCheckedListBox.SetItemChecked(CopyAssetsCheckedListBox.Items.Count - 1, True)
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
            CopyAssetsDestinationTextBox.Text = DestinationFolderName
            For Each AssetFolder As String In My.Computer.FileSystem.GetDirectories(SourceFolderName)
                Dim FolderName As New System.IO.DirectoryInfo(AssetFolder)
                CopyAssetsCheckedListBox.Items.Add(FolderName.Name)
                CopyAssetsCheckedListBox.SetItemChecked(CopyAssetsCheckedListBox.Items.Count - 1, True)
            Next
        End If
    End Sub

    Private Sub CopyAssetsDestinationBrowseButton_Click(sender As Object, e As EventArgs) Handles CopyAssetsDestinationBrowseButton.Click
        CopyAssetsDestinationBrowserDialog.ShowDialog()
        CopyAssetsDestinationTextBox.Text = CopyAssetsDestinationBrowserDialog.SelectedPath
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

        IsSourceFolderNameValid = IsValidPathName(SourceFolderName)
        DoesSourceFolderExist = System.IO.Directory.Exists(SourceFolderName)

        IsDestinationFolderNameValid = IsValidPathName(DestinationFolderName)

        If SourceFolderName <> "" And IsSourceFolderNameValid And DoesSourceFolderExist Then
            If DestinationFolderName <> "" And IsDestinationFolderNameValid Then
                If CopyAssetsCheckedListBox.CheckedItems.Count >= 1 Then
                    OutputForm.OutputTextBox.Text = ""
                    OutputForm.Show()
                    OutputForm.BringToFront()
                    Message = "### Starting selected folders " & DateTime.Now & "." & vbCrLf & vbCrLf
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
                        CopyAssetsSub(CopySource, CopyDestination, CreateTagFile, Portals, CreateLog, LogFileName, SubIndent)
                        Message = Indent & "Finished " & AssetFolder & " at " & DateTime.Now & vbCrLf & vbCrLf
                        OutputForm.OutputTextBox.AppendText(Message)
                        If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
                    Next
                    Message = "### Finished selected folders " & DateTime.Now & "." & vbCrLf
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

    '###### Tag Assets Group Box ######
    Private Sub TagAssetsTextBox_LostFocus(sender As Object, e As EventArgs) Handles TagAssetsSourceTextBox.LostFocus
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
                        TagAssetsCheckedListBox.SetItemChecked(TagAssetsCheckedListBox.Items.Count - 1, True)
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
                        TagAssetsCheckedListBox.SetItemChecked(TagAssetsCheckedListBox.Items.Count - 1, True)
                    End If
                Next
            End If
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
                Message = "### Starting selected folders " & DateTime.Now & "." & vbCrLf & vbCrLf
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
                Message = "### Finished selected folders " & DateTime.Now & "." & vbCrLf
                OutputForm.OutputTextBox.AppendText(Message)
                If CreateLog Then My.Computer.FileSystem.WriteAllText(LogFileName, Message, True)
            Else
                MsgBox("Nothing was selected.")
            End If
        Else
            MsgBox("Source folder name is either invalid or does not exist.")
        End If

    End Sub

    '###### Unpack Group Box ######
    Private Sub UnpackSourceTextBox_LostFocus(sender As Object, e As EventArgs) Handles UnpackSourceTextBox.LostFocus
        UnpackCheckedListBox.Items.Clear()

        Dim UserFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim SourceFolderName As String
        Dim IsFolderNameValid As Boolean
        Dim DoesFolderExist As Boolean
        Dim SourceFolder As System.IO.DirectoryInfo
        Dim DestinationFolderName As String

        SourceFolderName = UnpackSourceTextBox.Text
        IsFolderNameValid = IsValidPathName(SourceFolderName)
        DoesFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If IsFolderNameValid And DoesFolderExist Then
            SourceFolder = New System.IO.DirectoryInfo(SourceFolderName)
            DestinationFolderName = UserFolder & "\Dungeondraft"
            UnpackDestinationTextBox.Text = DestinationFolderName
            For Each PackFile As String In My.Computer.FileSystem.GetFiles(SourceFolderName)
                Dim PackName As New System.IO.DirectoryInfo(PackFile)
                If PackName.Extension = ".dungeondraft_pack" Then
                    UnpackCheckedListBox.Items.Add(PackName.Name)
                    UnpackCheckedListBox.SetItemChecked(UnpackCheckedListBox.Items.Count - 1, True)
                End If
            Next
        End If
    End Sub

    Private Sub UnpackSourceBrowseButton_Click(sender As Object, e As EventArgs) Handles UnpackSourceBrowseButton.Click
        UnpackCheckedListBox.Items.Clear()
        UnpackSourceBrowserDialog.ShowDialog()
        UnpackSourceTextBox.Text = UnpackSourceBrowserDialog.SelectedPath

        Dim UserFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim SourceFolderName As String
        Dim IsFolderNameValid As Boolean
        Dim DoesFolderExist As Boolean
        Dim SourceFolder As System.IO.DirectoryInfo
        Dim DestinationFolderName As String

        SourceFolderName = UnpackSourceTextBox.Text
        IsFolderNameValid = IsValidPathName(SourceFolderName)
        DoesFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If IsFolderNameValid And DoesFolderExist Then
            SourceFolder = New System.IO.DirectoryInfo(SourceFolderName)
            DestinationFolderName = UserFolder & "\Dungeondraft\Unpacked\" & SourceFolder.Name
            UnpackDestinationTextBox.Text = DestinationFolderName
            For Each PackFile As String In My.Computer.FileSystem.GetFiles(SourceFolderName)
                Dim PackName As New System.IO.DirectoryInfo(PackFile)
                If PackName.Extension = ".dungeondraft_pack" Then
                    UnpackCheckedListBox.Items.Add(PackName.Name)
                    UnpackCheckedListBox.SetItemChecked(UnpackCheckedListBox.Items.Count - 1, True)
                End If
            Next
        End If
    End Sub

    Private Sub UnpackDestinationBrowseButton_Click(sender As Object, e As EventArgs) Handles UnpackDestinationBrowseButton.Click
        UnpackDestinationBrowserDialog.ShowDialog()
        UnpackDestinationTextBox.Text = UnpackDestinationBrowserDialog.SelectedPath
    End Sub

    Private Sub UnpackSelectAllButton_Click(sender As Object, e As EventArgs) Handles UnpackSelectAllButton.Click
        Dim Count As Integer
        For Count = 0 To UnpackCheckedListBox.Items.Count - 1
            UnpackCheckedListBox.SetItemChecked(Count, True)
        Next
    End Sub

    Private Sub UnpackSelectNoneButton_Click(sender As Object, e As EventArgs) Handles UnpackSelectNoneButton.Click
        Dim Count As Integer
        For Count = 0 To UnpackCheckedListBox.Items.Count - 1
            UnpackCheckedListBox.SetItemChecked(Count, False)
        Next
    End Sub

    Private Sub UnpackStartButton_Click(sender As Object, e As EventArgs) Handles UnpackStartButton.Click
        Dim UnpackEXE As String = "dungeondraft-unpack.exe"
        Dim SourceFolderName As String = UnpackSourceTextBox.Text
        Dim DestinationFolderName As String = UnpackDestinationTextBox.Text
        Dim IsSourceFolderNameValid As String
        Dim DoesSourceFolderExist As Boolean
        Dim IsDestinationFolderNameValid As String
        Dim UnpackEXEexists As Boolean
        Dim Message As String
        Dim Indent As String = "    " '4 spaces
        Dim PackBaseName As String

        IsSourceFolderNameValid = IsValidPathName(SourceFolderName)
        DoesSourceFolderExist = System.IO.Directory.Exists(SourceFolderName)
        IsDestinationFolderNameValid = IsValidPathName(DestinationFolderName)

        UnpackEXEexists = My.Computer.FileSystem.FileExists(UnpackEXE)

        If UnpackEXEexists Then
            If SourceFolderName <> "" And IsSourceFolderNameValid And DoesSourceFolderExist Then
                If DestinationFolderName <> "" And IsDestinationFolderNameValid Then

                    If UnpackCheckedListBox.CheckedItems.Count >= 1 Then
                        OutputForm.OutputTextBox.Text = ""
                        OutputForm.Show()
                        OutputForm.BringToFront()
                        Message = "### Starting selected pack files " & DateTime.Now & "." & vbCrLf
                        OutputForm.OutputTextBox.AppendText(Message)
                        For Each PackFile As String In UnpackCheckedListBox.CheckedItems
                            PackBaseName = Path.GetFileNameWithoutExtension(SourceFolderName & "\" & PackFile)
                            Message = Indent & "Unpacking " & PackFile & " to " & DestinationFolderName & "\" & PackBaseName & vbCrLf
                            OutputForm.OutputTextBox.AppendText(Message)
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
                        Message = "### Finished selected pack files " & DateTime.Now & "." & vbCrLf
                        OutputForm.OutputTextBox.AppendText(Message)
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

    '###### Pack Group Box ######
    Private Sub PackSourceTextBox_LostFocus(sender As Object, e As EventArgs) Handles PackSourceTextBox.LostFocus
        PackCheckedListBox.Items.Clear()

        Dim UserFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim SourceFolderName As String
        Dim IsFolderNameValid As Boolean
        Dim DoesFolderExist As Boolean
        Dim SourceFolder As System.IO.DirectoryInfo
        Dim DestinationFolderName As String

        SourceFolderName = PackSourceTextBox.Text
        IsFolderNameValid = IsValidPathName(SourceFolderName)
        DoesFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If IsFolderNameValid And DoesFolderExist Then
            SourceFolder = New System.IO.DirectoryInfo(SourceFolderName)
            DestinationFolderName = UserFolder & "\Dungeondraft"
            PackDestinationTextBox.Text = DestinationFolderName
            For Each PackFolder As String In My.Computer.FileSystem.GetDirectories(SourceFolderName)
                Dim PackName As New System.IO.DirectoryInfo(PackFolder)
                Dim JSONFile As String = PackName.FullName & "\" & "pack.json"
                If My.Computer.FileSystem.FileExists(JSONFile) Then
                    PackCheckedListBox.Items.Add(PackName.Name)
                    PackCheckedListBox.SetItemChecked(PackCheckedListBox.Items.Count - 1, True)
                End If
            Next
        End If
    End Sub

    Private Sub PackSourceBrowseButton_Click(sender As Object, e As EventArgs) Handles PackSourceBrowseButton.Click
        PackCheckedListBox.Items.Clear()
        PackSourceBrowserDialog.ShowDialog()
        PackSourceTextBox.Text = PackSourceBrowserDialog.SelectedPath

        Dim UserFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim SourceFolderName As String
        Dim IsFolderNameValid As Boolean
        Dim DoesFolderExist As Boolean
        Dim SourceFolder As System.IO.DirectoryInfo
        Dim DestinationFolderName As String

        SourceFolderName = PackSourceTextBox.Text
        IsFolderNameValid = IsValidPathName(SourceFolderName)
        DoesFolderExist = System.IO.Directory.Exists(SourceFolderName)
        If IsFolderNameValid And DoesFolderExist Then
            SourceFolder = New System.IO.DirectoryInfo(SourceFolderName)
            DestinationFolderName = UserFolder & "\Dungeondraft\Repacked\" & SourceFolder.Name
            PackDestinationTextBox.Text = DestinationFolderName
            For Each PackFolder As String In My.Computer.FileSystem.GetDirectories(SourceFolderName)
                Dim PackName As New System.IO.DirectoryInfo(PackFolder)
                Dim JSONFile As String = PackName.FullName & "\" & "pack.json"
                If My.Computer.FileSystem.FileExists(JSONFile) Then
                    PackCheckedListBox.Items.Add(PackName.Name)
                    PackCheckedListBox.SetItemChecked(PackCheckedListBox.Items.Count - 1, True)
                End If
            Next
        End If
    End Sub

    Private Sub PackDestinationBrowseButton_Click(sender As Object, e As EventArgs) Handles PackDestinationBrowseButton.Click
        PackDestinationBrowserDialog.ShowDialog()
        PackDestinationTextBox.Text = PackDestinationBrowserDialog.SelectedPath
    End Sub

    Private Sub PackSelectAllButton_Click(sender As Object, e As EventArgs) Handles PackSelectAllButton.Click
        Dim Count As Integer
        For Count = 0 To PackCheckedListBox.Items.Count - 1
            PackCheckedListBox.SetItemChecked(Count, True)
        Next
    End Sub

    Private Sub PackSelectNoneButton_Click(sender As Object, e As EventArgs) Handles PackSelectNoneButton.Click
        Dim Count As Integer
        For Count = 0 To PackCheckedListBox.Items.Count - 1
            PackCheckedListBox.SetItemChecked(Count, False)
        Next
    End Sub

    Private Sub PackStartButton_Click(sender As Object, e As EventArgs) Handles PackStartButton.Click
        Dim PackEXE As String = "dungeondraft-pack.exe"
        Dim SourceFolderName As String = PackSourceTextBox.Text
        Dim DestinationFolderName As String = PackDestinationTextBox.Text
        Dim IsSourceFolderNameValid As String
        Dim DoesSourceFolderExist As Boolean
        Dim IsDestinationFolderNameValid As String
        Dim PackEXEexists As Boolean
        Dim Message As String
        Dim Indent As String = "    " '4 spaces

        IsSourceFolderNameValid = IsValidPathName(SourceFolderName)
        DoesSourceFolderExist = System.IO.Directory.Exists(SourceFolderName)
        IsDestinationFolderNameValid = IsValidPathName(DestinationFolderName)

        PackEXEexists = My.Computer.FileSystem.FileExists(PackEXE)

        If PackEXEexists Then
            If SourceFolderName <> "" And IsSourceFolderNameValid And DoesSourceFolderExist Then
                If DestinationFolderName <> "" And IsDestinationFolderNameValid Then

                    If PackCheckedListBox.CheckedItems.Count >= 1 Then
                        OutputForm.OutputTextBox.Text = ""
                        OutputForm.Show()
                        OutputForm.BringToFront()
                        Message = "### Starting selected asset folders " & DateTime.Now & "." & vbCrLf
                        OutputForm.OutputTextBox.AppendText(Message)
                        For Each AssetFolder As String In PackCheckedListBox.CheckedItems
                            Message = Indent & "Packing " & AssetFolder & " to " & DestinationFolderName & "\" & AssetFolder & ".dungeondraft_pack" & vbCrLf
                            OutputForm.OutputTextBox.AppendText(Message)
                            Dim PackArguments As String = """" & SourceFolderName & "\" & AssetFolder & """ " & """" & DestinationFolderName & """"
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
                        Next
                        Message = "### Finished selected asset folders " & DateTime.Now & "." & vbCrLf
                        OutputForm.OutputTextBox.AppendText(Message)
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
            MsgBox(PackEXE & " not found.")
        End If

    End Sub

End Class
