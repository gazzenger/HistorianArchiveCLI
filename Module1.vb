Imports System.IO
Imports System.IO.Compression
Module Module1
    ' Declared variables for iHistorian_SDK
    Public ConnectedServer As iHistorian_SDK.Server
    'Input Variables
    Public ParameterFlags(11, 2) As String
    Public ServerName As String = System.Net.Dns.GetHostName
    Public FilePath As String
    Public Username As String = Nothing
    Public Password As String = Nothing
    Public Datastore As String = Nothing
    Public Overwrite As Boolean = False
    Public TimeStamp As String = DateTime.Now.ToString("yyyyMMddhhmmss")
    'Other Variables
    Public ArchivePath As String = Nothing
    Public BackupPath As String = Nothing
    Public ProgramFilesDirs(1) As String
    Public ArchiveBackupApp As String = "C:\Program Files\Proficy\Proficy Historian\x86\Server\ihArchiveBackup.exe"
    Public RestoreAction As Boolean = False
    Public BackupAction As Boolean = False
    Public RemoveAction As Boolean = False

    Public OlderThan? As Integer = Nothing
    Public ExportPath As String = Nothing

    Public ExistingArchiveFiles As Object
    Public OldThanArchives As Object

    Dim mydatastores As iHistorian_SDK.DataStores
    Dim mydatastore As iHistorian_SDK.DataStore
    Dim myarchives As iHistorian_SDK.Archives
    Dim myarchive As iHistorian_SDK.Archive = Nothing

    Sub Main(args As String())
        ParameterFlags = {
            {"-S", "The hostname of the server for connecting to", " servername"},
            {"-U", "The username for authenticating with historian (leave blank to use AD group)", " username"},
            {"-P", "The password for authenticating with historian (leave blank to use AD group)", " password"},
            {"--RESTORE", "Restoring archives listed in the file", ""},
            {"--BACKUP", "Backing up archives, either from a file or based on an age", ""},
            {"--REMOVE", "Removing archives, either from a file or based on an age", ""},
            {"-F", "The file path to the file containing the list of the archive paths to import for restoring (only IHAs at the moment), one item per line", " filename"},
            {"-OLDERTHAN", "Archives older than a specified age (in days from today) to backup and/or remove", " days"},
            {"-EXPORTPATH", "The path to move backed up or removed archives", " path"},
            {"-O", "Overwrite any existing IHA files in the default archive path", ""},
            {"-D", "The specified datastore (leave blank to use the default datastore)", " datastore"},
            {"-NC", "Skip the backing up of configuration file (IHC) before restoration", ""},
            {"-H", "", ""},
            {"-HELP", "", ""},
            {"--HELP", "", ""}
        }
        'Extract Arguments for application
        If args.Count() = 0 Then
            OutputHelp()
            Threading.Thread.Sleep(2000)
            GoTo errc
        End If
        For i As Integer = 0 To args.Count() - 1

            Select Case UCase(args(i))
                Case ParameterFlags(0, 0) '-s
                    If (i < args.Count() - 1) Then
                        If (IsInArray(CStr(args(i + 1)), ParameterFlags, True, False) = -1) Then
                            ServerName = args(i + 1)
                        End If
                    End If
                Case ParameterFlags(1, 0) '-u
                    If (i < args.Count() - 1) Then
                        If (IsInArray(CStr(args(i + 1)), ParameterFlags, True, False) = -1) Then
                            Username = args(i + 1)
                        End If
                    End If
                Case ParameterFlags(2, 0) '-p
                    If (i < args.Count() - 1) Then
                        If (IsInArray(CStr(args(i + 1)), ParameterFlags, True, False) = -1) Then
                            Password = args(i + 1)
                        End If
                    End If
                Case ParameterFlags(3, 0) '--restore
                    RestoreAction = True
                Case ParameterFlags(4, 0) '--backup
                    BackupAction = True
                Case ParameterFlags(5, 0) '--remove
                    RemoveAction = True
                Case ParameterFlags(6, 0) '-f
                    If (i < args.Count() - 1) Then
                        If (IsInArray(CStr(args(i + 1)), ParameterFlags, True, False) = -1) Then
                            FilePath = args(i + 1)
                        End If
                    End If
                Case ParameterFlags(7, 0) '-olderthan
                    If (i < args.Count() - 1) Then
                        If (IsInArray(CStr(args(i + 1)), ParameterFlags, True, False) = -1) Then
                            OlderThan = 0
                            If Not Int32.TryParse(args(i + 1), OlderThan) Then
                                Console.WriteLine("The provided age using the -olderthan flag is invalid, this may be due to invalid characters")
                                GoTo errc
                            End If
                        End If
                    End If
                Case ParameterFlags(8, 0) '-exportpath
                    If (i < args.Count() - 1) Then
                        If (IsInArray(CStr(args(i + 1)), ParameterFlags, True, False) = -1) Then
                            ExportPath = args(i + 1)
                        End If
                    End If
                Case ParameterFlags(9, 0) '-o
                    Overwrite = True
                Case ParameterFlags(10, 0) '-d
                    If (i < args.Count() - 1) Then
                        If (IsInArray(CStr(args(i + 1)), ParameterFlags, True, False) = -1) Then
                            Datastore = args(i + 1)
                        End If
                    End If
                Case ParameterFlags(11, 0) '-nc
                    ArchiveBackupApp = ""
                Case ParameterFlags(12, 0), ParameterFlags(13, 0), ParameterFlags(14, 0) '-help
                    OutputHelp()
                    GoTo errc
                Case Else
                    ' Do Nothing
            End Select
        Next

        'check for the Historian SDK DLL Depedencies
        If (Not System.IO.File.Exists("C:\Windows\SYSWOW64\iHSDK.dll")) And (Not System.IO.File.Exists("C:\Windows\System32\iHSDK.dll")) Then
            Console.WriteLine("The Historian SDK has not been installed, please install this from the Historian Installation Disk under Client Tools")
            Console.WriteLine("The ISO image can be downloaded from the GE Website, https://digitalsupport.ge.com/en_US/Download/Historian-7-1")
            GoTo errc
        End If

        'check that if an action is given as a param, that only one param is specified
        If (BackupAction And RestoreAction) Or (RestoreAction And RemoveAction) Then
            Console.WriteLine("Cannot have more than one action in a single commmand - except backing up and removing archives at the same time")
            GoTo errc
        End If

        'check a filelist is provided a/o the olderthan flag
        If RestoreAction Then
            If String.IsNullOrEmpty(FilePath) Then
                If (OlderThan Is Nothing) Then
                    Console.WriteLine("No file provided, please provide a file list, and use the -f flag")
                    GoTo errc
                ElseIf OlderThan IsNot Nothing Then
                    Console.WriteLine("The -olderthan flag cannot be used with the --restore flag")
                    GoTo errc
                End If
            End If

            If Not String.IsNullOrEmpty(ExportPath) Then
                Console.WriteLine("The -exportpath flag cannot be used with the --restore flag")
                GoTo errc
            End If

        ElseIf BackupAction Or RemoveAction Then
            If String.IsNullOrEmpty(FilePath) Then
                If (OlderThan Is Nothing) Then
                    Console.WriteLine("No file provided, please provide a file list, and use the -f flag")
                    GoTo errc
                ElseIf OlderThan <= 0 Then
                    Console.WriteLine("The provided age using the -olderthan flag is invalid, typically due to age being zero or negative")
                    GoTo errc
                End If
            ElseIf (OlderThan IsNot Nothing) Then
                Console.WriteLine("The use of both the -f as well as the -olderthan flags at the same time is prohibited, please select just 1 of these flags")
                GoTo errc
            End If

            'check the export path
            If Not String.IsNullOrEmpty(ExportPath) Then
                If Not Directory.Exists(ExportPath) Then
                    Console.WriteLine("The export path " + ExportPath + " doesn't exist")
                    GoTo errc
                Else
                    'create sub directory for export
                    ExportPath = ExportPath + "\Backup-" + TimeStamp + "\"
                    Directory.CreateDirectory(ExportPath)
                End If
            End If

        End If

        'Replace default location of ArchiveBackupApp
        ArchiveBackupApp = ConvertToUNC(ArchiveBackupApp)
        If Not String.IsNullOrEmpty(ArchiveBackupApp) And Not System.IO.File.Exists(ArchiveBackupApp) Then
            ArchiveBackupApp = ""
            'Environment.ExpandEnvironmentVariables("%ProgramW6432%"),
            'Environment.ExpandEnvironmentVariables("%ProgramFiles(x86)%"),
            ProgramFilesDirs = {
                ConvertToUNC(Environment.ExpandEnvironmentVariables("%ProgramW6432%")),
                ConvertToUNC(Environment.ExpandEnvironmentVariables("%ProgramFiles(x86)%"))
            }
            Dim dirInfo
            Dim matches As New List(Of IO.FileInfo)
            For Each ProgramFileDir As String In ProgramFilesDirs
                dirInfo = New IO.DirectoryInfo(ProgramFileDir)
                SearchForFileInDirectory(matches, dirInfo, "ihArchiveBackup.exe")
                If matches.Count > 0 Then
                    ArchiveBackupApp = matches(0).FullName
                    Exit For
                End If
            Next
            If String.IsNullOrEmpty(ArchiveBackupApp) Then
                Console.WriteLine("Warning! No backup program detected for backing up the server configuration.")
                Console.WriteLine("Press any key to continue.")
                Console.ReadKey()
            End If
        End If

        'Establish connection to server
        On Error GoTo errc
        Dim lStartTime&, lEndTime&
        ' connect to the server
        ConnectedServer = New iHistorian_SDK.Server

        ' Measure the performance of the connect.
        ' Measuring connect time may be useful if you
        '   are using remote archivers and/or if you are
        '   using Domain security groups.

        lStartTime = Timer
        If Not ConnectedServer.Connect(ServerName, Username, Password) Then
            Console.WriteLine("Connect To Server: " + ServerName + " Failed.")
            GoTo errc
        End If
        lEndTime = Timer
        Console.WriteLine("Connection completed in " & lEndTime - lStartTime & " seconds")
        If Not CheckConnection() Then
            Console.WriteLine("Connection to server " + ServerName + " is bad.")
        Else
            Console.WriteLine("HistorianArchiveCLI - Connected to: " + ServerName)
        End If

        'Check datastore archive path, backup path, and loading archives
        mydatastores = ConnectedServer.DataStores
        For Each mydatastore In mydatastores.Item    ' Iterate through each element.  
            If String.IsNullOrEmpty(Datastore) And mydatastore.IsDefault Then
                ArchivePath = mydatastore.Archives.ArchivingOptions("ArchiveDefaultPath")
                BackupPath = mydatastore.Archives.ArchivingOptions("ArchiveBackupPath")
                myarchives = ConnectedServer.Archives
                Exit For
            ElseIf mydatastore.Name = Datastore Then    ' If datastore name equals the datastore we're after
                ArchivePath = mydatastore.Archives.ArchivingOptions("ArchiveDefaultPath")
                BackupPath = mydatastore.Archives.ArchivingOptions("ArchiveBackupPath")
                myarchives = mydatastore.Archives
                Exit For
            End If
        Next

        'check the archive path
        If String.IsNullOrEmpty(ConvertToUNC(ArchivePath)) Then
            Console.WriteLine("The archive path could not be found, please check you've provided the correct datastore name," +
                              " or that there is only 1 datastore on the server")
            GoTo errc
        End If

        'check the backup path
        If (BackupAction And String.IsNullOrEmpty(ConvertToUNC(BackupPath))) Then
            Console.WriteLine("The backup path could not be found, please check you've provided the correct datastore name," +
                              " or that there is only 1 datastore on the server")
            GoTo errc
        End If

        'Test Archive file list
        If (OlderThan Is Nothing) Then
            If Not System.IO.File.Exists(FilePath) Then
                Console.WriteLine("The file " + FilePath + " doesn't exist")
                GoTo errc
            End If
        End If

        'Perform a backup of the IHC file
        If Not String.IsNullOrEmpty(ArchiveBackupApp) Then
            Console.WriteLine("Backing up IHC file before commencing restoration")
            Dim p As New Process
            Dim psi As New ProcessStartInfo(ArchiveBackupApp, " -s " + ServerName + " -t 30 -c")
            psi.CreateNoWindow = False
            psi.UseShellExecute = False
            p.StartInfo = psi
            p.Start()
            p.WaitForExit()
            'rename the config backup with date/time stamp
            System.IO.File.Move(ConvertToUNC(BackupPath) + ServerName + "_Config.ihc", ConvertToUNC(BackupPath) + ServerName + "_Config" + TimeStamp + ".ihc")
            'move the config file over to the export path
            If Not String.IsNullOrEmpty(ExportPath) Then
                If ExportFile(ConvertToUNC(BackupPath) + ServerName + "_Config" + TimeStamp + ".ihc", ConvertToUNC(ExportPath), True) Then
                    Console.WriteLine("Successfully exported the IHC file " + ServerName + "_Config" + TimeStamp + ".ihc" + " to the export path " + ConvertToUNC(ExportPath))
                Else
                    Console.WriteLine("Failed to export the IHC file " + ServerName + "_Config" + TimeStamp + ".ihc" + " to the export path " + ConvertToUNC(ExportPath))
                End If
            End If
        End If

        'check if backup or remove action are to be performed, if so get the list of existing archives, and place in variable
        If (BackupAction Or RemoveAction) Then
            ReDim ExistingArchiveFiles(myarchives.Count)
            ReDim OldThanArchives(myarchives.Count)
            Dim i As Integer = 0
            Dim j As Integer = 0
            For Each myarchive In myarchives.Item
                ExistingArchiveFiles(i) = myarchive.FileName.ToUpper()
                'if older than flag is defined, check for archive end times to shortlist archives for backup a/o removal
                If (OlderThan IsNot Nothing) Then
                    If ((DateTime.Today - myarchive.EndTime).TotalDays >= OlderThan) Then
                        OldThanArchives(j) = myarchive.FileName.ToUpper()
                        j = j + 1
                    End If
                End If
                i = i + 1
            Next
        End If

        If (OlderThan IsNot Nothing) Then
            'cycle through each line in the input file
            For Each Line As String In OldThanArchives
                If Not String.IsNullOrEmpty(Line) Then
                    If Not System.IO.File.Exists(ConvertToUNC(Line)) Then
                        Console.WriteLine("The archive " + ConvertToUNC(Line) + " doesn't exist")
                    Else
                        If BackupAction Then BackupArchive(Line)
                        If RemoveAction Then RemoveArchive(Line)
                    End If
                End If
            Next
        ElseIf Not (String.IsNullOrEmpty(FilePath)) Then
            'cycle through each line in the input file
            For Each Line As String In File.ReadLines(FilePath)
                If Not System.IO.File.Exists(ConvertToUNC(Line)) Then
                    Console.WriteLine("The archive " + ConvertToUNC(Line) + " doesn't exist")
                Else
                    If RestoreAction Then
                        RestoreArchive(Line)
                    Else
                        If BackupAction Then BackupArchive(Line)
                        If RemoveAction Then RemoveArchive(Line)
                    End If
                End If
            Next
        End If

errc:
        If Err.Number Then Console.WriteLine(Err.Number)
        If Not ConnectedServer Is Nothing Then
            ConnectedServer.Disconnect()
            ConnectedServer = Nothing
        End If
        myarchives = Nothing
        mydatastores = Nothing
        Console.WriteLine("End.")

    End Sub

    ' make sure that we are connected to a server
    Public Function CheckConnection() As Boolean
        On Error GoTo errc
        If ConnectedServer Is Nothing Then
            CheckConnection = False
            Exit Function
        End If
        If Not ConnectedServer.Connected Then
            CheckConnection = False
            Exit Function
        End If
        If ConnectedServer.ServerTime < CDate("1/1/1970") Then
            CheckConnection = False
            Exit Function
        End If
        CheckConnection = True
        Exit Function
errc:
        CheckConnection = False
    End Function

    Private Function IsInArray(valToBeFound As Object, arr As Object, TwoDim As Boolean, ConvertUNC As Boolean) As Integer
        'DEVELOPER: Ryan Wells (wellsr.com)
        'DESCRIPTION: Function to check if a value is in an array of values
        'INPUT: Pass the function a value to search for and an array of values of any data type.
        'OUTPUT: True if is in array, false otherwise
        Dim i As Integer = -1
        On Error GoTo IsInArrayError 'array is empty
        IsInArray = -1

        For i = LBound(arr, 1) To UBound(arr, 1)
            If TwoDim Then
                If arr(i, 0) = valToBeFound Then
                    IsInArray = i
                    Exit Function
                End If
            Else
                If ConvertUNC Then
                    If ConvertToUNC(arr(i)) = ConvertToUNC(valToBeFound) Then
                        IsInArray = i
                        Exit Function
                    End If
                Else
                    If arr(i) = valToBeFound Then
                        IsInArray = i
                        Exit Function
                    End If
                End If
            End If
        Next i

        Exit Function
IsInArrayError:
        On Error GoTo 0
        IsInArray = -1
    End Function

    Public Function IsFileInUse(sFile As String) As Boolean
        Try
            Using f As New IO.FileStream(sFile, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
            End Using
        Catch Ex As Exception
            Return True
        End Try
        Return False
    End Function

    Public Sub OutputHelp()
        Console.WriteLine()
        Console.WriteLine(My.Application.Info.AssemblyName)
        Console.WriteLine("v " & My.Application.Info.Version.ToString)
        Console.WriteLine("By " + My.Application.Info.CompanyName)
        Console.WriteLine()
        Console.WriteLine(My.Application.Info.Description)
        Console.WriteLine()
        Dim i As Integer
        Dim ConcatStr As String = ""
        For i = LBound(ParameterFlags, 1) To UBound(ParameterFlags, 1) - 3
            ConcatStr += " [" + LCase(ParameterFlags(i, 0)) + LCase(ParameterFlags(i, 2)) + "]"
        Next i
        Console.WriteLine(Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetExecutingAssembly.GetModules()(0).FullyQualifiedName) +
            ConcatStr)
        Console.WriteLine()
        For i = LBound(ParameterFlags, 1) To UBound(ParameterFlags, 1) - 3
            Console.WriteLine(vbTab + LCase(ParameterFlags(i, 0)) + vbTab + ParameterFlags(i, 1))
        Next i
    End Sub

    Private Sub SearchForFileInDirectory(ByVal matches As List(Of IO.FileInfo), ByVal directory As IO.DirectoryInfo, ByVal fileName As String)
        Try
            For Each subDirectory As IO.DirectoryInfo In directory.GetDirectories()
                SearchForFileInDirectory(matches, subDirectory, fileName)
            Next
        Catch
            'MsgBox("Error iterating : " & directory.FullName)
        End Try
        Try
            For Each file As IO.FileInfo In directory.GetFiles()
                If file.Name.ToLower = fileName.ToLower() Then
                    matches.Add(file)
                End If
            Next
        Catch
            'MsgBox("Error iterating : " & directory.FullName)
        End Try
    End Sub

    'Restore Archive Function
    Sub RestoreArchive(fileName As String)
        If (Path.GetExtension(ConvertToUNC(fileName)).ToUpper() = ".ZIP") Then
            Console.WriteLine("The file " + ConvertToUNC(fileName) + " is a ZIP, uncompressing to Archive folder.")
            ZipFile.ExtractToDirectory(ConvertToUNC(fileName), ConvertToUNC(ArchivePath))
            'check the uncompressed file exists
            If System.IO.File.Exists(ConvertToUNC(ArchivePath) + Path.GetFileNameWithoutExtension(ConvertToUNC(fileName)) + ".iha") Then
                Console.WriteLine("The file has been uncompressed to " + ConvertToUNC(ArchivePath) + Path.GetFileNameWithoutExtension(ConvertToUNC(fileName)) + ".iha" + " successfully")
                fileName = ArchivePath + Path.GetFileNameWithoutExtension(ConvertToUNC(fileName)) + ".iha"
            Else
                Console.WriteLine("The file could not be uncompressed to " + ConvertToUNC(ArchivePath) + Path.GetFileNameWithoutExtension(ConvertToUNC(fileName)) + ".iha")
                myarchive = Nothing
                Exit Sub
            End If
        End If

        'check if the archive path is inside the archive path (or if it needs to be copied)
        If ConvertToUNC(fileName).Contains(ConvertToUNC(ArchivePath)) Then
            If Not IsFileInUse(ConvertToUNC(ArchivePath) + Path.GetFileName(ConvertToUNC(fileName))) Then
                myarchive = myarchives.Add(Path.GetFileNameWithoutExtension(ConvertToUNC(fileName)), fileName, 0, Datastore)
            Else
                Console.WriteLine("The archive " + ConvertToUNC(fileName) + " is currently in use, and cannot be overwritten.")
            End If

        Else 'copy the file
            If Not Overwrite Then
                'check if a file already exists in the archive path
                If Not System.IO.File.Exists(ConvertToUNC(ArchivePath) + Path.GetFileName(ConvertToUNC(fileName))) Then
                    System.IO.File.Copy(ConvertToUNC(fileName), ConvertToUNC(ArchivePath) + Path.GetFileName(ConvertToUNC(fileName)))
                    myarchive = myarchives.Add(Path.GetFileNameWithoutExtension(ConvertToUNC(fileName)), ArchivePath + Path.GetFileName(ConvertToUNC(fileName)), 0, Datastore)
                    'check if file is a zip file

                    If (Path.GetExtension(ConvertToUNC(fileName)).ToUpper() = ".ZIP") Then
                        If System.IO.File.Exists(ConvertToUNC(ArchivePath) + Path.GetFileName(ConvertToUNC(fileName))) Then
                            If Not IsFileInUse(ConvertToUNC(ArchivePath) + Path.GetFileName(ConvertToUNC(fileName))) Then
                                System.IO.File.Delete(ConvertToUNC(ArchivePath) + Path.GetFileName(ConvertToUNC(fileName)))
                            End If
                        End If
                    End If

                Else
                    Console.WriteLine("File " + ConvertToUNC(fileName) + " already exists in the default archive path, " + ConvertToUNC(ArchivePath))
                End If
            Else
                'check if the file is in use
                If Not IsFileInUse(ConvertToUNC(ArchivePath) + Path.GetFileName(ConvertToUNC(fileName))) Then
                    System.IO.File.Copy(ConvertToUNC(fileName), ConvertToUNC(ArchivePath) + Path.GetFileName(ConvertToUNC(fileName)), True)
                    myarchive = myarchives.Add(Path.GetFileNameWithoutExtension(ConvertToUNC(fileName)), ArchivePath + Path.GetFileName(ConvertToUNC(fileName)), 0, Datastore)
                Else
                    Console.WriteLine("The archive " + ConvertToUNC(ArchivePath) + Path.GetFileName(ConvertToUNC(fileName)) + " is currently in use, and cannot be overwritten.")
                End If
            End If
        End If

        If myarchive Is Nothing Then
            Console.WriteLine("An Error Occured Trying To Restore The Archive. " + ConvertToUNC(fileName) +
            " The Details Of The Error Follow: " + ConnectedServer.Archives.LastError)
        Else
            Console.WriteLine("The archive " + ConvertToUNC(fileName) + " has been successfully added")
        End If
        myarchive = Nothing
    End Sub

    'Backup Archive Function
    Sub BackupArchive(fileName As String)
        'Find the filename in the existing archive list
        Dim idx = IsInArray(CStr(fileName).ToUpper(), ExistingArchiveFiles, False, True)
        'if it does exist
        If (idx <> -1) Then
            'verify the archive filename directly
            myarchive = myarchives.Item(idx)
            'if the filename also matches
            If (myarchive.FileName.ToUpper() = CStr(fileName).ToUpper()) Then
                If myarchive.Backup(BackupPath + Path.GetFileNameWithoutExtension(ConvertToUNC(fileName)), Datastore) Then
                    Console.WriteLine("The archive " + ConvertToUNC(fileName) + " has been successfully backed up")

                    'check if exporting is required (moving files to another location) but not removing the archive
                    If Not String.IsNullOrEmpty(ExportPath) And Not RemoveAction Then
                        ExportFile(ConvertToUNC(BackupPath) + Path.GetFileNameWithoutExtension(ConvertToUNC(fileName)) + ".zip", ConvertToUNC(ExportPath), True)
                    End If

                Else
                    Console.WriteLine("The archive " + ConvertToUNC(fileName) + " was unable to be backed up")
                End If
            Else
                Console.WriteLine("The archive " + ConvertToUNC(fileName) + " does not exist on the server, or has encountered an error, try rerunning")
            End If
        Else
            Console.WriteLine("The archive " + ConvertToUNC(fileName) + " does not exist on the server")
        End If
    End Sub

    'Remove Archive Function
    Sub RemoveArchive(fileName As String)
        'Find the filename in the existing archive list
        Dim idx = IsInArray(CStr(fileName).ToUpper(), ExistingArchiveFiles, False, True)
        'if it does exist
        If (idx <> -1) Then
            'verify the archive filename directly
            myarchive = myarchives.Item(idx)
            'if the filename also matches
            If (myarchive.FileName.ToUpper() = CStr(fileName).ToUpper()) Then
                'check if the archive is current
                If (myarchive.IsCurrent) Then
                    Console.WriteLine("The archive " + ConvertToUNC(fileName) + " is set to current, and cannot be removed")
                Else
                    If myarchive.Delete(Datastore) Then
                        Console.WriteLine("The archive " + ConvertToUNC(fileName) + " has been successfully removed")
                        'check if exporting is required (moving files to another location)
                        If Not String.IsNullOrEmpty(ExportPath) Then
                            'move the backed up archive
                            If ExportFile(ConvertToUNC(BackupPath) + "\Offline\" + Path.GetFileNameWithoutExtension(ConvertToUNC(fileName)) + ".zip", ConvertToUNC(ExportPath), True) Then
                                'delete the iha file
                                System.IO.File.Delete(ConvertToUNC(fileName))
                            End If
                        End If
                    Else
                        Console.WriteLine("The archive " + ConvertToUNC(fileName) + " was unable to be removed")
                    End If
                End If
            Else
                Console.WriteLine("The archive " + ConvertToUNC(fileName) + " does not exist on the server, or has encountered an error, try rerunning")
            End If
        Else
            Console.WriteLine("The archive " + ConvertToUNC(fileName) + " does not exist on the server")
        End If
    End Sub

    Public Function ConvertToUNC(LocalPath As String) As String
        'check if already a UNC path
        If (LocalPath.Substring(0, 2) = "\\") Then
            Return LocalPath
        Else
            Return "\\" + ServerName + "\" + LocalPath.Replace(":", "$")
        End If
    End Function

    '
    'Input source and destination in UNC
    'Input is a filepath
    'Destination is a Folder path
    '
    Public Function ExportFile(Source As String, Destination As String, DeleteSourceOnSuccess As Boolean) As Boolean
        'verify the source file exists
        If (Not System.IO.File.Exists(Source)) Then
            Console.WriteLine("Cannot export file " + Source + " as the file does not exist.")
            Return False
        End If
        System.IO.File.Copy(Source, Destination + "\" + Path.GetFileName(ConvertToUNC(Source)))
        'verify new file
        If (Not System.IO.File.Exists(Destination + "\" + Path.GetFileName(ConvertToUNC(Source)))) Then
            Console.WriteLine("Export failed to copy the file " + Source + " to the destination " + Destination)
            Return False
        Else
            Console.WriteLine("Successfully exported the file " + Source + " to the destination " + Destination + "\" + Path.GetFileName(ConvertToUNC(Source)))
            If DeleteSourceOnSuccess Then
                System.IO.File.Delete(Source)
                Console.WriteLine("Successfully deleted the file " + Source)
            End If
            Return True
        End If
    End Function

End Module
