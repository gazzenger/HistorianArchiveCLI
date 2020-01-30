Imports System.IO

Module Module1
    ' Declared variables for iHistorian_SDK
    Public ConnectedServer As iHistorian_SDK.Server
    'Input Variables
    Public ParameterFlags(11, 1) As String
    Public ServerName As String = System.Net.Dns.GetHostName
    Public FilePath As String
    Public Username As String = Nothing
    Public Password As String = Nothing
    Public Datastore As String = Nothing
    Public Overwrite As Boolean = False
    'Other Variables
    Public ArchivePath As String = Nothing
    Public BackupPath As String = Nothing
    Public ProgramFilesDirs(1) As String
    Public ArchiveBackupApp As String = "C:\Program Files\Proficy\Proficy Historian\x86\Server\ihArchiveBackup.exe"
    Public RestoreAction As Boolean = False
    Public BackupAction As Boolean = False
    Public RemoveAction As Boolean = False
    Public ExistingArchiveFiles As Object

    Dim mydatastores As iHistorian_SDK.DataStores
    Dim mydatastore As iHistorian_SDK.DataStore
    Dim myarchives As iHistorian_SDK.Archives
    Dim myarchive As iHistorian_SDK.Archive = Nothing

    Sub Main(args As String())
        ParameterFlags = {
            {"-S", "The hostname of the server for connecting to"},
            {"-U", "The username for authenticating with historian (leave blank to use AD group)"},
            {"-P", "The password for authenticating with historian (leave blank to use AD group)"},
            {"--RESTORE", "Restoring archives listed in the file"},
            {"--BACKUP", "Backing up archives listed in the file"},
            {"--REMOVE", "Removing archives listed in the file"},
            {"-F", "The file path to the file containing the list of the archive paths to import for restoring (only IHAs at the moment), one item per line"},
            {"-O", "Overwrite any existing IHA files in the default archive path"},
            {"-D", "The specified datastore (leave blank to use the default datastore)"},
            {"-NC", "Skip the backing up of configuration file (IHC) before restoration"},
            {"-H", ""},
            {"-HELP", ""},
            {"--HELP", ""}
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
                        If (IsInArray(CStr(args(i + 1)), ParameterFlags, True) = -1) Then
                            ServerName = args(i + 1)
                        End If
                    End If
                Case ParameterFlags(1, 0) '-u
                    If (i < args.Count() - 1) Then
                        If (IsInArray(CStr(args(i + 1)), ParameterFlags, True) = -1) Then
                            Username = args(i + 1)
                        End If
                    End If
                Case ParameterFlags(2, 0) '-p
                    If (i < args.Count() - 1) Then
                        If (IsInArray(CStr(args(i + 1)), ParameterFlags, True) = -1) Then
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
                        If (IsInArray(CStr(args(i + 1)), ParameterFlags, True) = -1) Then
                            FilePath = args(i + 1)
                        End If
                    End If
                Case ParameterFlags(7, 0) '-o
                    Overwrite = True
                Case ParameterFlags(8, 0) '-d
                    If (i < args.Count() - 1) Then
                        If (IsInArray(CStr(args(i + 1)), ParameterFlags, True) = -1) Then
                            Datastore = args(i + 1)
                        End If
                    End If
                Case ParameterFlags(9, 0) '-nc
                    ArchiveBackupApp = ""
                Case ParameterFlags(10, 0), ParameterFlags(8, 0), ParameterFlags(9, 0) '-help
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

        'check a filelist is provided
        If String.IsNullOrEmpty(FilePath) Then
            Console.WriteLine("No file provided, please provide a file list, and use the -f flag")
            GoTo errc
        End If

        If Not String.IsNullOrEmpty(ArchiveBackupApp) And Not System.IO.File.Exists(ArchiveBackupApp) Then
            ArchiveBackupApp = ""
            ProgramFilesDirs = {
                Environment.ExpandEnvironmentVariables("%ProgramW6432%"),
                Environment.ExpandEnvironmentVariables("%ProgramFiles(x86)%")
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
                Console.WriteLine("Press any key to continue with the restoration")
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
        If String.IsNullOrEmpty(ArchivePath) Then
            Console.WriteLine("The archive path could not be found, please check you've provided the correct datastore name," +
                              " or that there is only 1 datastore on the server")
            GoTo errc
        End If

        'check the backup path
        If (BackupAction And String.IsNullOrEmpty(BackupPath)) Then
            Console.WriteLine("The backup path could not be found, please check you've provided the correct datastore name," +
                              " or that there is only 1 datastore on the server")
            GoTo errc
        End If

        'Test Archive file list
        If Not System.IO.File.Exists(FilePath) Then
            Console.WriteLine("The file " + FilePath + " doesn't exist")
            GoTo errc
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
        End If

        'check if backup or remove action are to be performed, if so get the list of files, and place in variable
        If (BackupAction Or RemoveAction) Then
            ReDim ExistingArchiveFiles(myarchives.Count)
            Dim i As Integer = 0
            For Each myarchive In myarchives.Item
                ExistingArchiveFiles(i) = myarchive.FileName.ToUpper()
                i = i + 1
            Next
        End If

        For Each Line As String In File.ReadLines(FilePath)
            If Not System.IO.File.Exists(Line) Then
                Console.WriteLine("The archive " + Line + " doesn't exist")
            Else
                If RestoreAction Then
                    RestoreArchive(Line)
                Else
                    If BackupAction Then BackupArchive(Line)
                    If RemoveAction Then RemoveArchive(Line)
                End If
            End If
        Next

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

    Private Function IsInArray(valToBeFound As Object, arr As Object, TwoDim As Boolean) As Integer
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
                If arr(i) = valToBeFound Then
                    IsInArray = i
                    Exit Function
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
            If Not ParameterFlags(i, 0) = "-F" Then
                ConcatStr += " [" + LCase(ParameterFlags(i, 0)) + "]"
            Else
                ConcatStr += " " + LCase(ParameterFlags(i, 0))
            End If
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
        'check if the archive path is inside the archive path (or if it needs to be copied)
        If fileName.Contains(ArchivePath) Then
            If Not IsFileInUse(ArchivePath + Path.GetFileName(fileName)) Then
                myarchive = myarchives.Add(Path.GetFileNameWithoutExtension(fileName), fileName, 0, Datastore)
            Else
                Console.WriteLine("The archive " + fileName + " is currently in use, and cannot be overwritten.")
            End If

        Else 'copy the file
            If Not Overwrite Then
                'check if a file already exists in the archive path
                If Not System.IO.File.Exists(ArchivePath + Path.GetFileName(fileName)) Then
                    System.IO.File.Copy(fileName, ArchivePath + Path.GetFileName(fileName))
                    myarchive = myarchives.Add(Path.GetFileNameWithoutExtension(fileName), ArchivePath + Path.GetFileName(fileName), 0, Datastore)
                    'check if file is a zip file

                    If (Path.GetExtension(fileName).ToUpper() = ".ZIP") Then
                        If System.IO.File.Exists(ArchivePath + Path.GetFileName(fileName)) Then
                            If Not IsFileInUse(ArchivePath + Path.GetFileName(fileName)) Then
                                System.IO.File.Delete(ArchivePath + Path.GetFileName(fileName))
                            End If
                        End If
                    End If

                Else
                    Console.WriteLine("File " + fileName + " already exists in the default archive path, " + ArchivePath)
                End If
            Else
                'check if the file is in use
                If Not IsFileInUse(ArchivePath + Path.GetFileName(fileName)) Then
                    System.IO.File.Copy(fileName, ArchivePath + Path.GetFileName(fileName), True)
                    myarchive = myarchives.Add(Path.GetFileNameWithoutExtension(fileName), ArchivePath + Path.GetFileName(fileName), 0, Datastore)
                Else
                    Console.WriteLine("The archive " + ArchivePath + Path.GetFileName(fileName) + " is currently in use, and cannot be overwritten.")
                End If
            End If
        End If

        If myarchive Is Nothing Then
            Console.WriteLine("An Error Occured Trying To Restore The Archive. " + fileName +
            " The Details Of The Error Follow: " + ConnectedServer.Archives.LastError)
        Else
            Console.WriteLine("The archive " + fileName + " has been successfully added")
        End If
        myarchive = Nothing
    End Sub

    'Backup Archive Function
    Sub BackupArchive(fileName As String)
        'Find the filename in the existing archive list
        Dim idx = IsInArray(CStr(fileName).ToUpper(), ExistingArchiveFiles, False)
        'if it does exist
        If (idx <> -1) Then
            'verify the archive filename directly
            myarchive = myarchives.Item(idx)
            'if the filename also matches
            If (myarchive.FileName.ToUpper() = CStr(fileName).ToUpper()) Then
                If myarchive.Backup(BackupPath + Path.GetFileNameWithoutExtension(fileName), Datastore) Then
                    Console.WriteLine("The archive " + fileName + " has been successfully backed up")
                Else
                    Console.WriteLine("The archive " + fileName + " was unable to be backed up")
                End If
            Else
                Console.WriteLine("The archive " + fileName + " does not exist on the server, or has encountered an error, try rerunning")
            End If
        Else
            Console.WriteLine("The archive " + fileName + " does not exist on the server")
        End If
    End Sub

    'Remove Archive Function
    Sub RemoveArchive(fileName As String)
        'Find the filename in the existing archive list
        Dim idx = IsInArray(CStr(fileName).ToUpper(), ExistingArchiveFiles, False)
        'if it does exist
        If (idx <> -1) Then
            'verify the archive filename directly
            myarchive = myarchives.Item(idx)
            'if the filename also matches
            If (myarchive.FileName.ToUpper() = CStr(fileName).ToUpper()) Then
                'check if the archive is current
                If (myarchive.IsCurrent) Then
                    Console.WriteLine("The archive " + fileName + " is set to current, and cannot be removed")
                Else
                    If myarchive.Delete(Datastore) Then
                        Console.WriteLine("The archive " + fileName + " has been successfully removed")
                    Else
                        Console.WriteLine("The archive " + fileName + " was unable to be removed")
                    End If
                End If
            Else
                Console.WriteLine("The archive " + fileName + " does not exist on the server, or has encountered an error, try rerunning")
            End If
        Else
            Console.WriteLine("The archive " + fileName + " does not exist on the server")
        End If
    End Sub

End Module
