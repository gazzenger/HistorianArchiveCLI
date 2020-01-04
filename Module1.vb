Imports System.IO

Module Module1
    ' Declared variables for iHistorian_SDK
    Public ConnectedServer As iHistorian_SDK.Server
    'Input Variables
    Public ParameterFlags(8, 1) As String
    Public ServerName As String = System.Net.Dns.GetHostName
    Public FilePath As String
    Public Username As String = Nothing
    Public Password As String = Nothing
    Public Datastore As String = Nothing
    Public Overwrite As Boolean = False
    'Other Variables
    Public ArchivePath As String = Nothing

    Sub Main(args As String())
        ParameterFlags = {
            {"-S", "The hostname of the server for connecting to"},
            {"-U", "The username for authenticating with historian (leave blank to use AD group)"},
            {"-P", "The password for authenticating with historian (leave blank to use AD group)"},
            {"-F", "The file path to the file containing the list of the archive paths to import for restoring (only IHAs at the moment), one item per line"},
            {"-O", "Overwrite any existing IHA files in the default archive path"},
            {"-D", "The specified datastore (leave blank to use the default datastore)"},
            {"-H", ""},
            {"-HELP", ""},
            {"--HELP", ""}
        }
        'Extract Arguments for application
        For i As Integer = 0 To args.Count() - 1
            Select Case UCase(args(i))
                Case ParameterFlags(0, 0) '-s
                    If (i < args.Count() - 1) Then
                        If Not IsInArray(CStr(args(i + 1)), ParameterFlags) Then
                            ServerName = args(i + 1)
                        End If
                    End If
                Case ParameterFlags(1, 0) '-u
                    If (i < args.Count() - 1) Then
                        If Not IsInArray(CStr(args(i + 1)), ParameterFlags) Then
                            Username = args(i + 1)
                        End If
                    End If
                Case ParameterFlags(2, 0) '-p
                    If (i < args.Count() - 1) Then
                        If Not IsInArray(CStr(args(i + 1)), ParameterFlags) Then
                            Password = args(i + 1)
                        End If
                    End If
                Case ParameterFlags(3, 0) '-f
                    If (i < args.Count() - 1) Then
                        If Not IsInArray(CStr(args(i + 1)), ParameterFlags) Then
                            FilePath = args(i + 1)
                        End If
                    End If
                Case ParameterFlags(4, 0) '-o
                    Overwrite = True
                Case ParameterFlags(5, 0) '-d
                    If (i < args.Count() - 1) Then
                        If Not IsInArray(CStr(args(i + 1)), ParameterFlags) Then
                            Datastore = args(i + 1)
                        End If
                    End If
                Case ParameterFlags(6, 0), ParameterFlags(7, 0), ParameterFlags(8, 0) '-help
                    OutputHelp()
                    GoTo errc
                Case Else
                    ' Do Nothing
            End Select
        Next

        'check for the Historian SDK DLL Depedencies
        If Not System.IO.File.Exists("%SYSTEMROOT%\SYSWOW64\iHSDK.dll") Then
            Console.WriteLine("The Historian SDK has not been installed, please run this from the Historian Install Disk under Client Tools")
            GoTo errc
        End If

        'check a filelist is provided
        If String.IsNullOrEmpty(FilePath) Then
            Console.WriteLine("No file provided, please provide a file list, and use the -f flag")
            GoTo errc
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
        Console.WriteLine("Connect completed in " & lEndTime - lStartTime & " seconds")
        If Not CheckConnection() Then
            Console.WriteLine("Connection to server " + ServerName + " is bad.")
        Else
            Console.WriteLine("Proficy Historian SDK Sample - Connected to: " + ServerName)
        End If

        'Check datastore archive path
        Dim mydatastores As iHistorian_SDK.DataStores
        Dim mydatastore As iHistorian_SDK.DataStore
        mydatastores = ConnectedServer.DataStores
        For Each mydatastore In mydatastores.Item    ' Iterate through each element.  
            If String.IsNullOrEmpty(Datastore) And mydatastore.IsDefault Then
                ArchivePath = mydatastore.Archives.ArchivingOptions("ArchiveDefaultPath")
            ElseIf mydatastore.Name = Datastore Then    ' If datastore name equals the datastore we're after
                ArchivePath = mydatastore.Archives.ArchivingOptions("ArchiveDefaultPath")
                Exit For    ' Exit loop. 
            End If
        Next

        'check the archive path
        If String.IsNullOrEmpty(ArchivePath) Then
            Console.WriteLine("The archive path could not be found, please check you've provided the correct datastore name," +
                              " or that there is only 1 datastore on the server")
            GoTo errc
        End If

        'Test Archive file list
        If Not System.IO.File.Exists(FilePath) Then
            Console.WriteLine("The file " + FilePath + " doesn't exist")
            GoTo errc
        End If

        'Perform a backup of the IHC file
        Console.WriteLine("Backing up IHC file before commencing restoration")
        Dim p As New Process
        Dim psi As New ProcessStartInfo("C:\Program Files\Proficy\Proficy Historian\x86\Server\ihArchiveBackup.exe", " -s " + ServerName + " -t 30 -c")
        psi.CreateNoWindow = False
        psi.UseShellExecute = False
        p.StartInfo = psi
        p.Start()
        p.WaitForExit()

        'Loop through each line of the file
        Dim myarchives As iHistorian_SDK.Archives
        Dim myarchive As iHistorian_SDK.Archive = Nothing
        myarchives = ConnectedServer.Archives
        For Each Line As String In File.ReadLines(FilePath)
            If Not System.IO.File.Exists(Line) Then
                Console.WriteLine("The archive " + Line + " doesn't exist")
            Else
                'check if the archive path is inside the archive path (or if it needs to be copied)
                If Line.Contains(ArchivePath) Then
                    If Not IsFileInUse(ArchivePath + Path.GetFileName(Line)) Then
                        myarchive = myarchives.Add(Path.GetFileNameWithoutExtension(Line), Line, 0, Datastore)
                    Else
                        Console.WriteLine("The archive " + Line + " is currently in use, and cannot be overwritten.")
                    End If

                Else 'copy the file

                    If Not Overwrite Then
                        'check if a file already exists in the archive path
                        If Not System.IO.File.Exists(ArchivePath + Path.GetFileName(Line)) Then
                            System.IO.File.Copy(Line, ArchivePath + Path.GetFileName(Line))
                            myarchive = myarchives.Add(Path.GetFileNameWithoutExtension(Line), ArchivePath + Path.GetFileName(Line), 0, Datastore)
                        Else
                            Console.WriteLine("File " + Line + " already exists in the default archive path, " + ArchivePath)
                        End If
                    Else
                        'check if the file is in use
                        If Not IsFileInUse(ArchivePath + Path.GetFileName(Line)) Then
                            System.IO.File.Copy(Line, ArchivePath + Path.GetFileName(Line), True)
                            myarchive = myarchives.Add(Path.GetFileNameWithoutExtension(Line), ArchivePath + Path.GetFileName(Line), 0, Datastore)
                        Else
                            Console.WriteLine("The archive " + ArchivePath + Path.GetFileName(Line) + " is currently in use, and cannot be overwritten.")
                        End If
                    End If
                End If

                If myarchive Is Nothing Then
                    Console.WriteLine("An Error Occured Trying To Restore The Archive. " + Line +
                    " The Details Of The Error Follow: " + ConnectedServer.Archives.LastError)
                Else
                    Console.WriteLine("The archive " + Line + " has been successfully added")
                End If
                myarchive = Nothing
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

    Private Function IsInArray(valToBeFound As Object, arr As Object) As Boolean
        'DEVELOPER: Ryan Wells (wellsr.com)
        'DESCRIPTION: Function to check if a value is in an array of values
        'INPUT: Pass the function a value to search for and an array of values of any data type.
        'OUTPUT: True if is in array, false otherwise
        Dim i As Integer
        On Error GoTo IsInArrayError 'array is empty

        For i = LBound(arr, 1) To UBound(arr, 1)
            If arr(i, 0) = valToBeFound Then
                IsInArray = True
                Exit Function
            End If
        Next i

        Exit Function
IsInArrayError:
        On Error GoTo 0
        IsInArray = False
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
End Module
