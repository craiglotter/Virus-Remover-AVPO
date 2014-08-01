Imports System.IO
Imports System.Threading
Imports System.ComponentModel
Imports System.Text
Imports Microsoft.Win32



Public Class Main_Screen

    Private busyworking As Boolean = False

    Private lastinputline As String = ""
    Private inputlines As Long = 0
    Private highestPercentageReached As Integer = 0
    Private inputlinesprecount As Long = 0
    Private pretestdone As Boolean = False
    Private primary_PercentComplete As Integer = 0
    Private percentComplete As Integer

    Private SelectedIndex As Integer = 0

    Private backupdirectory As String = ""
    Private savedirectory As String = ""

    Private AlertMessage As String = ""




    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message()
                Display_Message1.Message_Textbox.Text = "The Application encountered the following problem: " & vbCrLf & identifier_msg & ": " & ex.Message.ToString

                Display_Message1.Timer1.Interval = 1000
                Display_Message1.ShowDialog()
                Dim dir As System.IO.DirectoryInfo = New System.IO.DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                dir = Nothing
                Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg & ": " & ex.ToString)
                filewriter.WriteLine("")
                filewriter.Flush()
                filewriter.Close()
                filewriter = Nothing
            End If
            ex = Nothing
            identifier_msg = Nothing
        Catch exc As Exception
            MsgBox("An error occurred in the application's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub Activity_Handler(Optional ByVal identifier_msg As String = "")
        Try
            Dim dir As System.IO.DirectoryInfo = New System.IO.DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            dir = Nothing
            Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & Format(Now(), "yyyyMMdd") & "_Activity_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg)
            filewriter.Flush()
            filewriter.Close()
            filewriter = Nothing
            identifier_msg = Nothing
        Catch ex As Exception
            Error_Handler(ex, "Activity Logger")
        End Try
    End Sub


    Private Sub cancelAsyncButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cancelAsyncButton.Click

        ' Cancel the asynchronous operation.
        Me.BackgroundWorker1.CancelAsync()

        ' Disable the Cancel button.
        cancelAsyncButton.Enabled = False
        sender = Nothing
        e = Nothing
    End Sub 'cancelAsyncButton_Click



    Private Sub startAsyncButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles startAsyncButton.Click
        Try
            If busyworking = False Then
            

                busyworking = True


                inputlines = 0
                lastinputline = ""
                highestPercentageReached = 0
                inputlinesprecount = 0

                backupdirectory = ""
                savedirectory = ""
                pretestdone = False
                startAsyncButton.Enabled = False
                cancelAsyncButton.Enabled = True
                ' Start the asynchronous operation.
                AlertMessage = ""

                BackgroundWorker1.RunWorkerAsync()
            
            End If
        Catch ex As Exception
            Error_Handler(ex, "StartWorker")
        End Try
    End Sub

    ' This event handler is where the actual work is done.
    Private Sub backgroundWorker1_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        ' Get the BackgroundWorker object that raised this event.
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)

        ' Assign the result of the computation
        ' to the Result property of the DoWorkEventArgs
        ' object. This is will be available to the 
        ' RunWorkerCompleted eventhandler.
        e.Result = MainWorkerFunction(worker, e)
        sender = Nothing
        e = Nothing
        worker.Dispose()
        worker = Nothing
    End Sub 'backgroundWorker1_DoWork

    ' This event handler deals with the results of the
    ' background operation.
    Private Sub backgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        busyworking = False


        ' First, handle the case where an exception was thrown.
        If Not (e.Error Is Nothing) Then
            Error_Handler(e.Error, "backgroundWorker1_RunWorkerCompleted")
        ElseIf e.Cancelled Then
            ' Next, handle the case where the user canceled the 
            ' operation.
            ' Note that due to a race condition in 
            ' the DoWork event handler, the Cancelled
            ' flag may not have been set, even though
            ' CancelAsync was called.
            Me.ToolStripStatusLabel1.Text = "Operation Cancelled" & "   (" & inputlines & " of " & inputlinesprecount & ")"
            Me.ProgressBar1.Value = 0

        Else
            ' Finally, handle the case where the operation succeeded.
            Me.ToolStripStatusLabel1.Text = "Operation Completed" & "   (" & inputlines & " of " & inputlinesprecount & ")"
            Me.ProgressBar1.Value = 100
            If AlertMessage.Length > 0 Then
                'MsgBox("The following alerts were raised during the operation. If you wish to save these alerts, press Ctrl+C and paste it into NotePad." & vbCrLf & vbCrLf & "********************" & vbCrLf & vbCrLf & AlertMessage, MsgBoxStyle.Information, "Raised Alerts")
                'MsgBox(AlertMessage & " copies of 7za.exe were distributed", MsgBoxStyle.Information, "Copies Distributed")
            End If
        End If

       
        startAsyncButton.Enabled = True
        cancelAsyncButton.Enabled = False

        sender = Nothing
        e = Nothing


    End Sub 'backgroundWorker1_RunWorkerCompleted

    Private Sub backgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged


        Me.ProgressBar1.Value = e.ProgressPercentage
        'If lastinputline.StartsWith("Operation Completed") Then
        'Me.ToolStripStatusLabel1.Text = lastinputline
        'Else
        Me.ToolStripStatusLabel1.Text = lastinputline & "   (" & inputlines & " of " & inputlinesprecount & ")"
        'End If


        sender = Nothing
        e = Nothing
    End Sub

    Private Sub PreCount_Function(ByVal worker As BackgroundWorker)
        Try
            inputlinesprecount = 8
            inputlines = 8
            worker.ReportProgress(0)
        Catch ex As Exception
            Error_Handler(ex, "PreCount_Function")
        End Try
    End Sub

    Function MainWorkerFunction(ByVal worker As BackgroundWorker, ByVal e As DoWorkEventArgs) As Boolean
        Dim result As Boolean = False
        Try
            If Me.pretestdone = False Then
                primary_PercentComplete = 0
                worker.ReportProgress(0)
                PreCount_Function(worker)
                Me.pretestdone = True
            End If

            If worker.CancellationPending Then
                e.Cancel = True
                Return False
            End If

            primary_PercentComplete = 0
            worker.ReportProgress(0)

            inputlines = 0
            Activity_Handler("***********************************************************")
            Activity_Handler("Launching System Scan and Removal Operation")
            Activity_Handler("")

            '************************************************************'
            'Kill Processes
            '************************************************************'
            Try
                lastinputline = "Killing Processes"
                 PercentageCalculation(worker)
                RemoveProcess("wscript")
                inputlines = inputlines + 1
                lastinputline = "Processes Killed"
                 PercentageCalculation(worker)
            Catch ex As Exception
                Error_Handler(ex, "Kill Processes")
            End Try
            '************************************************************'
            'End Kill Processes
            '************************************************************'

            '************************************************************'
            'Update Windows 'Hidden' Registry Entry
            '************************************************************'
            Try
                lastinputline = "Updating Windows 'Hidden' Registry Entry"
                PercentageCalculation(worker)
                Dim key As RegistryKey
                Try
                    key = Registry.CurrentUser.OpenSubKey("Software", False)
                    key = key.OpenSubKey("Microsoft")
                    key = key.OpenSubKey("Windows")
                    key = key.OpenSubKey("CurrentVersion")
                    key = key.OpenSubKey("Explorer")
                    key = key.OpenSubKey("Advanced", True)
                    key.SetValue("Hidden", 1, RegistryValueKind.DWord)
                    key.Close()
                Catch ex As Exception
                    Error_Handler(ex, "Updating Windows 'Hidden' Registry Entry")
                End Try
                inputlines = inputlines + 1
                lastinputline = "Windows 'Hidden' Registry Entry Updated"
                PercentageCalculation(worker)
            Catch ex As Exception
                Error_Handler(ex, "Updating Windows 'Hidden' Registry Entry")
            End Try
            '************************************************************'
            'End Update Windows 'Hidden' Registry Entry
            '************************************************************'

            '************************************************************'
            'Remove Autorun.inf
            '************************************************************'
            Try
                lastinputline = "Searching and Removing Autorun.inf"
                PercentageCalculation(worker)
                Dim runner As IEnumerator
                Dim fso As New Scripting.FileSystemObject
                runner = fso.Drives.GetEnumerator()
                While runner.MoveNext() = True
                    Dim d As Scripting.Drive
                    d = runner.Current()
                    Select Case d.DriveType
                        Case Scripting.DriveTypeConst.Fixed
                            RemoveFILE((d.DriveLetter & ":\autorun.inf").Replace("\\", "\"))
                        Case Scripting.DriveTypeConst.Removable
                            RemoveFILE((d.DriveLetter & ":\autorun.inf").Replace("\\", "\"))
                        Case Scripting.DriveTypeConst.Remote
                            RemoveFILE((d.DriveLetter & ":\autorun.inf").Replace("\\", "\"))
                    End Select
                    d = Nothing
                End While
                runner = Nothing
                inputlines = inputlines + 1
                lastinputline = "Autorun.inf instances Removed"
                PercentageCalculation(worker)
            Catch ex As Exception
                Error_Handler(ex, "Remove Autorun.inf")
            End Try
            '************************************************************'
            'End Remove Autorun.inf
            '************************************************************'

            '************************************************************'
            'Remove Avpo*
            '************************************************************'
            Try
                lastinputline = "Searching and Removing Avpo*"
                PercentageCalculation(worker)
                Dim runner As IEnumerator
                Dim fso As New Scripting.FileSystemObject
                runner = fso.Drives.GetEnumerator()
                While runner.MoveNext() = True
                    Dim d As Scripting.Drive
                    d = runner.Current()
                    If d.DriveType = Scripting.DriveTypeConst.Fixed Or d.DriveType = Scripting.DriveTypeConst.Removable Or d.DriveType = Scripting.DriveTypeConst.Remote Then
                        Try
                            If My.Computer.FileSystem.DirectoryExists((d.DriveLetter & ":\Windows\system32").Replace("\\", "\")) Then
                                Dim dinfo As DirectoryInfo = New DirectoryInfo((d.DriveLetter & ":\Windows\system32").Replace("\\", "\"))
                                For Each finfo As FileInfo In dinfo.GetFiles("avp*.*")
                                    If finfo.Name.ToLower.StartsWith("avp") Then
                                        Dim fname As String = finfo.FullName
                                        finfo = Nothing
                                        Select Case fname.ToLower.Substring(fname.Length - 3, 3)
                                            Case "exe"
                                                RemoveEXE(fname)
                                            Case "dll"
                                                RemoveDLL(fname)
                                            Case Else
                                                RemoveFILE(fname)
                                        End Select
                                    End If
                                Next
                            End If
                        Catch ex As Exception
                            Error_Handler(ex, "Remove Avpo*")
                        End Try
                    End If
                    d = Nothing
                End While
                runner = Nothing
                inputlines = inputlines + 1
                lastinputline = "Avpo* instances Removed"
                PercentageCalculation(worker)
            Catch ex As Exception
                Error_Handler(ex, "Remove Avpo*")
            End Try
            '************************************************************'
            'End Remove Avpo*
            '************************************************************'


            '************************************************************'
            'Remove ntde1ect*
            '************************************************************'
            Try
                lastinputline = "Searching and Removing ntde1ect*"
                PercentageCalculation(worker)
                Dim runner As IEnumerator
                Dim fso As New Scripting.FileSystemObject
                runner = fso.Drives.GetEnumerator()
                While runner.MoveNext() = True
                    Dim d As Scripting.Drive
                    d = runner.Current()
                    If d.DriveType = Scripting.DriveTypeConst.Fixed Or d.DriveType = Scripting.DriveTypeConst.Removable Or d.DriveType = Scripting.DriveTypeConst.Remote Then
                        Try
                            If My.Computer.FileSystem.DirectoryExists((d.DriveLetter & ":\").Replace("\\", "\")) Then
                                Dim dinfo As DirectoryInfo = New DirectoryInfo((d.DriveLetter & ":\").Replace("\\", "\"))
                                For Each finfo As FileInfo In dinfo.GetFiles("ntde1ect.*")
                                    If finfo.Name.ToLower.StartsWith("ntde1ect") Then
                                        Dim fname As String = finfo.FullName
                                        finfo = Nothing
                                        Select Case fname.ToLower.Substring(fname.Length - 3, 3)
                                            Case "exe"
                                                RemoveEXE(fname)
                                            Case "dll"
                                                RemoveDLL(fname)
                                            Case Else
                                                RemoveFILE(fname)
                                        End Select
                                    End If
                                Next
                            End If
                        Catch ex As Exception
                            Error_Handler(ex, "Remove ntde1ect*")
                        End Try
                    End If
                    d = Nothing
                End While
                runner = Nothing
                inputlines = inputlines + 1
                lastinputline = "ntde1ect* instances Removed"
                 PercentageCalculation(worker)
            Catch ex As Exception
                Error_Handler(ex, "Remove ntde1ect*")
            End Try
            '************************************************************'
            'End Remove ntde1ect
            '************************************************************'


            '************************************************************'
            'Remove Avp0* Registry Entries
            '************************************************************'
            Try
                lastinputline = "Searching and Removing Avp0* Registry Entries"
                PercentageCalculation(worker)
                Dim key As RegistryKey
                Try
                    key = Registry.CurrentUser.OpenSubKey("Software", False)
                    key = key.OpenSubKey("Microsoft")
                    key = key.OpenSubKey("Windows")
                    key = key.OpenSubKey("CurrentVersion")
                    key = key.OpenSubKey("Run", True)
                    For Each val As String In key.GetValueNames
                        If val.ToLower.StartsWith("avp0") Or val.ToLower.StartsWith("avpo") Then
                            key.DeleteValue(val)
                            Activity_Handler("Removed " & key.ToString & "\" & val)
                        End If
                    Next
                    key.Close()
                Catch ex As Exception
                    Error_Handler(ex, "Remove Avp0* Registry Entry")
                End Try
                Try
                    key = Registry.CurrentUser.OpenSubKey("Software", False)
                    key = key.OpenSubKey("Microsoft")
                    key = key.OpenSubKey("Windows")
                    key = key.OpenSubKey("CurrentVersion")
                    key = key.OpenSubKey("RunOnce", True)
                    For Each val As String In key.GetValueNames
                        If val.ToLower.StartsWith("avp0") Or val.ToLower.StartsWith("avpo") Then
                            key.DeleteValue(val)
                            Activity_Handler("Removed " & key.ToString & "\" & val)
                        End If
                    Next
                    key.Close()
                Catch ex As Exception
                    Error_Handler(ex, "Remove Avp0* Registry Entry")
                End Try
                inputlines = inputlines + 1
                lastinputline = "Avp0* Registry Entries Removed"
               PercentageCalculation(worker)
            Catch ex As Exception
                Error_Handler(ex, "Remove Avp0* Registry Entry")
            End Try
            '************************************************************'
            'End Remove Avp0* Registry Entries
            '************************************************************'

            '************************************************************'
            'Remove ntde1ect Registry Entries
            '************************************************************'
            Try
                lastinputline = "Searching and Removing ntde1ect Registry Entries"

              PercentageCalculation(worker)

                Try
                    Dim key As RegistryKey
                    key = Registry.CurrentUser
                    RecursiveNTDE1ECTRemover(key)

                Catch ex As Exception
                    Error_Handler(ex, "Remove ntde1ect Registry Entry")
                End Try


                inputlines = inputlines + 1
                lastinputline = "ntde1ect Registry Entries Removed"
           PercentageCalculation(worker)
            Catch ex As Exception
                Error_Handler(ex, "Remove ntde1ect Registry Entry")
            End Try
            '************************************************************'
            'End Remove ntde1ect Registry Entries
            '************************************************************'

            '************************************************************'
            'Update Windows 'Hidden' Registry Entry
            '************************************************************'
            Try
                lastinputline = "Updating Windows 'Hidden' Registry Entry"
                PercentageCalculation(worker)
                Dim key As RegistryKey
                Try
                    key = Registry.CurrentUser.OpenSubKey("Software", False)
                    key = key.OpenSubKey("Microsoft")
                    key = key.OpenSubKey("Windows")
                    key = key.OpenSubKey("CurrentVersion")
                    key = key.OpenSubKey("Explorer")
                    key = key.OpenSubKey("Advanced", True)
                    key.SetValue("Hidden", 1, RegistryValueKind.DWord)
                    key.Close()
                Catch ex As Exception
                    Error_Handler(ex, "Updating Windows 'Hidden' Registry Entry")
                End Try
                inputlines = inputlines + 1
                lastinputline = "Windows 'Hidden' Registry Entry Updated"
                PercentageCalculation(worker)
            Catch ex As Exception
                Error_Handler(ex, "Updating Windows 'Hidden' Registry Entry")
            End Try
            '************************************************************'
            'End Update Windows 'Hidden' Registry Entry
            '************************************************************'

            Activity_Handler("")
            Activity_Handler("Operation Completed")
            Activity_Handler("***********************************************************")
            Activity_Handler("")

        Catch ex As Exception
            Error_Handler(ex, "MainWorkerFunction")
        End Try
        worker.Dispose()
        worker = Nothing
        e = Nothing
        Return result

    End Function

    Private Sub RecursiveNTDE1ECTRemover(ByVal key As RegistryKey)
        Try
            'Activity_Handler("Recording: " & key.ToString)
            Dim subkey As RegistryKey
            For Each ke1 As String In key.GetSubKeyNames()
                Try
                    subkey = key.OpenSubKey(ke1, True)
                    Label1.Text = "Checking " & key.ToString & "\" & subkey.ToString
                    For Each val As String In subkey.GetValueNames
                        Try
                            If val.ToLower.StartsWith("ntde1ect") Then
                                subkey.DeleteValue(val)
                                Activity_Handler("Removed " & subkey.ToString & "\" & val)
                            End If
                        Catch ex As Exception
                            Error_Handler(ex, "Remove ntde1ect Registry Entry")
                        End Try
                    Next
                    If subkey.Name.ToLower.StartsWith("ntde1ect") Then
                        key.DeleteSubKey(ke1, True)
                        Activity_Handler("Removed " & key.ToString & "\" & ke1)
                    End If
                    RecursiveNTDE1ECTRemover(subkey)
                    subkey.Close()
                Catch ex As Exception
                    Error_Handler(ex, "Remove ntde1ect Registry Entry")
                End Try
            Next
            key.Close()
        Catch ex As Exception
            Error_Handler(ex, "Remove ntde1ect Registry Entry")
        End Try

    End Sub

    Private Sub Form1_Close(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            Me.ToolStripStatusLabel1.Text = "Application Closing"
            SaveSettings()
        Catch ex As Exception
            Error_Handler(ex, "Application Close")
        End Try
    End Sub

    Private Sub LoadSettings()
        Try
            Dim configfile As String = (Application.StartupPath & "\config.sav").Replace("\\", "\")
            If My.Computer.FileSystem.FileExists(configfile) Then
                Dim reader As StreamReader = New StreamReader(configfile)
                Dim lineread As String
                Dim variablevalue As String
                While reader.Peek <> -1
                    lineread = reader.ReadLine
                    If lineread.IndexOf("=") <> -1 Then

                        variablevalue = lineread.Remove(0, lineread.IndexOf("=") + 1)

                        'If lineread.StartsWith("ImageFolder=") Then
                        '    Dim dinfo As DirectoryInfo = New DirectoryInfo(variablevalue)
                        '    If dinfo.Exists Then
                        '        FolderBrowserDialog1.SelectedPath = variablevalue
                        '        TextBox1.Text = variablevalue
                        '    End If
                        '    dinfo = Nothing
                        'End If

                        'If lineread.StartsWith("SetVariable=") Then
                        '    ComboBox1.SelectedIndex = variablevalue
                        'End If

                        'If lineread.StartsWith("PixelValue=") Then
                        '    NumericUpDown2.Value = variablevalue
                        'End If
                    
                    End If
                End While
                reader.Close()
                reader = Nothing
            End If
        Catch ex As Exception
            Error_Handler(ex, "Load Settings")
        End Try
    End Sub

    Private Sub SaveSettings()
        Try
            Dim configfile As String = (Application.StartupPath & "\config.sav").Replace("\\", "\")

            Dim writer As StreamWriter = New StreamWriter(configfile, False)

            'If TextBox1.Text.Length > 0 Then
            '    Dim dinfo As DirectoryInfo = New DirectoryInfo(TextBox1.Text)
            '    If dinfo.Exists Then
            '        writer.WriteLine("ImageFolder=" & TextBox1.Text)
            '    End If
            '    dinfo = Nothing
            'End If
            'If ComboBox1.SelectedIndex <> -1 Then
            '    writer.WriteLine("SetVariable=" & ComboBox1.SelectedIndex)
            'End If

            'writer.WriteLine("PixelValue=" & NumericUpDown2.Value)

            writer.Flush()
            writer.Close()
            writer = Nothing

        Catch ex As Exception
            Error_Handler(ex, "Save Settings")
        End Try
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            Me.Text = My.Application.Info.ProductName & " " & Format(My.Application.Info.Version.Major, "0000") & Format(My.Application.Info.Version.Minor, "00") & Format(My.Application.Info.Version.Build, "00") & "." & Format(My.Application.Info.Version.Revision, "00") & ""
            LoadSettings()
            Me.ToolStripStatusLabel1.Text = "Application Loaded"
        Catch ex As Exception
            Error_Handler(ex, "Application Load")
        End Try

    End Sub





    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        Try
            Me.ToolStripStatusLabel1.Text = "About displayed"
            AboutBox1.ShowDialog()
        Catch ex As Exception
            Error_Handler(ex, "Display About Screen")
        End Try
    End Sub

    Private Sub HelpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem.Click
        Try
            Me.ToolStripStatusLabel1.Text = "Help displayed"
            HelpBox1.ShowDialog()
        Catch ex As Exception
            Error_Handler(ex, "Display Help Screen")
        End Try
    End Sub


    Private Function RemoveFILE(ByVal filename As String) As Boolean
        Dim result As Boolean = False
        Try
            If My.Computer.FileSystem.FileExists(filename) Then
                File.SetAttributes(filename, FileAttributes.Normal)
                My.Computer.FileSystem.DeleteFile(filename, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                Activity_Handler("Removed File: " & filename)
                result = True
            End If
        Catch ex As Exception
            Error_Handler(ex, "Remove File: (" & filename & ")")
        End Try
        Return result
    End Function

    Private Function RemoveProcess(ByVal processname As String) As Boolean
        Dim result As Boolean = False
        Try
         
            Dim existing As Process() = Process.GetProcesses
            Dim eproc As Process
            For Each eproc In existing
                Try
                    Dim eprocname As String = eproc.ProcessName

                    If eprocname.ToLower = processname.ToLower Then
                        eproc.Kill()
                        Activity_Handler("Killed " & processname & " Process")
                    End If
                Catch ex As Exception
                    Error_Handler(ex, "Process Killer")
                End Try
                eproc = Nothing
            Next
            existing = Nothing
            result = True
        Catch ex As Exception
            Error_Handler(ex, "Rempve Process: (" & processname & ")")
        End Try
        Return result
    End Function

    Private Function RemoveEXE(ByVal filename As String) As Boolean
        Dim result As Boolean = False
        Try
            If My.Computer.FileSystem.FileExists(filename) Then
                Dim finfo As FileInfo = New FileInfo(filename)
                Dim processname As String = finfo.Name.Substring(0, finfo.Name.LastIndexOf("."))
                finfo = Nothing
                Dim existing As Process() = Process.GetProcesses
                Dim eproc As Process
                For Each eproc In existing
                    Try
                        Dim eprocname As String = eproc.ProcessName
                       
                        If eprocname.ToLower = processname.ToLower Then
                            eproc.Kill()
                            Activity_Handler("Killed " & processname & " Process")
                        End If
                    Catch ex As Exception
                        Error_Handler(ex, "Process Killer")
                    End Try
                    eproc = Nothing
                Next
                existing = Nothing
                File.SetAttributes(filename, FileAttributes.Normal)
                My.Computer.FileSystem.DeleteFile(filename, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                Activity_Handler("Removed EXE: " & filename)
                result = True
            End If
        Catch ex As Exception
            Error_Handler(ex, "Remove EXE: (" & filename & ")")
        End Try
        Return result
    End Function


    Private Function RemoveDLL(ByVal filename As String) As Boolean
        Dim result As Boolean = False
        Try
            If My.Computer.FileSystem.FileExists(filename) Then
                DosShellCommand("regsvr32 /u /s """ & filename & """")
                File.SetAttributes(filename, FileAttributes.Normal)
                My.Computer.FileSystem.DeleteFile(filename, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                Activity_Handler("Removed DLL: " & filename)
                result = True
            End If
        Catch ex As Exception
            Activity_Handler("Failed to Remove DLL: " & filename)
            Error_Handler(ex, "Remove DLL: (" & filename & ")")
        End Try
        Return result
    End Function

    Private Function DosShellCommand(ByVal AppToRun As String) As String
        Dim s As String = ""
        Try
            Dim myProcess As Process = New Process

            myProcess.StartInfo.FileName = "cmd.exe"
            myProcess.StartInfo.UseShellExecute = False
            myProcess.StartInfo.CreateNoWindow = True
            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.StartInfo.RedirectStandardError = True
            myProcess.Start()
            Dim sIn As StreamWriter = myProcess.StandardInput
            sIn.AutoFlush = True

            Dim sOut As StreamReader = myProcess.StandardOutput
            Dim sErr As StreamReader = myProcess.StandardError
            sIn.Write(AppToRun & _
               System.Environment.NewLine)
            sIn.Write("exit" & System.Environment.NewLine)
            s = sOut.ReadToEnd()
            If Not myProcess.HasExited Then
                myProcess.Kill()
            End If

            'MessageBox.Show("The 'dir' command window was closed at: " & myProcess.ExitTime & "." & System.Environment.NewLine & "Exit Code: " & myProcess.ExitCode)

            sIn.Close()
            sOut.Close()
            sErr.Close()
            myProcess.Close()
            'MessageBox.Show(s)
        Catch ex As Exception
            Error_Handler(ex)
        End Try
        Return s
    End Function

    Private Function PercentageCalculation(ByVal worker As BackgroundWorker) As Boolean
        Dim result As Boolean = False
        Try
            ' Report progress as a percentage of the total task.
            percentComplete = 0
            If inputlinesprecount > 0 Then
                percentComplete = CSng(inputlines) / CSng(inputlinesprecount) * 100
            Else
                percentComplete = 100
            End If
            primary_PercentComplete = percentComplete
            If percentComplete > 100 Then
                percentComplete = 100
            End If
            If percentComplete = 100 Then
                lastinputline = "Operation Completed"
            End If
            If percentComplete > highestPercentageReached Then
                highestPercentageReached = percentComplete
                worker.ReportProgress(percentComplete)
                result = True
            End If
        Catch ex As Exception
            Error_Handler(ex, "Percentage Calculation")
        End Try
        Return result
    End Function

End Class
