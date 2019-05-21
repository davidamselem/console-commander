


'//
Imports System.IO
Imports System.IO.Directory







''' <summary>
''' 
''' </summary>
Module Main

    '/**/
    Sub Main()

        '//
        Dim CmdMsg As String = Nothing
        Dim CmdLine As String

        '//
        Dim Source As String
        Dim SourceFolder As String
        Dim SourceCount As Long

        '//
        Dim Target As String
        Dim TargetFolder As String
        Dim TargetCount As Long

        '//
        Dim TimeSpanObject As TimeSpan
        Dim StartTime As Date
        Dim EndTime As Date

        '//
        Dim DirectoryInfoObject As DirectoryInfo

        '//
        Dim CmdMenu As String = Nothing
        Dim CmdMenuItems() As String
        Dim CmdIndex As Integer = 1

        '//
        Console.WriteLine("Welcome to Console Commander.")

        '//
        WL("")

        Try

            '//
            Console.Write("Source:       ")

            '//
            CmdLine = Console.ReadLine()

            '//
            Source = CmdLine

            If Source.Length = 0 Then

                '//
                GoTo PRESS_ANY_KEY_TO_EXIT

            ElseIf Source.Length = 1 Then

                '//
                Source = Source.ToUpper & ":\"

            Else

                '//
                If Source.LastIndexOf("\") < Source.Length - 1 Then

                    '//
                    Source = Source & "\"

                End If

            End If

            '//
            WL("")

            '//Menu     
            DirectoryInfoObject = New DirectoryInfo(Source)

            '//
            For Each DirectoryInfoObj In DirectoryInfoObject.GetDirectories

                '//
                Console.WriteLine(Format(CmdIndex, "0#") & " - " & DirectoryInfoObj.Name)

                '//
                CmdMenu &= "," & DirectoryInfoObj.Name
                CmdIndex += 1

            Next

            '//
            CmdMenuItems = CmdMenu.Split(",")

            '//
            WL("")

            '//Folder   
            Console.Write("SourceFolder: ")
            CmdLine = Console.ReadLine()

            '//
            SourceFolder = CmdMenuItems(CmdLine)

            '//
            WL("")

            '//Target
            Console.Write("Target:       ")

            '//
            CmdLine = Console.ReadLine()

            '//
            Target = CmdLine

            '//
            If Target.Length = 0 Then

                '//
                GoTo PRESS_ANY_KEY_TO_EXIT

            ElseIf Target.Length = 1 Then

                '//
                Target = Target.ToUpper & ":\"

            Else

                '//
                If Target.LastIndexOf("\") < Target.Length - 1 Then

                    '//
                    Target = Target & "\"

                End If

            End If

            '//
            'Console.Write("TargetFolder: ")
            'CmdLine = Console.ReadLine()

            '//
            'If CmdLine.Length > 0 Then
            '    TargetFolder = CmdLine
            'Else

            'End If

            '//
            TargetFolder = SourceFolder

            '//
            WL("")

            '//Backup   
            StartTime = Now

            '//
            If Backup(Source, SourceFolder, Target, TargetFolder) Then

                '//
                EndTime = Now

                '//
                TimeSpanObject = EndTime - StartTime

                '//Compare  
                ReadFolderInfo(Source & "\" & SourceFolder, True, SourceCount, CmdMsg)
                WL("SourceCount:  " & CapacityString(SourceCount) & " (" & SourceCount & " bytes)")

                '//
                ReadFolderInfo(Target & "\" & TargetFolder, True, TargetCount, CmdMsg)
                WL("TargetCount:  " & CapacityString(TargetCount) & " (" & TargetCount & " bytes)" + vbCrLf)

                '//TimeSpan 
                WL("TimeSpan:     " & TimeSpanObject.Minutes & " min " & TimeSpanObject.Seconds & " s" + vbCrLf)

            Else

                '//
                WL("An error has occured." + vbCrLf)

            End If

        Catch ex As Exception

            '//
            WL(ex.Message + vbCrLf)

        End Try

PRESS_ANY_KEY_TO_EXIT:

        '//
        Console.Write("Press any key to exit.")

        '//
        Console.ReadKey()

    End Sub

    '/**/
    Private Function WL(ByVal msg As String) As Boolean

        '//
        Console.WriteLine(msg)

        '//
        Return True

    End Function

    '/**/
    Public Function Backup(ByRef Source As String,
                           ByRef SourceFolder As String,
                           ByRef Target As String,
                           ByRef TargetFolder As String) As Boolean

        '//
        Dim StartupPath As String
        Dim PathName As String
        Dim Arguments As String

        '//
        Dim ProcessId As Integer
        Dim ProcessCode As Integer = -1
        Dim ProcessObject As Diagnostics.Process

        '//
        Try

            '//
            StartupPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)
            StartupPath = StartupPath.Replace("file:\", "")

            '//
            If Is64BitOS() Then

                '//
                PathName = StartupPath & "\lib\64bit\xxcopy.exe"

            Else

                '//
                PathName = StartupPath & "\lib\32bit\xxcopy.exe"

            End If

            '//
            Arguments = Source & SourceFolder & " " & Target & TargetFolder & " /CLONE /FF"

            '//Start the process.                                                   
            ProcessObject = Diagnostics.Process.Start(PathName, Arguments)

            '//Read the process id.                                                 
            ProcessId = ProcessObject.Id

            '//Wait for the process.                                                
            ProcessObject.WaitForExit()

            '//Read the exit code.                                                  
            If ProcessObject.HasExited Then

                '//
                ProcessCode = ProcessObject.ExitCode

            End If

            '//
            'Console.WriteLine("ProcessId " & ProcessId & " terminated with exit code " & ProcessCode & ".")

            '//
            Return True

        Catch ex As Exception

            '//
            Return False

        End Try

    End Function

    '/**/
    Private Function CapacityString(ByVal Capacity As Long) As String

        '//
        Dim CapacitySingle As Single

        '//
        Select Case Capacity

            '//
            Case Is >= 1000000000

                '//
                CapacitySingle = Capacity / 1024 / 1024 / 1024

                '//
                CapacityString = Math.Round(CapacitySingle, 2) & " GB"

            Case Is >= 1000000

                '//
                CapacitySingle = Capacity / 1024 / 1024

                '//
                CapacityString = Math.Round(CapacitySingle, 2) & " MB"

            Case Is >= 100000

                '//
                CapacitySingle = Capacity / 1024

                '//
                CapacityString = Math.Round(CapacitySingle, 2) & " KB"

            Case Else

                '//
                CapacityString = Capacity & " Bytes"

        End Select

    End Function

    '/**/
    Private Function ReadFolder(ByVal PathName As String, ByVal Recursive As Boolean) As Long

        '//
        Dim FolderSize As Long
        Dim FolderCount As Long

        '//
        Dim FileInfoObject As FileInfo

        '//
        Dim DirectoryInfoObject As DirectoryInfo
        Dim SubDirectoryInfoObject As DirectoryInfo

        '//
        FolderSize = 0

        '//
        DirectoryInfoObject = New DirectoryInfo(PathName)

        '//
        Try

            '//
            For Each FileInfoObject In DirectoryInfoObject.GetFiles()

                '//
                FolderSize += FileInfoObject.Length

            Next

            '//
            If Recursive = True Then

                '//
                FolderCount = 0

                '//
                For Each SubDirectoryInfoObject In DirectoryInfoObject.GetDirectories()

                    '//
                    FolderSize += ReadFolder(SubDirectoryInfoObject.FullName, True)

                    '//
                    FolderCount += 1

                Next

            End If

            '//
            Return FolderSize

        Catch FileNotFoundExceptionObject As System.IO.FileNotFoundException

            '...

        Catch ex As Exception

            '//
            Return -1

        End Try

        '//
        Return 0

    End Function

    '/**/
    Private Function ReadFolderInfo(ByVal PathName As String,
                                    ByVal Recursive As Boolean,
                                    ByRef FolderSize As Long,
                                    ByRef CmdMsg As String) As Boolean

        '//
        Try

            '//
            FolderSize = ReadFolder(PathName, Recursive)

            '//
            If FolderSize > -1 Then

                '//
                Return True

            Else

                '//
                CmdMsg = "An unknown error has occured."

                '//
                Return False

            End If

        Catch ex As Exception

            '//
            CmdMsg = ex.Message

            '//
            Return False

        End Try

    End Function

End Module