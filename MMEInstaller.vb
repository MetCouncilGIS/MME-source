Imports System.ComponentModel
Imports System.Configuration.Install
Imports System.Runtime.InteropServices


''' <summary>
''' Installer class to register/unregister with ArcGIS Desktop.
''' </summary>
''' <remarks></remarks>
Public Class MMEInstaller

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add initialization code after the call to InitializeComponent

    End Sub
    ''' <summary>
    ''' Register assembly upon installation.
    ''' </summary>
    ''' <param name="stateSaver"></param>
    ''' <remarks>Stock ESRI example works by a custom install action telling which assembly to register here. We modififed it for the assembly to self register. 
    ''' ESRI method would fail if the user installed somewhere other than the default path.</remarks>
    Public Overrides Sub Install(ByVal stateSaver As System.Collections.IDictionary)
        MyBase.Install(stateSaver)
        Dim regsrv As New RegistrationServices
        regsrv.RegisterAssembly(MyBase.GetType().Assembly, AssemblyRegistrationFlags.SetCodeBase)


        'Register the custom component.
        '-----------------------------

        'The default location of the ESRIRegAsm utility.
        'Note how the whole string is embedded in quotes because of the spaces in the path.
        Dim cmd1 As String = """" + Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles) + "\ArcGIS\bin\ESRIRegAsm.exe" + """"

        'Obtain the path to the running assembly
        Dim part1 As String = Utils.getAssemblyPath

        'Add the appropriate command line switches when invoking the ESRIRegAsm utility.
        'In this case: /p:Desktop = means the ArcGIS Desktop product, /s = means a silent install.
        Dim part2 As String = " /p:Desktop /s"

        ' For debugging, uncomment the following line
        'part2 = " /e"

        'It is important to embed the part1 in quotes in case there are any spaces in the path.
        Dim cmd2 As String = """" + part1 + """" + part2

        'Call the routing that will execute the ESRIRegAsm utility.
        Dim exitCode As Integer = ExecuteCommand(cmd1, cmd2, 20000)

    End Sub

    ''' <summary>
    ''' Unregister assembly upon uninstallation.
    ''' </summary>
    ''' <param name="savedState"></param>
    ''' <remarks>Stock ESRI example works by a custom uninstall action telling which assembly to unregister here. We modififed it for the assembly to self unregister. 
    ''' ESRI method would fail if the user installed somewhere other than the default path.</remarks>
    Public Overrides Sub Uninstall(ByVal savedState As System.Collections.IDictionary)
        MyBase.Uninstall(savedState)
        Dim regsrv As New RegistrationServices
        regsrv.UnregisterAssembly(MyBase.GetType().Assembly)


        'Unregister the custom component.
        '---------------------------------

        'The default location of the ESRIRegAsm utility.
        'Note how the whole string is embedded in quotes because of the spaces in the path.
        Dim cmd1 As String = """" + Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles) + "\ArcGIS\bin\ESRIRegAsm.exe" + """"

        Dim part1 As String = Utils.getAssemblyPath

        'Add the appropriate command line switches when invoking the ESRIRegAsm utility.
        'In this case: /p:Desktop = means the ArcGIS Desktop product, /u = means unregister the Custom Component, /s = means a silent install.
        Dim part2 As String = " /p:Desktop /u /s"

        ' For debugging, uncomment the following line
        'part2 = " /u /e"

        'It is important to embed the part1 in quotes in case there are any spaces in the path.
        Dim cmd2 As String = """" + part1 + """" + part2

        'Call the routing that will execute the ESRIRegAsm utility.
        Dim exitCode As Integer = ExecuteCommand(cmd1, cmd2, 20000)
    End Sub

    Public Shared Function ExecuteCommand(ByVal Command1 As String, ByVal Command2 As String, ByVal Timeout As Integer) As Integer
        'System.IO.File.WriteAllText("c:\eme_out.txt", Command1 + vbCrLf + Command2)
        'MsgBox(Command1)
        'MsgBox(Command2)
        'Set up a ProcessStartInfo using your path to the executable (Command1) and the command line arguments (Command2).
        Dim ProcessInfo As ProcessStartInfo = New ProcessStartInfo(Command1, Command2)
        ProcessInfo.CreateNoWindow = True
        ProcessInfo.UseShellExecute = False
        If System.Environment.OSVersion.Version.Major >= 6 Then
            ProcessInfo.Verb = "runas"
        End If

        'Invoke the process.
        Dim Process As Process = Process.Start(ProcessInfo)
        Process.WaitForExit(Timeout)

        'Finish.
        Dim ExitCode As Integer = Process.ExitCode
        Process.Close()
        Return ExitCode
    End Function

End Class
