Imports System.ComponentModel
Imports System.Configuration.Install
Imports System.Runtime.InteropServices

''' <summary>
''' Installer class.
''' </summary>
''' <remarks></remarks>
Public Class Installer

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
    ''' <remarks></remarks>
    Public Overrides Sub Install(ByVal stateSaver As System.Collections.IDictionary)
        MyBase.Install(stateSaver)
        Dim regsrv As New RegistrationServices
        regsrv.RegisterAssembly(MyBase.GetType().Assembly, AssemblyRegistrationFlags.SetCodeBase)
    End Sub

    ''' <summary>
    ''' Unregister assembly upon uninstallation.
    ''' </summary>
    ''' <param name="savedState"></param>
    ''' <remarks></remarks>
    Public Overrides Sub Uninstall(ByVal savedState As System.Collections.IDictionary)
        MyBase.Uninstall(savedState)
        Dim regsrv As New RegistrationServices
        regsrv.UnregisterAssembly(MyBase.GetType().Assembly)
    End Sub

End Class
