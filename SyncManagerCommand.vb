Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.CatalogUI


''' <summary>
''' Class that implements a user interface for controlling (turning on/off) synchronizers
''' and also controlling the EPA Synchronizer options.
''' </summary>
''' <remarks></remarks>
<ComClass(SyncManagerCommand.ClassId, SyncManagerCommand.InterfaceId, SyncManagerCommand.EventsId), _
 ProgId("MnGeoMetadataEditor.SyncManagerCommand")> _
Public NotInheritable Class SyncManagerCommand
    Inherits BaseCommand

    ''' <summary>
    ''' Reference to ArcCatalog instance
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared gxApp As IGxApplication

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.

    ' Be sure to generate and insert your own GUIDs here
    'Public Const ClassId As String = "FD222E78-0A44-44DC-8A33-104F18B97303"
    'Public Const InterfaceId As String = "F7B096EA-977E-4B0F-A6D6-C00500E6A26F"
    'Public Const EventsId As String = "8A42E40E-4574-428A-9210-C01DDA1D643E"
    Public Const ClassId As String = "11e6eeb2-7047-4c88-9270-df89bdbb2a64"
    Public Const InterfaceId As String = "7908c310-6900-418c-aadb-32d8e45369fc"
    Public Const EventsId As String = "ad3570d3-c6ea-4458-b04a-de8a70b7fb50"
#End Region


#Region "COM Registration Function(s)"
    <ComRegisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub RegisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryRegistration(registerType)

        'Add any COM registration code after the ArcGISCategoryRegistration() call

    End Sub

    <ComUnregisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub UnregisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryUnregistration(registerType)

        'Add any COM unregistration code after the ArcGISCategoryUnregistration() call

    End Sub

#Region "ArcGIS Component Category Registrar generated code"
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        GxCommands.Register(regKey)

    End Sub
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        GxCommands.Unregister(regKey)

    End Sub

#End Region
#End Region

    ''' <summary>
    ''' Create and initialize.
    ''' </summary>
    ''' <remarks>A creatable COM class must have a Public Sub New() 
    ''' with no parameters, otherwise, the class will not be 
    ''' registered in the COM registry and cannot be created 
    ''' via CreateObject.
    ''' </remarks>
    Public Sub New()
        MyBase.New()

        MyBase.m_category = "MGMG Tools"  'localizable text 
        MyBase.m_caption = "MGMG Synchronizer Manager"   'localizable text 
        MyBase.m_message = "Manage metadata synchronizers including MGMG Synchronizer"   'localizable text 
        MyBase.m_toolTip = "Manage metadata synchronizers including MGMG Synchronizer" 'localizable text 
        MyBase.m_name = "MnGeoMetadataEditor_SynchronizerManagerCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcCatalogCommand")

        Try
            MyBase.m_bitmap = My.Resources.SyncManagerBitmap
        Catch ex As Exception
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
            ErrorHandler(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Event handler that gets called when an instance is created. Used to get a hold of ArcCatalog.
    ''' </summary>
    ''' <param name="hook"></param>
    ''' <remarks></remarks>
    Public Overrides Sub OnCreate(ByVal hook As Object)
        If Not hook Is Nothing Then
            'Disable if it is not ArcCatalog
            If TypeOf hook Is IGxApplication Then
                gxApp = hook
                MyBase.m_enabled = True
            Else
                MyBase.m_enabled = False
            End If
        End If
    End Sub

    ''' <summary>
    ''' Event handler that opens the form for managing syncronizers.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub OnClick()
        Try
            Dim frm As New SyncManagerForm
            frm.ShowDialog()
        Catch ex As Exception
            ErrorHandler(ex)
        End Try

        ' Explicitly, save settings as we are a DLL - not EXE.
        My.Settings.Save()
    End Sub


End Class


