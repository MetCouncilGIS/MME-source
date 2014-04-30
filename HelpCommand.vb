Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
'Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.CatalogUI

<ComClass(HelpCommand.ClassId, HelpCommand.InterfaceId, HelpCommand.EventsId), _
 ProgId("MnGeoMetadataEditor.HelpCommand")> _
Public NotInheritable Class HelpCommand
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.

    ' Be sure to generate and insert your own GUIDs here.
    'Public Const ClassId As String = "F973C64B-D0BB-4DEF-9C7A-8FBD968E35C7"
    'Public Const InterfaceId As String = "6C345252-C092-4833-9B4C-660A15BD9288"
    'Public Const EventsId As String = "69EE1461-6FA8-4C40-B09C-B7522B86840C"
    Public Const ClassId As String = "5a62c40d-d1f2-4a14-9a4c-8f1fd4acaa8c"
    Public Const InterfaceId As String = "daff2173-4c44-45ba-8dee-c0e089e60ca6"
    Public Const EventsId As String = "941d32d7-0034-4d09-adac-60720cb55b15"
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

    'Private m_application As IApplication

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        MyBase.m_category = "MGMG Tools"  'localizable text 
        MyBase.m_caption = "MME Help"   'localizable text 
        MyBase.m_message = "MME Help"   'localizable text 
        MyBase.m_toolTip = "MME Help" 'localizable text 
        MyBase.m_name = "MnGeoMetadataEditor_HelpCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcCatalogCommand")

        Try
            MyBase.m_bitmap = My.Resources.Help
        Catch ex As Exception
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try
    End Sub

    ''' <summary>
    ''' Event handler to initialize the command when created.
    ''' </summary>
    ''' <param name="hook"></param>
    ''' <remarks></remarks>
    Public Overrides Sub OnCreate(ByVal hook As Object)
        If Not hook Is Nothing Then
            'm_application = CType(hook, IApplication)

            'Disable if it is not ArcCatalog
            If TypeOf hook Is IGxApplication Then
                MyBase.m_enabled = True
            Else
                MyBase.m_enabled = False
            End If
        End If

        ' TODO:  Add other initialization code
    End Sub

    ''' <summary>
    ''' Event handler to open help when button is clicked.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub OnClick()
        Try
            If Not GlobalVars.enabled Then Return

            Utils.HelpSeeker("/Help_Main.html", GlobalVars.proc)
        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Sub
End Class



