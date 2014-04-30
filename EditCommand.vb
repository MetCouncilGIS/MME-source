Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem


<ComClass(EditCommand.ClassId, EditCommand.InterfaceId, EditCommand.EventsId), _
 ProgId("MnGeoMetadataEditor.EditCommand")> _
Public NotInheritable Class EditCommand
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.

    ' Be sure to generate and insert your own GUIDs here.
    'Public Const ClassId As String = "CE9F0491-D5AC-47D8-9507-34BFE27EEDFC"
    'Public Const InterfaceId As String = "9D8A82C0-DAFD-4433-A8B8-23ED86431B00"
    'Public Const EventsId As String = "EF0F0407-C74E-4C5B-AA25-2D417E634911"
    Public Const ClassId As String = "B99A7012-EAAA-48CC-B33F-4BFE8492E1CB"
    Public Const InterfaceId As String = "08BDAE8D-2B2B-409A-8233-80D19FFCA715"
    Public Const EventsId As String = "A04CA783-21D8-4B9B-8C80-285417B9D96C"
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
    Private Shared gxApp As IGxApplication

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        MyBase.m_category = "MGMG Tools"  'localizable text 
        MyBase.m_caption = "Edit MGMG metadata"   'localizable text 
        MyBase.m_message = "Edit MGMG metadata"   'localizable text 
        MyBase.m_toolTip = "Edit MGMG metadata" 'localizable text 
        MyBase.m_name = "MnGeoMetadataEditor_EditCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcCatalogCommand")

        Try
            MyBase.m_bitmap = New Bitmap(My.Resources.MetadataEdit16)
        Catch ex As Exception
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
            ErrorHandler(ex)
        End Try


    End Sub

    ''' <summary>
    ''' Event handler to initialize the command when created.
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

        ' TODO:  Add other initialization code
    End Sub

    ''' <summary>
    ''' Indicate whether the command (and its button) should be enabled.
    ''' </summary>
    ''' <value></value>
    ''' <returns>True to enable. False otherwise.</returns>
    ''' <remarks></remarks>
    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            Return gxApp IsNot Nothing AndAlso gxApp.SelectedObject IsNot Nothing 'AndAlso selectionHasMetadata()
        End Get
    End Property

    ''' <summary>
    ''' Event handler to start editor when button is clicked.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub OnClick()
        Try
            If Not GlobalVars.enabled Then Return

            Dim objects As List(Of IGxObject) = getSelectedObjects(gxApp, metability.CanHaveMetadata)
            If objects.Count = 0 Then
                MsgBox("Your selection does not include any valid objects that can have metadata!", , "MME Edit Metadata")
            ElseIf objects.Count > 1 Then
                MsgBox("You can edit only one metadata record at a time!", , "MME Edit Metadata")
            Else
                Dim md As IMetadata = objects(0)
                Dim ips As IPropertySet = md.Metadata
                Dim mc As New MasterController()
                If mc.Edit(ips, 0) Then
                    'save
                    md.Metadata = ips
                End If
            End If
        Catch ex As Exception
            ErrorHandler(ex)
        End Try

    End Sub
End Class



