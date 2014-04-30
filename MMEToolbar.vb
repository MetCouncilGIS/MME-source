Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ADF.BaseClasses
'Imports ESRI.ArcGIS.ArcCatalogUI
Imports System.Runtime.InteropServices

''' <summary>
''' An ArcCatalog toolbar for EME, EPA Synchronizer and related tools.
''' </summary>
''' <remarks></remarks>
<ComClass(MMEToolbar.ClassId, MMEToolbar.InterfaceId, MMEToolbar.EventsId), _
 ProgId("MnGeoMetadataEditor.MMEToolbar")> _
Public NotInheritable Class MMEToolbar
    Inherits BaseToolbar

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
    ''' <summary>
    ''' Required method for ArcGIS Component Category registration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        GxCommandBars.Register(regKey)

    End Sub
    ''' <summary>
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        GxCommandBars.Unregister(regKey)

    End Sub

#End Region
#End Region

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.

    ' Be sure to generate and insert your own GUIDs here.
    Public Const ClassId As String = "B92DAB77-2972-488B-B06D-F9328B83ECC0"
    Public Const InterfaceId As String = "9FDE2664-987A-413E-9CA8-B5DBEC887C5D"
    Public Const EventsId As String = "C42C3A2C-6A17-4657-AC3E-1E8DBE12A357"
#End Region

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks>A creatable COM class must have a Public Sub New() 
    ''' with no parameters, otherwise, the class will not be 
    ''' registered in the COM registry and cannot be created 
    ''' via CreateObject.
    ''' </remarks>
    Public Sub New()
        'BeginGroup() 'Separator
        ' We include some commands of our own
        AddItem("MnGeoMetadataEditor.EditCommand")
        'AddItem("MnGeoMetadataEditor.BatchValidatorCommand")
        AddItem("MnGeoMetadataEditor.SyncManagerCommand")
        AddItem("MnGeoMetadataEditor.SyncCommand")
        'AddItem("MnGeoMetadataEditor.ClearAllMetadataCommand")
        AddItem("MnGeoMetadataEditor.ImportMetadataCommand")
        AddItem("MnGeoMetadataEditor.ExportMetadataCommand")
        AddItem("MnGeoMetadataEditor.HelpCommand")

        ' See this list for other metadata related ArcCatalog commands you could add
        ' http://help.arcgis.com/en/sdk/10.0/arcobjects_net/conceptualhelp/index.html#//000100000020000000
    End Sub

    ''' <summary>
    ''' Caption for the toolbar.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides ReadOnly Property Caption() As String
        Get
            Return "MGMG Metadata Toolbar"
        End Get
    End Property

    ''' <summary>
    ''' Name of the toolbar.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides ReadOnly Property Name() As String
        Get
            Return "MMEToolbar"
        End Get
    End Property
End Class


