Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.esriSystem
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.CatalogUI


''' <summary>
''' EME is implemented as an extension to ArcCatalog using the IExtension interface
''' </summary>
''' <remarks></remarks>
<ComClass(ArcCatalogExtension.ClassId, ArcCatalogExtension.InterfaceId, ArcCatalogExtension.EventsId), _
 ProgId("MnGeoMetadataEditor.ArcCatalogExtension")> _
Public Class ArcCatalogExtension
    Implements IExtension

    Private dGxSelChE As IGxSelectionEvents_OnSelectionChangedEventHandler
    Private Shared WithEvents gxApp As IGxApplication


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
        GxExtensions.Register(regKey)

    End Sub
    ''' <summary>
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        GxExtensions.Unregister(regKey)

    End Sub

#End Region
#End Region

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.

    ' Be sure to generate and insert your own GUIDs here.
    'Public Const ClassId As String = "492D566D-E1CB-49F7-A392-2AB50DCC5FF2"
    'Public Const InterfaceId As String = "78C3FC99-54EA-454B-877F-8D2449B67E08"
    'Public Const EventsId As String = "2F450914-5ADC-4CD0-873C-D80223C6B6E7"
    Public Const ClassId As String = "d8878039-0afa-42b3-aeaf-9d966abc378f"
    Public Const InterfaceId As String = "8b052f4f-e7d6-49c9-8f33-b7cea78df71d"
    Public Const EventsId As String = "495bbc79-1b1a-42de-97a6-53cac01ee927"
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
        MyBase.New()
    End Sub

    ''' <summary>
    ''' Name of extension. Do not exceed 31 characters
    ''' </summary>
    Public ReadOnly Property Name() As String Implements ESRI.ArcGIS.esriSystem.IExtension.Name
        Get
            'TODO: Modify string to uniquely identify extension
            Return "MMESyncHelperExtension"
        End Get
    End Property

    ''' <summary>
    ''' Clean up while ArcCatalog is shutting down this extension.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Shutdown() Implements ESRI.ArcGIS.esriSystem.IExtension.Shutdown
        'TODO: Clean up resources
        gxApp = Nothing
    End Sub

    ''' <summary>
    ''' Initialization activities while ArcCatalog is starting up this extension.
    ''' </summary>
    ''' <param name="initializationData"></param>
    ''' <remarks></remarks>
    Public Sub Startup(ByRef initializationData As Object) Implements ESRI.ArcGIS.esriSystem.IExtension.Startup
        If TypeOf initializationData Is IGxApplication Then
            gxApp = CType(initializationData, IGxApplication)
        End If
    End Sub

End Class


