Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.Geodatabase


<ComClass(ClearAllMetadataCommand.ClassId, ClearAllMetadataCommand.InterfaceId, ClearAllMetadataCommand.EventsId), _
 ProgId("MnGeoMetadataEditor.ClearAllMetadataCommand")> _
Public NotInheritable Class ClearAllMetadataCommand
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.

    ' Be sure to generate and insert your own GUIDs here.
    'Public Const ClassId As String = "BDA5A413-2F3B-49A9-93D6-DCA23D9E93A0"
    'Public Const InterfaceId As String = "C5DEEBF3-5472-4CE7-B089-4EECF67F70E1"
    'Public Const EventsId As String = "C6070934-F812-4885-8332-402C71B2DC14"
    Public Const ClassId As String = "4f439339-3ae1-417b-bae2-ce9d1a8e5ef1"
    Public Const InterfaceId As String = "be049c4e-4db8-40d6-9106-ef37ef3c5559"
    Public Const EventsId As String = "04e5a59b-77bd-4be7-ada4-8703204bd016"
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

    Private Shared gxApp As IGxApplication

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        MyBase.m_category = "MGMG Tools"  'localizable text 
        MyBase.m_caption = "MME Batch Clear Metadata"   'localizable text 
        MyBase.m_message = "Clear metadata for selected object(s)"   'localizable text 
        MyBase.m_toolTip = "Clear metadata for selected object(s)" 'localizable text 
        MyBase.m_name = "MnGeoMetadataEditor_ClearMetadataCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcCatalogCommand")

        Try
            MyBase.m_bitmap = My.Resources.ClearMetadata
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
            gxApp = CType(hook, IGxApplication)

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
    ''' Event handler to carry out batch metadata clearing when button is clicked.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub OnClick()
        Try
            If Not GlobalVars.enabled Then Return

            Dim objects As List(Of IGxObject) = getSelectedObjects(gxApp, metability.CanWriteMetadata)

            If objects.Count = 0 Then
                MsgBox("Your selection does not include any valid objects that you can write metadata to!", , "MME Batch Clear Metadata")
            Else
                If MessageBox.Show("You are about to clear metadata for " + objects.Count.ToString + " object(s). Proceed?", GlobalVars.nameStr, MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button2) = DialogResult.OK Then
                    For Each md As IMetadata In objects
                        Dim srcXml As String = "<?xml version=""1.0""?>" + vbCrLf + "<metadata></metadata>"
                        Dim iXPSv As IXmlPropertySet2 = CType(md.Metadata, IXmlPropertySet2)
                        iXPSv.SetXml(srcXml)
                        md.Metadata = CType(iXPSv, ESRI.ArcGIS.esriSystem.IPropertySet)
                    Next
                    MessageBox.Show("Metadata for " + objects.Count.ToString + " object(s) have been cleared.", GlobalVars.nameStr, MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Sub

End Class



