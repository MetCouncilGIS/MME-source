Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Catalog


<ComClass(ExportMetadataCommand.ClassId, ExportMetadataCommand.InterfaceId, ExportMetadataCommand.EventsId), _
 ProgId("MnGeoMetadataEditor.ExportMetadataCommand")> _
Public NotInheritable Class ExportMetadataCommand
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.

    ' Be sure to generate and insert your own GUIDs here.
    'Public Const ClassId As String = "59B1533F-6492-43B2-9269-A3B7ABCBD005"
    'Public Const InterfaceId As String = "D1F3966E-EC5B-4E42-9544-6CC74DD0E1BD"
    'Public Const EventsId As String = "53DE4756-AC9A-4B2A-8C83-FBF967BD0A55"
    Public Const ClassId As String = "811970dc-e724-4b5b-a5ac-da4a3a8102b3"
    Public Const InterfaceId As String = "69ddfed7-cf41-4724-9f0d-50e7e2cb9287"
    Public Const EventsId As String = "7d787b96-e555-4a67-82c5-7dce57f2ed32"
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
        MyBase.m_caption = "Export Metadata"   'localizable text 
        MyBase.m_message = "Export metadata from selected object(s)"   'localizable text 
        MyBase.m_toolTip = "Export metadata from selected object(s)" 'localizable text 
        MyBase.m_name = "MnGeoMetadataEditor_ExportMetadataCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcCatalogCommand")

        Try
            MyBase.m_bitmap = My.Resources.ExportMetadata
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
    ''' Event handler to carry out batch export when button is clicked.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub OnClick()
        Try
            If Not GlobalVars.enabled Then Return

            Dim objects As List(Of IGxObject) = getSelectedObjects(gxApp, metability.HasMetadata)

            If objects.Count = 0 Then
                MsgBox("Your selection does not include any objects that have metadata!", , "MME Export Metadata")
            ElseIf objects.Count = 1 Then
                Dim dlg As New SaveFileDialog()
                dlg.DefaultExt = ".xml"
                dlg.AddExtension = True
                dlg.RestoreDirectory = True
                dlg.FileName = makeMdFilename(objects(0))
                dlg.Filter = "XML file (*.xml)|*.xml|All files (*.*)|*.*"
                Dim res As DialogResult = dlg.ShowDialog()
                If res = Windows.Forms.DialogResult.OK Then
                    saveMdToFile(objects(0), dlg.FileName)
                End If
            Else
                Dim dlg As New SaveFileDialog()
                dlg.FileName = "folder selection only - no filename required"
                dlg.Filter = "folders|folders_only"
                dlg.CheckFileExists = False
                dlg.CheckPathExists = True
                'fbd.ValidateNames = False
                dlg.RestoreDirectory = True
                Dim res As DialogResult = dlg.ShowDialog()
                If res = Windows.Forms.DialogResult.OK Then
                    Dim i As Integer = 0
                    For Each md As IMetadata In objects
                        Dim filename As String = IO.Path.GetFullPath(IO.Path.GetDirectoryName(dlg.FileName)) + "\" + makeMdFilename(objects(i), i)
                        saveMdToFile(objects(i), filename)
                        i += 1
                    Next
                    MessageBox.Show("Metadata has been exported from " + objects.Count.ToString + " object(s).", GlobalVars.nameStr, MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Generate a filename for the metadata record
    ''' </summary>
    ''' <param name="o">The IGxObject for which a metadata filename is to be generated</param>
    ''' <param name="idx">An Integer index to use in filename. Useful when generating metadata files in batch.</param>
    ''' <returns>String containing filename generated.</returns>
    ''' <remarks>Currently uses object name, metadata title and optionally an index number to generate a filename.</remarks>
    Function makeMdFilename(ByVal o As IGxObject, Optional ByVal idx As Integer = -1) As String
        makeMdFilename = o.Name
        Dim mdTitle As String = AoUtils.getMdProperty(o, "idinfo/citation/citeinfo/title")
        If mdTitle <> "" Then makeMdFilename += " - " + mdTitle
        If idx > -1 Then makeMdFilename = (idx + 1).ToString.Trim.PadLeft(4, "0") + " - " + makeMdFilename
        makeMdFilename += ".xml"
    End Function

    ''' <summary>
    ''' Write metadata record to file.
    ''' </summary>
    ''' <param name="imd">IMetadata object with metadata to write to file.</param>
    ''' <param name="filename">The filename to use for the file.</param>
    ''' <remarks></remarks>
    Sub saveMdToFile(ByVal imd As IMetadata, ByVal filename As String)
        Dim iXPSv As IXmlPropertySet2 = CType(imd.Metadata, IXmlPropertySet2)
        Dim srcXml As String = iXPSv.GetXml("")
        System.IO.File.WriteAllText(filename, srcXml)
    End Sub

End Class



