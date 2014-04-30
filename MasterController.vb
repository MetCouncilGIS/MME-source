Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.CatalogUI
Imports System.Runtime.InteropServices


''' <summary>
''' MasterController implements the IMetadataEditor interface and acts as the glue between ArcCatalog and EME.
''' </summary>
''' <remarks></remarks>
<ComClass(MasterController.ClassId, MasterController.InterfaceId, MasterController.EventsId), _
 ProgId("MnGeoMetadataEditor.MasterController")> _
Public Class MasterController
    Implements IMetadataEditor


#Region "COM Registration Function(s)"
    ''' <summary>
    ''' Register EME with COM
    ''' </summary>
    ''' <param name="registerType"></param>
    ''' <remarks></remarks>
    <ComRegisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub RegisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryRegistration(registerType)

        'Add any COM registration code after the ArcGISCategoryRegistration() call

    End Sub

    ''' <summary>
    ''' Unregister EME with COM
    ''' </summary>
    ''' <param name="registerType"></param>
    ''' <remarks></remarks>
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
        GxCommands.Register(regKey)
        GxExtensions.Register(regKey)
        MetadataEditor.Register(regKey)

    End Sub
    ''' <summary>
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        GxCommands.Unregister(regKey)
        GxExtensions.Unregister(regKey)
        MetadataEditor.Unregister(regKey)

    End Sub

#End Region
#End Region

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.

    ' Be sure to generate and insert your own GUIDs here.
    'Public Const ClassId As String = "59390CD0-5534-46C3-8E42-E0822C51DC00"
    'Public Const InterfaceId As String = "DE2F0A40-56FC-4C88-B929-B88C3840DCE5"
    'Public Const EventsId As String = "D3692AF9-C33C-4776-96D7-68CF8E2A0E85"
    Public Const ClassId As String = "920ee5bf-6a6d-48c3-a0b9-6d85b0d69572"
    Public Const InterfaceId As String = "01992999-ab9a-40a1-add4-00aac2bb34b5"
    Public Const EventsId As String = "26ffd4e6-2fc9-4fbd-98c3-9bd9feeca4ce"
#End Region


    ''' <summary>
    ''' Main entry point into metadata editor when initiated by ArcCatalog
    ''' </summary>
    ''' <param name="metadata">Metadata record copy to be edited.</param>
    ''' <param name="hWnd">Handle to the metadata editor window.</param>
    ''' <returns>
    ''' Boolean return value indicates if the metadata record was modified by the editor
    ''' </returns>
    ''' <remarks>With the advent of standalone editor, this is now a wrapper around it.</remarks>
    Public Function Edit(ByVal metadata As IPropertySet, ByVal hWnd As Integer) As Boolean Implements IMetadataEditor.Edit
        Try
            ' We immediately convert the metadata record from IPropertySet to XmlMetadata 
            ' since the latter has a richer interface to manipulate the record. 
            iXPS = New XmlMetadata
            Dim AoMdHolder As IXmlPropertySet2 = CType(metadata, IXmlPropertySet2)
            iXPS.SetXml(AoMdHolder.GetXml(""))

            Dim editorPath As String
            If Utils.iAmInstalled Then
                ' app installed 
                editorPath = Utils.getAppFolder + "\MetadataEditor.exe"
            Else
                ' app in dev env - AE: Change this if projects are reorganized...
                editorPath = Utils.getAppFolder.Replace("MME", "MME\MetadataEditor") + "\MetadataEditor.exe"
            End If

            Dim editorProc As New Process
            editorProc.StartInfo.FileName = editorPath
            Dim tmpMdFile As String = Utils.textToTempFile(iXPS.GetXml(""))
            Dim oldMd5 As String = Utils.stringMd5(iXPS.GetXml(""))
            editorProc.StartInfo.Arguments = """" + tmpMdFile + """"

            editorProc.Start()
            editorProc.WaitForExit()

            iXPS.SetXml(IO.File.ReadAllText(tmpMdFile))
            Dim newMd5 As String = Utils.stringMd5(iXPS.GetXml(""))
            ' Save if md5 hash changed after editing
            Dim saveHint As Boolean = (oldMd5 <> newMd5)

            AoMdHolder.SetXml(iXPS.GetXml(""))
            ' Explicitly releasing the metadata object.
            AoMdHolder = Nothing

            ' Explicitly save settings as we are a DLL - not EXE.
            My.Settings.Save()

            ' When the form is closed, we return a boolean value indicating
            ' whether or not the metadata was modified. That information is
            ' passed back to the ArcCatalog -- if edited, the metadata at the source 
            ' will be updated and metadata view is refreshed so you can see the changes.
            Return saveHint
        Catch ex As Exception
            ' No well-defined mechanism in place for reporting errors.
            Utils.ErrorHandler(ex)
        End Try
    End Function

    ''' <summary>
    ''' Sets the name for this editor - this name will appear in the
    ''' list of metadata editors in ArcCatalog's Options dialog box
    ''' </summary>
    ''' <returns>String containing the name of this metadata editor.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Name1() As String Implements IMetadataEditor.Name
        Get
            Return "Minnesota Metadata Editor Build: " + Utils.GetVersion()
        End Get
    End Property

End Class


