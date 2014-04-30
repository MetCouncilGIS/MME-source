Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Catalog


<ComClass(ImportMetadataCommand.ClassId, ImportMetadataCommand.InterfaceId, ImportMetadataCommand.EventsId), _
 ProgId("MnGeoMetadataEditor.ImportMetadataCommand")> _
Public NotInheritable Class ImportMetadataCommand
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.

    ' Be sure to generate and insert your own GUIDs here.
    'Public Const ClassId As String = "F58331D0-0760-4A0F-8B19-B71F7085A891"
    'Public Const InterfaceId As String = "8969F8A4-EAEA-4462-B1C2-32A09614D8F8"
    'Public Const EventsId As String = "44387B21-C65F-48FC-96E8-9CFEF54851F0"
    Public Const ClassId As String = "16ba0e83-7559-4e07-971c-822eb25cfb92"
    Public Const InterfaceId As String = "cc6d5586-0d26-4794-a74c-fdfdc8b670ca"
    Public Const EventsId As String = "05e9cd71-cb10-4480-ab1f-ff86d45c6bf4"
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
        MyBase.m_caption = "Import Metadata"   'localizable text 
        MyBase.m_message = "Import metadata into selected object(s)"   'localizable text 
        MyBase.m_toolTip = "Import metadata into selected object(s)" 'localizable text 
        MyBase.m_name = "MnGeoMetadataEditor_ImportMetadataCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcCatalogCommand")

        Try
            MyBase.m_bitmap = My.Resources.ImportMetadata
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
    ''' Event handler to carry out batch metadata import when button is clicked.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub OnClick()
        Try
            If Not GlobalVars.enabled Then Return

            Dim objects As List(Of IGxObject) = getSelectedObjects(gxApp, metability.CanWriteMetadata)

            If objects.Count = 0 Then
                MsgBox("Your selection does not include any valid objects that can have metadata!", , "MME Import Metadata")
            Else

                Dim o As IGxObject = gxbrowse_for_geodatabase(gxApp.View.hWnd)

                If o Is Nothing Then Exit Sub

                If Not canHaveMetadata(o) Then
                    MsgBox("Selected object is not a metadata source.")
                    Exit Sub
                End If

                If o.Category = "XML Document" AndAlso Not isWellFormedXmlFile(o.FullName) Then
                    MsgBox("Selected file does not contain well-formed XML.")
                    Exit Sub
                Else
                    If Not hasMetadata(o) Then
                        MsgBox("Selected object does not have metadata.")
                        Exit Sub
                    End If
                End If

                Dim srcXml As String = CType(CType(o, IMetadata).Metadata, IXmlPropertySet2).GetXml("")

                If MessageBox.Show("You are about to import metadata into " + objects.Count.ToString + " object(s). Proceed?", GlobalVars.nameStr, MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button2) = DialogResult.OK Then
                    For Each md As IMetadata In objects
                        Dim iXPSv As IXmlPropertySet2 = CType(md.Metadata, IXmlPropertySet2)
                        iXPSv.SetXml(srcXml)
                        md.Metadata = CType(iXPSv, ESRI.ArcGIS.esriSystem.IPropertySet)
                    Next
                    MessageBox.Show("Metadata has been imported into " + objects.Count.ToString + " object(s).", GlobalVars.nameStr, MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Prompt user for selecting an ArcCatalog object
    ''' </summary>
    ''' <param name="hParentHwnd">Parent hwnd</param>
    ''' <returns>IGxObject representing the user selected ArcCatalog object. Returns Nothing if no object was selected.</returns>
    ''' <remarks></remarks>
    Public Shared Function gxbrowse_for_geodatabase(ByVal hParentHwnd As Integer) As ESRI.ArcGIS.Catalog.IGxObject
        gxbrowse_for_geodatabase = Nothing

        Dim pGxDialog As ESRI.ArcGIS.CatalogUI.IGxDialog
        Dim pGxObject As ESRI.ArcGIS.Catalog.IGxObject
        Dim pEnumGxObj As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        'Set variables
        pGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog

        'Set dialog properties
        pGxDialog.Title = "Select the object that you want to import metadata from"
        pGxDialog.ButtonCaption = "Select"

        Try
            If Not pGxDialog.DoModalOpen(hParentHwnd, pEnumGxObj) = True Then
                Return Nothing
            End If
            Try
                pEnumGxObj.Reset()
                pGxObject = pEnumGxObj.Next
            Catch ex As Exception
                ErrorHandler(ex)
                Return Nothing
            End Try
        Catch ex As Exception
            ErrorHandler(ex)
            Return Nothing
        End Try

        Return pGxObject
    End Function


End Class



