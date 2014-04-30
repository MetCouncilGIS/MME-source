Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.Geodatabase


''' <summary>
''' Class that implements the command for the batch validator.
''' Batch validator allows users to select multiple objects and validate their metadata (if any)
''' without opening EME. A button for the command is available on EME toolbar.
''' </summary>
''' <remarks></remarks>
<ComClass(BatchValidatorCommand.ClassId, BatchValidatorCommand.InterfaceId, BatchValidatorCommand.EventsId), _
 ProgId("MnGeoMetadataEditor.BatchValidatorCommand")> _
Public NotInheritable Class BatchValidatorCommand
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.

    ' Be sure to generate and insert your own GUIDs here.
    'Public Const ClassId As String = "0F15B028-B2F4-4787-9A00-22FEA8F9BB4C"
    'Public Const InterfaceId As String = "509508AE-9DB3-44CA-8968-21CF802969E6"
    'Public Const EventsId As String = "3633D0FD-A295-41E7-8985-826B107E6550"
    Public Const ClassId As String = "f44b5f23-dce2-4b9b-935e-4a7a01c40cb6"
    Public Const InterfaceId As String = "c3eab5ac-76c7-4a11-943c-78fd7c32dfb7"
    Public Const EventsId As String = "357f7947-8b6d-47ee-8c8c-a1bf2ead6fd5"
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

        MyBase.m_category = "MGMG Tools"  'localizable text 
        MyBase.m_caption = "MME Batch Validator"   'localizable text 
        MyBase.m_message = "Validate multiple metadata records using MnGeo's metadata validation service"   'localizable text 
        MyBase.m_toolTip = "Validate selected metadata records using MnGeo's metadata validation service" 'localizable text 
        MyBase.m_name = "MnGeoMetadataEditor_BatchValidatorCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcCatalogCommand")

        Try
            MyBase.m_bitmap = My.Resources.BatchValidatorBitmap
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
    ''' Event handler to carry out batch validation when button is clicked.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub OnClick()
        Try
            If Not GlobalVars.enabled Then Return

            Dim files As New List(Of String)
            Dim i As Integer = 1


            For Each md As IMetadata In getSelectedObjects(gxApp, metability.HasMetadata)
                'Dim md As IMetadata = CType(o, IMetadata)
                Dim args As New Hashtable
                Dim iXPSv As IXmlPropertySet2 = CType(md.Metadata, IXmlPropertySet2)

                ' We currently use object's name (typically its filename) as HTML page title.
                Dim title As String = CType(md, IGxObject).Name.Trim
                ' If a name cannot be obtained, we fall back to generating one here.
                If title Is Nothing OrElse title = "" Then
                    title = "Untitled " + i.ToString
                End If

                args("Metadata") = iXPSv.GetXml("")
                args("ObjectName") = title

                i += 1

                ' AEAE
                ' Try webservice...
                args("ValidationMode") = ValidationMode.Webservice
                If Not My.Settings.ValidationEnabled OrElse Not Utils.ValidateActionDo(args) Then
                    ' if disabled or doesn't work, try local...
                    args("ValidationMode") = ValidationMode.Local
                    If Not Utils.ValidateActionDo(args) Then
                        Continue For
                    End If
                End If

                Dim u As New Uri(args("Filename"))
                files.Add(u.AbsoluteUri)

                ' Display in browser
                Utils.OpenInIE(args("Filename"))
            Next

            If files.Count = 0 Then
                If i > 1 Then
                    MsgBox("Your selection does not include any valid objects that have metadata!", , "MME Batch Validation")
                Else
                    MsgBox("You must select one or more objects that have metadata!", , "MME Batch Validation")
                End If
            End If
        Catch ex As Exception
            ErrorHandler(ex)
        End Try

    End Sub

    ''' <summary>
    ''' Indicate whether the command (and its button) should be enabled.
    ''' </summary>
    ''' <value></value>
    ''' <returns>True to enable. False otherwise.</returns>
    ''' <remarks></remarks>
    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            ' We decided not to check if there are any items with metadata, in the interest of not slowing down the GUI
            Return gxApp IsNot Nothing 'AndAlso selectionHasMetadata()
        End Get
    End Property

End Class



