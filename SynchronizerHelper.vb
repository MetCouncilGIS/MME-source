Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem
Imports System.Runtime.InteropServices


''' <summary>
''' EPA Synchronizer implementation helper.
''' Implements IMetadataSynchronizer interface.
''' You never see or intereact with this directly.
''' </summary>
''' <remarks>Relies on the FGDC Synchronizer to do much of its work.</remarks>
<ComClass(SynchronizerHelper.ClassId, SynchronizerHelper.InterfaceId, SynchronizerHelper.EventsId), _
 ProgId("MnGeoMetadataEditor.SynchronizerHelper")> _
Public Class SynchronizerHelper
    Implements IMetadataSynchronizer


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
        MetadataSynchronizers.Register(regKey)

    End Sub
    ''' <summary>
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MetadataSynchronizers.Unregister(regKey)

    End Sub

#End Region
#End Region


#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.

    ' Be sure to generate and insert your own GUIDs here.
    'Public Const ClassId As String = "553783EC-C0BD-429E-8B59-E007699E6BCC"
    'Public Const InterfaceId As String = "882DF4AA-C4E5-4255-94DD-C3799FFF7E23"
    'Public Const EventsId As String = "FA71DC96-1733-481A-AA8A-F6FC0F1214BF"
    Public Const ClassId As String = "097F568F-F50E-4CC8-A12A-C27A55423BF2"
    Public Const InterfaceId As String = "B8A78D02-2674-451E-82BD-9808F6C68312"
    Public Const EventsId As String = "CA6F056E-2519-46CE-A4C2-97710CFA798C"
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
    ''' Implements IMetadataSynchronizer.ClassID
    ''' </summary>
    ''' <returns>A unique identifier for this IMetadataSynchronizer.</returns>
    ''' <remarks></remarks>
    ReadOnly Property ClassID1() As UID Implements IMetadataSynchronizer.ClassID
        Get
            'An ID is required to distinguish this synchronizer uniquely.
            'Safest is to use the UID of this VB Class module
            Dim myUID As New UID
            myUID.Value = ClassId
            Return myUID
        End Get
    End Property

    ''' <summary>
    ''' Implements IMetadataSynchronizer.Name
    ''' </summary>
    ''' <returns>A string containing the name of this IMetadataSynchronizer.</returns>
    ''' <remarks></remarks>
    ReadOnly Property Name() As String Implements IMetadataSynchronizer.Name
        Get
            Return "MGMG Synchronizer Helper"
        End Get
    End Property

    ''' <summary>
    ''' List of names of items being synchronized
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared scoutingData As List(Of String)

    ''' <summary>
    ''' Implements IMetadataSynchronizer.Update
    ''' </summary>
    ''' <param name="iXPS">IXmlPropertySet that contains the metadata being synchronized.</param>
    ''' <param name="itemDesc">Name of the metadata element being synchronized.</param>
    ''' <param name="value">Value for the metadata element being synchronized.</param>
    ''' <remarks>Please refer to ESRI metadata synchronization documents to understand the synchronization API.</remarks>
    Sub Update(ByVal iXPS As IXmlPropertySet, ByVal itemDesc As String, ByVal value As Object) Implements IMetadataSynchronizer.Update

        ' Collect scouting data if that's why we are called
        If SyncCommand.currentSyncLevel = SyncCommand.syncMode.helperSyncerScoutingRun Then
            scoutingData.Add(itemDesc)
            Return
        End If

        ' Do nothing if we are not called by mothership
        If SyncCommand.currentSyncLevel <> SyncCommand.syncMode.helperSyncerExecutionRun Then Return

        Dim lBracketPos As Long
        Dim sEltDesc As String
        Dim sNumPlus As String
        Dim lEltCount As Long


        'Debug.Print(itemDesc)

        'Some item descriptions have bracketed counts, e.g. Entity[0]. Strip the count off into a variable
        lBracketPos = itemDesc.IndexOf("[")

        If lBracketPos > -1 Then
            sNumPlus = itemDesc.Substring(lBracketPos + 1, itemDesc.Length - lBracketPos - 2)
            lEltCount = CLng(sNumPlus)
            sEltDesc = itemDesc.Substring(0, lBracketPos)
        Else
            sEltDesc = itemDesc
        End If

        Dim FGDCSync As New FGDCSynchronizer
        Dim delegateToFGDCSync As Boolean = False

        Select Case sEltDesc
            Case "OperatingSystem"
            Case "Software"
            Case "Environment"
            Case "Language"
            Case "CountryName"
            Case "MetadataStandard"
            Case "Boilerplate"
                ' We get here only when metadata is newly created
            Case "DatasetName"
            Case "GeoForm"
            Case "NativeForm"
            Case "GeometryType"
            Case "NativeExtent"
            Case "DDExtent"
                If My.Settings.SyncExtent Then
                    iXPS.DeleteProperty("idinfo/spdom")
                    delegateToFGDCSync = True
                End If
            Case "SpatialReference"
                If My.Settings.SyncCSI Then
                    iXPS.DeleteProperty("spref")
                    delegateToFGDCSync = True
                End If
            Case "FeatureClass"
            Case "Entity"
                If My.Settings.SyncAttrInfo Then
                    iXPS.DeleteProperty("eainfo")
                    delegateToFGDCSync = True
                End If
            Case "EntityBrief"
                If My.Settings.SyncAttrInfo Then
                    delegateToFGDCSync = True
                End If
            Case "DatasetLocation"
                If My.Settings.SyncOnlink Then
                    iXPS.DeleteProperty("idinfo/citation/citeinfo/onlink")
                    delegateToFGDCSync = True
                End If
            Case "DatasetSize"
            Case "MetadataDate"
        End Select

        If delegateToFGDCSync Then
            FGDCSync.Update(iXPS, itemDesc, value)
        End If

    End Sub

End Class


