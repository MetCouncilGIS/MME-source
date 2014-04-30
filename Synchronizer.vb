Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem
Imports System.Runtime.InteropServices


''' <summary>
''' EPA Synchronizer implementation.
''' Implements IMetadataSynchronizer interface.
''' </summary>
''' <remarks>This is more of a command control module that leaves the real work to a helper syncer.</remarks>
<ComClass(Synchronizer.ClassId, Synchronizer.InterfaceId, Synchronizer.EventsId), _
 ProgId("MnGeoMetadataEditor.Synchronizer")> _
Public Class Synchronizer
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
    'Public Const ClassId As String = "F53571D3-0BBC-4608-A0D6-34B802127401"
    'Public Const InterfaceId As String = "47A6B531-21C1-486E-A30C-47B378460A8F"
    'Public Const EventsId As String = "F7827C03-2534-410A-B1FF-67C16CBBE758"
    Public Const ClassId As String = "e3c81059-01a2-4aa2-b5a7-2370302e5b31"
    Public Const InterfaceId As String = "d0fcc47f-58d8-4539-acdb-12f14aaa6c8b"
    Public Const EventsId As String = "d998fa2e-fd55-4658-bf40-4897de6b1720"
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
            Return "MGMG Synchronizer"
        End Get
    End Property

    ''' <summary>
    ''' List of tags registered as existing prior to syncing
    ''' </summary>
    ''' <remarks>This list only includes tags at risk of being removed during sync.</remarks>
    Shared preExisting As List(Of String) = Nothing

    ''' <summary>
    ''' Stores the xml declaration (if any) prior to syncing
    ''' </summary>
    ''' <remarks></remarks>
    Shared xmlDecl As String

    ''' <summary>
    ''' Original XML prior to syncing
    ''' </summary>
    ''' <remarks></remarks>
    Shared origXml As New XmlMetadata


    ''' <summary>
    ''' Implements IMetadataSynchronizer.Update
    ''' </summary>
    ''' <param name="iXPS">IXmlPropertySet that contains the metadata being synchronized.</param>
    ''' <param name="itemDesc">Name of the metadata element being synchronized.</param>
    ''' <param name="value">Value for the metadata element being synchronized.</param>
    ''' <remarks>Please refer to ESRI metadata synchronization documents to understand the synchronization API.</remarks>
    Sub Update(ByVal iXPS As IXmlPropertySet, ByVal itemDesc As String, ByVal value As Object) Implements IMetadataSynchronizer.Update
        Debug.Print(itemDesc)

        Try
            If SyncCommand.imd Is Nothing Then
                UpdateAuto(iXPS, itemDesc, value)
            Else
                UpdateManual(iXPS, itemDesc, value)
            End If

        Catch argEx As ArgumentException
            handleExceptionInner("Argument Exception", argEx)
            Exit Try
        Catch comex As COMException
            handleExceptionInner("COM Exception", comex)
            Exit Try
        Catch ex As System.Exception
            handleExceptionInner("System Exception", ex)
            Exit Try
        Finally
        End Try
    End Sub

    ''' <summary>
    ''' Does the actual IMetadataSynchronizer.Update work if syncing was initiated by user from the toolbar.
    ''' </summary>
    ''' <param name="iXPS">IXmlPropertySet that contains the metadata being synchronized.</param>
    ''' <param name="itemDesc">Name of the metadata element being synchronized.</param>
    ''' <param name="value">Value for the metadata element being synchronized.</param>
    ''' <remarks>Please refer to ESRI metadata synchronization documents to understand the synchronization API.</remarks>
    Sub UpdateManual(ByVal iXPS As IXmlPropertySet, ByVal itemDesc As String, ByVal value As Object)
        ' Do nothing if we are not called by mothership
        If SyncCommand.currentSyncLevel <> SyncCommand.syncMode.mainSyncer AndAlso SyncCommand.currentSyncLevel <> SyncCommand.syncMode.none Then Exit Sub

        Dim xmd As New XmlMetadata
        Dim tmp As IXmlPropertySet2 = iXPS
        xmd.SetXml(tmp.GetXml(""))

        ' Do the following one time at the beginning
        If preExisting Is Nothing Then
            xmlDecl = captureXmlDecl(iXPS)
            preSyncBuildup(xmd, iXPS.IsNew)

            Try
                'DBG tmp = SyncCommand.imd.Metadata
                SyncCommand.currentSyncLevel = SyncCommand.syncMode.helperSyncerScoutingRun
                SynchronizerHelper.scoutingData = New List(Of String)
                SyncCommand.imd.Synchronize(esriMetadataSyncAction.esriMSAOverwrite, 11)
            Finally
                SyncCommand.currentSyncLevel = SyncCommand.syncMode.mainSyncer
            End Try
        End If

        If itemDesc = SynchronizerHelper.scoutingData(SynchronizerHelper.scoutingData.Count - 1) Then
            Try
                'If SyncCommand.imd Is Nothing Then Exit Sub
                Try
                    SyncCommand.currentSyncLevel = SyncCommand.syncMode.helperSyncerExecutionRun
                    SyncCommand.imd.Synchronize(esriMetadataSyncAction.esriMSAOverwrite, 11)
                Finally
                    SyncCommand.currentSyncLevel = SyncCommand.syncMode.mainSyncer
                End Try

                ' Recapture metadata force-synced by us
                tmp = SyncCommand.imd.Metadata
                xmd.SetXml(tmp.GetXml(""))
                ' Post-process
                slurp(xmd)
                postSyncCleanup(xmd)

                ' Convert back to AO API
                Dim ixps2 As IXmlPropertySet2
                ixps2 = iXPS
                ixps2.SetXml(xmlDecl + xmd.GetXml("/metadata"))
            Finally
                preExisting = Nothing
                SyncCommand.imd = Nothing
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Does the actual IMetadataSynchronizer.Update work if syncing was initiated automatically by ArcCatalog.
    ''' </summary>
    ''' <param name="iXPS">IXmlPropertySet that contains the metadata being synchronized.</param>
    ''' <param name="itemDesc">Name of the metadata element being synchronized.</param>
    ''' <param name="value">Value for the metadata element being synchronized.</param>
    ''' <remarks>Please refer to ESRI metadata synchronization documents to understand the synchronization API.</remarks>
    Sub UpdateAuto(ByVal iXPS As IXmlPropertySet, ByVal itemDesc As String, ByVal value As Object)
        Dim xmd As New XmlMetadata
        Dim tmp As IXmlPropertySet2 = iXPS
        xmd.SetXml(tmp.GetXml(""))

        xmlDecl = captureXmlDecl(iXPS)
        preSyncBuildup(xmd, iXPS.IsNew)


        Try
            Try
                SyncCommand.currentSyncLevel = SyncCommand.syncMode.helperSyncerExecutionRun
                Dim sh As New SynchronizerHelper
                sh.Update(iXPS, itemDesc, value)
            Finally
                SyncCommand.currentSyncLevel = SyncCommand.syncMode.mainSyncer
            End Try

            ' Recapture metadata force-synced by us
            xmd.SetXml(tmp.GetXml(""))
            ' Post-process
            slurp(xmd)
            postSyncCleanup(xmd)

            ' Convert back to AO API
            Dim ixps2 As IXmlPropertySet2
            ixps2 = iXPS
            ixps2.SetXml(xmlDecl + xmd.GetXml("/metadata"))
        Finally
            preExisting = Nothing
        End Try
    End Sub



    ''' <summary>
    ''' This appears to be the only way to capture the xml declaration header fully.
    ''' </summary>
    ''' <param name="iXPS"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function captureXmlDecl(ByVal iXPS As IXmlPropertySet) As String
        captureXmlDecl = ""
        Dim tempFile As String = IO.Path.GetTempFileName() ' & ".xml"
        iXPS.SaveAsFile(Nothing, Nothing, False, tempFile)
        Dim xml As String = My.Computer.FileSystem.ReadAllText(tempFile)
        IO.File.Delete(tempFile)
        If xml.Contains("<metadata") Then captureXmlDecl = xml.Substring(0, xml.IndexOf("<metadata"))
        If Not captureXmlDecl.Contains("<?xml") Then captureXmlDecl = ""
        captureXmlDecl = captureXmlDecl.Trim + vbCrLf
    End Function

    ''' <summary>
    ''' Initialize the list of pre-existing elements at risk of being removed.
    ''' </summary>
    ''' <param name="ixps"></param>
    ''' <param name="isNew"></param>
    ''' <remarks></remarks>
    Private Sub preSyncBuildup(ByVal ixps As XmlMetadata, Optional ByVal isNew As Boolean = False)
        ' Save original xml
        origXml.SetXml(ixps.GetXml(""))
        'Debug.Print("distinfo: " + ixps.CountX("distinfo").ToString)
        'Debug.Print("distInfo: " + ixps.CountX("distInfo").ToString)
        'Debug.Print("---------------------------------------")
        preExisting = New List(Of String)
        ' If new metadata, nothing should be registered
        If isNew Then Return
        ' Anything we may want to delete later, we need to make sure it's not already there.
        ' We shouldn't delete stuff that we didn't create during sync.
        registerIfPreExisting(ixps, "distinfo/stdorder/digform/digtinfo/transize")
        registerIfPreExisting(ixps, "distinfo/stdorder/digform/digtinfo/dssize")
        registerIfPreExisting(ixps, "distinfo/stdorder/digform/digtinfo")
        registerIfPreExisting(ixps, "distinfo/stdorder/digform")
        registerIfPreExisting(ixps, "distinfo/stdorder")
        registerIfPreExisting(ixps, "distinfo")
        registerIfPreExisting(ixps, "eainfo/detailed/enttyp/enttypt")
        registerIfPreExisting(ixps, "eainfo/detailed/enttyp/enttypc")
        registerIfPreExisting(ixps, "eainfo/detailed/attr/attalias")
        registerIfPreExisting(ixps, "eainfo/detailed/attr/attrtype")
        registerIfPreExisting(ixps, "eainfo/detailed/attr/attwidth")
        registerIfPreExisting(ixps, "eainfo/detailed/attr/atprecis")
        registerIfPreExisting(ixps, "eainfo/detailed/attr/attscale")
        registerIfPreExisting(ixps, "eainfo/detailed/attr/atnumdec")
        registerIfPreExisting(ixps, "idinfo/natvform")
        ' Esri element
        registerIfPreExisting(ixps, "Esri")
        ' Esri ISO elements
        registerIfPreExisting(ixps, "dataIdInfo")
        registerIfPreExisting(ixps, "distInfo")
        registerIfPreExisting(ixps, "refSysInfo")
        registerIfPreExisting(ixps, "spatRefInfo")
        registerIfPreExisting(ixps, "spatRepInfo")
        registerIfPreExisting(ixps, "mdChar")
        registerIfPreExisting(ixps, "mdDateSt")
        registerIfPreExisting(ixps, "mdHrLv")
        registerIfPreExisting(ixps, "mdHrLvName")
        registerIfPreExisting(ixps, "mdFileID")
        registerIfPreExisting(ixps, "mdLang")
        registerIfPreExisting(ixps, "mdParentID")
        registerIfPreExisting(ixps, "mdContact")
        registerIfPreExisting(ixps, "mdStanName")
        registerIfPreExisting(ixps, "mdStanVer")
        registerIfPreExisting(ixps, "appSchInfo")
        registerIfPreExisting(ixps, "porCatInfo")
        registerIfPreExisting(ixps, "mdMaint")
        registerIfPreExisting(ixps, "mdConst")
        registerIfPreExisting(ixps, "dqInfo")
        registerIfPreExisting(ixps, "contInfo")
        registerIfPreExisting(ixps, "mdExtInfo")

    End Sub

    ''' <summary>
    ''' Unwanted side effects of the FGDC Synchronizer are cleaned up here (for FGDC compliance)
    ''' </summary>
    ''' <param name="ixps">IXmlPropertySet that contains the metadata being synchronized.</param>
    ''' <remarks>We only clean up unwanted tags that were generated during our sync operations.</remarks>
    Private Sub postSyncCleanup(ByVal ixps As XmlMetadata)
        'Debug.Print("@@ " & CType(ixps, IXmlPropertySet2).GetXml("/metadata").Length)

        ' We don't want transize or dssize, so delete them including any parents that did not exist before sync
        deleteIfNotPreExisting(ixps, "distinfo")
        deleteIfNotPreExisting(ixps, "distinfo/stdorder")
        deleteIfNotPreExisting(ixps, "distinfo/stdorder/digform")
        deleteIfNotPreExisting(ixps, "distinfo/stdorder/digform/digtinfo")
        deleteIfNotPreExisting(ixps, "distinfo/stdorder/digform/digtinfo/transize")
        deleteIfNotPreExisting(ixps, "distinfo/stdorder/digform/digtinfo/dssize")

        ' We don't want ESRI profile additions to attribute information
        deleteIfNotPreExisting(ixps, "eainfo/detailed/enttyp/enttypt")
        deleteIfNotPreExisting(ixps, "eainfo/detailed/enttyp/enttypc")
        deleteIfNotPreExisting(ixps, "eainfo/detailed/attr/attalias")
        deleteIfNotPreExisting(ixps, "eainfo/detailed/attr/attrtype")
        deleteIfNotPreExisting(ixps, "eainfo/detailed/attr/attwidth")
        deleteIfNotPreExisting(ixps, "eainfo/detailed/attr/atprecis")
        deleteIfNotPreExisting(ixps, "eainfo/detailed/attr/attscale")
        deleteIfNotPreExisting(ixps, "eainfo/detailed/attr/atnumdec")
        deleteIfNotPreExisting(ixps, "idinfo/natvform")
        'ixps.SetPropertyX("spref/horizsys/planar/planci/coordrep/absres[.=""0.000000""]", "0.000001", esriXmlPropertyType.esriXPTText, esriXmlSetPropertyAction.esriXSPAReplaceIfExists, True)
        'ixps.SetPropertyX("spref/horizsys/planar/planci/coordrep/ordres[.=""0.000000""]", "0.000001", esriXmlPropertyType.esriXPTText, esriXmlSetPropertyAction.esriXSPAReplaceIfExists, True)
        'ixps.SetPropertyX("latres[.=""8.9831528411952117e-009""]", "0.000001", esriXmlPropertyType.esriXPTText, esriXmlSetPropertyAction.esriXSPAReplaceIfExists, True)
        'ixps.SetPropertyX("longres[.=""8.9831528411952117e-009""]", "0.000001", esriXmlPropertyType.esriXPTText, esriXmlSetPropertyAction.esriXSPAReplaceIfExists, True)

        ' Esri element
        deleteIfNotPreExisting(ixps, "Esri")
        ' Esri ISO elements
        deleteIfNotPreExisting(ixps, "dataIdInfo")
        deleteIfNotPreExisting(ixps, "distInfo")
        deleteIfNotPreExisting(ixps, "refSysInfo")
        deleteIfNotPreExisting(ixps, "spatRefInfo")
        deleteIfNotPreExisting(ixps, "spatRepInfo")
        deleteIfNotPreExisting(ixps, "mdChar")
        deleteIfNotPreExisting(ixps, "mdDateSt")
        deleteIfNotPreExisting(ixps, "mdHrLv")
        deleteIfNotPreExisting(ixps, "mdHrLvName")
        deleteIfNotPreExisting(ixps, "mdFileID")
        deleteIfNotPreExisting(ixps, "mdLang")
        deleteIfNotPreExisting(ixps, "mdParentID")
        deleteIfNotPreExisting(ixps, "mdContact")
        deleteIfNotPreExisting(ixps, "mdStanName")
        deleteIfNotPreExisting(ixps, "mdStanVer")
        deleteIfNotPreExisting(ixps, "appSchInfo")
        deleteIfNotPreExisting(ixps, "porCatInfo")
        deleteIfNotPreExisting(ixps, "mdMaint")
        deleteIfNotPreExisting(ixps, "mdConst")
        deleteIfNotPreExisting(ixps, "dqInfo")
        deleteIfNotPreExisting(ixps, "contInfo")
        deleteIfNotPreExisting(ixps, "mdExtInfo")

        If My.Settings.WipeEsriTagsOnSync Then Utils.wipeEsriTags(ixps)
        'Debug.Print("@@ " & CType(ixps, IXmlPropertySet2).GetXml("/metadata").Length)
    End Sub

    ''' <summary>
    ''' Registers the name as a tag that existed before sync operation made any modifications to metadata.
    ''' </summary>
    ''' <param name="ixps">IXmlPropertySet that contains the metadata being synchronized.</param>
    ''' <param name="name">The xsl pattern identifying the element being synchronized.</param>
    ''' <remarks></remarks>
    Private Sub registerIfPreExisting(ByVal ixps As XmlMetadata, ByVal name As String)
        If ixps.CountX(name) > 0 AndAlso Not preExisting.Contains(name) Then
            preExisting.Add(name)
        End If
    End Sub

    ''' <summary>
    ''' Deletes the name if it did not exist before sync operation made any modifications to metadata.
    ''' </summary>
    ''' <param name="ixps">IXmlPropertySet that contains the metadata being synchronized.</param>
    ''' <param name="name">The xsl pattern identifying the element being synchronized.</param>
    ''' <remarks></remarks>
    Private Sub deleteIfNotPreExisting(ByVal ixps As XmlMetadata, ByVal name As String)
        If Not preExisting.Contains(name) Then
            'Debug.Print("distinfo: " + ixps.CountX("distinfo").ToString)
            'Debug.Print("distInfo: " + ixps.CountX("distInfo").ToString)
            'Debug.Print("Deleting: " + name)
            'Debug.Print(CType(ixps, IXmlPropertySet2).GetXml("/metadata").Length)
            ixps.DeleteProperty(name)
            'Debug.Print("distinfo: " + ixps.CountX("distinfo").ToString)
            'Debug.Print("distInfo: " + ixps.CountX("distInfo").ToString)
            'Debug.Print("---------------------------------------")
            'Debug.Print(CType(ixps, IXmlPropertySet2).GetXml("/metadata").Length)
            'Debug.Print("")
        End If
    End Sub

    ''' <summary>
    ''' Exception handler. Used only for debugging.
    ''' </summary>
    ''' <param name="typ">A string describing the type of exception caught.</param>
    ''' <param name="ex">Exception object.</param>
    ''' <remarks></remarks>
    Private Shared Sub handleExceptionInner(ByVal typ As String, ByVal ex As Exception)
        Dim msgStr As String = typ & vbCrLf & ex.ToString
        'DSASUtility.log(TraceLevel.Info, "********EXCEPTION INFO BEGIN*********")
        'DSASUtility.log(TraceLevel.Info, typ)
        'DSASUtility.log(TraceLevel.Info, "Exception Source:")
        'DSASUtility.log(TraceLevel.Info, ex.Source)
        'DSASUtility.log(TraceLevel.Info, "Exception Target Site:")
        'DSASUtility.log(TraceLevel.Info, ex.TargetSite.ToString)
        'DSASUtility.log(TraceLevel.Info, "Exception Message:")
        'DSASUtility.log(TraceLevel.Info, ex.Message)
        'DSASUtility.log(TraceLevel.Error, msgStr)
        'DSASUtility.log(TraceLevel.Info, "********EXCEPTION INFO END*********")
        'Debug.Print(msgStr)
        'MsgBox(msgStr)
    End Sub

    ''' <summary>
    ''' Extract FGDC elements from synced metadata by applying ArcGIS to FGDC transformation.
    ''' </summary>
    ''' <param name="mdSync">Synced metadata record</param>
    ''' <remarks></remarks>
    Private Shared Sub slurp(ByVal mdSync As XmlMetadata)
        Dim xslt As New Xml.Xsl.XslCompiledTransform
        Dim xSettings As New Xml.Xsl.XsltSettings
        xSettings.EnableDocumentFunction = True

        Try
            'xslt.Load("C:\Program Files (x86)\ArcGIS\Desktop10.0\Metadata\Translator\Transforms\ArcGIS2FGDC.xsl", xSettings, Nothing)
            'xslt.Load(System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + "\ArcGIS\Desktop10.0\Metadata\Translator\Transforms\ArcGIS2FGDC.xsl", xSettings, Nothing)
            xslt.Load(AoUtils.getArcgisInstallDir + "Metadata\Translator\Transforms\ArcGIS2FGDC.xsl", xSettings, Nothing)

            'http://support.microsoft.com/kb/323370

            ' Instantiate an XsltArgumentList object.
            ' An XsltArgumentList object is used to supply extension object instances
            ' and values for XSLT paarmeters required for an XSLT transformation	    
            Dim xsltArgList As New System.Xml.Xsl.XsltArgumentList()

            ' Instantiate and add an instance of the extension object to the XsltArgumentList.
            ' The AddExtensionObject method is used to add the Extension object instance to the
            ' XsltArgumentList object. The namespace URI specified as the first parameter 
            ' should match the namespace URI used to reference the extension object in the
            ' XSLT style sheet.
            Dim esri As New EsriXslExtSim()
            xsltArgList.AddExtensionObject("http://www.esri.com/metadata/", esri)


            Dim insr As New System.IO.StringReader(mdSync.GetXml(""))
            Dim inr As New System.Xml.XmlTextReader(insr)
            Dim outsb As New System.Text.StringBuilder()
            Dim outw As System.Xml.XmlWriter = System.Xml.XmlTextWriter.Create(outsb, xslt.OutputSettings)

            xslt.Transform(inr, xsltArgList, outw)

            'Dim mdMain As New XmlMetadata
            Dim mdSlurp As New XmlMetadata
            'mdMain.SetXml(Xml)
            mdSlurp.SetXml(outsb.ToString)

            ' Bring in elements of interest from slurped md into current md

            If My.Settings.SyncExtent Then mdSync.copyFrom(mdSlurp, "idinfo/spdom/bounding")

            If My.Settings.SyncAttrInfo Then syncEntInfo(mdSync, mdSlurp)

            If My.Settings.SyncOnlink Then mdSync.copyFrom(mdSync, "Esri/DataProperties/itemProps[1]/itemLocation[1]/linkage[1]", "idinfo/citation/citeinfo/onlink[1]")

            If My.Settings.SyncCSI Then
                mdSync.copyFrom(mdSlurp, "spref")
                If mdSync.CountX("spref/horizsys/geograph/geogunit[.=""Decimal Degree""]") > 0 Then
                    mdSync.SetPropertyX("spref/horizsys/geograph/geogunit[.=""Decimal Degree""]", "Decimal Degrees")
                End If
            End If

        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Merge entity (and attribute) info in two metadata records into one.
    ''' </summary>
    ''' <param name="mdTarget">The metadata record that will receive merged entity info.</param>
    ''' <param name="mdSource">The metadata record that provides new entity info (typically from syncing)</param>
    ''' <remarks>The method we employ is only suitable for single entity datasets - which is probably 99.9% of our world.</remarks>
    Private Shared Sub syncEntInfo(ByVal mdTarget As XmlMetadata, ByVal mdSource As XmlMetadata)
        mdTarget.copyFrom(origXml, "eainfo")

        Dim adds As New List(Of String)
        Dim dels As New List(Of String)
        Dim keeps As New List(Of String)

        'Dim entNo As Integer = 1 ' mdTarget.CountX("eainfo/detailed")
        'Dim enttypl As String = origXml.SimpleGetProperty("eainfo/detailed[1]/enttyp/enttypl")

        If mdTarget.CountX("eainfo/detailed/enttyp") = 1 AndAlso mdSource.CountX("eainfo/detailed/enttyp") = 1 Then
            ' Single entity sync
            mdTarget.copyFrom(mdSource, "eainfo/detailed/enttyp[1]/enttypl")
        Else
            ' Multi-entity sync
        End If

        computeChanges(mdTarget.GetProperty("eainfo/detailed/enttyp/enttypl"), mdSource.GetProperty("eainfo/detailed/enttyp/enttypl"), adds, dels, keeps)
        'computeChanges(mdTarget.GetProperty("eainfo/detailed[enttyp/enttypl=""" + enttypl + """]/attr/attrlabl"), mdSource.GetProperty("eainfo/detailed[enttyp/enttypl=""" + enttypl + """]/attr/attrlabl"), adds, dels, keeps)

        For Each enttypl As String In dels
            mdTarget.DeleteProperty("eainfo/detailed[enttyp/enttypl=""" + enttypl + """]")
        Next

        Dim entNo As Integer = mdTarget.CountX("eainfo/detailed")

        For Each enttypl As String In adds
            entNo += 1
            mdTarget.copyFrom(mdSource, "eainfo/detailed[enttyp/enttypl=""" + enttypl + """]", "eainfo/detailed[" + entNo.ToString + "]")
        Next

        For Each enttypl As String In keeps
            syncAttrInfo(mdTarget, mdSource, enttypl)
            'mdTarget.SetPropertyX("eainfo/detailed[enttyp/enttypl=""" + enttypl + """]/attr[attrlabl=""" + attrlabl + """]", mdSource.SimpleGetProperty("eainfo/detailed[enttyp/enttypl=""" + enttypl + """]/attr/attrlabl"))
        Next
        Debug.Print("source: " + Str(mdSource.CountX("//attrdef")))
        Debug.Print("target: " + Str(mdTarget.CountX("//attrdef")))
    End Sub

    ''' <summary>
    ''' Merge attribute info in two metadata records into one.
    ''' </summary>
    ''' <param name="mdTarget">The metadata record that will receive merged entity info.</param>
    ''' <param name="mdSource">The metadata record that provides new entity info (typically from syncing)</param>
    ''' <param name="enttypl">The label of the entity to merge.</param>
    ''' <remarks>The method we employ is only suitable for single entity datasets - which is probably 99.9% of our world.</remarks>
    Private Shared Sub syncAttrInfo(ByVal mdTarget As XmlMetadata, ByVal mdSource As XmlMetadata, ByVal enttypl As String)
        Dim adds As New List(Of String)
        Dim dels As New List(Of String)
        Dim keeps As New List(Of String)

        Dim entNo As Integer = mdTarget.CountX("eainfo/detailed[enttyp/enttypl=""" + enttypl + """]")
        ' Dim enttypl As String = origXml.SimpleGetProperty("eainfo/detailed[1]/enttyp/enttypl") ' "Group5_VerImp_May_09"

        'computeChanges(mdTarget.GetProperty("eainfo/detailed/enttyp/enttypl"), mdSource.GetProperty("eainfo/detailed/enttyp/enttypl"), adds, dels, keeps)
        computeChanges(mdTarget.GetProperty("eainfo/detailed[enttyp/enttypl=""" + enttypl + """]/attr/attrlabl"), mdSource.GetProperty("eainfo/detailed[enttyp/enttypl=""" + enttypl + """]/attr/attrlabl"), adds, dels, keeps)

        For Each attrlabl As String In dels
            mdTarget.DeleteProperty("eainfo/detailed[enttyp/enttypl=""" + enttypl + """]/attr[attrlabl=""" + attrlabl + """]")
        Next

        Dim attrNo As Integer = mdTarget.CountX("eainfo/detailed/attr") 'AE TODO: Will this work if multiple entities?

        For Each attrlabl As String In adds
            attrNo += 1
            'mdTarget.copyFrom(mdSource, "eainfo/detailed[enttyp/enttypl=""" + enttypl + """]/attr[attrlabl=""" + attrlabl + """]", "eainfo/detailed[enttyp/enttypl=""" + enttypl + """]/attr[" + attrNo.ToString + "]")
            mdTarget.copyFrom(mdSource, "eainfo/detailed[enttyp/enttypl=""" + enttypl + """]/attr[attrlabl=""" + attrlabl + """]", "eainfo/detailed[" + entNo.ToString + "]/attr[" + attrNo.ToString + "]")
        Next

        For Each attrlabl As String In keeps
            'mdTarget.SetPropertyX("eainfo/detailed[enttyp/enttypl=""" + enttypl + """]/attr[attrlabl=""" + attrlabl + """]", mdSource.SimpleGetProperty("eainfo/detailed[enttyp/enttypl=""" + enttypl + """]/attr/attrlabl"))
        Next
    End Sub

    ''' <summary>
    ''' Compute the changes between two lists, recording the names to add to, delete from or kept as-is in old list to derive the new list.
    ''' </summary>
    ''' <param name="oldList">Original list of names</param>
    ''' <param name="newList">Current list of names</param>
    ''' <param name="adds">List of names to add (to be computed) - passed as reference</param>
    ''' <param name="dels">List of names to delete (to be computed) - passed as reference</param>
    ''' <param name="keeps">List of names to keep (to be computed) - passed as reference</param>
    ''' <remarks></remarks>
    Shared Sub computeChanges(ByVal oldList As List(Of String), ByVal newList As List(Of String), ByRef adds As List(Of String), ByRef dels As List(Of String), ByRef keeps As List(Of String))
        If newList Is Nothing Then newList = New List(Of String)
        If oldList Is Nothing Then oldList = New List(Of String)

        adds.AddRange(newList)
        dels.AddRange(oldList)

        For Each item As String In oldList
            If adds.Contains(item) Then
                adds.Remove(item)
                keeps.Add(item)
            End If
        Next

        For Each item As String In newList
            If dels.Contains(item) Then
                dels.Remove(item)
            End If
        Next
    End Sub

End Class


