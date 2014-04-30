Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.Geodatabase
Imports Microsoft.Win32


''' <summary>
''' Metadata utilities that need to reference ArcObjects libraries are defined here.
''' </summary>
''' <remarks></remarks>
Module AoUtils

    ''' <summary>
    ''' Values representing an object's metadata capability (from the user's perspective).
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum metability
        CanHaveMetadata
        CanWriteMetadata
        HasMetadata
    End Enum

    ''' <summary>
    ''' Get the objects that are selected in ArcCatalog UI
    ''' </summary>
    ''' <param name="gxApp"></param>
    ''' <param name="filter"></param>
    ''' <returns>List of IGxObject objects that are selected.</returns>
    ''' <remarks>This works whether selection in made in contents pane or TOC view.</remarks>
    Function getSelectedObjects(ByRef gxApp As IGxApplication, Optional ByVal filter As metability = Nothing) As List(Of IGxObject)
        Dim tmp = New List(Of IGxObject)

        Dim selEnum As IEnumGxObject = gxApp.Selection.SelectedObjects
        Dim o As IGxObject

        ' Try to get selected objects from the contents view
        Do
            o = selEnum.Next
            If o Is Nothing Then Exit Do
            'MsgBox(o.FullName)

            Dim md As IMetadata = o 'CType(o, IMetadata)
            tmp.Add(md)
            'Dim iXPSv As IXmlPropertySet2 = CType(md.Metadata, IXmlPropertySet2)
        Loop

        ' If we didn't get any selected objects from the contents view, try to get from the tree view
        If tmp.Count = 0 AndAlso gxApp.SelectedObject IsNot Nothing Then
            tmp.Add(gxApp.SelectedObject)
        End If

        If filter = Nothing Then
            getSelectedObjects = tmp
        Else
            getSelectedObjects = New List(Of IGxObject)

            ' Pick the ones that fit the criteria
            For Each o In tmp
                If (filter = metability.HasMetadata AndAlso hasMetadata(o)) OrElse (filter = metability.CanHaveMetadata AndAlso canHaveMetadata(o)) OrElse (filter = metability.CanWriteMetadata AndAlso canWriteMetadata(o)) Then
                    getSelectedObjects.Add(o)
                End If
            Next
        End If

    End Function

    ''' <summary>
    ''' Determine if we can write metadata to the given object.
    ''' </summary>
    ''' <param name="pGxObj">The object being queired for metadata capability.</param>
    ''' <returns>True if we can write metadata to the goven object, False otherwise.</returns>
    ''' <remarks></remarks>
    Public Function canWriteMetadata(ByVal pGxObj As IGxObject) As Boolean
        Return canHaveMetadata(pGxObj) AndAlso CType(pGxObj, IMetadataEdit).CanEditMetadata
    End Function


    ''' <summary>
    ''' Determine if an object has metadata.
    ''' </summary>
    ''' <param name="pGxObj">The object being queried for metadata.</param>
    ''' <returns>True if the object has metadata. False otherwise.</returns>
    ''' <remarks></remarks>
    Public Function hasMetadata(ByVal pGxObj As IGxObject) As Boolean
        Try
            If canHaveMetadata(pGxObj) Then
                Dim g_pMD As IMetadata = CType(pGxObj, IMetadata)
                Dim g_pXPS As IXmlPropertySet = g_pMD.Metadata

                'If the metadata was not created just now, we are interested.
                Return Not g_pXPS.IsNew
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Determine if an object can have metadata.
    ''' </summary>
    ''' <param name="pGxObj">The object being queried.</param>
    ''' <returns>True if the object can have metadata. False otherwise.</returns>
    ''' <remarks></remarks>
    Public Function canHaveMetadata(ByVal pGxObj As IGxObject) As Boolean
        Try
            Return TypeOf pGxObj Is IMetadata AndAlso pGxObj.IsValid
        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Function

    ''' <summary>
    ''' Get the location of ArcGIS installation
    ''' </summary>
    ''' <returns>The full path of the directory where ArcGIS Desktop has been installed.</returns>
    ''' <remarks></remarks>
    Public Function getArcgisInstallDir() As String
        getArcgisInstallDir = Nothing
        Dim keyObj As RegistryKey
        Dim keyStr As String = "SOFTWARE\ESRI\Desktop"
        Dim keyStr1 As String = "SOFTWARE\ESRI\Desktop10.0"
        Dim keyStr2 As String = "SOFTWARE\ESRI\Desktop10.1"
        'Dim keyStr As String = "SOFTWARE\Wow6432Node\ESRI\Desktop10.0"
        'Dim keyStr As String = "SOFTWARE\Microsoft\Windows\CurrentVersion"
        keyObj = Registry.LocalMachine.OpenSubKey(keyStr, False)
        If (keyObj Is Nothing) Then
            keyObj = Registry.LocalMachine.OpenSubKey(keyStr1, False)
            If (keyObj Is Nothing) Then
                keyObj = Registry.LocalMachine.OpenSubKey(keyStr2, False)
            End If
        End If
        If (Not keyObj Is Nothing) Then
            getArcgisInstallDir = keyObj.GetValue("InstallDir", "")
            keyObj.Close()
        End If
    End Function

    ''' <summary>
    ''' Get a property value.
    ''' </summary>
    ''' <param name="imd">The IMetadata object beign queried.</param>
    ''' <param name="propName">The name (effectively the xpath expression) for the property</param>
    ''' <returns>Value of the property.</returns>
    ''' <remarks></remarks>
    Public Function getMdProperty(ByVal imd As IMetadata, ByVal propName As String) As String
        Dim o() As Object = imd.Metadata.GetProperty(propName)
        If o Is Nothing Then Return ""
        Return o(0).ToString.Trim
    End Function


End Module
