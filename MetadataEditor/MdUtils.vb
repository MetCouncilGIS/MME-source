Imports System.Runtime.InteropServices
Imports System.Xml
Imports System.Threading

''' <summary>
''' Loose collection of metadata file utilities
''' </summary>
''' <remarks></remarks>
Module MdUtils

    ''' <summary>
    ''' Quit hint. Set to true when the application is expected to exit based on user interaction
    ''' </summary>
    ''' <remarks></remarks>
    Public quit As Boolean = False
    ''' <summary>
    ''' The File to load next. Full file path name.
    ''' </summary>
    ''' <remarks></remarks>
    Public loadNext As String = Nothing
    ''' <summary>
    ''' File currently being edited. Full file path name.
    ''' </summary>
    ''' <remarks></remarks>
    Public currentlyEditing As String = Nothing

    ''' <summary>
    ''' Load a default metadata record which is essentially empty.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub loadDefaultMd()
        editMd("<?xml version=""1.0""?>" + vbCrLf + "<metadata></metadata>")
    End Sub

    ''' <summary>
    ''' Extract metadata filename from command-line arguments if given.
    ''' First argument is assumed to be the filename (full path).
    ''' </summary>
    ''' <returns>The metadata filename if passed as command-line argument, Nothing otherwise</returns>
    ''' <remarks></remarks>
    Function getFilenameFromCommandLine() As String
        If Environment.GetCommandLineArgs.GetLength(0) > 1 Then
            Dim mdFile As String = Environment.GetCommandLineArgs(1).Trim
            If mdFile <> "" Then
                Return mdFile
            End If
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' Open a file dialog to prompt for metadata file selection.
    ''' Selected filename (full path) is stored in loadNext variable, if successful.
    ''' </summary>
    ''' <returns>True if a file selection is made, False otherwise.</returns>
    ''' <remarks></remarks>
    Public Function promptForMdFile() As Boolean
        Dim fdlg As OpenFileDialog = New OpenFileDialog()
        With fdlg
            .Title = "Load Metadata File"
            .Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
            .RestoreDirectory = True
        End With
        If fdlg.ShowDialog() = DialogResult.OK Then
            loadNext = fdlg.FileName
            Return True
        End If
    End Function

    ''' <summary>
    ''' Load the metadata file at the given path.
    ''' </summary>
    ''' <param name="mdFilename">Full file path name to the metadata file to be loaded</param>
    ''' <remarks>Checks are performed on the file to give meaningful messages if file can not be loaded.</remarks>
    Public Sub loadMdFile(ByVal mdFilename As String)
        Dim mdXml As String = "<metadata></metadata>"
        If mdFilename IsNot Nothing Then
            mdFilename = IO.Path.GetFullPath(mdFilename)
            If IO.File.Exists(mdFilename) Then
                Dim readPermission As New System.Security.Permissions.FileIOPermission(System.Security.Permissions.FileIOPermissionAccess.Read, mdFilename)

                If System.Security.SecurityManager.IsGranted(readPermission) Then
                    mdXml = System.IO.File.ReadAllText(mdFilename)
                    If mdXml.Contains("<metadata") Then
                        Dim writePermission As New System.Security.Permissions.FileIOPermission(System.Security.Permissions.FileIOPermissionAccess.Write, mdFilename)
                        If System.Security.SecurityManager.IsGranted(writePermission) Then
                            editMd(mdXml)
                        Else
                            MsgBox("You do not have write access to the specified file. Opening read only...")
                            editMd(mdXml, True)
                        End If
                    Else
                        MsgBox("Specified file does not appear to be a metadata file. Exiting...")
                    End If
                Else
                    MsgBox("You do not have access to the specified file. Exiting...")
                End If
            Else
                MsgBox("Specified file does not exist. Exiting...")
            End If
        Else
            ' This is the case where we start with empty metadata
        End If
    End Sub

    ''' <summary>
    ''' Edit the metadata record whose content is given.
    ''' </summary>
    ''' <param name="mdXml">The XML metadata record</param>
    ''' <param name="readOnlyFile">Hint to indicate that the metadata record was opened in read-only mode if true. False (read-write) by default.</param>
    ''' <remarks></remarks>
    Public Sub editMd(ByVal mdXml As String, Optional ByVal readOnlyFile As Boolean = False)
        Dim mdXPS As New XmlMetadata
        mdXPS.SetXml(mdXml)
        currentlyEditing = loadNext
        loadNext = Nothing
        If Edit(mdXPS) AndAlso Not readOnlyFile AndAlso currentlyEditing IsNot Nothing Then
            System.IO.File.WriteAllText(currentlyEditing, mdXPS.GetXml(""))
        End If
        currentlyEditing = Nothing
    End Sub

    ''' <summary>
    ''' Main entry point into metadata editor.
    ''' </summary>
    ''' <param name="metadata">Metadata record copy to be edited.</param>
    ''' <returns>
    ''' Boolean return value indicates if the metadata record was modified by the editor
    ''' </returns>
    ''' <remarks></remarks>
    Public Function Edit(ByVal metadata As XmlMetadata) As Boolean
        Try
            'Disable splash screen 20081030
            'Dim ts As Thread = New Thread(New ThreadStart(AddressOf splash))
            'ts.Start()


            If Not GlobalVars.enabled Then
                Return False
            End If

            If GlobalVars.mdbPath Is Nothing Then
                MsgBox( _
                    My.Settings.CommonEditorAbbreviation + " is unable to find its database in either the configured location:" & _
                    vbCrLf & My.Settings.MdbFilepathname & vbCrLf & _
                    "or the default location:" & vbCrLf & Utils.getAppFolder() & "\metadata.mdb" & _
                    vbCrLf & My.Settings.CommonEditorAbbreviation & " will not function until this is corrected.")
                Return False
            End If

            iXPS = metadata
            GlobalVars.init()

            ' See if user wants to edit a saved session.
            Utils.promptAndRecoverSavedSession()

            '' Ensure that keyword theasarus tags are there (inserting if not).
            'iXPS.checkCreateKTTags()

            ' Open the editor form as a modal dialog
            Dim frm As New EditorForm
            frm.ShowDialog()

            ''Ensure that unused keyword theasarus tags are removed when editor is closed
            'iXPS.checkDeleteKTTags()

            ' Explicitly releasing the metadata object.
            iXPS = Nothing


            ' When the form is closed, we return a boolean value indicating
            ' whether or not the metadata was modified. That information is
            ' passed back to the ArcCatalog -- if edited, the metadata at the source 
            ' will be updated and metadata view is refreshed so you can see the changes.

            Return frm.saveHint

            'Return frm.FormChanged = Modified.Dirty OrElse GlobalVars.savedSession OrElse GlobalVars.recoveredSession
            'Return True
        Catch ex As Exception
            ' No well-defined mechanism in place for reporting errors.
            ' Uncomment following line when debugging
            'Utils.ErrorHandler(ex)
        End Try
    End Function


    ''' <summary>
    ''' Put up a splash screen.
    ''' </summary>
    ''' <remarks>Disabled but left here for other users of EME source code.</remarks>
    Private Sub splash()
        Dim ss As SplashScreen
        Try
            ss = New SplashScreen
            ss.Show()
            ss.Refresh()
            Thread.Sleep(4000)
            If ss IsNot Nothing AndAlso Not ss.Disposing Then ss.Close()
        Finally
            ss = Nothing
        End Try
    End Sub

End Module
