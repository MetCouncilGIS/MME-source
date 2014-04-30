Imports System.Xml
Imports System.Xml.XmlTextReader
Imports System.IO
Imports System.IO.Path
Imports System.Reflection.Assembly
Imports System.Collections.Generic

Module LocalValidator
    ''' <summary>
    ''' Generate an XML snippet (for use in metadata validation error reporting) with the provided information.
    ''' </summary>
    ''' <param name="type">Type of validation issue</param>
    ''' <param name="message">Message describing the issue</param>
    ''' <param name="linenum">Metadata XML file line number if available</param>
    ''' <param name="errid">An identifier for the element with the reported issue if available. Constructed using the XPath for the element</param>
    ''' <returns>Returns a string containing the XML snippet.</returns>
    ''' <remarks></remarks>
    Private Function errorxml(ByVal type As String, ByVal message As String, ByVal linenum As String, ByVal errid As String) As String
        Return "  <err>" & vbCrLf & _
        "    <type>" & type & "</type>" & vbCrLf & _
        "    <message>" & escapeXml(message) & "</message>" & vbCrLf & _
        "    <linenum>" & linenum & "</linenum>" & vbCrLf & _
        "    <errid>" & errid & "</errid>" & vbCrLf & _
        "  </err>" & vbCrLf
    End Function

    ''' <summary>
    ''' Replace special XML characters with their escape codes
    ''' </summary>
    ''' <param name="txt">Text to be XML-escaped</param>
    ''' <returns>A String with content XML-escaped.</returns>
    ''' <remarks></remarks>
    Public Function escapeXml(ByVal txt As String) As String
        txt = txt.Replace("&", "&amp;")
        txt = txt.Replace("'", "&apos;")
        txt = txt.Replace("""", "&quot;")
        txt = txt.Replace(">", "&gt;")
        txt = txt.Replace("<", "&lt;")
        Return txt
    End Function

    ''' <summary>
    ''' Perform MP validation on the provided XML metadata record.
    ''' </summary>
    ''' <param name="root_dir">Root directory for EME.</param>
    ''' <param name="inxml">Metadata XML to be validated.</param>
    ''' <returns>Returns an XML snippet containing issues found with the metadata record</returns>
    ''' <remarks></remarks>
    Function validate_mp(ByVal root_dir As String, ByVal inxml As String) As String
        Dim ans As String = ""
        Dim mycmd As New System.Diagnostics.Process
        'Dim currDir As String
        Dim fname As String

        fname = System.IO.Path.GetRandomFileName

        If inxml.IndexOf(Chr(10)) < 1 OrElse inxml.IndexOf(Chr(13)) < 1 Then
            Return "<errs>" & vbCrLf & errorxml("Global", "Mp requires a newline", "", "") & "</errs>"
        End If

        Try
            file_in("mp", fname, inxml)

            Dim workingDir As String = type_to_dir("mp")
            Dim inputFile As String = workingDir & fname & ".errid.xml"
            Dim errorFileName As String = workingDir & fname & ".errid.err"

            'Parse the errors/warnings in .err file
            If runMP(workingDir, inputFile, errorFileName, ans) Then
                ans = parseErrors(errorFileName, inputFile, ans)
            End If

            ' Get rid of the temp files we created.
            delete_files("mp", fname)
            
            return ans
            
        Catch ex As Exception
            ans = "<errs>" & vbCrLf & errorxml("Global", "Problem in mp " & ex.ToString, "", "") & "</errs>"
        End Try

        Return ans
    End Function

    ''' <summary>
    ''' Perform EPA validation on the provided XML metadata record.
    ''' </summary>
    ''' <param name="root_dir">Root directory for EME</param>
    ''' <param name="inxml">Metadata XML to be validated</param>
    ''' <returns>Returns an XML snippet containing issues found with the metadata record</returns>
    ''' <remarks>Utilizes a Schematron generated XSL transform to identify issues with respect to metadata rules that EPA enforces. 
    ''' Needs to be made optional to better accomodate non-EPA users.</remarks>
    Function validate_epa(ByVal root_dir As String, ByVal inxml As String) As String

        Dim ans As String
        Dim epadir As String
        Dim fname As String

        fname = System.IO.Path.GetRandomFileName

        epadir = type_to_dir("mgmg")
        ans = errorxml("Global", "Unimplemented", "", "")
        file_in("mgmg", fname, inxml)

        Try
            transform_epa(epadir, "epa_validator.xsl", fname & ".errid.xml", fname & "err.svrl")
            transform_epa(epadir, "svrl_to_xsl.xsl", fname & "err.svrl", fname & "svrl_to_err.xsl")
            transform_epa(epadir, fname & "svrl_to_err.xsl", fname & ".errid.xml", fname & ".errs.xml")
        Catch ex As Exception
            ans = errorxml("Global", ex.ToString, "", "")
            Return ans
        End Try

        Try
            ans = file_out("mgmg", fname)

            delete_files("mgmg", fname)
        Catch ex As Exception
            ans = errorxml("Global", "Problem with reading answer" & ex.ToString, "", "")
        End Try

        Return ans

    End Function

    ''' <summary>
    ''' Apply an XSL transform to an XML file.
    ''' </summary>
    ''' <param name="xslfile">Full filepathname of the XSLT file</param>
    ''' <param name="fin">Full filepathname of the input file</param>
    ''' <param name="fout">Full filepathname of the output file</param>
    ''' <remarks></remarks>
    Private Sub transform_epa(ByVal workdir As String, ByVal xslfile As String, ByVal fin As String, ByVal fout As String)
        Dim xslt As New Xml.Xsl.XslCompiledTransform
        Dim xSettings As New Xml.Xsl.XsltSettings

        xSettings.EnableDocumentFunction = True
        Try
            xslt.Load(workdir & xslfile, xSettings, Nothing)
            xslt.Transform(workdir & fin, workdir & fout)
        Catch ex As Exception
            Throw New ApplicationException("Transform problem with " & xslfile & ", " & fin & ", " & fout & ": " & ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' Given the type of validation, return the subdirectory name to use.
    ''' </summary>
    ''' <param name="type">Validation type</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function type_to_dir(ByVal type As String) As String
        'Dim ans As String
        'ans = "mp_interface\"
        'If type = "epa" Then
        '    ans = "epa_interface\"
        'End If
        'Return GetDirectoryName(GetExecutingAssembly().Location) & "\" & ans
        If type = "mgmg" Then
            Return Utils.getEpaFolder
        Else
            Return Utils.getMpFolder
        End If
    End Function

    ''' <summary>
    ''' Write file with given name and contents under appropriate validation subdirectory to feed into the validation system.
    ''' </summary>
    ''' <param name="type">Validation type</param>
    ''' <param name="fname">Filename</param>
    ''' <param name="contents">File contents</param>
    ''' <remarks></remarks>
    Private Sub file_in(ByVal type As String, ByVal fname As String, ByVal contents As String)
        My.Computer.FileSystem.WriteAllText(type_to_dir(type) & fname & ".errid.xml", contents, False)
    End Sub

    ''' <summary>
    ''' Read file with given name under appropriate validation subdirectory to get feed back from the validation system.
    ''' </summary>
    ''' <param name="type">Validation type</param>
    ''' <param name="fname">Filename</param>
    ''' <remarks></remarks>
    Private Function file_out(ByVal type As String, ByVal fname As String) As String
        Return My.Computer.FileSystem.ReadAllText(type_to_dir(type) & fname & ".errs.xml", System.Text.Encoding.UTF8)
    End Function

    ''' <summary>
    ''' Return date that validation rules were last updated
    ''' </summary>
    ''' <param name="root_dir">Root directory for EME install</param>
    ''' <returns></returns>
    ''' <remarks>Relies on "date_changed.xml" file. Only useful on the server side.</remarks>
    Function date_changed(ByVal root_dir As String) As String
        Dim ans As String
        Try
            ans = My.Computer.FileSystem.ReadAllText("date_changed.xml", System.Text.Encoding.UTF8)
        Catch ex As Exception
            ans = "<errs>" & vbCrLf & errorxml("Global", "Problem finding date file " & ex.ToString, "", "") & "</errs>"
            Return ans
        End Try

        Return ans
    End Function

    ''' <summary>
    ''' Delete files in MP or EPA subdirectory matching the given name specification.
    ''' </summary>
    ''' <param name="type">Type ("mp" or "epa")</param>
    ''' <param name="fname">Filename starts with this string</param>
    ''' <remarks></remarks>
    Sub delete_files(ByVal type As String, ByVal fname As String)
        Dim mydir As String = type_to_dir(type)
        Dim mydirinfo As New System.IO.DirectoryInfo(mydir)
        Dim files_to_kill As FileInfo() = mydirinfo.GetFiles(fname & "*")
        Dim finfo As System.IO.FileInfo

        For Each finfo In files_to_kill
            Try
                System.IO.File.Delete(finfo.FullName)
            Catch ex As Exception
                Throw New ApplicationException("Could not delete " & finfo.FullName & " : " & ex.ToString)
            End Try
        Next

    End Sub
    
    ''' <summary>
    ''' Parse errors and warnings in MP output.
    ''' </summary>
    ''' <param name="errorFileName">Fully qualified filename for the file that contains MP output.</param>
    ''' <param name="inputFile">Fully qualified filename for the metadata file.</param>
    ''' <param name="ans">This parameter is not used. Should be removed and locally declared.</param>
    ''' <returns>Returns an XML snippet containing errors/warnings found with the metadata record </returns>
    ''' <remarks></remarks>
    Function parseErrors(ByVal errorFileName As String, _
     ByVal inputFile As String, ByRef ans As String) As String
        Dim readStream As IO.FileStream = Nothing
        Dim textReader As XmlTextReader = Nothing
        Try
            Dim errorType As New List(Of String)
            Dim errorLine As New List(Of Integer)
            Dim errorMessage As New List(Of String)

            readStream = New IO.FileStream(errorFileName, IO.FileMode.Open, IO.FileAccess.Read)
            'MsgBox(workingDir & "\" & errorFileName)
            Dim sr As New IO.StreamReader(readStream)
            Dim line As String
            Dim temp As String

            'Go to beginning of the file
            sr.BaseStream.Seek(0, IO.SeekOrigin.Begin)
            'Read and ignore the first line
            line = sr.ReadLine

            'The latest mp.exe as of 20080605 adds an "Info" line that displays the file name
            'Ignore this line if it is present
            If line.Contains("Info: ") Then
                line = sr.ReadLine
            End If

            'Read the rest of the file to obtain error info
            While sr.Peek() > -1
                'Read the 3rd line onwards (2nd line if "Info" is not present)
                line = sr.ReadLine

                Dim currErrorType As String = ""
                'Add the Errors to the Arraylist
                If line.StartsWith("Error ") Then
                    'Add the errorType to the ArrayList
                    currErrorType = "Error"
                    'Add the Warnings to the Arraylist
                ElseIf line.StartsWith("Warning ") Then
                    'Add the errorType to the ArrayList
                    currErrorType = "Warning"
                End If
                If currErrorType <> "" Then
                    errorType.Add(currErrorType)

                    'Read the line # of the error
                    temp = line
                    temp = temp.Remove(0, temp.IndexOf("line ") + 4)
                    temp = temp.Remove(temp.IndexOf("):"), temp.Length - temp.IndexOf("):"))
                    'Add the line # to the ArrayList
                    errorLine.Add(Convert.ToInt32(temp))

                    'Add the error message to the ArrayList
                    temp = line
                    temp = temp.Remove(0, temp.IndexOf(": ") + 2)
                    errorMessage.Add(temp.Trim)
                End If
            End While

            sr.Close()

            'If records exist in the Arraylist, Generate the error xml return string
            If errorType.Count > 0 Then
                Dim i As Integer
                Dim startString As String
                Dim midString As String
                midString = ""
                Dim endString As String


                Dim linenum As Integer

                startString = "<errs>" & vbCrLf

                For i = 0 To errorType.Count - 1
                    Dim errid As String = Nothing

                    Dim errorElementName As String = ""
                    Dim elementArray As New List(Of String)

                    linenum = errorLine.Item(i)

                    ' Create an isntance of XmlTextReader and call Read method to read the file
                    textReader = New XmlTextReader(inputFile)
                    ' Dim nType As XmlNodeType = textReader.NodeType

                    'Read lines until the error/warning line
                    'Store the element names in an ArrayList for the errid
                    While textReader.LineNumber <> linenum
                        textReader.Read()
                        If textReader.IsEmptyElement Then
                            'skip
                        ElseIf textReader.NodeType = XmlNodeType.Element Then
                            elementArray.Add(textReader.Name)
                        ElseIf textReader.NodeType = XmlNodeType.EndElement Then
                            elementArray.Remove(textReader.Name)
                        End If
                    End While
                    'Read the element name for the error/warning line
                    textReader.MoveToElement()
                    errorElementName = textReader.Name
                    textReader.Close()

                    errid = "/" & String.Join("/", elementArray.ToArray)

                    ' Try to perfect error path, if possible
                    If errorMessage(i).Contains(" is required in ") Then
                        Dim errid2 As String = Utils.getFromSingleValuedSQL("SELECT ChildTag FROM MPHelper WHERE MPMessage='" & errorMessage(i) & "' AND ChildTag LIKE '" & errid & "%' ORDER BY len(ChildTag)")
                        If errid2 IsNot Nothing Then errid = errid2
                    End If

                    midString = midString & _
                        "  <err>" & vbCrLf & _
                        "    <type>" & CStr(errorType(i)) & "</type>" & vbCrLf & _
                        "    <message>" & CStr(errorMessage(i)) & "</message>" & vbCrLf & _
                        "    <linenum>" & CStr(errorLine(i)) & "</linenum>" & vbCrLf & _
                        "    <errid>" & errid & "</errid>" & vbCrLf & _
                        "  </err>" & vbCrLf
                Next

                endString = "</errs>" & vbCrLf

                Return startString & midString & endString

            Else
                ans = "<errs></errs>"
                Return ans
            End If
        Catch ex As Exception
            ans = "<errs>" & vbCrLf & errorxml("Global", "Error while parsing error file " _
                  & inputFile & " - " & ex.ToString, "", "") & "</errs>"
            Return ans
        Finally
            If readStream IsNot Nothing Then readStream.Close()
            If textReader IsNot Nothing Then textReader.Close()
            readStream = Nothing
            textReader = Nothing
        End Try
    End Function


    ''' <summary>
    ''' Run the mp_win.exe application.
    ''' </summary>
    ''' <param name="workdir">Working directory where mp_win.exe resides and temporary files are created.</param>
    ''' <param name="outputFile">Output filename where MP output will go.</param>
    ''' <param name="errorFileName">Error filename where MP errors will go.</param>
    ''' <param name="ans">Anny errors that occur will be reported in this parameter passed by reference.</param>
    ''' <returns>True if successful. False otherwise.</returns>
    ''' <remarks></remarks>
    Function runMP(ByVal workdir As String, ByVal outputFile As String, ByVal errorFileName As String, ByRef ans As String) As Boolean
        Try
            'Create the Process
            Dim startInfo As System.Diagnostics.ProcessStartInfo
            Dim pStart As New System.Diagnostics.Process

            'Create the arguments for mp_win.exe. The arguments need to be the cleaned 
            'filename (without errids) and the error filename: i.e "temp.xml -e temp.err"
            'Dim processArguments As String = outputFilename & " -e " & errorFileName
            Dim processArguments As String = """" & outputFile & """" & " -e " & """" & errorFileName & """"

            'Feed the executable file path and arguments
            startInfo = New System.Diagnostics.ProcessStartInfo("""" & workdir & "mp_win.exe" & """", processArguments)
            pStart.StartInfo = startInfo

            'Start the process
            pStart.Start()

            'Code will halt until the exe file has executed
            pStart.WaitForExit()
            Return True
        Catch ex As Exception
            ans = "<errs>" & vbCrLf & errorxml("Global", "Error while running mp_win.exe for " _
            & outputFile & " - " & ex.ToString, "", "") & "</errs>"
            Return False
        End Try
    End Function
    
End Module
