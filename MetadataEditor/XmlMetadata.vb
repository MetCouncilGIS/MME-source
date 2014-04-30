Imports System.Xml

''' <summary>
''' XmlMetadata represents a metadata document that is serialized from/to XML and provides basic operations to manipulate metadata.
''' </summary>
''' <remarks>XmlMetadata class was designed to remove reliance on several ArcObjects metadata objects such as IXMLPropertySet2.
''' It does not attempt to be 100% compatible in function or API coverage but implements enough to fulfill its design goal.</remarks>
Public Class XmlMetadata
    ''' <summary>
    ''' Internal structure to hold the metadata document
    ''' </summary>
    ''' <remarks></remarks>
    Private dom As New XmlDocument

    Sub New()
    End Sub

    ''' <summary>
    ''' Set the element with the given xpath to the given value.
    ''' </summary>
    ''' <param name="xpathStr">xpath expression of element to update. Does not have to pre-exist in the metadata document.</param>
    ''' <param name="value">New content for element</param>
    ''' <remarks>Assumes that the xpath points at most to a single target. Only the first occurence is updated.</remarks>
    Public Sub SetPropertyX(ByVal xpathStr As String, ByVal value As String)
        Try
            If Not xpathStr.StartsWith("/") AndAlso Not xpathStr.StartsWith("//") Then
                xpathStr = "/metadata/" + xpathStr
            End If

            If makeXpath(xpathStr) Then
                Dim node As XmlNode = dom.SelectSingleNode(xpathStr)
                Try
                    node.FirstChild.Value = value
                Catch ex As Exception
                    Dim aaa = 1
                End Try
            End If
        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Create the given element hierarchy as dictated by the given xpath. All elements leading up to the leaf tag will be created if not already there.
    ''' </summary>
    ''' <param name="xpathStr">xpath expression for the element hierarchy to create</param>
    ''' <param name="nodeType">The type of leaf node to create. Default is a text node which is currently the only supported node.</param>
    ''' <returns>True if succesful, False otherwise.</returns>
    ''' <remarks>Logic does handle xpaths involving integer indexes.</remarks>
    Public Function makeXpath(ByVal xpathStr As String, Optional ByVal nodeType As XmlNodeType = XmlNodeType.Text) As Boolean
        makeXpath = True

        If Not xpathStr.StartsWith("/") AndAlso Not xpathStr.StartsWith("//") Then
            xpathStr = "/metadata/" + xpathStr
        End If

        If xpathStr Is Nothing OrElse xpathStr.Trim = "" OrElse xpathStr.StartsWith("//") OrElse Not xpathStr.StartsWith("/") Then
            Return False
        End If

        Dim parts As String() = xpathStr.Split("/")
        Dim parent As String = ""
        Dim parentNode As XmlNode
        Dim current As String = ""
        Dim nodes As XmlNodeList

        For i As Integer = 1 To parts.Length - 2
            parent += "/" + parts(i)
            parentNode = dom.SelectSingleNode(parent)
            current = parts(i + 1)
            nodes = dom.SelectNodes(parent + "/" + current)

            If current.Contains("[") Then
                Dim idxStr As String = current.Substring(current.IndexOf("[") + 1)
                idxStr = Left(idxStr, idxStr.LastIndexOf("]"))
                If Utils.IsNumeric(idxStr) Then
                    Dim idx As Integer = Integer.Parse(idxStr)
                    current = Left(current, current.IndexOf("["))
                    ' Create as many as needed
                    'For j As Integer = nodes.Count + 1 To idx
                    '    parentNode.AppendChild(dom.CreateElement(current))
                    'Next

                    If nodes.Count = 0 Then
                        parentNode.AppendChild(dom.CreateElement(current))
                    End If
                End If
            ElseIf nodes.Count = 0 Then
                'create
                Try
                    parentNode.AppendChild(dom.CreateElement(current))
                Catch ex As Exception
                    Debug.Print(xpathStr)
                End Try
            ElseIf nodes.Count = 1 Then
                'nothing to do
            ElseIf nodes.Count > 1 Then
                'error, no index but multiple nodes
                Return False
            End If
        Next

        ' Leaf node should be what the caller wanted (text node by default)
        If nodeType = XmlNodeType.Text Then
            For Each node As XmlNode In dom.SelectNodes(xpathStr)
                If node.FirstChild Is Nothing OrElse node.FirstChild.NodeType <> XmlNodeType.Text Then
                    node.AppendChild(dom.CreateTextNode(""))
                End If
            Next
        End If

    End Function



    ''' <summary>
    ''' Get the XML snippet pointed to by the given xpath expression.
    ''' </summary>
    ''' <param name="xpathStr">xpath expression of the XML snippet targeted.</param>
    ''' <param name="normalize">Normalize the xpath expression to start with root element (/metadata/) if not already.</param>
    ''' <param name="outerXml">Return outer XML for the xpath target if True, return inner XML otherwise.</param>
    ''' <returns>Returns the inner/outer XML snippet targeted by the xpath expression.</returns>
    ''' <remarks>Some xpaths have special meaning: empty xpath return entire XML including any XML declaration; slash returns starting at root element without any headers.</remarks>
    Public Function GetXml(ByVal xpathStr As String, Optional ByVal normalize As Boolean = True, Optional ByVal outerXml As Boolean = True) As String
        ' Special case to return the entire xml including xml declaration header
        If xpathStr = "" Then Return dom.OuterXml

        If normalize Then xpathStr = Me.normalize(xpathStr)
        Try
            Dim node As XmlNode = dom.SelectSingleNode(xpathStr)
            If node Is Nothing Then
                GetXml = ""
            Else
                If outerXml Then
                    GetXml = node.OuterXml
                Else
                    GetXml = node.InnerXml
                End If
            End If
        Catch ex As Exception
            ErrorHandler(ex)
            GetXml = ""
        End Try
    End Function

    ''' <summary>
    ''' Set the XML document to the given xml content.
    ''' </summary>
    ''' <param name="xmlStr">String with XML content.</param>
    ''' <remarks></remarks>
    Public Sub SetXml(ByVal xmlStr As String)
        dom.LoadXml(xmlStr)
    End Sub

    ''' <summary>
    ''' Copy an XML snippet from one document and graft into current.
    ''' </summary>
    ''' <param name="src">Souce XmlMetadata object to copy from</param>
    ''' <param name="xpathStrSource">Source xpath expression to copy from</param>
    ''' <param name="xpathStrTarget">Traget xpath expression denoting where the XML snippet will be copied to. If Nothing, then same as source xpath.</param>
    ''' <remarks></remarks>
    Public Sub copyFrom(ByVal src As XmlMetadata, ByVal xpathStrSource As String, Optional ByVal xpathStrTarget As String = Nothing)
        ' Default to the same xpath as source for target - if not specified
        If xpathStrTarget Is Nothing Then xpathStrTarget = xpathStrSource
        ' Get source xml
        Dim srcXml As String = src.GetXml(xpathStrSource)
        ' Copy to target xpath - only if there is anything to copy.
        If srcXml IsNot Nothing AndAlso srcXml.Trim <> "" Then Me.SetXml(xpathStrTarget, srcXml)
    End Sub

    ''' <summary>
    ''' Copy the given XML snippet and graft at the element pointed by given xpath expression
    ''' </summary>
    ''' <param name="xpathStr">xpath target to receive XML snippet</param>
    ''' <param name="xmlStr">The XML snippet to be received.</param>
    ''' <remarks></remarks>
    Public Sub SetXml(ByVal xpathStr As String, ByVal xmlStr As String)
        xpathStr = normalize(xpathStr)
        Me.DeleteProperty(xpathStr)
        Me.SetPropertyX(xpathStr, "")
        Dim node As XmlNode = dom.SelectSingleNode(xpathStr)
        If node IsNot Nothing Then
            Dim tmp As New XmlMetadata
            tmp.SetXml(xmlStr)
            node.InnerXml = tmp.GetXml(tmp.getRootTag(), False, False)
        End If
    End Sub

    ' ''' <summary>
    ' ''' Historical artifact
    ' ''' </summary>
    ' ''' <param name="xpathStr"></param>
    ' ''' <param name="mdSnippetDoc"></param>
    ' ''' <param name="overwrite"></param>
    ' ''' <remarks></remarks>
    'Public Sub SetXml(ByVal xpathStr As String, ByVal mdSnippetDoc As XmlMetadata, Optional ByVal overwrite As Boolean = True)
    '    ' Delete tag if ovewriting
    '    If overwrite Then
    '        Me.DeleteProperty(xpathStr)
    '    End If
    '    ' Make sure tag does not exist before writing given snippet to it.
    '    Dim node As XmlNode = dom.SelectSingleNode(xpathStr)
    '    If node Is Nothing Then
    '        ' Set tag to empty content simply to force its creation
    '        Me.SetPropertyX(xpathStr, "")
    '        ' Find the node for the tag
    '        node = dom.SelectSingleNode(xpathStr)
    '        ' Set its inner xml to the snippet
    '        node.InnerXml = mdSnippetDoc.dom.FirstChild.InnerXml
    '    End If
    'End Sub


    ''' <summary>
    ''' Delete the element(s) pointed by the given xpath expression
    ''' </summary>
    ''' <param name="xpathStr">xpath expression for element(s) to be deleted</param>
    ''' <remarks>All elements targeted by the xpath expression are deleted.</remarks>
    Public Sub DeleteProperty(ByVal xpathStr As String)
        'Debug.Print(xpathStr)
        Try
            If Not xpathStr.StartsWith("/") Then
                xpathStr = "/metadata/" + xpathStr
            End If

            Dim nodes As XmlNodeList = dom.SelectNodes(xpathStr)
            If nodes IsNot Nothing AndAlso nodes.Count > 0 Then
                Dim parent As XmlNode = nodes(0).ParentNode
                If parent IsNot Nothing Then
                    For Each node As XmlNode In nodes
                        parent.RemoveChild(node)
                    Next
                End If
            End If

        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Get element value
    ''' </summary>
    ''' <param name="xpathStr">xpath expression for element targeted</param>
    ''' <returns>String value of the element at given xpath.</returns>
    ''' <remarks>If element is compound, then value is compound values of all subelements - though it's unusual to use that way.</remarks>
    Public Function SimpleGetProperty(ByVal xpathStr As String) As String
        xpathStr = xpathStr.Replace("[?]", "")
        If Not xpathStr.StartsWith("/") Then
            xpathStr = "/metadata/" + xpathStr
        End If

        'Debug.Print(xpathStr)
        SimpleGetProperty = ""
        Try
            Dim node As XmlNode = dom.SelectSingleNode(xpathStr)
            If node IsNot Nothing Then
                SimpleGetProperty = node.InnerText
            End If
        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Function

    ''' <summary>
    ''' Get element values
    ''' </summary>
    ''' <param name="xpathStr">xpath expression for elements targeted</param>
    ''' <returns>List of String values with one entry for each element matching the given xpath expression. Nothing is returned if no matches.</returns>
    ''' <remarks></remarks>
    Public Function GetProperty(ByVal xpathStr As String) As List(Of String)
        GetProperty = New List(Of String)
        If Not xpathStr.StartsWith("/") AndAlso Not xpathStr.StartsWith("//") Then
            xpathStr = "/metadata/" + xpathStr
        End If
        'Debug.Print(xpathStr)
        Try
            Dim nodes As XmlNodeList = dom.SelectNodes(xpathStr)
            Dim tags As New List(Of String)
            Dim uniqueTags As New Hashtable()
            For Each node As XmlNode In nodes
                For Each nodeC As XmlNode In node.ChildNodes
                    tags.Add(nodeC.Name)
                    uniqueTags(nodeC.Name) = Nothing
                    If nodeC.NodeType = XmlNodeType.Text Then
                        GetProperty.Add(nodeC.InnerText)
                    ElseIf nodeC.NodeType = XmlNodeType.Element Then
                        GetProperty.Add(nodeC.Name)
                    End If
                Next
            Next
            If uniqueTags.Count = 0 Then
                GetProperty = Nothing
            ElseIf uniqueTags.Count > 1 Then
                GetProperty.Clear()
                GetProperty = tags
            End If

        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Function

    ''' <summary>
    ''' Determine the no of occurences of the given xpath expression.
    ''' </summary>
    ''' <param name="xpathStr">xpath expression to count the occurrences of</param>
    ''' <returns>The number of time that the given xpath expression is encountered.</returns>
    ''' <remarks></remarks>
    Public Function CountX(ByVal xpathStr As String) As Integer
        If Not xpathStr.StartsWith("/") AndAlso Not xpathStr.StartsWith("//") Then
            xpathStr = "/metadata/" + xpathStr
        End If
        Return dom.SelectNodes(xpathStr).Count
    End Function

    ''' <summary>
    ''' Get the name of the root element of the XML document.
    ''' </summary>
    ''' <returns>Name of the root element as String</returns>
    ''' <remarks></remarks>
    Public Function getRootTag() As String
        Return dom.FirstChild.Name
    End Function

    'Public Function getOfficialCopy(Optional ByVal root As String = "") As String
    '    Dim iXPSv As New XmlMetadata
    '    iXPSv.SetXml(iXPS.GetXml(root))

    '    ' Delete ESRI elements to avoid unnecessary validation warnings. Handle ISO elements better next time.
    '    iXPSv.DeleteProperty("Esri")
    '    iXPSv.DeleteProperty("mdDateSt")


    '    iXPSv.DeleteProperty("innovate")

    '    'AE_xml
    '    '' Delete keyword themes that don't have keywords
    '    iXPSv.checkDeleteKTTags()
    '    Return iXPSv.GetXml(root)
    'End Function

    ''' <summary>
    ''' Delete XSL pattern indexed tags if not used in the metadata record.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub checkDeleteKTTags()
        If CountX("idinfo/keywords/theme[themekt='ISO 19115 Topic Category']/themekey") = 0 Then
            DeleteProperty("idinfo/keywords/theme[themekt='ISO 19115 Topic Category']")
        End If
        If CountX("idinfo/keywords/theme[themekt='" + "EPA GIS" + " Keyword Thesaurus']/themekey") = 0 Then
            DeleteProperty("idinfo/keywords/theme[themekt='" + "EPA GIS" + " Keyword Thesaurus']")
        End If
        'If CountX("idinfo/keywords/theme[themekt='" + My.Settings.CommonAgencyAbbreviation + " Keyword Thesaurus']/themekey") = 0 Then
        '    DeleteProperty("idinfo/keywords/theme[themekt='" + My.Settings.CommonAgencyAbbreviation + " Keyword Thesaurus']")
        'End If
        If CountX("idinfo/keywords/place[placekt='None']/placekey") = 0 Then
            DeleteProperty("idinfo/keywords/place[placekt='None']")
        End If
        ' AE:
        If CountX("idinfo/keywords/theme[themekt='" & GlobalVars.userThesaurus & "']/themekey") = 0 Then
            DeleteProperty("idinfo/keywords/theme[themekt='" & GlobalVars.userThesaurus & "']")
        End If
    End Sub


    ''' <summary>
    ''' Create XSL pattern indexed tags if not already in the metadata record.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub checkCreateKTTags()
        If CountX("idinfo/keywords/theme[themekt='ISO 19115 Topic Category']/themekt") = 0 Then
            SetPropertyX("idinfo/keywords/theme[" & (CountX("idinfo/keywords/theme") + 1) & "]/themekt", "ISO 19115 Topic Category")
        End If
        'AE: This may not always work properly
        'If CountX("idinfo/keywords/theme[themekt='" + "EPA GIS" + " Keyword Thesaurus']/themekt") = 0 Then
        '    SetPropertyX("idinfo/keywords/theme[" & (CountX("idinfo/keywords/theme") + 1) & "]/themekt", "EPA GIS" + " Keyword Thesaurus")
        'End If
        'If CountX("idinfo/keywords/theme[themekt='" + My.Settings.CommonAgencyAbbreviation + " Keyword Thesaurus']/themekt") = 0 Then
        '    SetPropertyX("idinfo/keywords/theme[" & (CountX("idinfo/keywords/theme") + 1) & "]/themekt", My.Settings.CommonAgencyAbbreviation + " Keyword Thesaurus")
        'End If
        If CountX("idinfo/keywords/place[placekt='None']/placekt") = 0 Then
            SetPropertyX("idinfo/keywords/place[" & (CountX("idinfo/keywords/place") + 1) & "]/placekt", "None")
        End If
        ' AE:
        If CountX("idinfo/keywords/theme[themekt='" & GlobalVars.userThesaurus & "']/themekt") = 0 Then
            SetPropertyX("idinfo/keywords/theme[" & (CountX("idinfo/keywords/theme") + 1) & "]/themekt", GlobalVars.userThesaurus)
        End If
    End Sub

    ''' <summary>
    ''' Normalize the given xpath expression to start with the root element (/metadata/), if not already.
    ''' </summary>
    ''' <param name="xpathStr">xpath expression to normalize</param>
    ''' <returns>Normalized xpath expression</returns>
    ''' <remarks>Handles only simple xpath expressions</remarks>
    Function normalize(ByVal xpathStr As String) As String
        If Not xpathStr.StartsWith("/") AndAlso Not xpathStr.StartsWith("//") Then
            'AE: Temp hack. Why did this work in .net 2.0 but not in .net 3.5 ??
            'xpathStr = "/metadata/" + xpathStr
            xpathStr = "/metadata/" + xpathStr
            If xpathStr.EndsWith("/") Then xpathStr = xpathStr.Substring(0, xpathStr.Length - 1)
        End If
        Return xpathStr
    End Function
End Class
