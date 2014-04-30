Imports System.Collections.Generic
Imports System.Xml

''' <summary>
''' An XmlFragment holds a fragment of a larger XML document.
''' Structurally, XmlFragment is a cross between a DOM node and a specialized object. 
''' All XMLFragment instances share a DOM and each one points to itw own DOM node as 
''' well as a set of its children that we care about (conrolled children). Controlled
''' children are extracted out of the DOM during XMLFragment creation and all other 
''' children (if any) are left behind in the DOM. When serialized, DOM is reconstructed
''' from controlled children and DOM children at every level.
''' </summary>
''' <remarks>Current implementation allows for only one DOM shared by all class members. 
''' This implementation is geared towards handling eainfo FGDC element and its children.</remarks>
Public Class XmlFragment
    ''' <summary>
    ''' Name of the FGDC element this XMLFragment represents.
    ''' </summary>
    ''' <remarks></remarks>
    Public name As String
    ''' <summary>
    ''' The DOM node (from the shared DOM) that corresponds to the FGDC element that this XMLFragment represents.
    ''' </summary>
    ''' <remarks></remarks>
    Public node As XmlNode
    ''' <summary>
    ''' Shared DOM representing the entire eainfo FGDC element.
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared dom As New XmlDocument()
    ''' <summary>
    ''' Controlled children dictionary. Keyed by FGDC element names of children. 
    ''' Since children with a given name may be repeated, values are list of XMLFrament objects.
    ''' </summary>
    ''' <remarks></remarks>
    Public controlledChildren As New Dictionary(Of String, List(Of XmlFragment))

    ''' <summary>
    ''' Construct an XMLFragment from an xml string.
    ''' </summary>
    ''' <param name="xml">A string containing well-formed XML.</param>
    ''' <param name="edomFix">Boolean value indicating if edom fix should be applied (see edomFixApply).</param>
    ''' <returns>The newly created XMLFragment instance representing the given XML string.</returns>
    ''' <remarks></remarks>
    Public Shared Function fromXml(ByRef xml As String, Optional ByVal edomFix As Boolean = False) As XmlFragment
        dom.LoadXml(xml)
        If edomFix Then edomFixApply(dom)
        Return XmlFragment.fromXmlNode(dom.DocumentElement)
    End Function

    ''' <summary>
    ''' Construct an XMLFragment from a DOM node (and its children).
    ''' </summary>
    ''' <param name="xn">A DOM node.</param>
    ''' <returns>The newly created XMLFragment instance representing the given DOM node (and its children).</returns>
    ''' <remarks></remarks>
    Public Shared Function fromXmlNode(ByVal xn As XmlNode) As XmlFragment
        Dim xf As New XmlFragment
        xf.node = xn
        xf.name = xf.node.Name
        xf.extract()
        Return xf
    End Function

    ''' <summary>
    ''' The display name of an XMLFragment is the text content of its DOM node.
    ''' </summary>
    ''' <value></value>
    ''' <returns>A string with the contents of the XMLFragment.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property displayName() As String
        Get
            Return Me.node.InnerText
        End Get
    End Property

    ''' <summary>
    ''' Extract the shared DOM into XMLFragments. We only extract names we might manipulate, not everything.
    ''' </summary>
    ''' <remarks>This is where this implementation get eainfo specific.</remarks>
    Sub extract()
        'Debug.Print("-------------------")
        'Debug.Print(Me.name)
        'Debug.Print(Me.node.InnerText)
        Select Case Me.name
            Case "eainfo"
                extract(New String() {"detailed", "overview"})
            Case "detailed"
                extract(New String() {"enttyp", "attr"})
            Case "enttyp"
                extract(New String() {"enttypl", "enttypd", "enttypds"})
            Case "attr"
                extract(New String() {"attrlabl", "attrdef", "attrdefs", "attrdomv"})
            Case "attrdomv"
                extract(New String() {"edom", "rdom", "codesetd", "udom"})
            Case "edom"
                extract(New String() {"edomv", "edomvd", "edomvds"})
            Case "rdom"
                extract(New String() {"rdommin", "rdommax"})
            Case "codesetd"
                extract(New String() {"codesetn", "codesets"})
        End Select
    End Sub

    ''' <summary>
    ''' Extract each of the given names. 
    ''' </summary>
    ''' <param name="nameArray">Array of string containing the FGDC element names that need to be extracted.</param>
    ''' <remarks>extract(String()) is just the middleman between extract() and extract(String)</remarks>
    Sub extract(ByVal nameArray As String())
        For Each name As String In nameArray
            extract(name)
        Next
    End Sub

    ''' <summary>
    ''' Extract instances of the given FGDC element. They are moved from being DOM children to controlled children.
    ''' </summary>
    ''' <param name="name">String containing the FGDC element name that need to be extracted.</param>
    ''' <remarks></remarks>
    Sub extract(ByVal name As String)
        Dim xnl As XmlNodeList
        Dim xfl As New List(Of XmlFragment)

        ' Find all nodes with name of interest.
        xnl = node.SelectNodes(name)
        For Each xn As XmlNode In xnl
            ' Remove each one from DOM children.
            Me.node.RemoveChild(xn)
            ' and add to our temp list here.
            xfl.Add(XmlFragment.fromXmlNode(xn))
        Next
        If xfl.Count = 0 Then
            ' Create an empty entry if no children mathed the name of interest.
            xfl.Add(XmlFragment.fromXmlNode(dom.CreateElement(name)))
        End If
        ' Assign the temp list the controlled children list for the name of interest.
        controlledChildren(name) = xfl
    End Sub

    ''' <summary>
    ''' Get the value of the XMLFragment at the given XSL pattern.
    ''' </summary>
    ''' <param name="name">String containing the XSL pattern for the FGDC element of interest.</param>
    ''' <value>Not used.</value>
    ''' <returns></returns>
    ''' <remarks>We should have combined setValue() sub with the set() part of this property.</remarks>
    Public Property value(ByVal name As String) As String
        Get
            ' Is this node's text being asked?
            If name = "" Then Return Me.node.InnerText

            ' Find the child to pass the request on to...
            Dim nameParts As String() = name.Split("/")
            ' child here is the leaf child
            Dim childNameParts As String() = nameParts(nameParts.Length - 1).Split("[]".ToCharArray)
            Dim childName As String = childNameParts(0)

            Dim idx As Integer = CInt(childNameParts(1))     ' attr[0] -> {"attr", "0"}
            Dim listName As String = String.Join("/", nameParts, 0, nameParts.Length - 1) & "/" & childName
            'Return Me.valueList(listName)(idx).value("")
            Dim xfl As List(Of XmlFragment) = Me.valueList(listName)
            If xfl.Count > idx AndAlso idx > -1 Then
                Return xfl(idx).value("")
            Else
                Return ""
            End If
        End Get
        Set(ByVal val As String)
        End Set
    End Property

    ''' <summary>
    ''' Get the list of XMLFragments identified by an XSL pattern.
    ''' </summary>
    ''' <param name="name">The XSL pattern identifying the values of interest.</param>
    ''' <value>Not used.</value>
    ''' <returns>List of XMLFragment.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property valueList(ByVal name As String) As List(Of XmlFragment)
        Get
            ' Find the child to pass the request on to...
            Dim nameParts As String() = name.Split("/")
            Dim childNameParts As String() = nameParts(0).Split("[]".ToCharArray)
            Dim childName As String = childNameParts(0)

            If Me.controlledChildren.ContainsKey(childName) Then
                If nameParts.Length = 1 Then
                    Return Me.controlledChildren(childName)
                Else
                    Dim idx As Integer = CInt(childNameParts(1))     ' attr[0] -> {"attr", "0"}
                    If Me.controlledChildren(childName).Count > idx AndAlso idx > -1 Then
                        Return Me.controlledChildren(childName)(idx).valueList(String.Join("/", nameParts, 1, nameParts.Length - 1))
                    End If
                End If
            End If
            Return New List(Of XmlFragment)
        End Get
    End Property

    ''' <summary>
    ''' Set the value of the XMLFragment identified by an XSL pattern.
    ''' </summary>
    ''' <param name="name">The XSL pattern identifying the values of interest.</param>
    ''' <param name="val">String containing the value to set.</param>
    ''' <remarks>We should have combined setValue() sub with the set() part of this property.</remarks>
    Sub setValue(ByVal name As String, ByVal val As String)
        If name = "" Then
            Me.node.InnerText = val
            Exit Sub
        End If

        ' Find the child to pass the request on to...
        Dim nameParts As String() = name.Split("/")
        Dim childNameParts As String() = nameParts(0).Split("[]".ToCharArray)
        Dim childName As String = childNameParts(0)
        Dim idx As Integer = CInt(childNameParts(1))     ' attr[0] -> {"attr", "0"}

        'If we don't have a list for childName, create one
        If Not Me.controlledChildren.ContainsKey(childName) Then
            Me.controlledChildren(childName) = New List(Of XmlFragment)
        End If

        'If the list for childName is empty, add one
        If Me.controlledChildren(childName).Count <= idx Then
            Me.controlledChildren(childName).Add(XmlFragment.fromXml("<" & childName & "></" & childName & ">"))
        End If

        If Me.controlledChildren(childName).Count > idx AndAlso idx > -1 Then
            Me.controlledChildren(childName)(idx).setValue(String.Join("/", nameParts, 1, nameParts.Length - 1), val)
        End If
    End Sub

    '''' <summary>
    '''' 
    '''' </summary>
    '''' <param name="childName"></param>
    '''' <param name="txt"></param>
    '''' <remarks></remarks>
    'Sub addChild(ByVal childName As String, Optional ByVal txt As String = "")
    '    Dim xml As String = "<" & childName & ">" & txt & "</" & childName & ">"
    '    Dim xf As XmlFragment = XmlFragment.fromXml(xml)
    '    If Not Me.controlledChildren.ContainsKey(childName) Then
    '        Me.controlledChildren(childName) = New List(Of XmlFragment)
    '    End If
    '    Me.controlledChildren(childName).Add(xf)
    'End Sub

    ''' <summary>
    ''' Fix edom (if necessary) to comply with CSDGM.
    ''' </summary>
    ''' <param name="dom"></param>
    ''' <remarks>While a no of sources (NOAA XSD and FGDC DTD for CSDGM as well as FGDC Editor in ArcGIS) 
    ''' appear to allow for multiple edom elements under an attrdomv element, 
    ''' these are divergent from CSDGM itself.</remarks>
    Shared Sub edomFixApply(ByVal dom As XmlDocument)
        Dim xnd As New Dictionary(Of XmlNode, List(Of XmlNode))
        ' Find attrdomv elements that have more than one edom element.
        For Each xn_attrdomv As XmlNode In dom.SelectNodes("//attrdomv[count(edom) > 1]")
            xnd(xn_attrdomv) = New List(Of XmlNode)
            ' We will leave the first edom element undisturbed.
            Dim skipNext As Boolean = True
            For Each xn_edom As XmlNode In xn_attrdomv.SelectNodes("edom")
                If skipNext Then
                    ' Leave the first edom alone
                    skipNext = False
                Else
                    ' Mark others for migration to under their own attrdomv elements
                    xnd(xn_attrdomv).Add(xn_attrdomv.RemoveChild(xn_edom))
                End If
            Next
        Next

        ' Now, we create a new attrdomv element for each edom element to be migrated and then put it
        ' alongside the original attrdomv element that contained the edom element.
        For Each xn_attrdomv As XmlNode In xnd.Keys
            For Each xn_edom As XmlNode In xnd(xn_attrdomv)
                Dim xn_new As XmlNode = dom.CreateElement("attrdomv")
                xn_new.AppendChild(xn_edom)
                xn_attrdomv.ParentNode.AppendChild(xn_new)
            Next
        Next
    End Sub

    ''' <summary>
    ''' Collect a list of XMLFragment objects representing repeated elements at two levels deep, optionally keeping duplicates.
    ''' </summary>
    ''' <param name="level1">Name of first level to go down from this XMLFragment.</param>
    ''' <param name="level2">Name of second level to go down from the first level.</param>
    ''' <param name="name">The name from which to collect values.</param>
    ''' <param name="keepDuplicates">Boolean indicating if the duplicate values need to be kept.</param>
    ''' <returns>List of XMLFragment.</returns>
    ''' <remarks></remarks>
    Public Function collectFromTwoLevels(ByVal level1 As String, ByVal level2 As String, ByVal name As String, Optional ByVal keepDuplicates As Boolean = False) As XmlFragment()
        Dim lst As New List(Of XmlFragment)
        For Each xf1 As XmlFragment In Me.valueList(level1)
            For Each xf2 As XmlFragment In xf1.valueList(level2)
                collectInner(name, xf2, lst, keepDuplicates)
            Next
        Next
        Return lst.ToArray
    End Function

    ''' <summary>
    ''' Collect a list of XMLFragment objects representing repeated elements at two levels deep, optionally keeping duplicates.
    ''' </summary>
    ''' <param name="level1">Name of first level to go down from this XMLFragment.</param>
    ''' <param name="level2">Name of second level to go down from the first level.</param>
    ''' <param name="level3">Name of third level to go down from the second level.</param>
    ''' <param name="name">The name from which to collect values.</param>
    ''' <param name="keepDuplicates">Boolean indicating if the duplicate values need to be kept.</param>
    ''' <returns>List of XMLFragment.</returns>
    ''' <remarks></remarks>
    Private Function collectFromThreeLevels(ByVal level1 As String, ByVal level2 As String, ByVal level3 As String, ByVal name As String, Optional ByVal keepDuplicates As Boolean = False) As XmlFragment()
        Dim lst As New List(Of XmlFragment)
        For Each xf1 As XmlFragment In Me.valueList(level1)
            For Each xf2 As XmlFragment In xf1.valueList(level2)
                For Each xf3 As XmlFragment In xf2.valueList(level3)
                    collectInner(name, xf3, lst, keepDuplicates)
                Next
            Next
        Next
        Return lst.ToArray
    End Function

    ''' <summary>
    ''' Collect the XMLFragment at the given name.
    ''' </summary>
    ''' <param name="name">The name of interest.</param>
    ''' <param name="xfa">XMLFragment object to collect from.</param>
    ''' <param name="lst">List of previously collected XMLFragment objects.</param>
    ''' <param name="keepDuplicates">Boolean indicating if the duplicate values need to be kept.</param>
    ''' <remarks>This should've been designed to operate on the target object.</remarks>
    Private Sub collectInner(ByVal name As String, ByVal xfa As XmlFragment, ByVal lst As List(Of XmlFragment), Optional ByVal keepDuplicates As Boolean = False)
        Dim v As XmlFragment = xfa.valueList(name)(0)
        If v.displayName.Trim > "" Then
            Dim add As Boolean = True
            If Not keepDuplicates Then
                ' Check for duplicates
                For Each xf As XmlFragment In lst
                    If xf.displayName = v.displayName Then
                        add = False
                        Exit For
                    End If
                Next
            End If
            If add Then
                lst.Add(v)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Front-end function to collect a repeated value from every entity.
    ''' </summary>
    ''' <param name="name">Name of repeated element to collect.</param>
    ''' <returns>Array of XMLFragment.</returns>
    ''' <remarks></remarks>
    Public Function collectFromEntities(ByVal name As String) As XmlFragment()
        Return collectFromTwoLevels("detailed", "enttyp", name, name = "enttypl")
    End Function

    ''' <summary>
    ''' Front-end function to collect a repeated value from every attribute.
    ''' </summary>
    ''' <param name="name"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function collectFromAttributes(ByVal name As String) As XmlFragment()
        Return collectFromTwoLevels("detailed", "attr", name, name = "attrlabl")
    End Function

    ''' <summary>
    ''' Front-end function to collect a repeated value from every attrdomv element.
    ''' </summary>
    ''' <param name="name"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function collectFromAttributeDomains(ByVal name As String) As XmlFragment()
        Return collectFromThreeLevels("detailed", "attr", "attrdomv", name)
    End Function

    ''' <summary>
    ''' Get the no of names at the given XSL pattern.
    ''' </summary>
    ''' <param name="name">An XSL pattern identifying the name of of interest.</param>
    ''' <returns>An integer value.</returns>
    ''' <remarks></remarks>
    Public Function count(ByVal name As String) As Integer
        Return Me.valueList(name).Count
    End Function

    ''' <summary>
    ''' Delete the name identified by the given XSL pattern and index.
    ''' </summary>
    ''' <param name="name">An XSL pattern identifying the name of of interest.</param>
    ''' <param name="deleteIndex">The index of the XMLFragment to be deleted.</param>
    ''' <remarks></remarks>
    Public Sub delete(ByVal name As String, ByVal deleteIndex As Integer)
        Dim nameParts As String() = name.Split("/")
        Dim childNameParts As String() = nameParts(0).Split("[]".ToCharArray)
        Dim childName As String = childNameParts(0)

        If Me.controlledChildren.ContainsKey(childName) Then
            If nameParts.Length = 1 Then
                Me.controlledChildren(childName).RemoveAt(deleteIndex)
            Else
                Dim idx As Integer = CInt(childNameParts(1))     ' attr[0] -> {"attr", "0"}
                If Me.controlledChildren(childName).Count > idx Then
                    Me.controlledChildren(childName)(idx).delete(String.Join("/", nameParts, 1, nameParts.Length - 1), deleteIndex)
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Add the given XMLFragment at the given XSL pattern and index.
    ''' </summary>
    ''' <param name="name">An XSL pattern identifying the name of of interest.</param>
    ''' <param name="addIndex">The index at which to add the XMLFragment.</param>
    ''' <param name="xf">The XMLFragment to be added.</param>
    ''' <remarks></remarks>
    Public Sub add(ByVal name As String, ByVal addIndex As Integer, ByVal xf As XmlFragment)
        Dim nameParts As String() = name.Split("/")
        Dim childNameParts As String() = nameParts(0).Split("[]".ToCharArray)
        Dim childName As String = childNameParts(0)

        If Me.controlledChildren.ContainsKey(childName) Then
            If nameParts.Length = 1 Then
                Me.controlledChildren(childName).Insert(addIndex, xf)
            Else
                Dim idx As Integer = CInt(childNameParts(1))     ' attr[0] -> {"attr", "0"}
                If Me.controlledChildren(childName).Count > idx Then
                    Me.controlledChildren(childName)(idx).delete(String.Join("/", nameParts, 1, nameParts.Length - 1), addIndex)
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Reconstruct the XML representation of an XMLFragment from its name, controlled children and DOM children.
    ''' </summary>
    ''' <param name="topLevel">Boolean indicating if this XMLFragment is a top-level element.</param>
    ''' <returns>A string containing the XML representation.</returns>
    ''' <remarks></remarks>
    Public Function construct(Optional ByVal topLevel As Boolean = True) As String
        construct = ""
        ' First put together XML representations of controlled children
        For Each xfl As List(Of XmlFragment) In Me.controlledChildren.Values
            Dim childConstruct As String = ""
            For Each xf As XmlFragment In xfl
                childConstruct &= xf.construct(False)
            Next
            construct &= childConstruct
            'attrdomv will only use the first non-empty element (domain)
            If childConstruct > "" AndAlso Me.name = "attrdomv" Then
                Exit For
            End If
        Next
        ' Here, we add the XML representation of DOM children.
        construct &= Me.node.InnerXml.Trim
        ' If constructed string is not empty and this is a top-level node...
        If construct > "" AndAlso Not topLevel Then
            ' then we wrap it all with this XMLFragment's name.
            construct = "<" & Me.name & ">" & construct & "</" & Me.name & ">"
        End If
    End Function

End Class
