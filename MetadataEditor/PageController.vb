Imports System.Reflection
Imports System.Xml
Imports System.ComponentModel
Imports System.Text.RegularExpressions


''' <summary>
''' Class that manages all manipulations to and operations on metadata - within the editor.
''' </summary>
''' <remarks></remarks>
Public Class PageController
    Implements IComparable

    Private Shared HiveMind As New SortedDictionary(Of String, PageController)
    Private Shared HiveMindInitialized As Boolean = False
    Private Shared WithEvents ValidationWorker As New BackgroundWorker
    Private Shared WithEvents ValidationTimer As New Timer
    Private Shared EMEDataTables As New Hashtable
    Private Shared clusterTags As New Hashtable
    Private orderedId As Long
    Private tag As String
    Private srcTable As String
    Private srcField As String
    Private formFieldName_ As String
    Private tabNo As Integer
    Private spellcheck As Boolean
    Private clusterName As String
    Private clusterUpdate As Boolean
    Public help As String
    Private leafTag As String
    Private disabled As Boolean '= False


    ''' <summary>
    ''' Find the PageController oject that controls the form control with the given name.
    ''' </summary>
    ''' <param name="ctrlName">Name of form control</param>
    ''' <returns>PageController object that controls the form control</returns>
    ''' <remarks></remarks>
    Shared Function thatControls(ByVal ctrlName As String) As PageController
        Dim pc As PageController = Nothing
        HiveMind.TryGetValue(ctrlName, pc)
        Return pc
    End Function

    ''' <summary>
    ''' Create and initialize PageController objects based on information stored in the database.
    ''' </summary>
    ''' <remarks>EME table in the database drives this process.</remarks>
    Shared Sub readFromDb()
        Dim p As PageController
        If HiveMindInitialized Then
            Exit Sub
        End If
        HiveMindInitialized = True

        Dim dr As OleDb.OleDbDataReader = Utils.readerForSQL("SELECT * FROM MME ORDER BY OrderedId")

        Do While dr.Read()
            p = New PageController( _
                dr("OrderedId"), _
                dr("tag"), _
                IIf(dr("srcTable") Is DBNull.Value, Nothing, dr("srcTable")), _
                IIf(dr("srcField") Is DBNull.Value, dr("tag").Substring(dr("tag").LastIndexOf("/") + 1), dr("srcField")), _
                dr("tabNo"), _
                dr("spellcheck"), _
                IIf(dr("cluster") Is DBNull.Value, Nothing, dr("cluster")), _
                dr("clusterUpdate"), _
                IIf(dr("help") Is DBNull.Value, Nothing, dr("help")) _
            )
            If p.isClusterController Then
                p.srcField = "clusterInfo"
            End If

            'Debug.Print(p.formFieldName)
            'Debug.Print(p.srcField)
        Loop
        dr.Close()
    End Sub

    ''' <summary>
    ''' Create a new PageController object with given information.
    ''' </summary>
    ''' <param name="orderedId">Unique id for the PageController that also determines the order in which its associated FGDC element is read/written</param>
    ''' <param name="tag">Name of the FGDC element tag</param>
    ''' <param name="srcTable">Name of database table to read element info from</param>
    ''' <param name="srcField">Name of field in the database table to read info from</param>
    ''' <param name="tabNo">Tab number on the EME user interface where the element is displayed</param>
    ''' <param name="spellcheck">Whether the element will be a target of spellchecking</param>
    ''' <param name="cluster">The name of element cluster (parent tag) that this tag is part of, if applicable.</param>
    ''' <param name="help">Name of the help screen associated with this element, if different from default naming scheme</param>
    ''' <remarks>
    ''' A PageController is associated with form control whose value determines the value of an FGDC element's XML tag. 
    ''' This is critical to understanding the operation of EME.
    ''' </remarks>
    Public Sub New(ByVal orderedId As Long, ByVal tag As String, ByVal srcTable As String, ByVal srcField As String, ByVal tabNo As Integer, ByVal spellcheck As Boolean, ByVal cluster As String, ByVal clusterUpdate As Boolean, ByVal help As String)
        MyBase.New()
        Me.orderedId = orderedId
        Me.tag = tag
        Me.srcTable = srcTable
        Me.srcField = srcField
        If tag.Equals("idinfo/keywords/theme[themekt='" & GlobalVars.userThesaurus & "']/themekey") Then
            ' Special handling for listbox that contains user-specified thesaurus.
            Me.formFieldName_ = stripNonAlphanumeric("idinfo/keywords/theme[themekt='User']/themekey", "_")
        Else
            Me.formFieldName_ = stripNonAlphanumeric(tag, "_")
        End If
        HiveMind.Add(formFieldName_, Me)
        Me.tabNo = tabNo
        Me.spellcheck = spellcheck
        Me.clusterName = stripNonAlphanumeric(cluster, "_")
        Me.help = IIf(help Is Nothing, "/t" & Me.tabNo.ToString & "_" & Me.srcField & ".html", help)
        Me.addToCluster()
        Me.clusterUpdate = Me.isClusterController AndAlso clusterUpdate

        Dim parts As String() = Me.tag.Split("/")
        Me.leafTag = parts(parts.Length - 1)
        If Me.leafTag.Contains("[") Then Me.leafTag = Me.leafTag.Substring(0, Me.leafTag.IndexOf("["))
    End Sub

    ''' <summary>
    ''' A utility function to get the value of simple tags (non-repeated, no child tags) 
    ''' </summary>
    ''' <param name="name">XSL pattern for the element relative to the top-level "metadata" tag</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function SimpleGetProperty(ByVal name As String) As String
        Dim props As Object = iXPS.SimpleGetProperty("/metadata/" & name)
        Return IIf(props Is Nothing, "", props)
    End Function

    ''' <summary>
    ''' Return all text controls marked for spellchecking, optionally filtered by tab number
    ''' </summary>
    ''' <param name="frm"></param>
    ''' <param name="tabNo"></param>
    ''' <returns>Returns array of TextBox objects that will be spellchecked.</returns>
    ''' <remarks>
    ''' If <paramref name="tabNo">tabNo</paramref> is specified (in base-1 index),
    ''' only controls from the specified tab are returned.
    ''' </remarks>
    Public Shared Function GetTextControls(ByRef frm As EditorForm, Optional ByVal tabNo As Integer = 0) As TextBoxBase()
        Dim ctrlArr(HiveMind.Count) As TextBoxBase
        Dim ctrl As Control
        Dim i As Integer = 0
        For Each pc As PageController In HiveMind.Values()
            ctrl = frm.getControlForTag(pc.formFieldName)
            If (tabNo = 0 Or pc.tabNo = tabNo) And pc.spellcheck And TypeOf ctrl Is TextBoxBase Then
                ctrlArr(i) = ctrl
                i += 1
            End If
        Next
        Array.Resize(ctrlArr, i)
        Return ctrlArr
    End Function

    ''' <summary>
    ''' Iterate through all PageControllers and populate the elements/form controls that they control.
    ''' </summary>
    ''' <param name="frm"></param>
    ''' <remarks></remarks>
    Public Shared Sub ElementPopulator(ByRef frm As EditorForm)
        For Each pc As PageController In HiveMind.Values()
            pc.populate(frm)
        Next
    End Sub

    ''' <summary>
    ''' Determine the parent tag of a given semi-simple xpath to a tag.
    ''' </summary>
    ''' <value></value>
    ''' <returns>The xpath to parent tag</returns>
    ''' <remarks>Cheesy but works with the tag xpaths we deal with.</remarks>
    Private ReadOnly Property parentTag() As String
        Get
            Return Me.tag.Substring(0, Me.tag.LastIndexOf("/")).Replace("[?]", "")
        End Get
    End Property

    ''' <summary>
    ''' Populate this PageController object's form control
    ''' </summary>
    ''' <param name="frm">The editor form to find associated form control</param>
    ''' <remarks></remarks>
    Private Sub populate(ByRef frm As EditorForm)
        Dim ctrl As Control
        ctrl = frm.getControlForTag(formFieldName)
        'Debug.Print(formFieldName)
        'Debug.Print(ctrl.GetType.ToString)

        If Not Me.isClusterController Then
            ctrl.Text = SimpleGetProperty(tag)
        End If

        If srcTable Is Nothing Then
            ' Load from metadata
            If TypeOf ctrl Is ListBox Then
                listBoxLoaderFromMetadata(ctrl)
            ElseIf TypeOf ctrl Is DataGridView Then
                dataGridViewLoaderFromMetadata(ctrl)
            End If
        Else
            ' Load from DB
            If TypeOf ctrl Is ListBox Then
                listBoxLoader(ctrl)
            ElseIf TypeOf ctrl Is ComboBox Then
                comboBoxLoader(ctrl)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Load the ComboBox object associated with this PageController.
    ''' </summary>
    ''' <param name="ctrl">The ComboBox control</param>
    ''' <param name="reload">Whether this is a reload or an initial load (optional)</param>
    ''' <remarks>
    ''' Stay away if you haven't fully grokked workings of table-driven ComboBox objects, EME's database driven operation, 
    ''' complex FGDC elements  and their mapping to ComboBox objects.
    ''' </remarks>
    Sub comboBoxLoader(ByVal ctrl As ComboBox, Optional ByVal reload As Boolean = False)
        If Me.isClusterController Then
            ' Cluster controllers' content can't be typed in by user, so make them DropDownList.
            ctrl.DropDownStyle = ComboBoxStyle.DropDownList
        Else
            ctrl.AutoCompleteMode = AutoCompleteMode.Suggest
            ctrl.AutoCompleteSource = AutoCompleteSource.ListItems
        End If

        Dim SQLStr As String = ""
        If Me.isPartOfCluster Then
            SQLStr = "SELECT * FROM [" & Me.srcTable & "_cluster]"
            SQLStr &= " ORDER BY orderedId"
        Else
            ' Some entries may show up multiple times in the lookup table and it's not enough 
            ' to just say DISTINCT as we are also pulling 'default' which may have different values.
            SQLStr = "SELECT DISTINCT [" & Me.srcField & "], default FROM [" & Me.srcTable
            SQLStr &= "] AS t WHERE NOT EXISTS (SELECT * FROM [" & Me.srcTable & "] WHERE "
            SQLStr &= "[" & Me.srcField & "]=t.[" & Me.srcField & "] AND default AND NOT t.default) "
            SQLStr &= "AND [" & Me.srcField & "] > '' "
            SQLStr &= "ORDER BY [" & Me.srcField & "]"
        End If
        'Debug.Print(SQLStr)
        Dim dt As DataTable = Utils.datatableFromSQL(SQLStr)

        Dim tagValue As String
        ' If reloading...
        If reload Then
            'save current selection
            tagValue = ctrl.SelectedValue
        Else
            'otherwise, get selected value from metadata tag
            tagValue = SimpleGetProperty(tag).Trim()

            ctrl.DisplayMember = Me.srcField
            ctrl.ValueMember = Me.srcField
        End If

        ' Bind the ComboBox to a DataTable (well, its default view actually) that contains the options
        ctrl.DataSource = dt.DefaultView

        If Me.isPartOfCluster Then
            populateClusterEntry(ctrl)
            If reload Then
                If tagValue IsNot Nothing Then
                    ctrl.SelectedValue = tagValue
                End If
            Else
                AddHandler ctrl.SelectionChangeCommitted, AddressOf propagateClusterSelectionChanged
            End If
        Else
            'Debug.Print(ctrl.Name)
            If tagValue Is Nothing Then tagValue = ""
            Dim matchingRows As Array = dt.Select(ctrl.ValueMember & "='" & tagValue.Replace("'", "''") & "'")
            If matchingRows.Length = 0 Then
                ' Add an entry for value read from metadata only if it doesn't exist as an option (in the database)
                Dim dr As DataRow = dt.NewRow()
                dr(Me.srcField) = tagValue
                dr("default") = False
                dt.Rows.Add(dr)
                ctrl.SelectedValue = tagValue
            Else
                ' If the value exists, then select it.
                ctrl.SelectedValue = matchingRows(0)(Me.srcField)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Event handler for cluster controlling ComboBox objects that propagates the selection change to controlled cluster members.
    ''' </summary>
    ''' <param name="sender">Event sender. Cluster controlling ComboBox.</param>
    ''' <param name="e">Event arguments. Not used.</param>
    ''' <remarks>A cluster is a set of FGDC sibling tags whose values are determined by the selected item in ComboBox.</remarks>
    Public Sub propagateClusterSelectionChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim clusterController As ComboBox
        Dim ctrl As Control
        Dim pc As PageController
        Dim drv As DataRowView
        For Each formField As String In clusterTags(Me.clusterName)
            clusterController = DirectCast(sender, ComboBox)
            ctrl = DirectCast(clusterController.FindForm(), EditorForm).allControls(formField)
            'Debug.Print(ctrl.Name)
            pc = PageController.thatControls(formField)
            If Not pc.isClusterController Then
                drv = DirectCast(clusterController.SelectedItem, DataRowView)
                If drv Is Nothing Then
                    ctrl.Text = ""
                    'ElseIf Me.isClusterController AndAlso TypeOf ctrl Is ComboBox Then
                    '    'AndAlso DirectCast(ctrl, ComboBox).Items.Count > 1
                    '    ctrl.Text = ""
                Else
                    ctrl.Text = drv(pc.srcField).ToString()
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Populate cluster members based on applicable selections on the form.
    ''' </summary>
    ''' <param name="ctrl">Cluster controlling control</param>
    ''' <remarks>Too convoluted. Lenght of the sub is testament to that. Prime for refactoring.</remarks>
    Private Sub populateClusterEntry(ByVal ctrl As ComboBox)
        Dim primary As String
        Dim rb As RadioButton
        Dim dt As DataTable = DirectCast(ctrl.DataSource, DataView).Table
        Dim matchingRows As Array = dt.Select("OrderedId=-9999")
        ' If we have already created the cluster entry...
        If matchingRows.Length > 0 Then
            ' then nothing to do.
            Return
        End If

        Dim dr As DataRow = dt.NewRow()
        Dim pc As PageController
        Dim tagValue As String

        ' For each PageController that is part of the cluster...
        For Each formField As String In clusterTags(Me.clusterName)
            ' check to see if there is a DB column for the tag it controls...
            pc = PageController.thatControls(formField)
            If dr.Table.Columns.Contains(pc.srcField) Then
                ' and set the column value to tag value.
                tagValue = SimpleGetProperty(pc.tag)
                If tagValue.Trim() > "" Then
                    dr(pc.srcField) = tagValue

                    ' Hooks for cntorg and cntper (controlled by radio button). Could be better...
                    If pc.srcField = "cntorg" OrElse pc.srcField = "cntper" Then
                        primary = pc.formFieldName.Substring(0, pc.formFieldName.Length - (pc.srcField.Length + 1))
                        rb = DirectCast(ctrl.FindForm(), EditorForm).allControls(primary)
                        rb.Checked = True
                    End If

                End If
            End If
        Next

        ' Set mandatory fields
        dr("OrderedId") = -9999 'Negative ids reserved for our use
        dr("default") = False

        'Hooks for cluster controllers. clusterInfo computation method for metadata record should match 
        'the same for the source query in DB where applicable.
        If Me.srcTable = "Contact_Information" Then
            dr("clusterInfo") = dr("cntorg") & " | " & dr("cntper") & " (from metadata record)"
        ElseIf Me.srcTable = "1g_BoundingBox" Then
            dr("clusterInfo") = "N:" & dr("northbc") & " E:" & dr("eastbc") & _
            " S:" & dr("southbc") & " W:" & dr("westbc") & " (from metadata record)"
            ' Don't let ESRI inserted defaults creep in.
            If dr("clusterInfo").ToString.Contains("REQUIRED: ") Then dr("clusterInfo") = "N: E: S: W:"
        ElseIf Me.srcTable = "1k_Constraints" Then
            dr("clusterInfo") = dr("secclass") & " (from metadata record)"
        ElseIf Me.srcTable = "2b_CoordinateSystem" Then
            If Not tagIsEmpty("spref/horizsys/geograph") Then
                ctrl.SelectedValue = "Geographic"
            ElseIf Not tagIsEmpty("spref/horizsys/planar/gridsys/utm/utmzone") Then
                ctrl.SelectedValue = "Universal Transverse Mercator"
            ElseIf Not tagIsEmpty("spref/horizsys/planar/gridsys/spcs/spcszone") Then
                ctrl.SelectedValue = "State Plane Coordinate System"
            ElseIf Not tagIsEmpty("spref/horizsys/planar/gridsys/mgmg4coz") Then
                ctrl.SelectedValue = "Minnesota County Coordinate System"
            Else
                ctrl.SelectedValue = ""
            End If
            ' If no known coord sys, leave empty so we don't mess with it.
            dr("clusterInfo") = ""
        ElseIf Me.srcTable = "2b_Zone" Then
            Dim zone As String = ""
            Dim coordsys As String = ""
            zone = Utils.nv(SimpleGetProperty("spref/horizsys/planar/gridsys/spcs/spcszone"), "")
            If zone = "" Then
                zone = Utils.nv(SimpleGetProperty("spref/horizsys/planar/gridsys/utm/utmzone"), "")
                coordsys = "Universal Transverse Mercator"
            End If
            If zone = "" Then
                ' This seriously crappy system has multiple code values tha can return 
                '' i.e. one zone '39' returns 3 legitimate rows
                ' for our purposes, in the editor, we will display this as whatever row
                ' gets returned first.
                zone = Utils.nv(SimpleGetProperty("spref/horizsys/planar/gridsys/mgmg4coz"), "")
                coordsys = "Minnesota County Coordinate System"
            End If
            matchingRows = dt.Select("zone='" & zone.Replace("'", "''") & "'" & " AND CoordinateSystem='" & coordsys & "'")
            If matchingRows.Length > 0 Then
                ctrl.SelectedValue = matchingRows(0)("clusterInfo").ToString()
            Else
                ctrl.SelectedValue = ""
            End If
        ElseIf Me.srcTable = "2c_Units" Then
            If DirectCast(ctrl.FindForm(), EditorForm).spref_horizsys_CoordinateSystem.Text = "Geographic" Then
                ctrl.SelectedValue = Utils.nv(SimpleGetProperty("spref/horizsys/geograph/geogunit"), "")
            Else
                ctrl.SelectedValue = Utils.nv(SimpleGetProperty("spref/horizsys/planar/planci/plandu"), "")
            End If
        ElseIf Me.srcTable = "2d_Datum" Then
            ctrl.SelectedValue = Utils.nv(SimpleGetProperty("spref/horizsys/geodetic/horizdn"), "")
        Else
            dr("clusterInfo") = "Unrecognized value (from metadata record)"
        End If
        If ctrl.SelectedValue Is DBNull.Value OrElse dr("clusterInfo").ToString() > "" Then
            dr("default") = False
            dt.Rows.Add(dr)
            ctrl.SelectedIndex = ctrl.Items.Count - 1
        End If
    End Sub

    ''' <summary>
    ''' Load the ListBox object associated with this PageController.
    ''' </summary>
    ''' <param name="ctrl">The ListBox control</param>
    ''' <remarks></remarks>
    Sub listBoxLoader(ByVal ctrl As ListBox, Optional ByVal reload As Boolean = False)
        If Me.srcTable Is Nothing Then Exit Sub

        Dim selectList As New List(Of String)
        If reload Then
            ' Preserve current selections on the GUI
            For Each selectedIdx As Integer In ctrl.SelectedIndices
                selectList.Add(ctrl.Items(selectedIdx)(ctrl.ValueMember))
                'Dim sel As Object = Utils.nv(ctrl.Items(selectedIdx)(ctrl.ValueMember), "")
                'If sel > "" Then
                '    selectList.Add(sel)
                'End If
            Next
        Else
            ' Get entries from metadata
            'Dim props As Object = DirectCast(iXPS, ESRI.ArcGIS.esriSystem.IPropertySet).GetProperty(tag)
            Dim props As Object = iXPS.GetProperty(tag)
            ' If the tag exists in the metadata...
            If props IsNot Nothing Then
                For Each tagValue As String In props
                    selectList.Add(tagValue.Trim)
                Next
            End If
        End If

        Dim SQLStr As String = "SELECT DISTINCT [" & Me.srcField & "], default FROM " & Me.srcTable & " WHERE [" & Me.srcField & "] IS NOT NULL ORDER BY [" & Me.srcField & "]"
        'Debug.Print(SQLStr)
        Dim dt As DataTable = Utils.datatableFromSQL(SQLStr)

        ' Bind the ListBox to a DataTable that contains the options
        ctrl.DataSource = dt
        ctrl.DisplayMember = Me.srcField
        ctrl.ValueMember = Me.srcField

        ' Add any entries that do not exist in the values read from DB.
        For Each tagValue As Object In selectList
            ' and if its value is not blank...
            If tagValue <> "" Then
                ' check to see if each tag value exists in options read from DB...
                Dim matchingRows As Array = dt.Select(ctrl.ValueMember & "='" & tagValue.ToString().Trim().Replace("'", "''") & "'")
                ' and, if not, ...
                If matchingRows.Length = 0 Then
                    ' create an option for the tag value.
                    dt.Rows.Add(tagValue, False)
                End If
            End If
        Next

        ' Initialize selections
        ctrl.SelectedIndices.Clear()
        For j As Integer = 0 To ctrl.Items.Count - 1
            For Each tagValue As String In selectList
                If Utils.nv(ctrl.Items(j)(ctrl.ValueMember), "") = tagValue Then
                    ctrl.SelectedIndices.Add(j)
                End If
            Next
        Next
    End Sub

    ''' <summary>
    ''' Load the DataGridView object associated with this PageController.
    ''' </summary>
    ''' <param name="ctrl">The DataGridView control</param>
    ''' <remarks></remarks>
    Private Sub dataGridViewLoaderFromMetadata(ByVal ctrl As DataGridView, Optional ByVal reload As Boolean = False)
        If reload Then Return

        ' Get entries from metadata, what we're really looking for is the no if instances of the cluster controller in metadata
        ' AE: Afterthought: Can't we simply say Me.clusterName here?
        'Dim props As Object = DirectCast(iXPS, ESRI.ArcGIS.esriSystem.IPropertySet).GetProperty(Me.parentTag)
        Dim props As Object = iXPS.GetProperty(Me.parentTag)
        If props Is Nothing Then Return

        Dim cnt As Integer = 0
        For Each prop As String In props
            If prop = Me.leafTag Then cnt += 1
        Next
        For i As Integer = 0 To cnt - 1
            Dim rowValues As New List(Of String)
            For Each clusterTag As String In clusterTags(Me.clusterName)
                Dim pc As PageController = PageController.thatControls(clusterTag)
                If Not pc.isClusterController Then
                    'props = DirectCast(iXPS, ESRI.ArcGIS.esriSystem.IPropertySet).GetProperty(pc.tag.Replace("[?]", "[" & i & "]"))
                    props = iXPS.GetProperty(pc.tag.Replace("[?]", "[" & i + 1 & "]"))
                    If props IsNot Nothing Then
                        rowValues.Add(props(0))
                    Else
                        rowValues.Add("")
                    End If
                End If
            Next
            Dim row As New DataGridViewRow
            row.CreateCells(ctrl)
            For j As Integer = 0 To row.Cells.Count - 1
                row.Cells(j).Value = rowValues(j)
            Next
            ctrl.Rows.Add(row)
        Next
    End Sub

    ''' <summary>
    ''' Load the ListBox object associated with this PageController.
    ''' </summary>
    ''' <param name="ctrl">The ListBox control</param>
    ''' <remarks></remarks>
    Private Sub listBoxLoaderFromMetadata(ByVal ctrl As ListBox, Optional ByVal reload As Boolean = False)
        If reload Then Return

        ' Get entries from metadata
        'Dim props As Object = DirectCast(iXPS, ESRI.ArcGIS.esriSystem.IPropertySet).GetProperty(tag.Replace("[?]", ""))
        Dim props As Object = iXPS.GetProperty(tag.Replace("[?]", ""))
        ' If the tag exists in the metadata...
        If props IsNot Nothing Then
            For Each tagValue As String In props
                ctrl.Items.Add(tagValue.Trim)
                ctrl.SelectedItems.Add(tagValue.Trim)
            Next
        End If
    End Sub


    ''' <summary>
    ''' Set defaults for each PageController controlled form controls on the given tab.
    ''' </summary>
    ''' <param name="frm">Editor form</param>
    ''' <param name="tabNo">Tab number where default settings will be applied</param>
    ''' <remarks></remarks>
    Public Shared Sub DefaultSetter(ByRef frm As EditorForm, ByVal tabNo As Integer)
        If MessageBox.Show("This will set all elements on this page to their default values, as applicable. Do you really want to proceed?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
            Return
        End If
        For Each pc As PageController In PageController.orderedPageControllers()
            If pc.tabNo = frm.tcEME.SelectedIndex + 1 Then
                pc.setDefault(frm)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Set the form control controlled by this PageController to its default value.
    ''' </summary>
    ''' <param name="frm">Editor form</param>
    ''' <remarks></remarks>
    Sub setDefault(ByRef frm As EditorForm)
        Dim ctrl As Control
        ctrl = frm.getControlForTag(formFieldName)
        'Debug.Print("setDefault: " & formFieldName)
        'Debug.Print(ctrl.GetType.ToString)

        If srcTable IsNot Nothing Then
            If TypeOf ctrl Is ListBox Then
                listBoxSetDefault(ctrl)
            ElseIf TypeOf ctrl Is ComboBox Then
                comboBoxSetDefault(ctrl)
            ElseIf TypeOf ctrl Is TextBox Then
                textBoxSetDefault(ctrl)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Set the ListBox control controlled by this PageController to its default value(s).
    ''' </summary>
    ''' <param name="ctrl">Editor form</param>
    ''' <remarks></remarks>
    Private Sub listBoxSetDefault(ByVal ctrl As ListBox)
        ctrl.SelectedIndices.Clear()
        For j As Integer = 0 To ctrl.Items.Count - 1
            If ctrl.Items(j)("default") Then
                ctrl.SelectedIndices.Add(j)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Set the ComboBox control controlled by this PageController to its default value.
    ''' </summary>
    ''' <param name="ctrl">Editor form</param>
    ''' <remarks></remarks>
    Private Sub comboBoxSetDefault(ByVal ctrl As ComboBox)
        ctrl.SelectedText = ""
        For j As Integer = 0 To ctrl.Items.Count - 1
            If ctrl.Items(j)("default") Then
                ctrl.SelectedIndex = j
            End If
        Next
        If Me.isClusterController Then
            Me.propagateClusterSelectionChanged(ctrl, Nothing)
        End If
    End Sub

    ''' <summary>
    ''' Set the TextBox control controlled by this PageController to its default value.
    ''' </summary>
    ''' <param name="ctrl">Editor form</param>
    ''' <remarks></remarks>
    Private Sub textBoxSetDefault(ByVal ctrl As TextBox)
        Dim SQLStr As String = "SELECT [" & srcField & "] as fldVal, default FROM " & srcTable & IIf(Me.isPartOfCluster, "_cluster", "") & " WHERE default"
        Dim con As OleDb.OleDbConnection = Nothing
        Dim dr As OleDb.OleDbDataReader = readerForSQL(SQLStr, con)
        Do While dr.Read()
            ctrl.Text = dr("fldVal").ToString()
        Loop
        dr.Close()
        con.Close()
    End Sub

    ''' <summary>
    ''' Attempt webservice or local validation depending on current setting.
    ''' </summary>
    ''' <param name="frm">Editor form</param>
    ''' <remarks></remarks>
    Public Shared Sub ValidateAction(ByRef frm As EditorForm)
        If Not My.Settings.ViewValidationResultsInEME AndAlso Not My.Settings.ViewValidationResultsInBrowser Then
            MsgBox("You must select where you want to view validation results.")
            Exit Sub
        End If
        If My.Settings.ValidationTimeout > 0 Then
            ValidationTimer.Interval = My.Settings.ValidationTimeout * 1000
        Else
            ValidationTimer.Interval = 30 * 1000    'Default to 30 seconds
        End If
        frm.turnOffWarnings()
        PageController.PageSaver(frm)
        ValidationWorker.WorkerSupportsCancellation = True

        Dim args As New Hashtable(2)
        args("Form") = frm
        If My.Settings.ValidationEnabled Then
            ProgressDialog.ShowMessage("Performing validation via webservice. Timeout set to " & My.Settings.ValidationTimeout & " seconds...")
            ValidationTimer.Start()
            args("ValidationMode") = ValidationMode.Webservice
            ValidationWorker.RunWorkerAsync(args)
        Else
            ProgressDialog.ShowMessage("Performing validation via local service.")
            args("ValidationMode") = ValidationMode.Local
            ValidationWorker.RunWorkerAsync(args)
        End If
    End Sub

    ''' <summary>
    ''' Display validation results using the application associated with .xml files.
    ''' This is typically Internet Explorer, though not necessarily.
    ''' </summary>
    ''' <param name="filename">String containing name of file to be displayed.</param>
    ''' <remarks></remarks>
    Shared Sub showValidationResultsInBrowser(ByVal filename As String)
        Utils.OpenInIE(filename)
    End Sub

    ''' <summary>
    ''' Display validation results on EME GUI. 
    ''' </summary>
    ''' <param name="frm">The EditorForm object used for EME GUI.</param>
    ''' <param name="dom">The DOM containing the results of validation.</param>
    ''' <remarks>Real action is performed by showValidationResultsInEMEDo().</remarks>
    Private Shared Sub showValidationResultsInEME(ByVal frm As EditorForm, ByVal dom As XmlDocument)
        Dim errNodes As XmlNodeList = dom.GetElementsByTagName("err")
        For Each errNode As XmlNode In errNodes
            Try
                ' We try to construct the name of the form control that corresponds to the element that had an error/warning.
                ' Note that we typically use a variation of the XSL pattern for an FGDC element as the name of the control associated with it.
                Dim msg As String = errNode.SelectSingleNode("type").InnerText.Trim & " : " & errNode.SelectSingleNode("message").InnerText.Trim
                Dim attnStr As String = errNode.SelectSingleNode("errid").InnerText.Trim
                If attnStr.Trim = "" Then Continue For
                If attnStr.StartsWith("/metadata/") Then
                    attnStr = attnStr.Substring(10, attnStr.Length - 10)
                ElseIf attnStr.StartsWith("/") Then
                    attnStr = attnStr.Substring(1, attnStr.Length - 1)
                End If
                attnStr = stripNonAlphanumeric(attnStr, "_")
                'Debug.Print("*********************************")
                'Debug.Print(attnStr)
                showValidationResultsInEMEDo(frm, attnStr, msg)
            Catch ex As Exception
                ErrorHandler(ex)
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Delegate for showValidationResultsInEMEDo().
    ''' </summary>
    ''' <param name="frm"></param>
    ''' <param name="ctrlName"></param>
    ''' <param name="msg"></param>
    ''' <remarks>See showValidationResultsInEMEDo() for parameters. The delegate is required cross-thread calling of target.</remarks>
    Delegate Sub showValidationResultsInEMEDoCallback(ByVal frm As EditorForm, ByVal ctrlName As String, ByVal msg As String)

    ''' <summary>
    ''' The workhorse subroutine to display validation results on the EME GUI using little blinking led lights and hover tooltips.
    ''' </summary>
    ''' <param name="frm">The EditorForm object used for EME GUI.</param>
    ''' <param name="ctrlName">Name of the control for which to display validation results.</param>
    ''' <param name="msg">The message to be displayed in a hover tooltip.</param>
    ''' <remarks></remarks>
    Private Shared Sub showValidationResultsInEMEDo(ByVal frm As EditorForm, ByVal ctrlName As String, ByVal msg As String)
        If frm.InvokeRequired Then
            ' If this is a cross-thread call, then we need a callback so that we execute on the same thread that created 
            ' the EditorForm and be allowed to manipulate it.
            Dim d As New showValidationResultsInEMEDoCallback(AddressOf showValidationResultsInEMEDo)
            frm.Invoke(d, New Object() {frm, ctrlName, msg})
        Else
            Dim done As Boolean = False
            While Not done
                'Debug.Print(ctrlName)
                ' If we can find the warning control associated with the control name...
                If frm.allControls.ContainsKey(ctrlName & GlobalVars.idSep & "warning") Then
                    'frm.allControls(ctrlName).Visible = True
                    ' We shouldn't need this check. Hmmmm...
                    If DirectCast(frm.allControls(ctrlName & GlobalVars.idSep & "warning"), PictureBox).InvokeRequired Then
                    Else
                        ' Make the warning control visible.
                        Dim pb As PictureBox = frm.allControls(ctrlName & GlobalVars.idSep & "warning")
                        pb.Visible = True
                        ' Prep the message, special case for lineage.
                        If msg = "Lineage is required in Data_Quality_Information" Then
                            msg += " (one processing step is required)"
                        End If
                        Dim curtt As String = frm.HoverToolTip.GetToolTip(pb).Trim
                        ' Append the current message.
                        frm.HoverToolTip.SetToolTip(pb, IIf(curtt = "", msg, curtt & vbCrLf & vbCrLf & msg))
                    End If
                    Exit Sub
                Else
                    ' Proceed with the parent control name if we couldn't find a warning control.
                    If ctrlName > "" Then
                        ctrlName = findParentName(ctrlName)
                    Else
                        ' No more parent to go when control name is empty string.
                        done = True
                    End If
                End If
            End While
        End If
    End Sub


    ''' <summary>
    ''' Event handler that cancels the executing validation thread when validation timeout occurs.
    ''' </summary>
    ''' <param name="sender">Event sender. Not used.</param>
    ''' <param name="e">Event arguments. Not used.</param>
    ''' <remarks></remarks>
    Private Shared Sub ValidationTimerFired(ByVal sender As Object, ByVal e As System.EventArgs) Handles ValidationTimer.Tick
        ValidationTimer.Stop()
        If ValidationWorker.IsBusy Then
            ValidationWorker.CancelAsync()
            ProgressDialog.ShowMessage("Validation via webservice has timed out...")
        End If
    End Sub

    ''' <summary>
    ''' The worker process that will run in a separate thread and perform the actual validation work.
    ''' </summary>
    ''' <param name="sender">Event sender. Not used.</param>
    ''' <param name="e">Event arguments. Used to return information about the type and outcome of validation work performed.</param>
    ''' <remarks></remarks>
    Private Shared Sub ValidateAction(ByVal sender As Object, ByVal e As DoWorkEventArgs) Handles ValidationWorker.DoWork
        Dim args As Hashtable = e.Argument
        args("Success") = False
        e.Result = args

        Dim iXPSv As New XmlMetadata
        iXPSv.SetXml(iXPS.GetXml(""))

        ' Delete ESRI elements to avoid unnecessary validation warnings. Handle ISO elements better next time.
        iXPSv.DeleteProperty("Esri")
        iXPSv.DeleteProperty("mdDateSt")
        ' Delete keyword themes that don't have keywords
        'iXPSv.checkDeleteKTTags()

        args("Metadata") = iXPSv.GetXml("")
        args("Success") = ValidateActionDo(args)
    End Sub

    ''' <summary>
    ''' Event handler that runs when validation worker thread completes its run.
    ''' Informs user if something went wrong.
    ''' </summary>
    ''' <param name="sender">Event sender. Not used.</param>
    ''' <param name="e">Event arguments. Used to examine information about the type and outcome of validation work performed.</param>
    ''' <remarks></remarks>
    Private Shared Sub ValidationWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs) Handles ValidationWorker.RunWorkerCompleted
        ProgressDialog.CancelMessage()
        Dim args As Hashtable = DirectCast(e.Result, Hashtable)
        If Not args("Success") Then
            If args("ValidationMode") = ValidationMode.Webservice Then
                'Webservice validation failed, try local
                ProgressDialog.ShowMessage("Validation via webservice did not succeed. Trying local validation...")
                'ValidationWorker = New BackgroundWorker
                args("ValidationMode") = ValidationMode.Local
                ValidationWorker.RunWorkerAsync(args)
            Else
                'Local validation failed as well
                MessageBox.Show("Validation attempts have failed. Please check to make sure that you have a working internet connection to access validation webservice." + vbCrLf + "You may need to reinstall this application to resolve local validation failure.")
            End If
            Exit Sub
        End If

        Dim frm As EditorForm = args("Form")
        If My.Settings.ViewValidationResultsInBrowser Then
            OpenInIE(args("Filename"))
        End If
        If My.Settings.ViewValidationResultsInEME Then
            showValidationResultsInEME(frm, args("DOM"))
        End If
    End Sub

    ''' <summary>
    ''' Determine the parent control's name.
    ''' </summary>
    ''' <param name="name">String containing a control's name whose parent is being sought,</param>
    ''' <returns>String containing parent control's name.</returns>
    ''' <remarks>This is an imperfect method for determining the parent of a control as it relies on removing 
    ''' the part after the last name separator chracter. Caller should not assume that the returned name corresponds 
    ''' to an existing control's name.</remarks>
    Public Shared Function findParentName(ByVal name As String) As String
        Dim idx As String = name.LastIndexOf("_")
        If idx < 0 Then
            Return ""
        Else
            Return name.Substring(0, idx)
        End If
    End Function

    ''' <summary>
    ''' Determine if this PageController is part of a cluster.
    ''' </summary>
    ''' <returns>Returns true if this PageController is part of a cluster. Otherwise, returns false.</returns>
    ''' <remarks></remarks>
    Private ReadOnly Property isPartOfCluster() As Boolean
        Get
            Return clusterName IsNot Nothing
        End Get
    End Property

    ''' <summary>
    ''' Determine the name of the form control associated with this PageController.
    ''' </summary>
    ''' <returns>Returns the name of the form control associated with this PageController.</returns>
    ''' <remarks></remarks>
    Private ReadOnly Property formFieldName() As String
        Get
            Return Me.formFieldName_
        End Get
    End Property

    ''' <summary>
    ''' Determine the sort property of this PageController.
    ''' </summary>
    ''' <returns>Returns the sort property of this PageController.</returns>
    ''' <remarks>Sort property determines the order in which PageControllers, hence their, FGDC elements are processed.</remarks>
    Public ReadOnly Property sortProperty() As Long
        Get
            Return Me.orderedId
        End Get
    End Property

    ''' <summary>
    ''' Determine if this PageController is a cluster controller.
    ''' </summary>
    ''' <returns>Returns true if this PageController is a cluster controller. Otherwise, returns false.</returns>
    ''' <remarks></remarks>
    Private ReadOnly Property isClusterController() As Boolean
        Get
            Return Me.clusterName = Me.formFieldName
        End Get
    End Property

    ''' <summary>
    ''' Add this PageController to its cluster, if applicable.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub addToCluster()
        ' If associated with a cluster ...
        If Me.isPartOfCluster Then
            ' Auto-create cluster if it doesn't exist already
            If Not clusterTags.ContainsKey(Me.clusterName) Then
                clusterTags.Add(Me.clusterName, New List(Of String))
            End If
            ' Add tag to the cluster
            clusterTags(Me.clusterName).Add(Me.formFieldName)
        End If
    End Sub

    ''' <summary>
    ''' Save all PageControllers back to metadata record.
    ''' </summary>
    ''' <param name="frm">Editor form</param>
    ''' <remarks></remarks>
    Shared Sub PageSaver(ByRef frm As EditorForm)
        eainfoSave(frm.eainfoXF.construct())

        For Each pc As PageController In PageController.orderedPageControllers()
            'Debug.Print(pc.formFieldName)
            pc.save(frm.allControls(pc.formFieldName))
        Next
    End Sub

    ''' <summary>
    ''' Save the eainfo section of the metadata.
    ''' </summary>
    ''' <param name="eainfoXml">The XML representation of eainfo FGDC element as manipulated by EME.</param>
    ''' <remarks>This is a divergence from the more straightforward mapping between form controls and 
    ''' CSDGM sections that they control implemented using PageController objects. This is a necessary evil as 
    ''' eainfo is too complicated to be mapped to a single form control.</remarks>
    Shared Sub eainfoSave(ByVal eainfoXml As String)
        Dim dom As New XmlDocument
        dom.LoadXml(iXPS.GetXml(""))
        Dim nodeMd As XmlNode = dom.SelectSingleNode("/metadata")
        ' Add eainfo element if not there
        If nodeMd.SelectSingleNode("eainfo") Is Nothing Then
            nodeMd.AppendChild(dom.CreateElement("eainfo"))
        End If
        nodeMd.SelectSingleNode("eainfo").InnerXml = eainfoXml
        iXPS.SetXml(dom.InnerXml)
    End Sub

    ''' <summary>
    ''' Save this PageController back to metadata record.
    ''' </summary>
    ''' <param name="ctrl">The control associated with this PageController.</param>
    ''' <remarks></remarks>
    Private Sub save(ByVal ctrl As Control)
        If Me.disabled Then Return
        If Me.isClusterController AndAlso Not TypeOf ctrl Is DataGridView Then
            ' If EME has been instructed to control entire compound element...
            If Me.clusterUpdate Then
                ' then wipe it off 
                iXPS.DeleteProperty(Me.tag)
            End If
        Else
            If TypeOf ctrl Is ListBox Then
                listBoxSaver(ctrl)
            ElseIf TypeOf ctrl Is ComboBox Then
                comboBoxSaver(ctrl)
            ElseIf TypeOf ctrl Is TextBox Then
                textBoxSaver(ctrl)
            ElseIf TypeOf ctrl Is DataGridView Then
                dataGridViewSaver(ctrl)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Update/delete the metadata tag associated with this PageController.
    ''' </summary>
    ''' <param name="tagValue">Value of the tag</param>
    ''' <param name="index">Index of the tag if a repeated tag. 1-based.</param>
    ''' <remarks>If the value is empty, the tag is deleted from the metadata record.</remarks>
    Private Sub updateMetadata(ByVal tagValue As String, Optional ByVal index As Integer = 0)
        Dim riggedTag As String = Me.tag
        ' If an index is provided, rig the tag path such that it points to the index-th occurence (0-based) of the element.
        If index > 0 Then
            ' If no subscript is provided, we stick one at the end
            If Not riggedTag.Contains("[?]") Then
                riggedTag &= "[?]"
            End If
            ' Replace subscript with actual index
            riggedTag = riggedTag.Replace("[?]", "[" & index & "]")
        ElseIf tagValue.Trim = "" AndAlso riggedTag.Contains("[?]") Then
            ' We are removing the tag, so get rid of subscript
            riggedTag = riggedTag.Replace("[?]", "")
        ElseIf Regex.IsMatch(riggedTag, "\[\d+]$") Then
            ' tag contains an index but we will overwrite it to fit reality
            Dim m As Match = Regex.Match(riggedTag, "^(.*)\[(\d+)]$")
            Dim intIdx As Integer = Integer.Parse(m.Groups(2).Value)
            If intIdx = 1 Then
                iXPS.DeleteProperty(m.Groups(1).Value)
            End If
            Dim dynIdx As Integer = iXPS.CountX(m.Groups(1).Value)
            If intIdx > dynIdx Then
                dynIdx += 1
                riggedTag = riggedTag.Substring(0, riggedTag.LastIndexOf("[")) + "[" + dynIdx.ToString + "]"
            End If
        End If

        Try
            'riggedTag = realityFixXSLPattern(riggedTag)
            'Debug.Print(riggedTag & " : " & tagValue)
            ' If tag value is not empty...
            If tagValue.Trim() > "" Then
                ' update/create the tag with new value.
                iXPS.SetPropertyX(riggedTag, tagValue)
            Else
                ' otherwise, delete the tag.
                'iXPS.DeleteProperty(riggedTag)
                Utils.deleteProperty(riggedTag, True)
            End If
        Catch ex As Exception
            MessageBox.Show("Error while updating metadata for: " & riggedTag & vbCrLf & "We will attempt to proceed. Please record this message and the error description that will follow for error resolution.")
            Utils.ErrorHandler(ex)
        End Try
    End Sub


    ''' <summary>
    ''' Save the contents of the TextBox object associated with this PageController to the metadata record.
    ''' </summary>
    ''' <param name="ctrl">The TextBox object associated with the PageController</param>
    ''' <remarks></remarks>
    Private Sub textBoxSaver(ByRef ctrl As TextBox)
        'AE: Generalize this for DataGridRowView cluster controllers
        If Me.isPartOfCluster AndAlso Me.parentTag = "dataqual/lineage/procstep" Then
            Return
        End If

        updateMetadata(ctrl.Text)

        ' Hooks for cntorg and cntper (controlled by radio button). Could be better...
        If Me.srcField = "cntorg" OrElse Me.srcField = "cntper" Then
            Dim primary As String = Me.formFieldName.Substring(0, Me.formFieldName.Length - (Me.srcField.Length + 1))
            Dim rb As RadioButton = DirectCast(ctrl.FindForm(), EditorForm).allControls(primary)
            'Debug.Print(ctrl.Name)
            'Debug.Print(iXPS.GetXml("idinfo/ptcontac"))
            If Not rb.Checked Then
                iXPS.DeleteProperty(primary.Replace("_", "/"))
                'Debug.Print(iXPS.GetXml("idinfo/ptcontac"))
                'Debug.Print("---------------------------------------------")
            End If
        End If
        'If Me.srcField = "cntorg" OrElse Me.srcField = "cntper" Then
        '    Debug.Print(iXPS.GetXml("idinfo/ptcontac"))
        'End If
    End Sub

    ''' <summary>
    ''' Save the contents of the ListBox object associated with this PageController to the metadata record.
    ''' </summary>
    ''' <param name="ctrl">The ListBox object associated with the PageController</param>
    ''' <remarks></remarks>
    Private Sub listBoxSaver(ByRef ctrl As ListBox)
        'Debug.Print(ctrl.Name)
        ' Force deletion of all repeated tags by passing empty string.
        updateMetadata("")
        For i As Integer = 0 To ctrl.SelectedItems.Count - 1
            If TypeOf ctrl.SelectedItems(i) Is String Then
                updateMetadata(ctrl.SelectedItems(i), i + 1)
            Else
                updateMetadata(ctrl.SelectedItems(i)(Me.srcField).ToString(), i + 1)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Save the contents of the ComboBox object associated with this PageController to the metadata record.
    ''' </summary>
    ''' <param name="ctrl">The ComboBox object associated with the PageController</param>
    ''' <remarks></remarks>
    Private Sub comboBoxSaver(ByRef ctrl As ComboBox)
        updateMetadata(ctrl.Text)
    End Sub

    ''' <summary>
    ''' Save the contents of the DataGridView object associated with this PageController to the metadata record.
    ''' </summary>
    ''' <param name="ctrl">The DataGridView object associated with the PageController</param>
    ''' <remarks></remarks>
    Private Sub dataGridViewSaver(ByRef ctrl As DataGridView)
        ' If we control the whole cluster...
        If Me.clusterUpdate Then
            ' wipe it out.
            Me.updateMetadata("")
        Else
            ' otherwise, only wipe tags in cluster. 
            ' Warning: this is not a good idea unless all possible tags are included in the cluster (potential for mispaired tags!)
            For Each clusterTag As String In clusterTags(Me.clusterName)
                Dim pc As PageController = PageController.thatControls(clusterTag)
                If Not pc.isClusterController Then
                    ' Force deletion of all repeated tags by passing empty string.
                    pc.updateMetadata("")
                End If
            Next
        End If
        ' Save each cell value to appropriate metadata tag
        For i As Integer = 0 To ctrl.RowCount - 2
            Dim j As Integer = 0
            For Each clusterTag As String In clusterTags(Me.clusterName)
                Dim pc As PageController = PageController.thatControls(clusterTag)
                If Not pc.isClusterController Then
                    Dim value As String = ctrl.Rows(i).Cells(j).Value
                    pc.updateMetadata(IIf(value Is Nothing, "None", value), i + 1)
                    j += 1
                End If
            Next
        Next
    End Sub

    ''' <summary>
    ''' Order all PageControllers.
    ''' </summary>
    ''' <returns>Returns an ArrayList of PageController objects sorted by their "orderedId"</returns>
    ''' <remarks></remarks>
    Private Shared Function orderedPageControllers() As ArrayList
        Dim opcs As New ArrayList
        opcs.AddRange(HiveMind.Values())
        opcs.Sort()
        Return opcs
    End Function

    ''' <summary>
    ''' Determine if the control associated with this PageController is on a given tab.
    ''' </summary>
    ''' <param name="tabNum">Tab number</param>
    ''' <returns>Returns true if the this PageController's form control is on the given tab. Otherwise, returns false.</returns>
    ''' <remarks></remarks>
    Public Function isOnTab(ByVal tabNum As Integer) As Boolean
        Return Me.tabNo = tabNum
    End Function

    ''' <summary>
    ''' Compares the PageController object to another based on orderedId
    ''' </summary>
    ''' <param name="obj">PageController object to compare against</param>
    ''' <returns>Returns a negative value if this PageController preceeds obj, 0 if equivalent (should not happen) or a positive value it succeeds obj.</returns>
    ''' <remarks></remarks>
    Public Function CompareTo(ByVal obj As Object) As Integer Implements System.IComparable.CompareTo
        Return Me.orderedId.CompareTo(DirectCast(obj, PageController).orderedId)
    End Function

    ''' <summary>
    ''' Removes text content from all controls that start with "REQUIRED: ", something typically inserted by ESRI tools.
    ''' </summary>
    ''' <param name="frm">The EditorForm object whose controls will be inspected.</param>
    ''' <returns>Boolean value indicating if any modifications were made.</returns>
    ''' <remarks></remarks>
    Public Shared Function wipeEsriRequired(ByVal frm As EditorForm) As Boolean
        Dim modified As Boolean = False
        For Each pc As PageController In HiveMind.Values()
            Dim ctrl As Control = frm.getControlForTag(pc.formFieldName)
            If TypeOf ctrl Is TextBoxBase OrElse TypeOf ctrl Is ComboBox Then
                If ctrl.Text.StartsWith("REQUIRED: ") Then
                    ctrl.Text = ""
                End If
            End If
        Next
        Return modified
    End Function

    ''' <summary>
    ''' Find and replace all occurrences of some string with another one for all controls.
    ''' </summary>
    ''' <param name="frm">The EditorForm object whose controls will be inspected.</param>
    ''' <param name="oldValue">String containing text to be replaced.</param>
    ''' <param name="newValue">String containing text to replace with.</param>
    ''' <returns>The number of replacements made.</returns>
    ''' <remarks></remarks>
    Public Shared Function findAndReplaceAll(ByVal frm As EditorForm, ByVal oldValue As String, ByVal newValue As String) As Integer
        Dim cnt As Integer = 0
        For Each pc As PageController In HiveMind.Values()
            cnt += pc.findAndReplace(frm, oldValue, newValue)
        Next

        ' Special treatment for DataGridView controls
        'cnt += dataGridViewFindAndReplace(frm.dataqual_lineage_procstep, oldValue, newValue)
        cnt += dataGridViewFindAndReplace(frm.dgv_edom, oldValue, newValue)
        Return cnt
    End Function

    ''' <summary>
    ''' Specialized find and replace for text controlled via a DataGridView control.
    ''' </summary>
    ''' <param name="ctrl">DataGridView control to operate on.</param>
    ''' <param name="oldValue">String containing text to be replaced.</param>
    ''' <param name="newValue">String containing text to replace with.</param>
    ''' <returns>The number of replacements made.</returns>
    ''' <remarks></remarks>
    Public Shared Function dataGridViewFindAndReplace(ByVal ctrl As DataGridView, ByVal oldValue As String, ByVal newValue As String) As Integer
        Dim cnt As Integer = 0
        For i As Integer = 0 To ctrl.RowCount - 1
            If ctrl.Rows(i).IsNewRow Then Continue For
            For j As Integer = 0 To ctrl.ColumnCount - 1
                cnt += Utils.findAndReplaceWithCount(ctrl.Rows(i).Cells(j).Value, oldValue, newValue)
                ' If this is a DataGridViewComboBoxCell, we should add the modified text to the list of items for the combobox.
                Dim cbc As DataGridViewComboBoxCell = TryCast(ctrl.Rows(i).Cells(j), DataGridViewComboBoxCell)
                If cbc IsNot Nothing AndAlso Not cbc.Items.Contains(cbc.Value) Then
                    cbc.Items.Add(cbc.Value)
                End If
            Next
        Next
        Return cnt
    End Function

    ''' <summary>
    ''' Find and replace for all controls that have a "Text" property.
    ''' </summary>
    ''' <param name="frm">The EditorForm object whose controls will be inspected.</param>
    ''' <param name="oldValue">String containing text to be replaced.</param>
    ''' <param name="newValue">String containing text to replace with.</param>
    ''' <returns>The number of replacements made.</returns>
    ''' <remarks></remarks>
    Public Function findAndReplace(ByVal frm As EditorForm, ByVal oldValue As String, ByVal newValue As String) As Integer
        Dim ctrl As Control = frm.getControlForTag(formFieldName)
        If (TypeOf ctrl Is TextBoxBase OrElse TypeOf ctrl Is ComboBox) AndAlso Not Me.isClusterController Then
            Return Utils.findAndReplaceWithCount(ctrl.Text, oldValue, newValue)
        End If
        ' By default, return 0
        Return 0
    End Function

    ''' <summary>
    ''' Disable all controls participating in a cluster.
    ''' </summary>
    ''' <param name="name">String containing name of cluster whose participating controls will be disabled.</param>
    ''' <remarks></remarks>
    Public Shared Sub disableCluster(ByVal name As String)
        name = stripNonAlphanumeric(name, "_")
        For Each pc As PageController In HiveMind.Values
            If pc.clusterName = name Then
                pc.disabled = True
            End If
        Next
    End Sub

    ''' <summary>
    ''' Try to find and open the help page for the metadata element with the given name/xpath
    ''' </summary>
    ''' <param name="name">The xpath for the metadata element/tag</param>
    ''' <param name="defaultHelp">Default hekp page to open if can't find the specific help page</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getHelpPageFor(ByVal name As String, Optional ByVal defaultHelp As String = "/Help_Main.html") As String
        Dim pc As PageController = PageController.thatControls(name)

        If pc Is Nothing Then
            'If we can't find a page controller for the given name, default to main page on help.
            Return defaultHelp
        Else
            Return pc.help
        End If
    End Function


End Class
