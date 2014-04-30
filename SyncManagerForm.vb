Imports ESRI.ArcGIS.Geodatabase

''' <summary>
''' Windows Forms form that provides a user interface for controlling
''' metadata synchronization options.
''' </summary>
''' <remarks></remarks>
Public Class SyncManagerForm
    ''' <summary>
    ''' Get a hold of ArcObjects' MetadataSynchronizer that controls all registered synchronizers.
    ''' </summary>
    ''' <remarks></remarks>
    Private msm As MetadataSynchronizerClass

    ''' <summary>
    ''' Simple wrapper class to use IMetadataSynchronizer objects as list items.
    ''' </summary>
    ''' <remarks></remarks>
    Private Class MdSyncAdapter
        ''' <summary>
        ''' The object being wrapped
        ''' </summary>
        ''' <remarks></remarks>
        Private innerObject As IMetadataSynchronizer

        ''' <summary>
        ''' Create a wrapped object.
        ''' </summary>
        ''' <param name="o">An IMetadataSynchronizer to be wrapped.</param>
        ''' <remarks></remarks>
        Sub New(ByVal o As IMetadataSynchronizer)
            innerObject = o
        End Sub

        ''' <summary>
        ''' Get String representation of wrapped object.
        ''' </summary>
        ''' <returns>A string containing the wrapped object's name.</returns>
        ''' <remarks></remarks>
        Overrides Function ToString() As String
            Return innerObject.Name
        End Function
    End Class

    ''' <summary>
    ''' Index of the EPA Synchronizer Helper in the list of syncers
    ''' </summary>
    ''' <remarks></remarks>
    Shared idxEpaSynchronizerHelper As Integer = -1

    ''' <summary>
    ''' Initialize the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Don't go smaller than initial window size
        Me.MinimumSize = Me.Size

        ' Initialize synchronizers list
        msm = New MetadataSynchronizerClass
        Me.clbSynchronizers.Items.Clear()

        For i As Integer = 0 To msm.NumSynchronizers - 1
            Dim ms As IMetadataSynchronizer = msm.GetSynchronizer(i)
            Debug.Print(ms.Name)
            Dim enabled As Boolean = msm.GetEnabled(i)
            If ms.Name.Contains("Helper") Then
                idxEpaSynchronizerHelper = i
            Else
                Me.clbSynchronizers.Items.Add(New MdSyncAdapter(ms), enabled)
            End If
        Next

        ' Initialize sync settings
        Me.cbAttrInfo.Checked = My.Settings.SyncAttrInfo
        Me.cbOnlink.Checked = My.Settings.SyncOnlink
        Me.cbProcSteps.Checked = My.Settings.SyncProcSteps
        Me.cbCSI.Checked = My.Settings.SyncCSI
        Me.cbExtent.Checked = My.Settings.SyncExtent
        Me.cbWipeEsriTagsOnSync.Checked = My.Settings.WipeEsriTagsOnSync
        Me.cbRetainPublishedDocID.Enabled = Me.cbWipeEsriTagsOnSync.Checked
        Me.cbRetainPublishedDocID.Checked = My.Settings.RetainPublishedDocID

        ' Switch to the tab selected in previous session
        Me.tcEPASync.SelectedIndex = My.Settings.SelectedTabIndexEPASync
    End Sub

    ''' <summary>
    ''' Event handler to close the form
    ''' </summary>
    ''' <param name="sender">Event sender. Not used.</param>
    ''' <param name="e">Event arguments. Not used.</param>
    ''' <remarks></remarks>
    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    ''' <summary>
    ''' Event handler to save settings and close the form.
    ''' </summary>
    ''' <param name="sender">Event sender. Not used.</param>
    ''' <param name="e">Event arguments. Not used.</param>
    ''' <remarks></remarks>
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        ' Save A sync settings
        Dim i As Integer = 0

        For j As Integer = 0 To msm.NumSynchronizers - 1
            Dim ms As IMetadataSynchronizer = msm.GetSynchronizer(j)
            Debug.Print(ms.Name)
            If Not ms.Name.Contains("Helper") Then
                msm.SetEnabled(j, Me.clbSynchronizers.GetItemChecked(i))
                i = +1
            End If
        Next

        ' Save sync settings
        My.Settings.SyncAttrInfo = Me.cbAttrInfo.Checked
        My.Settings.SyncOnlink = Me.cbOnlink.Checked
        My.Settings.SyncProcSteps = Me.cbProcSteps.Checked
        My.Settings.SyncCSI = Me.cbCSI.Checked
        My.Settings.SyncExtent = Me.cbExtent.Checked
        My.Settings.WipeEsriTagsOnSync = Me.cbWipeEsriTagsOnSync.Checked
        My.Settings.RetainPublishedDocID = Me.cbRetainPublishedDocID.Checked

        ' Close form
        Me.Close()
    End Sub

    ''' <summary>
    ''' Event handler to keep track of the EPA Synchronizer tab selected.
    ''' </summary>
    ''' <param name="sender">Event sender. Not used.</param>
    ''' <param name="e">Event arguments. Not used.</param>
    ''' <remarks></remarks>
    Private Sub tcEPASync_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tcEPASync.SelectedIndexChanged
        ' Save last selected tab's index for later use
        My.Settings.SelectedTabIndexEPASync = Me.tcEPASync.SelectedIndex
    End Sub

    ''' <summary>
    ''' Event handler to provide help about the form.
    ''' </summary>
    ''' <param name="sender">Event sender. Not used.</param>
    ''' <param name="e">Event arguments. Not used.</param>
    ''' <remarks></remarks>
    Private Sub Help_Synchronizer_____help_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Help_Synchronizer_____help.Click
        Utils.HelpSeeker("Help_Synchronizer.html", GlobalVars.proc)
    End Sub

    ''' <summary>
    ''' Event handler 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub clbSynchronizers_ItemCheck(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles clbSynchronizers.ItemCheck
        If msm.GetSynchronizer(e.Index).Name = "MGMG Synchronizer" AndAlso idxEpaSynchronizerHelper > -1 Then
            msm.SetEnabled(idxEpaSynchronizerHelper, e.NewValue = CheckState.Checked)
        End If
    End Sub

    ''' <summary>
    ''' Event handler to enable/disable "Retain PublishedDocID" setting
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cbWipeEsriTagsOnSync_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbWipeEsriTagsOnSync.CheckedChanged
        Me.cbRetainPublishedDocID.Enabled = Me.cbWipeEsriTagsOnSync.Checked
    End Sub
End Class