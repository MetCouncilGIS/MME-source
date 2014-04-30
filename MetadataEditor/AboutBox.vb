Public NotInheritable Class AboutBox

    Private Sub AboutBox_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = String.Format("About {0}", My.Settings.CommonEditorAbbreviation)
        ' Initialize all of the text displayed on the About Box.
        Me.LabelProductName.Text = My.Application.Info.ProductName
        Me.LabelVersion.Text = String.Format("Full Version {0} Sub {1}", My.Application.Info.Version.ToString, My.Settings.ReleaseNo.ToString)
        Me.LabelCopyright.Text = "Developed for Minnesota Geographic Metadata Guidelines 1.2 by " + My.Application.Info.CompanyName
    End Sub

    ''' <summary>
    ''' Open MnGeo website in a browser window.
    ''' </summary>
    ''' <param name="sender">Event sender. Not used.</param>
    ''' <param name="e">Event arguments. Not used.</param>
    ''' <remarks></remarks>
    Private Sub pbMnGeo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pbMnGeo.Click
        System.Diagnostics.Process.Start("http://www.mngeo.state.mn.us/index.html")
    End Sub

    ''' <summary>
    ''' Open EPA OEI website in a browser window.
    ''' </summary>
    ''' <param name="sender">Event sender. Not used.</param>
    ''' <param name="e">Event arguments. Not used.</param>
    ''' <remarks></remarks>
    Private Sub pbOEI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        System.Diagnostics.Process.Start("http://www.epa.gov/oei/")
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub
End Class
