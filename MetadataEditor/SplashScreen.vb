''' <summary>
''' Class that provides a splash screen for the MME extension.
''' Has been disabled by request but still here if anyone wants to use it.
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class SplashScreen

    ''' <summary>
    ''' Create and initialize the splash screen form.
    ''' </summary>
    ''' <remarks></remarks>
    Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        Dim b As Bitmap = New Bitmap(Me.BackgroundImage)
        b.MakeTransparent(b.GetPixel(1, 1))
        Me.BackgroundImage = b

        'Version.Text = System.String.Format(Version.Text, My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.Build, My.Application.Info.Version.Revision)
        Version.Text = System.String.Format(Version.Text, My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.Build)
    End Sub

    ''' <summary>
    ''' Event handler for the click on the form. Closes the form.
    ''' </summary>
    ''' <param name="sender">Not used.</param>
    ''' <param name="e">Not used.</param>
    ''' <remarks></remarks>
    Private Sub SplashScreen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Click, Version.Click
        Me.Close()
    End Sub

    Private Sub SplashScreen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
