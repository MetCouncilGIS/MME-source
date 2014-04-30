
''' <summary>
''' EME loads up slowly due to lots of parsing processing at start up. This form is displayed meanwhile.
''' It also acts as the parent window controlling the EditorForm.
''' </summary>
''' <remarks></remarks>
Public Class DummyForm

    Private Sub DummyForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.Label1.Text = GlobalVars.nameStr + " is loading..."
            Me.Text = GlobalVars.nameStr
            Me.ShowIcon = True
            Me.Show()
            Application.DoEvents()

            If Not checkAndCopyTemplate() Then
                MessageBox.Show("Problem while initializing the application. Exiting...", My.Settings.CommonEditorAbbreviation, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            MdUtils.loadNext = getFilenameFromCommandLine()
            Do While Not MdUtils.quit
                'We default to quitting because we can't catch form close when user clicks X
                MdUtils.quit = True
                If MdUtils.loadNext Is Nothing Then
                    MdUtils.loadDefaultMd()
                Else
                    MdUtils.loadMdFile(MdUtils.loadNext)
                End If
                Me.Visible = True
                Application.DoEvents()
            Loop
        Catch ex As Exception
            Utils.ErrorHandler(ex)
        Finally
            Me.Close()
        End Try

    End Sub

    ''' <summary>
    ''' Initialize essential files by copying from template if ncessary.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function checkAndCopyTemplate() As Boolean
        Try
            Return Utils.copyDir(Utils.getTemplateFolder(), Utils.getAppDataFolder())
        Catch ex As Exception
            Return False
        End Try
    End Function

End Class