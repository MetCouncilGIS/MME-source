''' <summary>
''' Class that provides a spacious entry form, i.e. a form that has more space for entering text
''' for the particular data element it is associated with.
''' </summary>
''' <remarks>Only one form is allowed to be open currently. Maybe we should relax this constraint 
''' to allow easier copy/paste, viewing/comparing content, etc.?</remarks>
Public Class SpaciousEntryForm
    ''' <summary>
    ''' The name of the input element for which the form is opened.
    ''' </summary>
    ''' <remarks></remarks>
    Private elt As String
    ''' <summary>
    ''' Original text that will be made available for editing on the form.
    ''' </summary>
    ''' <remarks></remarks>
    Private origText As String
    ''' <summary>
    ''' Flag to indicate if the element has been saved after last edit on the form.
    ''' </summary>
    ''' <remarks></remarks>
    Private saved As Boolean

    ''' <summary>
    ''' The name of the control on main EME form that this form is associated with.
    ''' </summary>
    ''' <value>A control's name string.</value>
    ''' <remarks></remarks>
    Public WriteOnly Property element()
        Set(ByVal value)
            elt = value
        End Set
    End Property

    ''' <summary>
    ''' The name to display for the element being edited on this form.
    ''' </summary>
    ''' <value>User friendly name string.</value>
    ''' <remarks></remarks>
    Public WriteOnly Property elementDisplayName() As String
        Set(ByVal value As String)
            Me.Text = IIf(value.EndsWith(":"), Microsoft.VisualBasic.Left(value, value.Length - 1), value)
            Me.llHelp.Text = Me.Text
        End Set
    End Property

    ''' <summary>
    ''' The content being edited on this form.
    ''' </summary>
    ''' <value>A string containing the text content.</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property content()
        Get
            If saved Then
                Return Me.txtText.Text
                'Return Me.txtText.Text.Replace(vbLf, vbCrLf)
            Else
                Return origText
            End If
        End Get
        Set(ByVal value)
            origText = value
            Me.txtText.Lines = value.Split(New String() {vbCrLf}, Nothing)
            'Me.txtText.Text = value
            Me.txtText.SelectedText = Nothing
        End Set
    End Property

    ''' <summary>
    ''' The background color for the form depending on optionality of the element being edited.
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property optionalityColor()
        Set(ByVal value)
            Me.panelMain.BackColor = value
        End Set
    End Property

    ''' <summary>
    ''' Indicates that the content has changed.
    ''' </summary>
    ''' <returns>True if the content is different from the original text when the form was opened or last saved.</returns>
    ''' <remarks></remarks>
    Private Function changed() As Boolean
        Return Me.content <> origText
    End Function

    ''' <summary>
    ''' Event handler for the "Close and save" button.
    ''' </summary>
    ''' <param name="sender">Not used.</param>
    ''' <param name="e">Not used.</param>
    ''' <remarks></remarks>
    Private Sub btnCloseSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCloseSave.Click
        'If MessageBox.Show("Do you really want to save and close?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
        saved = True
        Me.Close()
        'End If
    End Sub

    ''' <summary>
    ''' Event handler for the "Save" button. Updated origText field so changes aren't discarded.
    ''' </summary>
    ''' <param name="sender">Not used.</param>
    ''' <param name="e">Not used.</param>
    ''' <remarks></remarks>
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        saved = True
        origText = Me.content
    End Sub

    ''' <summary>
    ''' Event handler for the "Close and Discard" button. Prompts user to confirm.
    ''' </summary>
    ''' <param name="sender">Not used.</param>
    ''' <param name="e">Not used.</param>
    ''' <remarks></remarks>
    Private Sub btnCloseDiscard_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCloseDiscard.Click
        'If MessageBox.Show("Do you really want to cancel?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
        saved = False
        Me.Close()
        'End If
    End Sub


    ''' <summary>
    ''' Event handler for the help link. Opens help for the element being edited.
    ''' </summary>
    ''' <param name="sender">Not used.</param>
    ''' <param name="e">Not used.</param>
    ''' <remarks></remarks>
    Private Sub llHelp_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llHelp.LinkClicked
        EditorForm.HelpSeeker(Me.elt)
    End Sub

    ''' <summary>
    ''' Focus on the input area as soon as the form is loaded.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub SpaciousEntryForm_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Me.txtText.Focus()
    End Sub

    ''' <summary>
    ''' Hook into a win32 function
    ''' </summary>
    ''' <param name="hwnd"></param>
    ''' <param name="wMsg"></param>
    ''' <param name="wParam"></param>
    ''' <param name="lParam"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Declare Auto Function SendMessage Lib "user32" (ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr

    ''' <summary>
    ''' Override the Pressed Key Processing Routine of the MDI-Parent primarily to be able to pass down ctrl-x/c/v keypresses
    ''' </summary>
    ''' <param name="msg"></param>
    ''' <param name="keyData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message, ByVal keyData As System.Windows.Forms.Keys) As Boolean
        If Me.ActiveMdiChild IsNot Nothing Then SendMessage(Me.ActiveMdiChild.Handle, msg.Msg, msg.WParam, msg.LParam)
    End Function
End Class