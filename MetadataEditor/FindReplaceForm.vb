''' <summary>
''' A simple class to provide user interface for find/replace functionality.
''' </summary>
''' <remarks></remarks>
Public Class FindReplaceForm

    ''' <summary>
    ''' Carry out find/replace functionality when the user clicks the button for it.
    ''' </summary>
    ''' <param name="sender">Not used</param>
    ''' <param name="e">Not used</param>
    ''' <remarks></remarks>
    Private Sub btnReplaceAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReplaceAll.Click
        If Me.tbFind.Text = "" Then
            MsgBox("You must specify some text to find!")
            Exit Sub
        End If

        Dim cnt As Integer = PageController.findAndReplaceAll(Me.Owner, Me.tbFind.Text, Me.tbReplace.Text)
        If cnt > 0 Then
            MsgBox(CStr(cnt) & " substitutions made.")
        Else
            MsgBox("No occurrences found!")
        End If
        Me.Close()
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