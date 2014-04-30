Imports System.Windows.Forms

''' <summary>
''' Class that provides a progress dialog to display during validation.
''' </summary>
''' <remarks></remarks>
Public Class ProgressDialog
    Private Shared pd As ProgressDialog

    ''' <summary>
    ''' Display a progress dialog with provided message.
    ''' </summary>
    ''' <param name="msg">Message to display on the progress dialog.</param>
    ''' <remarks>If a progress dialog is already being displayed, it is replaced with a new one</remarks>
    Public Shared Sub ShowMessage(ByVal msg As String)
        CancelMessage()
        If pd Is Nothing OrElse pd.IsDisposed Then
            pd = New ProgressDialog
        End If

        pd.Text = GlobalVars.nameStr & " - " & msg
        pd.TextBox1.Text = msg
        pd.TextBox1.SelectionLength = 0
        pd.Show(EditorForm.ActiveForm)
    End Sub

    ''' <summary>
    ''' Close the currently displayed progress dialog if there is one.
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub CancelMessage()
        If pd IsNot Nothing Then
            pd.Close()
        End If
    End Sub
End Class
