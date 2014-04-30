Public Class TempForm

    Private Sub TempForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        MasterController.frm = New EditorForm
        MasterController.frm.Show(Me)
    End Sub

End Class