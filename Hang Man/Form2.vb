Public Class frmWelcome
    Public strUsername As String
    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub cmdPlay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPlay.Click
        strUsername = txtName.Text
        frmSnowman.Show()
        Me.Close()
    End Sub
End Class