Public Class frmISTP

    Private Sub btnLogo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogo.Click
        Me.Close()
        frmMain.Show()
    End Sub

    Private Sub btnPersonalityTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPersonalityTest.Click
        Me.Close()
        frmTest.Show()
    End Sub

    Private Sub btnPersonalityTypes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPersonalityTypes.Click
        Me.Close()
        frmPersonalityTypes.Show()
    End Sub

    Private Sub btnAuthorsRights_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAuthorsRights.Click
        Me.Close()
        frmAuthorsRights.Show()

    End Sub
End Class