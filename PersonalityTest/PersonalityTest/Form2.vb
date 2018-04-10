Public Class frmTest

    Private Sub btnLogo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogo.Click
        Me.Close()
        frmMain.Show()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
        frmQuestionsOne.Show()

    End Sub

    Private Sub btnPersonalityTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPersonalityTest.Click
        Me.Refresh()

    End Sub

    Private Sub btnAuthorsRights_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAuthorsRights.Click
        Me.Close()
        frmAuthorsRights.Show()
    End Sub
End Class