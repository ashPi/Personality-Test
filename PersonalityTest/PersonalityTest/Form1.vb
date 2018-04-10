Public Class frmMain

    Private Sub btnStartTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStartTest.Click
        Me.Hide()
        frmTest.Show()
    End Sub

    Private Sub btnPersonalityTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPersonalityTest.Click
        Me.Hide()
        frmTest.Show()
    End Sub

    Private Sub btnLogo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogo.Click
        Me.Refresh()

    End Sub

    Private Sub btnAuthorsRights_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAuthorsRights.Click
        Me.Hide()
        frmAuthorsRights.Show()

    End Sub

    Private Sub btnPersonalityTypes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPersonalityTypes.Click
        Me.Hide()
        frmPersonalityTypes.Show()
    End Sub


    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
