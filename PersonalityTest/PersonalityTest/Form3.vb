Public Class frmQuestionsOne
    Public ex, i, s, n, t, f, j, p As Integer

    Private Sub btnPersonalityTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPersonalityTest.Click
        Me.Close()
        frmTest.Show()

    End Sub

    Private Sub btnLogo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogo.Click
        Me.Close()
        frmMain.Show()

    End Sub

    Private Sub frmQuestionsOne_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ex = 0
        i = 0
        s = 0
        n = 0
        t = 0
        f = 0
        j = 0
        p = 0
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        
        If q1_plus1.Checked = True Then
            ex += 1
        ElseIf q1_plus2.Checked = True Then
            ex += 2
        ElseIf q1_plus3.Checked = True Then
            ex += 3
        ElseIf q1_minus1.Checked = True Then
            i += 1
        ElseIf q1_minus2.Checked = True Then
            i += 2
            
        ElseIf q1_minus3.Checked = True Then
            i += 3
           
        ElseIf q1_zero.Checked = True Then
            i = i
            ex = ex
           
        Else
            MsgBox("Моля, отговорете на всеки от въпросите!", vbOKOnly + vbExclamation, "Не сте отговорили на всички въпроси!")
            Exit Sub
        End If
        If q2_plus1.Checked = True Then
            t += 1
        ElseIf q2_plus2.Checked = True Then
            t += 2
        ElseIf q2_plus3.Checked = True Then
            t += 3
        ElseIf q2_minus1.Checked = True Then
            f += 1
        ElseIf q2_minus2.Checked = True Then
            f += 2
        ElseIf q2_minus3.Checked = True Then
            f += 3
        ElseIf q2_zero.Checked = True Then
            t = t
            f = f
        Else
            MsgBox("Моля, отговорете на всеки от въпросите!", vbOKOnly + vbExclamation, "Не сте отговорили на всички въпроси!")
            Exit Sub
        End If
        If q3_plus1.Checked = True Then
            j += 1
        ElseIf q3_plus2.Checked = True Then
            j += 2
        ElseIf q3_plus3.Checked = True Then
            j += 3
        ElseIf q3_minus1.Checked = True Then
            p += 1
        ElseIf q3_minus2.Checked = True Then
            p += 2
        ElseIf q3_minus3.Checked = True Then
            p += 3
        ElseIf q3_zero.Checked = True Then
            p = p
            j = j
        Else
            MsgBox("Моля, отговорете на всеки от въпросите!", vbOKOnly + vbExclamation, "Не сте отговорили на всички въпроси!")
            Exit Sub
        End If
        If q4_plus1.Checked = True Then
            i += 1
        ElseIf q4_plus2.Checked = True Then
            i += 2
        ElseIf q4_plus3.Checked = True Then
            i += 3
        ElseIf q4_minus1.Checked = True Then
            ex += 1
        ElseIf q4_minus2.Checked = True Then
            ex += 2
        ElseIf q4_minus3.Checked = True Then
            ex += 3
        ElseIf q4_zero.Checked = True Then
            i = i
            ex = ex
        Else
            MsgBox("Моля, отговорете на всеки от въпросите!", vbOKOnly + vbExclamation, "Не сте отговорили на всички въпроси!")
            Exit Sub
        End If
        If q5_plus1.Checked = True Then
            j += 1
        ElseIf q5_plus2.Checked = True Then
            j += 2
        ElseIf q5_plus3.Checked = True Then
            j += 3
        ElseIf q5_minus1.Checked = True Then
            p += 1
        ElseIf q5_minus2.Checked = True Then
            p += 2
        ElseIf q5_minus3.Checked = True Then
            p += 3
        ElseIf q5_zero.Checked = True Then
            p = p
            j = j
        Else
            MsgBox("Моля, отговорете на всеки от въпросите!", vbOKOnly + vbExclamation, "Не сте отговорили на всички въпроси!")
            Exit Sub
        End If
        If q6_plus1.Checked = True Then
            i += 1
        ElseIf q6_plus2.Checked = True Then
            i += 2
        ElseIf q6_plus3.Checked = True Then
            i += 3
        ElseIf q6_minus1.Checked = True Then
            ex += 1
        ElseIf q6_minus2.Checked = True Then
            ex += 2
        ElseIf q6_minus3.Checked = True Then
            ex += 3
        ElseIf q6_zero.Checked = True Then
            i = i
            ex = ex
        Else
            MsgBox("Моля, отговорете на всеки от въпросите!", vbOKOnly + vbExclamation, "Не сте отговорили на всички въпроси!")
            Exit Sub
        End If
        frmQuestionsTwo.ex = ex
        frmQuestionsTwo.i = i
        frmQuestionsTwo.s = s
        frmQuestionsTwo.n = n
        frmQuestionsTwo.t = t
        frmQuestionsTwo.f = f
        frmQuestionsTwo.j = j
        frmQuestionsTwo.p = p
        Me.Close()
        frmQuestionsTwo.Show()
    End Sub

    Private Sub btnAuthorsRights_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAuthorsRights.Click
        Me.Close()
        frmAuthorsRights.Show()
    End Sub
End Class