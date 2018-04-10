Public Class frmQuestionsTwo
    Public ex, i, s, n, t, f, j, p As Integer
    Private Sub btnLogo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogo.Click
        Me.Close()
        frmMain.Show()


    End Sub

    Private Sub btnPersonalityTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPersonalityTest.Click
        Me.Close()

        frmTest.Show()
    End Sub

    Private Sub btnAuthorsRights_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAuthorsRights.Click
        Me.Close()
        frmAuthorsRights.Show()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If q1_plus1.Checked = True Then
            j += 1
        ElseIf q1_plus2.Checked = True Then
            j += 2
        ElseIf q1_plus3.Checked = True Then
            j += 3
        ElseIf q1_minus1.Checked = True Then
            p += 1
        ElseIf q1_minus2.Checked = True Then
            p += 2
        ElseIf q1_minus3.Checked = True Then
            p += 3
        ElseIf q1_zero.Checked = True Then
            p = p
            j = j
        Else
            MsgBox("Моля, отговорете на всеки от въпросите!", vbOKOnly + vbExclamation, "Не сте отговорили на всички въпроси!")
            Exit Sub
        End If
        If q2_plus1.Checked = True Then
            f += 1
        ElseIf q2_plus2.Checked = True Then
            f += 2
        ElseIf q2_plus3.Checked = True Then
            f += 3
        ElseIf q2_minus1.Checked = True Then
            t += 1
        ElseIf q2_minus2.Checked = True Then
            t += 2
        ElseIf q2_minus3.Checked = True Then
            t += 3
        ElseIf q2_zero.Checked = True Then
            t = t
            f = f
        Else
            MsgBox("Моля, отговорете на всеки от въпросите!", vbOKOnly + vbExclamation, "Не сте отговорили на всички въпроси!")
            Exit Sub
        End If
        If q3_plus1.Checked = True Then
            n += 1
        ElseIf q3_plus2.Checked = True Then
            n += 2
        ElseIf q3_plus3.Checked = True Then
            n += 3
        ElseIf q3_minus1.Checked = True Then
            s += 1
        ElseIf q3_minus2.Checked = True Then
            s += 2
        ElseIf q3_minus3.Checked = True Then
            s += 3
        ElseIf q3_zero.Checked = True Then
            s = s
            n = n
        Else
            MsgBox("Моля, отговорете на всеки от въпросите!", vbOKOnly + vbExclamation, "Не сте отговорили на всички въпроси!")
            Exit Sub
        End If
        If q4_plus1.Checked = True Then
            j += 1
        ElseIf q4_plus2.Checked = True Then
            j += 2
        ElseIf q4_plus3.Checked = True Then
            j += 3
        ElseIf q4_minus1.Checked = True Then
            p += 1
        ElseIf q4_minus2.Checked = True Then
            p += 2
        ElseIf q4_minus3.Checked = True Then
            p += 3
        ElseIf q4_zero.Checked = True Then
            p = p
            j = j
        Else
            MsgBox("Моля, отговорете на всеки от въпросите!", vbOKOnly + vbExclamation, "Не сте отговорили на всички въпроси!")
            Exit Sub
        End If
        If q5_plus1.Checked = True Then
            i += 1
        ElseIf q5_plus2.Checked = True Then
            i += 2
        ElseIf q5_plus3.Checked = True Then
            i += 3
        ElseIf q5_minus1.Checked = True Then
            ex += 1
        ElseIf q5_minus2.Checked = True Then
            ex += 2
        ElseIf q5_minus3.Checked = True Then
            ex += 3
        ElseIf q5_zero.Checked = True Then
            ex = ex
            i = i
        Else
            MsgBox("Моля, отговорете на всеки от въпросите!", vbOKOnly + vbExclamation, "Не сте отговорили на всички въпроси!")
            Exit Sub
        End If
        If q6_plus1.Checked = True Then
            p += 1
        ElseIf q6_plus2.Checked = True Then
            p += 2
        ElseIf q6_plus3.Checked = True Then
            p += 3
        ElseIf q6_minus1.Checked = True Then
            j += 1
        ElseIf q6_minus2.Checked = True Then
            j += 2
        ElseIf q6_minus3.Checked = True Then
            j += 3
        ElseIf q6_zero.Checked = True Then
            j = j
            p = p
        Else
            MsgBox("Моля, отговорете на всеки от въпросите!", vbOKOnly + vbExclamation, "Не сте отговорили на всички въпроси!")
            Exit Sub
        End If
        frmQuestionsThree.ex = ex
        frmQuestionsThree.i = i
        frmQuestionsThree.s = s
        frmQuestionsThree.n = n
        frmQuestionsThree.t = t
        frmQuestionsThree.f = f
        frmQuestionsThree.j = j
        frmQuestionsThree.p = p
        Me.Close()
        frmQuestionsThree.Show()

    End Sub
End Class