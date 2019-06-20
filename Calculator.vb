Public Class frmCalculator

    'INCOMPLETE PROJECT

    'TODO: change font, simplify code, add more buttons, debug buttons, fix colors... and probably more

    'DISTANT TODO: add history, add light/dark modes

    Dim signChanged As Boolean = False

    Private Sub lblExit_MouseLeave(sender As Object, e As EventArgs) Handles lblExit.MouseLeave
        lblExit.BackColor = Color.Gray
    End Sub

    Private Sub lblExit_Click(sender As Object, e As EventArgs) Handles lblExit.Click
        Me.Close()
    End Sub

    Private Sub lblExit_MouseEnter(sender As Object, e As EventArgs) Handles lblExit.MouseEnter
        lblExit.BackColor = Color.Red
    End Sub

    Private Sub EnterNumber(number As Double)
        If lblWorkspace.Text = 0 Then
            lblWorkspace.Text = number
            lblWorkspace.Tag = "resetFalse"
        ElseIf lblWorkspace.Tag = "resetTrue" Then
            lblWorkspace.Text = number
            lblWorkspace.Tag = "resetFalse"
        ElseIf lblWorkspaceHold.Text <> "" And lblWorkspace.Tag = "resetFalse" Then
            lblWorkspace.Text = lblWorkspace.Text & number
        ElseIf lblWorkspace.Tag = "resetTrue" Then
            lblWorkspace.Text = lblWorkspace.Text & number
            lblWorkspace.Tag = "resetFalse"
        Else
            lblWorkspace.Text = lblWorkspace.Text & number
        End If

        signChanged = False
    End Sub



    Private Sub MulOrDiv(sign)
        If signChanged = True Then
            If sign = "*" Then
                lblWorkspaceHold.Text = lblWorkspaceHold.Text.Substring(0, lblWorkspaceHold.Text.Length - 1) & "x"
            Else
                lblWorkspaceHold.Text = lblWorkspaceHold.Text.Substring(0, lblWorkspaceHold.Text.Length - 1) & sign
            End If
        Else
            If sign = "*" Then
                lblWorkspaceHold.Text += lblWorkspace.Text & "x"
            Else
                lblWorkspaceHold.Text += lblWorkspace.Text & sign
            End If

            Dim bothParts As String = lblWorkspaceHold.Text & lblWorkspace.Text

            Dim result As Double = New DataTable().Compute(bothParts.Replace("x", "*"), Nothing)

            lblWorkspace.Tag = "resetTrue"

            signChanged = True
        End If


    End Sub

    Private Function Calculate(expression As String)

        Dim hello As String = ""
        Dim answer As Double = 0

        For index As Integer = 0 To expression.Length - 1
            hello = expression(index)
        Next

    End Function


    Private Sub CalculateWith(sign As String)
        If signChanged = True Then
            lblWorkspaceHold.Text = lblWorkspaceHold.Text.Substring(0, lblWorkspaceHold.Text.Length - 1) & sign
        Else
            lblWorkspaceHold.Text += lblWorkspace.Text & sign

            Dim result = New DataTable().Compute(lblWorkspaceHold.Text.Replace("x", "*") & "0", Nothing)

            lblWorkspace.Text = result

            lblWorkspace.Tag = "resetTrue"

            signChanged = True
        End If
    End Sub


    Private Sub btnEquals_Click(sender As Object, e As EventArgs) Handles btnEquals.Click
        Dim partOne As String = lblWorkspaceHold.Text
        Dim partTwo As String = lblWorkspace.Text
        Dim bothParts As String = partOne & partTwo

        If bothParts.Contains("x") Then
            bothParts = bothParts.Replace("x", "*")
        End If

        Dim result As Double = New DataTable().Compute(bothParts, Nothing)
        lblWorkspace.Tag = "resetTrue"

        lblWorkspace.Text = result
        lblWorkspaceHold.Text = ""
    End Sub


    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        lblWorkspace.Text = 0
        lblWorkspaceHold.Text = ""
        lblWorkspace.Tag = "resetFalse"
    End Sub








    Private Sub btnOne_Click(sender As Object, e As EventArgs) Handles btnOne.Click
        EnterNumber(1)
    End Sub

    Private Sub btnTwo_Click(sender As Object, e As EventArgs) Handles btnTwo.Click
        EnterNumber(2)
    End Sub

    Private Sub btnThree_Click(sender As Object, e As EventArgs) Handles btnThree.Click
        EnterNumber(3)
    End Sub

    Private Sub btnFour_Click(sender As Object, e As EventArgs) Handles btnFour.Click
        EnterNumber(4)
    End Sub

    Private Sub btnFive_Click(sender As Object, e As EventArgs) Handles btnFive.Click
        EnterNumber(5)
    End Sub

    Private Sub btnSix_Click(sender As Object, e As EventArgs) Handles btnSix.Click
        EnterNumber(6)
    End Sub

    Private Sub btnSeven_Click(sender As Object, e As EventArgs) Handles btnSeven.Click
        EnterNumber(7)
    End Sub

    Private Sub btnEight_Click(sender As Object, e As EventArgs) Handles btnEight.Click
        EnterNumber(8)
    End Sub

    Private Sub btnNine_Click(sender As Object, e As EventArgs) Handles btnNine.Click
        EnterNumber(9)
    End Sub

    Private Sub btnZero_Click(sender As Object, e As EventArgs) Handles btnZero.Click
        EnterNumber(0)
    End Sub

    Private Sub btnPlus_Click(sender As Object, e As EventArgs) Handles btnPlus.Click
        CalculateWith("+")
    End Sub

    Private Sub btnMinus_Click(sender As Object, e As EventArgs) Handles btnMinus.Click
        CalculateWith("-")
    End Sub

    Private Sub btnTimes_Click(sender As Object, e As EventArgs) Handles btnTimes.Click
        MulOrDiv("*")
    End Sub

    Private Sub btnDivide_Click(sender As Object, e As EventArgs) Handles btnDivide.Click
        MulOrDiv("/")
    End Sub
End Class
