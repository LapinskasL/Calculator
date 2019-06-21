Public Class frmCalculator

    'INCOMPLETE PROJECT

    '@@@ EXTREMELY IMPORTANT TODO: Debug all the calcualtion buttons before moving onto DISTANT TODO.

    'TODO: simplify code, implement decimal button, debug buttons,
    '      fix colors, fix component placement... And probably more




    'DISTANT TODO: add history, add light/dark modes


    Dim signChanged As Boolean = False





    Dim drag As Boolean
    Dim mousex As Integer
    Dim mousey As Integer

    Private Sub lblCalculator_MouseDown(sender As Object, e As MouseEventArgs) Handles lblCalculator.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Left Then
            drag = True
            mousex = Windows.Forms.Cursor.Position.X - Me.Left
            mousey = Windows.Forms.Cursor.Position.Y - Me.Top
        End If
    End Sub

    Private Sub lblCalculator_MouseMove(sender As Object, e As MouseEventArgs) Handles lblCalculator.MouseMove
        If drag Then
            Me.Top = Windows.Forms.Cursor.Position.Y - mousey
            Me.Left = Windows.Forms.Cursor.Position.X - mousex
        End If
    End Sub

    Private Sub lblCalculator_MouseUp(sender As Object, e As MouseEventArgs) Handles lblCalculator.MouseUp
        drag = False
    End Sub





    Private Sub btnEquals_Click(sender As Object, e As EventArgs) Handles btnEquals.Click
        Dim partOne As String = lblWorkspaceHold.Text
        Dim partTwo As String = txtWorkspace.Text
        Dim bothParts As String = partOne & partTwo

        If bothParts.Contains("x") Then
            bothParts = bothParts.Replace("x", "*")
        End If

        Dim result As Double = New DataTable().Compute(bothParts, Nothing)
        txtWorkspace.Tag = "resetTrue"

        txtWorkspace.Text = result
        lblWorkspaceHold.Text = ""
    End Sub


    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txtWorkspace.Text = 0
        lblWorkspaceHold.Text = ""
        txtWorkspace.Tag = "resetFalse"
    End Sub






    Private Sub btnZero_Click(sender As Object, e As EventArgs) Handles btnZero.Click
        EnterNumber(0)
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

    Private Sub btnClearEntry_Click(sender As Object, e As EventArgs) Handles btnClearEntry.Click
        txtWorkspace.Text = "0"
        txtWorkspace.Tag = "resetTrue"
        signChanged = False
    End Sub







    Private Sub First3ColumnsButtons_MouseEnter(sender As Object, e As EventArgs) Handles btnBackspace.MouseEnter, btnZero.MouseEnter,
                                            btnTwo.MouseEnter, btnThree.MouseEnter, btnSix.MouseEnter, btnSeven.MouseEnter,
                                            btnPlusMinus.MouseEnter, btnOne.MouseEnter, btnNine.MouseEnter, btnFour.MouseEnter,
                                            btnFive.MouseEnter, btnEight.MouseEnter, btnDecimal.MouseEnter, btnClearEntry.MouseEnter,
                                            btnClear.MouseEnter
        Dim EnteredButton As Button

        EnteredButton = CType(sender, Button)
        EnteredButton.BackColor = Color.FromArgb(60, 60, 60)

    End Sub

    Private Sub Column4Buttons_MouseEnter(sender As Object, e As EventArgs) Handles btnTimes.MouseEnter, btnPlus.MouseEnter,
                                            btnMinus.MouseEnter, btnEquals.MouseEnter, btnDivide.MouseEnter
        Dim EnteredButton As Button

        EnteredButton = CType(sender, Button)
        EnteredButton.BackColor = Color.FromArgb(255, 144, 0) 'orange
    End Sub

    Private Sub NonDigitButtons_MouseLeave(sender As Object, e As EventArgs) Handles btnBackspace.MouseLeave, btnTimes.MouseLeave,
                                            btnPlusMinus.MouseLeave, btnPlus.MouseLeave, btnMinus.MouseLeave, btnEquals.MouseLeave,
                                            btnDivide.MouseLeave, btnDecimal.MouseLeave, btnClearEntry.MouseLeave, btnClear.MouseLeave
        Dim LeftButton As Button

        LeftButton = CType(sender, Button)
        LeftButton.BackColor = Color.FromArgb(19, 19, 19)
    End Sub

    Private Sub lblMinimize_MouseEnter(sender As Object, e As EventArgs) Handles lblMinimize.MouseEnter
        lblMinimize.BackColor = Color.FromArgb(60, 60, 60)
    End Sub

    Private Sub MinimizeExitLabels_MouseLeave(sender As Object, e As EventArgs) Handles lblMinimize.MouseLeave, lblExit.MouseLeave
        Dim LabelExited As Label

        LabelExited = CType(sender, Label)
        LabelExited.BackColor = Color.FromArgb(31, 31, 31)
    End Sub

    Private Sub DigitButtons_MouseLeave(sender As Object, e As EventArgs) Handles btnZero.MouseLeave, btnOne.MouseLeave, btnTwo.MouseLeave,
                                               btnThree.MouseLeave, btnSix.MouseLeave, btnSeven.MouseLeave, btnNine.MouseLeave,
                                               btnFour.MouseLeave, btnFive.MouseLeave, btnEight.MouseLeave
        Dim LeftDigitButton As Button

        LeftDigitButton = CType(sender, Button)
        LeftDigitButton.BackColor = Color.FromArgb(6, 6, 6)
    End Sub




    Private Sub lblMinimize_Click(sender As Object, e As EventArgs) Handles lblMinimize.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub lblExit_Click(sender As Object, e As EventArgs) Handles lblExit.Click
        Me.Close()
    End Sub

    Private Sub lblExit_MouseEnter(sender As Object, e As EventArgs) Handles lblExit.MouseEnter
        lblExit.BackColor = Color.Red
    End Sub




    Private Sub btnBackspace_Click(sender As Object, e As EventArgs) Handles btnBackspace.Click
        If txtWorkspace.Text <> "0" And txtWorkspace.Text <> "" Then
            If txtWorkspace.Text.Length = 2 And txtWorkspace.Text.Contains("-"c) Then
                txtWorkspace.Text = "0"
            ElseIf txtWorkspace.Text.Length = 1 Then
                txtWorkspace.Text = "0"
            Else
                txtWorkspace.Text = txtWorkspace.Text.Substring(0, txtWorkspace.Text.Length - 1)
            End If
        End If
    End Sub

    Private Sub btnPlusMinus_Click(sender As Object, e As EventArgs) Handles btnPlusMinus.Click
        If txtWorkspace.Text <> "0" And txtWorkspace.Text <> "" Then
            If txtWorkspace.Text.First <> "-" Then
                txtWorkspace.Text = "-" & txtWorkspace.Text
            Else
                txtWorkspace.Text = txtWorkspace.Text.Substring(1, txtWorkspace.Text.Length - 1)
            End If
        End If
    End Sub








    Private Sub EnterNumber(number As Double)
        If txtWorkspace.Text = 0 Then
            txtWorkspace.Text = number
            txtWorkspace.Tag = "resetFalse"
        ElseIf txtWorkspace.Tag = "resetTrue" Then
            txtWorkspace.Text = number
            txtWorkspace.Tag = "resetFalse"
        ElseIf lblWorkspaceHold.Text <> "" And txtWorkspace.Tag = "resetFalse" Then
            txtWorkspace.Text = txtWorkspace.Text & number
        ElseIf txtWorkspace.Tag = "resetTrue" Then
            txtWorkspace.Text = txtWorkspace.Text & number
            txtWorkspace.Tag = "resetFalse"
        Else
            txtWorkspace.Text = txtWorkspace.Text & number
        End If

        signChanged = False
    End Sub



    Private Sub MulOrDiv(sign)
        If signChanged = True Then
            If sign = "*" Then
                lblWorkspaceHold.Text = lblWorkspaceHold.Text.Substring(0, lblWorkspaceHold.Text.Length - 1) & " x "
            Else
                lblWorkspaceHold.Text = lblWorkspaceHold.Text.Substring(0, lblWorkspaceHold.Text.Length - 1) & " " & sign & " "
            End If
        Else
            If sign = "*" Then
                lblWorkspaceHold.Text += txtWorkspace.Text & " x "
            Else
                lblWorkspaceHold.Text += txtWorkspace.Text & " " & sign & " "
            End If

            Dim bothParts As String = lblWorkspaceHold.Text & txtWorkspace.Text

            Dim result As Double = New DataTable().Compute(bothParts.Replace("x", "*"), Nothing)

            txtWorkspace.Tag = "resetTrue"

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
            lblWorkspaceHold.Text = lblWorkspaceHold.Text.Substring(0, lblWorkspaceHold.Text.Length - 1) & " " & sign & " "
        Else
            lblWorkspaceHold.Text += txtWorkspace.Text & " " & sign & " "

            Dim result = New DataTable().Compute(lblWorkspaceHold.Text.Replace("x", "*") & "0", Nothing)

            txtWorkspace.Text = result

            txtWorkspace.Tag = "resetTrue"

            signChanged = True
        End If
    End Sub
End Class
