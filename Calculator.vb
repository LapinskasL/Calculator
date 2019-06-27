

Public Class frmCalculator

    'INCOMPLETE PROJECT

    '@@@ EXTREMELY IMPORTANT TODO: Debug all the calcualtion buttons before moving onto DISTANT TODO.
    'TODO: simplify code, debug buttons... And probably more




    'DISTANT TODO: add history, add light/dark modes switch


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




    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txtWorkspace.Text = 0
        lblWorkspaceHold.Text = ""
        txtWorkspace.Tag = "resetFalse"
        signChanged = False
    End Sub



    Private Sub frmCalculator_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        If e.KeyCode = Keys.Add OrElse My.Computer.Keyboard.ShiftKeyDown AndAlso e.KeyCode = 187 Then
            btnPlus_Click(Nothing, Nothing)
        ElseIf e.KeyCode = Keys.Subtract OrElse e.KeyCode = 189 Then
            btnMinus_Click(Nothing, Nothing)
        ElseIf e.KeyCode = Keys.Multiply OrElse My.Computer.Keyboard.ShiftKeyDown AndAlso e.KeyCode = Keys.D8 Then
            btnTimes_Click(Nothing, Nothing)
        ElseIf e.KeyCode = Keys.Divide OrElse e.KeyCode = 191 Then
            btnDivide_Click(Nothing, Nothing)
        ElseIf e.KeyCode = Keys.Return OrElse e.KeyCode = 187 Then
            btnEquals_Click(Nothing, Nothing)
        ElseIf e.KeyCode = Keys.NumPad0 OrElse e.KeyCode = Keys.D0 Then
            EnterNumber(0)
        ElseIf e.KeyCode = Keys.NumPad1 OrElse e.KeyCode = Keys.D1 Then
            EnterNumber(1)
        ElseIf e.KeyCode = Keys.NumPad2 OrElse e.KeyCode = Keys.D2 Then
            EnterNumber(2)
        ElseIf e.KeyCode = Keys.NumPad3 OrElse e.KeyCode = Keys.D3 Then
            EnterNumber(3)
        ElseIf e.KeyCode = Keys.NumPad4 OrElse e.KeyCode = Keys.D4 Then
            EnterNumber(4)
        ElseIf e.KeyCode = Keys.NumPad5 OrElse e.KeyCode = Keys.D5 Then
            EnterNumber(5)
        ElseIf e.KeyCode = Keys.NumPad6 OrElse e.KeyCode = Keys.D6 Then
            EnterNumber(6)
        ElseIf e.KeyCode = Keys.NumPad7 OrElse e.KeyCode = Keys.D7 Then
            EnterNumber(7)
        ElseIf e.KeyCode = Keys.NumPad8 OrElse e.KeyCode = Keys.D8 Then
            EnterNumber(8)
        ElseIf e.KeyCode = Keys.NumPad9 OrElse e.KeyCode = Keys.D9 Then
            EnterNumber(9)
        ElseIf e.KeyCode = 190 OrElse e.KeyCode = 46 Then
            btnDecimal_Click(Nothing, Nothing)
        ElseIf e.KeyCode = Keys.Back Then
            btnBackspace_Click(Nothing, Nothing)
        Else

        End If

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
        AddOrSub("+")
    End Sub

    Private Sub btnMinus_Click(sender As Object, e As EventArgs) Handles btnMinus.Click
        AddOrSub("-")
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

    Private Sub btnEquals_Click(sender As Object, e As EventArgs) Handles btnEquals.Click
        Dim partOne As String = lblWorkspaceHold.Text
        Dim partTwo As String = txtWorkspace.Text
        Dim bothParts As String = partOne & partTwo

        If bothParts.Contains("x") Then
            bothParts = bothParts.Replace("x", "*")
        End If

        bothParts = "1.0 *" & bothParts

        Dim result As Decimal = 0
        Try
            result = New DataTable().Compute(bothParts & "*1.0", Nothing)
            txtWorkspace.Text = CleanUpNumber(result)
            txtWorkspace.Tag = "resetTrue"
            lblWorkspaceHold.Text = ""
            Exit Try
        Catch ex As DivideByZeroException
            txtWorkspace.Text = "Cannot divide by 0"
            txtWorkspace.Tag = "resetTrue"
            lblWorkspaceHold.Text = ""

        Catch ex As OverflowException
            txtWorkspace.Text = "‭Overflow‬"
        End Try

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
        EnteredButton.ForeColor = Color.FromArgb(31, 31, 31) 'orange
    End Sub

    Private Sub Column4Buttons_MouseLeave(sender As Object, e As EventArgs) Handles btnTimes.MouseLeave, btnPlus.MouseLeave,
                                            btnMinus.MouseLeave, btnEquals.MouseLeave, btnDivide.MouseLeave
        Dim LeftButton As Button

        LeftButton = CType(sender, Button)
        LeftButton.BackColor = Color.FromArgb(19, 19, 19)
        LeftButton.ForeColor = Color.FromArgb(255, 255, 255) 'orange
    End Sub

    Private Sub NonDigitButtons_MouseLeave(sender As Object, e As EventArgs) Handles btnBackspace.MouseLeave,
                                            btnPlusMinus.MouseLeave, btnMinus.MouseLeave, btnDecimal.MouseLeave,
                                            btnClearEntry.MouseLeave, btnClear.MouseLeave
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
        If txtWorkspace.Text <> "0" AndAlso txtWorkspace.Text <> "" Then
            If txtWorkspace.Text.Length = 2 AndAlso txtWorkspace.Text.Contains("-"c) Then
                txtWorkspace.Text = "0"
            ElseIf txtWorkspace.Text.Length = 1 Then
                txtWorkspace.Text = "0"
            Else
                txtWorkspace.Text = txtWorkspace.Text.Substring(0, txtWorkspace.Text.Length - 1)
            End If
        End If
    End Sub

    Private Sub btnPlusMinus_Click(sender As Object, e As EventArgs) Handles btnPlusMinus.Click
        If txtWorkspace.Text <> "0" AndAlso txtWorkspace.Text <> "" Then
            If txtWorkspace.Text.First <> "-" Then
                txtWorkspace.Text = "-" & txtWorkspace.Text
            Else
                txtWorkspace.Text = txtWorkspace.Text.Substring(1, txtWorkspace.Text.Length - 1)
            End If
        End If
    End Sub



    Private Function CleanUpNumber(numberString As String)
        Dim cleanNumber As String = numberString

        If numberString.Last = "."c Then
            cleanNumber = numberString.Substring(0, numberString.Length - 1)
        ElseIf numberString.Contains("."c) Then
            While numberString.Last = "0"c
                numberString = numberString.Substring(0, numberString.Length - 1)
            End While
            cleanNumber = numberString


            Dim isNotZero As Boolean = False

            Dim index = numberString.IndexOf("."c) + 1
            While index < numberString.Length AndAlso isNotZero = False
                If numberString(index) <> "0"c Then
                    isNotZero = True
                End If
                index += 1
            End While

            If isNotZero = False Then
                cleanNumber = ""
                For index = 0 To numberString.IndexOf("."c) - 1
                    cleanNumber += numberString(index)
                Next
            End If

        End If

        If cleanNumber.Length > 16 Then
            cleanNumber = Format(cleanNumber, "Scientific")
        End If

        Return cleanNumber
    End Function


    Private Sub EnterNumber(number As Decimal)
        If txtWorkspace.Text = "0" AndAlso Not txtWorkspace.Text = "0." Then
            txtWorkspace.Text = number
            txtWorkspace.Tag = "resetFalse"
        ElseIf txtWorkspace.Tag = "resetTrue" AndAlso Not txtWorkspace.Text = "0." Then
            txtWorkspace.Text = number
            txtWorkspace.Tag = "resetFalse"
        ElseIf lblWorkspaceHold.Text <> "" AndAlso txtWorkspace.Tag = "resetFalse" Then
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
                lblWorkspaceHold.Text = lblWorkspaceHold.Text.Substring(0, lblWorkspaceHold.Text.Length - 3) & " x "
            Else
                lblWorkspaceHold.Text = lblWorkspaceHold.Text.Substring(0, lblWorkspaceHold.Text.Length - 3) & " " & sign & " "
            End If
        Else
            If sign = "*" Then
                lblWorkspaceHold.Text += CleanUpNumber(txtWorkspace.Text) & " x "
            Else
                lblWorkspaceHold.Text += CleanUpNumber(txtWorkspace.Text) & " " & sign & " "
            End If

            Dim bothParts As String = lblWorkspaceHold.Text & CleanUpNumber(txtWorkspace.Text)

            Dim result As String = ""



            Try
                result = New DataTable().Compute(lblWorkspaceHold.Text.Replace("x", "*") & "* 1.0", Nothing)
                txtWorkspace.Text = CleanUpNumber(result)
                Exit Try
            Catch ex As SyntaxErrorException
            End Try

            txtWorkspace.Tag = "resetTrue"

            signChanged = True
        End If


    End Sub


    Private Sub AddOrSub(sign As String)
        If signChanged = True Then
            lblWorkspaceHold.Text = lblWorkspaceHold.Text.Substring(0, lblWorkspaceHold.Text.Length - 3) & " " & sign & " "
        Else
            lblWorkspaceHold.Text += CleanUpNumber(txtWorkspace.Text) & " " & sign & " "


            Dim decWorkSpaceHold As String = "1.0 * " & lblWorkspaceHold.Text

            Dim result = New DataTable().Compute(decWorkSpaceHold.Replace("x", "*") & "0", Nothing)

            txtWorkspace.Text = CleanUpNumber(result)

            txtWorkspace.Tag = "resetTrue"

            signChanged = True
        End If
    End Sub

    Private Sub btnDecimal_Click(sender As Object, e As EventArgs) Handles btnDecimal.Click
        If txtWorkspace.Text.Contains(".") = False AndAlso txtWorkspace.Tag = "resetFalse" Then
            txtWorkspace.Text = txtWorkspace.Text & "."
        ElseIf txtWorkspace.Tag = "resetTrue" Then
            txtWorkspace.Text = "0."
            txtWorkspace.Tag = "resetFalse"
        End If
    End Sub

End Class
