

Public Class frmCalculator

    'INCOMPLETE PROJECT

    'CURRENT BUGS: 
    'When the result is being calculated if the operation parameter in the Calculate PROC is any - or any +,
    'the result does not get converted to a Decimal value.
    '(COULD NOT REPLICATE)

    '@@@ EXTREMELY IMPORTANT TODO: Debug all the calcualtion buttons before moving onto DISTANT TODO.
    'TODO: debug buttons... And probably more

    'DISTANT TODO: add history, add light/dark modes switch

    Dim signChanged As Boolean = False
    Dim resetWorkspace As Boolean = True
    Dim btnsEnabled As Boolean = True
    Dim errorOccurred As Boolean = False



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




    Private Sub btnClearEntry_Click(sender As Object, e As EventArgs) Handles btnClearEntry.Click
        buttonsEnabled(True)
        txtWorkspace.Text = "0"
        resetWorkspace = True
        signChanged = False
        errorOccurred = False

        btnEquals.Select() 'refocus so the Enter key would work properly (Enter always clicks Equals button)

    End Sub


    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        buttonsEnabled(True)
        txtWorkspace.Text = "0"
        lblWorkspaceHold.Text = ""
        resetWorkspace = False
        signChanged = False
        errorOccurred = False

        btnEquals.Select() 'refocus so the Enter key would work properly (Enter always clicks Equals button)
    End Sub




    Private Sub frmCalculator_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        If btnsEnabled AndAlso (e.KeyCode = Keys.Add OrElse
           My.Computer.Keyboard.ShiftKeyDown AndAlso e.KeyCode = 187) Then
            Calculate("add")
        ElseIf btnsEnabled AndAlso (e.KeyCode = Keys.Subtract OrElse e.KeyCode = 189) Then
            Calculate("sub")
        ElseIf btnsEnabled AndAlso e.KeyCode = Keys.Multiply OrElse
            (My.Computer.Keyboard.ShiftKeyDown AndAlso e.KeyCode = Keys.D8) Then
            Calculate("mul")
        ElseIf btnsEnabled AndAlso (e.KeyCode = Keys.Divide OrElse e.KeyCode = 191) Then
            Calculate("div")
        ElseIf btnsEnabled AndAlso (e.KeyCode = 13 OrElse e.KeyCode = 187) Then
            btnEquals_Click(Nothing, Nothing)
        ElseIf btnsEnabled AndAlso (e.KeyCode = 190 OrElse e.KeyCode = 46) Then
            btnDecimal_Click(Nothing, Nothing)
        ElseIf btnsEnabled AndAlso e.KeyCode = Keys.Back Then
            btnBackspace_Click(Nothing, Nothing)
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
        Else

        End If

        btnEquals.Select() 'refocus so the Enter key would work properly (Enter always clicks Equals button)

    End Sub

    Private Sub NumberButtons_Click(sender As Object, e As EventArgs) Handles btnZero.Click, btnOne.Click, btnTwo.Click,
                                                    btnThree.Click, btnFour.Click, btnFive.Click, btnSix.Click,
                                                    btnSeven.Click, btnEight.Click, btnNine.Click

        Dim button As Button = CType(sender, Button)

        Try
            EnterNumber(button.Text)
        Catch ex As Exception
            MsgBox("ERROR in NumberButtons_Click PROC", MsgBoxStyle.Critical) 'DELETE
        End Try

        btnEquals.Select() 'refocus so the Enter key would work properly (Enter always clicks Equals button)

    End Sub



    Private Sub OperationButtons_Click(sender As Object, e As EventArgs) Handles btnPlus.Click, btnMinus.Click, btnTimes.Click, btnDivide.Click

        Dim button As Button = CType(sender, Button)

        Select Case button.Text
            Case "+", "-", "/"
                Calculate(button.Text)
            Case "x"
                Calculate("*")
            Case Else
                MsgBox("ERROR in OperationButtons_Click PROC: Could not identify button case.", MsgBoxStyle.Critical)
        End Select

        btnEquals.Select() 'refocus so the Enter key would work properly (Enter always clicks Equals button)
    End Sub


    Private Sub btnEquals_Click(sender As Object, e As EventArgs) Handles btnEquals.Click
        Dim partOne As String = lblWorkspaceHold.Text
        Dim partTwo As String = txtWorkspace.Text
        Dim bothParts As String = partOne & partTwo

        If bothParts.Contains("x") Then
            bothParts = bothParts.Replace("x", "*")
        End If

        bothParts = "1.0 * " & bothParts

        Dim result As Decimal = 0
        Try
            result = New DataTable().Compute(bothParts & " * 1.0", Nothing)
            txtWorkspace.Text = CleanUpNumber(result)
            resetWorkspace = True
            lblWorkspaceHold.Text = ""
            Exit Try
        Catch ex As DivideByZeroException
            txtWorkspace.Text = "Cannot divide by 0"
            ErrorCaught()
        Catch ex As OverflowException
            txtWorkspace.Text = "‭Overflow‬"
            ErrorCaught()
        End Try

        btnEquals.Select() 'refocus so the Enter key would work properly (Enter always clicks Equals button)
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

    Private Sub lblExit_MouseEnter(sender As Object, e As EventArgs) Handles lblExit.MouseEnter
        lblExit.BackColor = Color.Red
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







    Private Sub btnBackspace_Click(sender As Object, e As EventArgs) Handles btnBackspace.Click
        If CleanUpNumber(txtWorkspace.Text) <> "0" AndAlso txtWorkspace.Text <> "" Then
            If txtWorkspace.Text.Length = 2 AndAlso txtWorkspace.Text.Contains("-"c) Then
                txtWorkspace.Text = "0"
            ElseIf txtWorkspace.Text.Length = 1 Then
                txtWorkspace.Text = "0"
            Else
                txtWorkspace.Text = txtWorkspace.Text.Substring(0, txtWorkspace.Text.Length - 1)
            End If
        End If

        btnEquals.Select() 'refocus so the Enter key would work properly (Enter always clicks Equals button)
    End Sub

    Private Sub btnPlusMinus_Click(sender As Object, e As EventArgs) Handles btnPlusMinus.Click
        If txtWorkspace.Text <> "0" AndAlso txtWorkspace.Text <> "" Then
            If txtWorkspace.Text.First <> "-" Then
                txtWorkspace.Text = "-" & txtWorkspace.Text
            Else
                txtWorkspace.Text = txtWorkspace.Text.Substring(1, txtWorkspace.Text.Length - 1)
            End If
        End If

        btnEquals.Select() 'refocus so the Enter key would work properly (Enter always clicks Equals button)
    End Sub


    Private Sub btnDecimal_Click(sender As Object, e As EventArgs) Handles btnDecimal.Click
        If txtWorkspace.Text.Contains(".") = False AndAlso resetWorkspace = False Then
            txtWorkspace.Text = txtWorkspace.Text & "."
        ElseIf resetWorkspace = True Then
            txtWorkspace.Text = "0."
            resetWorkspace = False
        End If

        btnEquals.Select() 'refocus so the Enter key would work properly (Enter always clicks Equals button)
    End Sub







    ''' <summary>
    ''' Runs necessary instructions if an error is caught.
    ''' </summary>
    Private Sub ErrorCaught()
        resetWorkspace = True
        lblWorkspaceHold.Text = ""
        buttonsEnabled(False)
        errorOccurred = True
    End Sub



    ''' <param name="numberString">A string containing a properly formatted number.</param>
    ''' <summary>
    ''' Removes redundant zeros and decimal numbers (ex: 1.0200 would become 1.02).
    ''' </summary>
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
            cleanNumber = Format(cleanNumber, "scientific")
        End If

        Return cleanNumber

    End Function



    ''' <param name="number">A decimal value.</param>
    ''' <summary>
    ''' Adds number to the workspace based on current value in workspace. 
    ''' </summary>
    Private Sub EnterNumber(number As Decimal)
        If txtWorkspace.Text.Length < 16 Or resetWorkspace = True Then
            buttonsEnabled(True)

            If txtWorkspace.Text = "0" AndAlso Not txtWorkspace.Text = "0." Then
                txtWorkspace.Text = number
                resetWorkspace = False
            ElseIf resetWorkspace = True AndAlso Not txtWorkspace.Text = "0." Then
                txtWorkspace.Text = number
                resetWorkspace = False
            ElseIf lblWorkspaceHold.Text <> "" AndAlso resetWorkspace = False Then
                txtWorkspace.Text = txtWorkspace.Text & number
            ElseIf resetWorkspace = True Then
                txtWorkspace.Text = txtWorkspace.Text & number
                resetWorkspace = False
            Else
                txtWorkspace.Text = txtWorkspace.Text & number
            End If

            signChanged = False
            errorOccurred = False
        End If
    End Sub



    ''' <param name="operation">Operations available: "mul", "div", "sub", or "add".</param>
    ''' <summary>
    ''' Calculates the expression based on given operation type. 
    ''' </summary>
    Private Sub Calculate(operation As String)
        Dim sign As String = ""
        Dim result As String = 0

        Select Case operation
            Case "mul", "*"
                sign = "*"
            Case "div", "/"
                sign = "/"
            Case "add", "+"
                sign = "+"
            Case "sub", "-"
                sign = "-"
            Case Else
                MsgBox("ERROR in CalculationType PROC: Could not assign proper sign.", MsgBoxStyle.Critical)
        End Select

        If signChanged = True And lblWorkspaceHold.Text <> "" Then
            Select Case operation
                Case "mul", "*"
                    lblWorkspaceHold.Text = lblWorkspaceHold.Text.Substring(0, lblWorkspaceHold.Text.Length - 3) & " x "
                Case Else
                    lblWorkspaceHold.Text = lblWorkspaceHold.Text.Substring(0, lblWorkspaceHold.Text.Length - 3) & " " & sign & " "
            End Select
        Else
            Select Case operation
                Case "mul", "*"
                    lblWorkspaceHold.Text += CleanUpNumber(txtWorkspace.Text) & " x "
                Case Else
                    lblWorkspaceHold.Text += CleanUpNumber(txtWorkspace.Text) & " " & sign & " "
            End Select

            Try
                If operation = "mul" OrElse operation = "div" OrElse
                   operation = "*" OrElse operation = "/" Then
                    result = New DataTable().Compute("1.0 * " & lblWorkspaceHold.Text.Replace("x", "*") & "1.0", Nothing)
                Else
                    result = New DataTable().Compute("1.0 * " & lblWorkspaceHold.Text.Replace("x", "*") & "0.0", Nothing)
                End If
                Exit Try
            Catch ex As DivideByZeroException
                txtWorkspace.Text = "Cannot divide by 0"
                ErrorCaught()
            Catch ex As OverflowException
                txtWorkspace.Text = "‭Overflow‬"
                ErrorCaught()
            End Try

            If errorOccurred = False Then
                txtWorkspace.Text = CleanUpNumber(result)
                resetWorkspace = True
                signChanged = True
            End If
        End If

    End Sub



    ''' <param name="enable">True to enable, False to disable.</param>
    ''' <summary>
    ''' Enables or disables non-digit buttons, except for Clear and ClearEntry.
    ''' </summary>
    Private Sub buttonsEnabled(enable As Boolean)
        If enable And btnEquals.Enabled = False Then
            btnBackspace.Enabled = True
            btnDivide.Enabled = True
            btnTimes.Enabled = True
            btnMinus.Enabled = True
            btnPlus.Enabled = True
            btnEquals.Enabled = True
            btnDecimal.Enabled = True
            btnPlusMinus.Enabled = True

            btnsEnabled = True
        ElseIf enable = False And btnEquals.Enabled = True Then
            btnBackspace.Enabled = False
            btnDivide.Enabled = False
            btnTimes.Enabled = False
            btnMinus.Enabled = False
            btnPlus.Enabled = False
            btnEquals.Enabled = False
            btnDecimal.Enabled = False
            btnPlusMinus.Enabled = False

            btnsEnabled = False
        End If
    End Sub


End Class
