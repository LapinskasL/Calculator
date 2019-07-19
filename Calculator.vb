'Author: Lukas Lapinskas
'Summary: A basic calculator. Style closely resembles Windows 10 Calculator.

'DISTANT TODO: add history, add light/dark modes switch

'NOTE: when I say 'sign' or 'operation', I mean either +, -, /, or *

Public Class frmCalculator

    Dim alterSign As Boolean = False        'set to true when user presses +,-,*,or / button. Set to false when user presses a digit number.
    Dim resetWorkspace As Boolean = True    'set to false when the user 
    Dim btnsEnabled As Boolean = True       'set to false if an error occurs
    Dim errorOccurred As Boolean = False    'set to true if an error occurs. Set to false once the user presses a button


    Dim drag As Boolean = False     'true if mouse is being dragged, false if not
    Dim mousex As Integer = 0       'holds value of cursor's x axis 
    Dim mousey As Integer = 0       'holds value of cursor's y axis 

    '@ lblCalculator MouseDown event handler
    Private Sub lblCalculator_MouseDown(sender As Object, e As MouseEventArgs) Handles lblCalculator.MouseDown
        'if the left mouse button is held  down (clicked without unclicking), the position 
        'of the cursor relative to the form's top left edge is calculated.
        If e.Button = Windows.Forms.MouseButtons.Left Then
            drag = True
            mousex = Windows.Forms.Cursor.Position.X - Me.Left  'X coordinate of cursor's location relative to the form's left edge
            mousey = Windows.Forms.Cursor.Position.Y - Me.Top   'Y coordinate of cursor's location relative to the form's top edge
        End If
    End Sub

    '@ lblCalculator MouseMove event handler
    Private Sub lblCalculator_MouseMove(sender As Object, e As MouseEventArgs) Handles lblCalculator.MouseMove
        'when the mouse is moved and drag is true, the form's position is changed
        If drag Then
            'the form's position is offset by where the mouse was MouseDowned on the form
            Me.Top = Windows.Forms.Cursor.Position.Y - mousey
            Me.Left = Windows.Forms.Cursor.Position.X - mousex
        End If
    End Sub

    '@ lblCalculator MouseUp event handler
    Private Sub lblCalculator_MouseUp(sender As Object, e As MouseEventArgs) Handles lblCalculator.MouseUp
        'drag is disabled
        drag = False
    End Sub





    '@ form KeyDown event handler
    Private Sub frmCalculator_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        'appropriate subroutines are called for keys pressed
        If btnsEnabled AndAlso (e.KeyCode = Keys.Add OrElse
           My.Computer.Keyboard.ShiftKeyDown AndAlso e.KeyCode = 187) Then
            Calculate("add")
        ElseIf btnsEnabled AndAlso (e.KeyCode = Keys.Subtract OrElse e.KeyCode = 189) Then
            Calculate("sub")
        ElseIf btnsEnabled AndAlso (e.KeyCode = Keys.Multiply OrElse
            (My.Computer.Keyboard.ShiftKeyDown AndAlso e.KeyCode = Keys.D8)) Then
            Calculate("mul")
        ElseIf btnsEnabled AndAlso (e.KeyCode = Keys.Divide OrElse e.KeyCode = 191) Then
            Calculate("div")
        ElseIf btnsEnabled AndAlso (e.KeyCode = 13 OrElse e.KeyCode = 187) Then
            btnEquals_Click(Nothing, Nothing)
        ElseIf btnsEnabled AndAlso (e.KeyCode = 190 OrElse e.KeyCode = 110) Then
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
        End If

        btnEquals.Select() 'refocus so the Enter key would work properly (Enter always clicks Equals button)
    End Sub

    '@ minimize label event handler
    Private Sub lblMinimize_Click(sender As Object, e As EventArgs) Handles lblMinimize.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    '@ exit label event handler
    Private Sub lblExit_Click(sender As Object, e As EventArgs) Handles lblExit.Click
        Me.Close()
    End Sub

    '@ clear entry (CE) button Click event handler
    Private Sub btnClearEntry_Click(sender As Object, e As EventArgs) Handles btnClearEntry.Click
        ButtonsEnabled(True)    'enables certain buttons
        txtWorkspace.Text = "0" 'sets workspace text to "0"
        'resetWorkspace = True  'the next digit pressed will reset the "0" in the workspace
        alterSign = False       'since the user hasn't pressed any operator button, there is no operator to alter
        errorOccurred = False   'since the calculator is reset, there's no more error

        btnEquals.Select() 'refocus so the Enter key would work properly (Enter always clicks Equals button)
    End Sub

    '@ clear (C) button Click event handler
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ButtonsEnabled(True)    'enables certain buttons
        txtWorkspace.Text = "0" 'sets workspace text to "0"
        'resetWorkspace = True  'the next digit pressed will reset the "0" in the workspace
        alterSign = False       'since the user hasn't pressed any operator button, there is no operator to alter
        errorOccurred = False   'since the calculator is reset, there's no more error

        txtWorkspaceHold.Text = ""  'empty any text from workspaceHold texbox

        btnEquals.Select() 'refocus so the Enter key would work properly (Enter always clicks Equals button)
    End Sub

    '@ backspace button Click event handler
    Private Sub btnBackspace_Click(sender As Object, e As EventArgs) Handles btnBackspace.Click
        'if cleaned workspace text is not 0 AND is not empty, OR unclean text is "0."
        If (CleanUpNumber(txtWorkspace.Text) <> "0" AndAlso txtWorkspace.Text <> "") OrElse txtWorkspace.Text = "0." Then
            'if it is a single-digit negative number (2 chars)
            If txtWorkspace.Text.Length = 2 Then
                'if text contains "-"
                If txtWorkspace.Text.Contains("-"c) Then
                    txtWorkspace.Text = "0" 'reset workspace text to 0
                Else
                    'remove rightmost character from workspace textbox
                    txtWorkspace.Text = txtWorkspace.Text.Substring(0, txtWorkspace.Text.Length - 1)
                End If
            ElseIf txtWorkspace.Text.Length = 1 Then 'if there's only 1 char (1 digit)
                txtWorkspace.Text = "0" 'reset workspace text to 0
            Else
                'remove rightmost character from workspace textbox
                txtWorkspace.Text = txtWorkspace.Text.Substring(0, txtWorkspace.Text.Length - 1)
            End If
        End If

        btnEquals.Select() 'refocus so the Enter key would work properly (Enter always clicks Equals button)
    End Sub

    '@ number (0 through 9) buttons Click event handlers
    Private Sub NumberButtons_Click(sender As Object, e As EventArgs) Handles btnZero.Click, btnOne.Click, btnTwo.Click,
                                                    btnThree.Click, btnFour.Click, btnFive.Click, btnSix.Click,
                                                    btnSeven.Click, btnEight.Click, btnNine.Click

        Dim button As Button = CType(sender, Button) 'holds button clicked

        EnterNumber(button.Text) 'call EnterNumber with the button's text (one of the digits)

        btnEquals.Select() 'refocus so the Enter key would work properly (Enter always clicks Equals button)

    End Sub


    '@ operation buttons (- + x /) Click event handlers
    Private Sub OperationButtons_Click(sender As Object, e As EventArgs) Handles btnPlus.Click, btnMinus.Click, btnTimes.Click, btnDivide.Click

        Dim button As Button = CType(sender, Button) 'holds button clicked

        'selects case depending on text of the button and performs a calculation depending on text
        Select Case button.Text
            Case "+", "-", "/"
                Calculate(button.Text)
            Case "x"
                Calculate("*")
        End Select

        btnEquals.Select() 'refocus so the Enter key would work properly (Enter always clicks Equals button)
    End Sub

    '@ equals button Click event handler
    Private Sub btnEquals_Click(sender As Object, e As EventArgs) Handles btnEquals.Click
        'concatonate text in both textboxes
        Dim bothParts As String = txtWorkspaceHold.Text & txtWorkspace.Text
        Dim result As Decimal = 0

        'if text contains "x", then replace it with "*"
        If bothParts.Contains("x") Then
            bothParts = bothParts.Replace("x", "*")
        End If

        bothParts = "1.0 * " & bothParts 'when calculation takes place, this converts the answer to Decimal

        Try
            result = New DataTable().Compute(bothParts, Nothing) 'computes the expression
            txtWorkspace.Text = CleanUpNumber(result) 'clean up calculation
            resetWorkspace = True                     'workspace has the answer, so any digit button will reset the value
            txtWorkspaceHold.Text = ""                'clear text in workspaceHold
            Exit Try
        Catch ex As DivideByZeroException   'if user divides by zero, error is displayed
            txtWorkspace.Text = "Cannot divide by 0"
            ErrorCaught()
        Catch ex As OverflowException       'if the value is too large to fit, error is displayed
            txtWorkspace.Text = "‭Overflow‬"
            ErrorCaught()
        End Try

        btnEquals.Select() 'refocus so the Enter key would work properly (Enter always clicks Equals button)
    End Sub

    '@ plus-minus button Click event handler
    Private Sub btnPlusMinus_Click(sender As Object, e As EventArgs) Handles btnPlusMinus.Click
        'if workspace is not 0 and not empty
        If txtWorkspace.Text <> "0" AndAlso txtWorkspace.Text <> "" Then
            'if the first character in workspace is not a "-", then put a "-" in front
            If txtWorkspace.Text.First <> "-" Then
                txtWorkspace.Text = "-" & txtWorkspace.Text
            Else
                'otherwise, remove the first char (which must be a "-")
                txtWorkspace.Text = txtWorkspace.Text.Substring(1, txtWorkspace.Text.Length - 1)
            End If
        End If

        btnEquals.Select() 'refocus so the Enter key would work properly (Enter always clicks Equals button)
    End Sub

    '@ decimal (.) button Click event handler
    Private Sub btnDecimal_Click(sender As Object, e As EventArgs) Handles btnDecimal.Click
        'if workspace workspace doesn't contain a decimal and doesn't need to be reset
        If txtWorkspace.Text.Contains(".") = False AndAlso resetWorkspace = False Then
            txtWorkspace.Text = txtWorkspace.Text & "." 'add a decimal to the text
        ElseIf resetWorkspace = True Then   'if the workspace does need to be reset
            txtWorkspace.Text = "0." 'replace workspace text with "0."
            resetWorkspace = False   'workspace no longer needs to be reset
        End If

        btnEquals.Select() 'refocus so the Enter key would work properly (Enter always clicks Equals button)
    End Sub

    '@ MouseEnter event handlers for the buttons in the first three columns
    Private Sub First3ColumnsButtons_MouseEnter(sender As Object, e As EventArgs) Handles btnBackspace.MouseEnter, btnZero.MouseEnter,
                                            btnTwo.MouseEnter, btnThree.MouseEnter, btnSix.MouseEnter, btnSeven.MouseEnter,
                                            btnPlusMinus.MouseEnter, btnOne.MouseEnter, btnNine.MouseEnter, btnFour.MouseEnter,
                                            btnFive.MouseEnter, btnEight.MouseEnter, btnDecimal.MouseEnter, btnClearEntry.MouseEnter,
                                            btnClear.MouseEnter

        Dim EnteredButton = CType(sender, Button) 'holds the button pressed
        EnteredButton.BackColor = Color.FromArgb(60, 60, 60) 'changes button's color
    End Sub

    '@ MouseEnter event handlers for column 4 buttons
    Private Sub Column4Buttons_MouseEnter(sender As Object, e As EventArgs) Handles btnTimes.MouseEnter, btnPlus.MouseEnter,
                                            btnMinus.MouseEnter, btnEquals.MouseEnter, btnDivide.MouseEnter

        Dim EnteredButton As Button = CType(sender, Button) 'holds the button pressed

        EnteredButton.BackColor = Color.FromArgb(255, 144, 0) 'changes button's color
        EnteredButton.ForeColor = Color.FromArgb(31, 31, 31)  'changes buttons text color
    End Sub

    '@ MouseLeave event handlers for column 4 buttons
    Private Sub Column4Buttons_MouseLeave(sender As Object, e As EventArgs) Handles btnTimes.MouseLeave, btnPlus.MouseLeave,
                                            btnMinus.MouseLeave, btnEquals.MouseLeave, btnDivide.MouseLeave

        Dim LeftButton = CType(sender, Button) 'holds the button pressed

        LeftButton.BackColor = Color.FromArgb(19, 19, 19)    'changes button's color
        LeftButton.ForeColor = Color.FromArgb(255, 255, 255) 'changes button's text color
    End Sub

    '@ MouseLeave event handlers for the non-digit buttons
    Private Sub NonDigitButtons_MouseLeave(sender As Object, e As EventArgs) Handles btnBackspace.MouseLeave,
                                            btnPlusMinus.MouseLeave, btnMinus.MouseLeave, btnDecimal.MouseLeave,
                                            btnClearEntry.MouseLeave, btnClear.MouseLeave

        Dim LeftButton As Button = CType(sender, Button)  'holds the button pressed

        LeftButton.BackColor = Color.FromArgb(19, 19, 19) 'changes button's color
    End Sub

    '@ minimize label MouseEnter event handler
    Private Sub lblMinimize_MouseEnter(sender As Object, e As EventArgs) Handles lblMinimize.MouseEnter
        lblMinimize.BackColor = Color.FromArgb(60, 60, 60) 'changes button's color
    End Sub

    '@ minimize and exit labels MouseLeave event handlers
    Private Sub MinimizeExitLabels_MouseLeave(sender As Object, e As EventArgs) Handles lblMinimize.MouseLeave, lblExit.MouseLeave
        Dim LabelExitMinimize As Label = CType(sender, Label)    'holds label handled

        LabelExitMinimize.BackColor = Color.FromArgb(40, 40, 40) 'changes label's background color
    End Sub

    '@ exit label MouseEnter event handler
    Private Sub lblExit_MouseEnter(sender As Object, e As EventArgs) Handles lblExit.MouseEnter
        lblExit.BackColor = Color.Red 'changes color of the exit label
    End Sub

    '@ buttons 0 through 9 MouseLeave event handlers
    Private Sub DigitButtons_MouseLeave(sender As Object, e As EventArgs) Handles btnZero.MouseLeave, btnOne.MouseLeave, btnTwo.MouseLeave,
                                               btnThree.MouseLeave, btnFour.MouseLeave, btnFive.MouseLeave, btnSix.MouseLeave,
                                               btnSeven.MouseLeave, btnNine.MouseLeave, btnEight.MouseLeave

        Dim LeftDigitButton As Button = CType(sender, Button) 'holds button MouseLeave'd

        LeftDigitButton.BackColor = Color.FromArgb(6, 6, 6)   'changes color of the button
    End Sub



















    ''' <param name="numberString">A string containing a properly formatted number.</param>
    ''' <summary>
    ''' Removes redundant zeros and decimal numbers (ex: 1.0200 would become 1.02).
    ''' </summary>
    Private Function CleanUpNumber(numberString As String)
        Dim cleanNumber As String = numberString 'clean number initialized with parameter in case doesn't need "cleaning"

        'if last char in numberString is "."
        If numberString.Last = "."c Then
            cleanNumber = numberString.Substring(0, numberString.Length - 1) 'remove the "." char 
        ElseIf numberString.Contains("."c) Then 'else if numberString contains "." char
            'while the last character is "0"
            While numberString.Last = "0"c
                numberString = numberString.Substring(0, numberString.Length - 1) 'remove last character (which is "0")
            End While
            cleanNumber = numberString          'update cleanNumber variable

            Dim isNotZero As Boolean = False    'set to true if a character is not zero

            Dim index = numberString.IndexOf("."c) + 1 'find and store the the index of the character after "." char
            'while index is not out of bounds of numberString and character is not "0"
            '(this while loop basically checks if there are any non-zero numbers after the decimal point. So if there aren't,
            'that means it's only zeros, which can be removed along with the decimal point. ex: 1.0100 becomes 1.01)
            While index < numberString.Length AndAlso isNotZero = False
                If numberString(index) <> "0"c Then 'if the character at numberString(index) is not a zero
                    isNotZero = True                'set isNotZero to True
                End If
                index += 1                          'increment index.
            End While

            If isNotZero = False Then   'if isNotZero was False
                cleanNumber = ""        'set cleanNumber to empty
                For index = 0 To numberString.IndexOf("."c) - 1
                    cleanNumber += numberString(index)  'add each character up to the non-zero number
                Next
            End If
        End If

        'if the length of the cleanNumber is more than 16 digits
        If cleanNumber.Length > 16 Then
            'if the cleanNumber contains a decimal and doesn't contain exponent (char "E")
            If cleanNumber.Contains("."c) And Not cleanNumber.Contains("E") Then
                cleanNumber = cleanNumber.Substring(0, 16)      'extract only the first 16 chars of cleanNumber
            Else
                cleanNumber = Format(cleanNumber, "scientific") 'else, format the number to scientific notation
            End If
        End If

        Return cleanNumber  'return the cleaned number
    End Function



    ''' <param name="number">A decimal value.</param>
    ''' <summary>
    ''' Adds number to the workspace text. 
    ''' </summary>
    Private Sub EnterNumber(number As Decimal)
        'if workspace's text length is less than 16 chars OR workspace needs to be reset
        If txtWorkspace.Text.Length < 16 Or resetWorkspace = True Then
            ButtonsEnabled(True)
            'if workspace text is "0"
            If txtWorkspace.Text = "0" Then
                txtWorkspace.Text = number  'change "0" to value in number
                resetWorkspace = False

                'else if workspace text needs to be reset and it doesn't equal "0."
            ElseIf resetWorkspace = True AndAlso Not txtWorkspace.Text = "0." Then
                txtWorkspace.Text = number  'change "0" to value in number
                resetWorkspace = False

                'else if workspace text is not empty and does not need to be reset
            ElseIf txtWorkspaceHold.Text <> "" AndAlso resetWorkspace = False Then
                txtWorkspace.Text = txtWorkspace.Text & number  'add the number to text
            Else
                txtWorkspace.Text = txtWorkspace.Text & number
            End If

            alterSign = False       'because a digit was inputted, there is no sign (+, -, *, or /) to alter
            errorOccurred = False   'any error that had occured is irrelevant now
        End If
    End Sub



    ''' <param name="operation">Operations available: "mul", "div", "sub", or "add".</param>
    ''' <summary>
    ''' Calculates the expression based on given operation type. 
    ''' </summary>
    Private Sub Calculate(operation As String)
        Dim sign As String = ""     'holds the sign (either +, -, /, or *
        Dim result As String = 0    'holds the calculated result

        'based on the parameter value, select change the value of the "sign" variable
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

        'if the sign needs changing AND workspaceHold text isn't empty
        If alterSign = True And txtWorkspaceHold.Text <> "" Then
            'replace the most recent operation sign (-, +, /, or *) in workspaceHold
            Select Case operation
                Case "mul", "*" 'if it's multiplication, replace it with an "x"
                    txtWorkspaceHold.Text = txtWorkspaceHold.Text.Substring(0, txtWorkspaceHold.Text.Length - 3) & " x "
                Case Else
                    txtWorkspaceHold.Text = txtWorkspaceHold.Text.Substring(0, txtWorkspaceHold.Text.Length - 3) & " " & sign & " "
            End Select
        Else 'otherwise
            'add the operation sign (-, +, /, or *) to workspaceHold
            Select Case operation
                Case "mul", "*" 'if it's multiplication, replace it with an "x"
                    txtWorkspaceHold.Text += CleanUpNumber(txtWorkspace.Text) & " x "
                Case Else
                    txtWorkspaceHold.Text += CleanUpNumber(txtWorkspace.Text) & " " & sign & " "
            End Select

            'attempt to calculate, but catch arithmetic and overflow errors
            Try
                'if the operation at the end is multiplication or division
                If operation = "mul" OrElse operation = "div" OrElse
                   operation = "*" OrElse operation = "/" Then
                    'add "1.0" at the end to prevent crashes when trying to calculate (ex: 65 * 1.0 instead of 65. *), and
                    'to convert the result into decimal
                    result = New DataTable().Compute("1.0 * " & txtWorkspaceHold.Text.Replace("x", "*") & "1.0", Nothing)
                Else
                    'add 0.0 at the end to prevent crashes and convert result to decimal
                    result = New DataTable().Compute("1.0 * " & txtWorkspaceHold.Text.Replace("x", "*") & "0.0", Nothing)
                End If
                Exit Try
            Catch ex As DivideByZeroException               'if user attempts to divide by zero
                txtWorkspace.Text = "Cannot divide by 0"    'show an error
                ErrorCaught()                               'call ErrorCaught PROC
            Catch ex As OverflowException           'if value does not fit in Decimal
                txtWorkspace.Text = "‭Overflow‬"      'show an error
                ErrorCaught()                       'call ErrorCaught PROC
            End Try

            'if an eror was not
            If errorOccurred = False Then
                txtWorkspace.Text = CleanUpNumber(result)   'show result
                resetWorkspace = True                       'indicate that next digit entry will reset workspace text
                alterSign = True                            'indicate that the sign will need to be altered
            End If
        End If

    End Sub



    ''' <param name="enable">True to enable, False to disable.</param>
    ''' <summary>
    ''' Checks if buttons need to be enabled or disabled. If needed, it does so based on parameter request.
    ''' </summary>
    Private Sub ButtonsEnabled(enable As Boolean)
        If enable And btnEquals.Enabled = False Then
            ButtonsEnabled_2(enable)
        ElseIf enable = False And btnEquals.Enabled = True Then
            ButtonsEnabled_2(enable)
        End If
    End Sub



    ''' <param name="enable">True to enable, False to disable.</param>
    ''' <summary>
    ''' Enables or disables non-digit buttons, except for Clear and ClearEntry.
    ''' </summary>
    Private Sub ButtonsEnabled_2(enable As Boolean)
        btnBackspace.Enabled = enable
        btnDivide.Enabled = enable
        btnTimes.Enabled = enable
        btnMinus.Enabled = enable
        btnPlus.Enabled = enable
        btnEquals.Enabled = enable
        btnDecimal.Enabled = enable
        btnPlusMinus.Enabled = enable

        btnsEnabled = enable
    End Sub




    ''' <summary>
    ''' Runs necessary instructions if an error is caught.
    ''' </summary>
    Private Sub ErrorCaught()
        resetWorkspace = True
        txtWorkspaceHold.Text = ""
        ButtonsEnabled(False)
        errorOccurred = True
    End Sub

End Class
