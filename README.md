# VBAuserform
The following project is a VBA user form code to capture data entered
User Form codes:

Dim TextBoxHandlers As Collection

Private Sub UserForm_Initialize()
    ' Get the screen width and height
    Dim screenWidth As Single
    Dim screenHeight As Single
    screenWidth = Application.Width
    screenHeight = Application.Height

    ' Calculate the form width and height as a percentage of the screen size
    Dim formWidth As Single
    Dim formHeight As Single
    formWidth = screenWidth * 0.9
    formHeight = screenHeight * 0.9

    ' Code for when the form opens up- you can adjust the default size of the form through this code.
    Me.StartUpPosition = 0 ' Manual
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
    Me.Width = formWidth ' Set your desired width
    Me.Height = formHeight ' Set your desired height

    ' Allow users to resize the form
    Me.BorderStyle = 1 ' Single border
    Me.ScrollBars = 3 ' Vertical and horizontal scrollbar if needed

    ' Initialize the collection
    Set TextBoxHandlers = New Collection
End Sub
Private Sub TextBox1_Enter()
    ' Code for Text Box 1 displayed at AT-W worksite number
    lblBoldText.Caption = "ENTER AT-W WORKSITE NUMBER"
    lblBoldText.Visible = True
    lblInputMessage.Caption = "Example: AT-W123456"
    lblInputMessage.Visible = True
End Sub

Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Hides the input message when the text box loses focus
    lblBoldText.Visible = False
    lblInputMessage.Visible = False
End Sub

Private Sub TextBox1_Change()
    Dim inputText As String
    inputText = TextBox1.Text
    
    ' Data validation with msgbox for AT-W text box.
    If Len(inputText) >= 4 Then
        If Left(inputText, 4) <> "AT-W" Or Len(inputText) > 10 Then
            MsgBox "Paste Valid Worksite Number with the correct formatting" & vbCrLf & "Example: AT-W123456", vbCritical + vbOKOnly, "Paste Valid Worksite Number"
            TextBox1.Text = ""
        End If
    End If
End Sub

Private Sub TextBox2_Enter()
' Code for Text Box 2 displayed at AT-T TMP number
    lblBoldText.Caption = "ENTER AT-T TMP NUMBER"
    lblBoldText.Visible = True
    lblInputMessage.Caption = "Example: AT-T123456"
    lblInputMessage.Visible = True
End Sub

Private Sub TextBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Hides the input message when the text box loses focus
    lblBoldText.Visible = False
    lblInputMessage.Visible = False
End Sub

Private Sub TextBox2_Change()
    Dim inputText As String
    inputText = TextBox2.Text
    
    ' Data validation with msgbox for AT-T text box.
    If Len(inputText) >= 4 Then
        If Left(inputText, 4) <> "AT-T" Or Len(inputText) > 10 Then
            MsgBox "Paste Valid TMP Number with the correct formatting" & vbCrLf & "Example: AT-T123456", vbCritical + vbOKOnly, "Enter Valid TMP Number"
            TextBox2.Text = ""
        End If
    End If
End Sub

Private Sub ComboBox1_Enter()
' Code for List/Combo Box displayed for Reason for Closure.
    lblBoldText.Caption = "CHOOSE FROM DROP DOWN LIST"
    lblBoldText.Visible = True
    lblInputMessage.Caption = ""
    lblInputMessage.Visible = False
End Sub

Private Sub ComboBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Hides the input message when the text box loses focus
    lblBoldText.Visible = False
    lblInputMessage.Visible = False
End Sub

Private Sub ComboBox1_Change()
    Dim i As Integer
    Dim isValid As Boolean
    
    ' Initialize the validation flag
    isValid = False
    
    ' Loop through the items in the ComboBox to check for matches
    For i = 0 To ComboBox1.ListCount - 1
        If ComboBox1.Text = ComboBox1.List(i) Then
            isValid = True
            Exit For
        End If
    Next i
    
    ' If the input is not valid, display the error message
    If Not isValid Then
        MsgBox "Please select or write a valid item suggested from the list.", vbCritical + vbOKOnly, "Invalid Entry"
        ComboBox1.Text = ""
    End If
End Sub

Private Sub ComboBox2_Enter()
' Code for List/Combo Box displayed for Road Closed and Suburb.
    lblBoldText.Caption = "CHOOSE FROM DROP DOWN LIST"
    lblBoldText.Visible = True
    lblInputMessage.Caption = "Ensure the correct suburb is selected"
    lblInputMessage.Visible = True
End Sub

Private Sub ComboBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Hides the input message when the text box loses focus
    lblBoldText.Visible = False
    lblInputMessage.Visible = False
End Sub

Private Sub ComboBox2_Change()
    Dim isValid As Boolean
    Dim i As Integer

    isValid = False

    ' Checks the value if it matches in the dropdown list
    For i = 0 To ComboBox2.ListCount - 1
        If ComboBox2.Value = ComboBox2.List(i) Then
            isValid = True
            Exit For
        End If
    Next i

    ' if the result is false, then it will display an error message.
    If Not isValid Then
        MsgBox "Only Use the drop-down arrow to select an option.", vbCritical + vbOKOnly, "Select from Drop Down List"
        ComboBox2.Value = ""
    End If
End Sub

Private Sub ComboBox3_Enter()
' Code for List/Combo Box displayed for Affected Length.
    lblBoldText.Caption = "CHOOSE FROM DROP DOWN LIST"
    lblBoldText.Visible = True
    lblInputMessage.Caption = ""
    lblInputMessage.Visible = False
End Sub

Private Sub ComboBox3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Hides the input message when the text box loses focus
    lblBoldText.Visible = False
    lblInputMessage.Visible = False
End Sub

Private Sub ComboBox3_Change()
    Dim isValid As Boolean
    Dim i As Integer

    isValid = False

    ' Checks the value if it matches in the dropdown list
    For i = 0 To ComboBox3.ListCount - 1
        If ComboBox3.Value = ComboBox3.List(i) Then
            isValid = True
            Exit For
        End If
    Next i

    ' if the result is false, then it will display an error message.
    If Not isValid Then
        MsgBox "Only Use the drop-down arrow to select an option.", vbCritical + vbOKOnly, "Select from Drop Down List"
        ComboBox3.Value = ""
    End If
End Sub

Private Sub TextBox19_Enter()
' Code for Text Box selected for From Location
    lblBoldText.Caption = "ENTER VALID STREET NUMBER"
    lblBoldText.Visible = True
    lblInputMessage.Caption = "Designated Feild is optional" & vbCrLf & _
                              " " & vbCrLf & _
                              "Please leave it blank if nothing to enter."
    lblInputMessage.Visible = True

End Sub

Private Sub TextBox19_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Hides the input message when the text box loses focus
    lblBoldText.Visible = False
    lblInputMessage.Visible = False
End Sub

Private Sub TextBox19_Change()
    Dim inputText As String
    Dim isNumber As Boolean
    Dim i As Integer
    Dim alphabetCount As Integer
    
    ' Get the input text from TextBox19
    inputText = TextBox19.Text
    
    ' Initialize the validation flag and alphabet count
    isValid = True
    alphabetCount = 0
    
    ' Check if the text length is more than 6 characters
    If Len(inputText) > 6 Then
        isValid = False
    Else
    
    ' Loop through each character in the input text
    For i = 1 To Len(inputText)
        If Not IsNumeric(Mid(inputText, i, 1)) Then
            alphabetCount = alphabetCount + 1
        End If
    Next i
    
    ' Check if the input is valid (only two characters allowed)
    If alphabetCount > 2 Then
        isValid = False
    End If
    End If
    ' If the input is not valid, display the error message and clear the text box
    If Not isValid Then
        MsgBox "Please enter a valid street number." & vbCrLf & "You must enter a number in designated field, no other charecters allowed", vbCritical + vbOKOnly, "Invalid Input"
        TextBox19.Text = ""
    End If
End Sub

Private Sub ComboBox7_Enter()
    ' Code for Text Box selected for From Location
    lblBoldText.Caption = "PLEASE SELECT OR WRITE VALID ADDRESS"
    lblBoldText.Visible = True
    lblInputMessage.Caption = "Enter the address in right format-" & vbCrLf & _
                              "ABBREVIATIONS ARE NOT ALLOWED" & vbCrLf & _
                              "Example: 20 Viaduct Harbour Avenue"
    lblInputMessage.Visible = True
End Sub

Private Sub ComboBox7_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Hides the input message when the text box loses focus
    lblBoldText.Visible = False
    lblInputMessage.Visible = False
End Sub
Private Sub ComboBox7_Change()
    Dim isValid As Boolean
    Dim i As Integer

    isValid = False

    ' Checks the value if it matches in the dropdown list
    For i = 0 To ComboBox7.ListCount - 1
        If ComboBox7.Value = ComboBox7.List(i) Then
            isValid = True
            Exit For
        End If
    Next i

    ' if the result is false, then it will display an error message.
    If Not isValid Then
        MsgBox "Only Use the drop-down arrow to select an option.", vbCritical + vbOKOnly, "Select from Drop Down List"
        ComboBox7.Value = ""
    End If
End Sub
Private Sub TextBox20_Enter()
' Code for Text Box selected for From Location
    lblBoldText.Caption = "ENTER VALID STREET NUMBER"
    lblBoldText.Visible = True
    lblInputMessage.Caption = "Designated Feild is optional" & vbCrLf & _
                              " " & vbCrLf & _
                              "Please leave it blank if nothing to enter."
    lblInputMessage.Visible = True

End Sub

Private Sub TextBox20_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Hides the input message when the text box loses focus
    lblBoldText.Visible = False
    lblInputMessage.Visible = False
End Sub

Private Sub TextBox20_Change()
    Dim inputText As String
    Dim isNumber As Boolean
    Dim i As Integer
    Dim alphabetCount As Integer
    
    ' Get the input text from TextBox20
    inputText = TextBox20.Text
    
    ' Initialize the validation flag and alphabet count
    isValid = True
    alphabetCount = 0
    
    ' Check if the text length is more than 6 characters
    If Len(inputText) > 6 Then
        isValid = False
    Else
    
    ' Loop through each character in the input text
    For i = 1 To Len(inputText)
        If Not IsNumeric(Mid(inputText, i, 1)) Then
            alphabetCount = alphabetCount + 1
        End If
    Next i
    
    ' Check if the input is valid (only two characters allowed)
    If alphabetCount > 2 Then
        isValid = False
    End If
    End If
    ' If the input is not valid, display the error message and clear the text box
    If Not isValid Then
        MsgBox "Please enter a valid street number." & vbCrLf & "You must enter a number in designated field, no other charecters allowed", vbCritical + vbOKOnly, "Invalid Input"
        TextBox20.Text = ""
    End If
End Sub

Private Sub ComboBox8_Enter()
' Code for Text Box selected for To Location
    lblBoldText.Caption = "PLEASE SELECT OR WRITE VALID ADDRESS"
    lblBoldText.Visible = True
    lblInputMessage.Caption = "Enter the address in right format-" & vbCrLf & _
                              "ABBREVIATIONS ARE NOT ALLOWED" & vbCrLf & _
                              "Example: 20 Viaduct Harbour Avenue"
    lblInputMessage.Visible = True
End Sub

Private Sub ComboBox8_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Hides the input message when the text box loses focus
    lblBoldText.Visible = False
    lblInputMessage.Visible = False
End Sub

Private Sub ComboBox8_Change()
    Dim isValid As Boolean
    Dim i As Integer

    isValid = False

    ' Checks the value if it matches in the dropdown list
    For i = 0 To ComboBox8.ListCount - 1
        If ComboBox8.Value = ComboBox8.List(i) Then
            isValid = True
            Exit For
        End If
    Next i

    ' if the result is false, then it will display an error message.
    If Not isValid Then
        MsgBox "Only Use the drop-down arrow to select an option.", vbCritical + vbOKOnly, "Select from Drop Down List"
        ComboBox8.Value = ""
    End If
End Sub
Private Sub OptionButtonYes_Click()
    ' Code for Label Box when the Yes option is selected
    Frame1.Caption = "Additional Road Closed 1 Details"
    Frame1.Visible = True
    Label19.Caption = "From Location"
    Label19.Visible = True
    Label20.Caption = "To Location"
    Label20.Visible = True
    TextBox21.Visible = True
    TextBox22.Visible = True
    ComboBox9.Visible = True
    ComboBox10.Visible = True
    CommandButton3.Caption = "Add More"
    CommandButton3.Visible = True
    CommandButton4.Caption = "Remove"
    CommandButton4.Visible = True
End Sub

Private Sub OptionButtonNo_Click()
    ' Code for when the No option is selected
    Frame1.Visible = False
    Label19.Visible = False
    Label20.Visible = False
    TextBox21.Visible = False
    TextBox22.Visible = False
    ComboBox9.Visible = False
    ComboBox10.Visible = False
    CommandButton3.Visible = False
    CommandButton4.Visible = False
End Sub

Private Sub CommandButton3_Click()
    On Error GoTo ErrorHandler
    
    ' Code for adding more frames
    Dim newFrame As MSForms.Frame
    Dim lblFrom As MSForms.Label
    Dim lblTo As MSForms.Label
    Dim newTextBoxFrom As MSForms.TextBox
    Dim newComboBoxFrom As MSForms.ComboBox
    Dim newTextBoxTo As MSForms.TextBox
    Dim newComboBoxTo As MSForms.ComboBox
    Dim frameCount As Integer
    Dim TextBoxHandler As ClsTextBoxHandler
    
    ' Count existing frames to determine the new frame number
    frameCount = GetFrameCount() + 1
    
    ' Create new frame
    Set newFrame = Me.Controls.Add("Forms.Frame.1", "Frame" & frameCount, True)
    With newFrame
        .Caption = "Additional Road Closed " & frameCount & " Details"
        .Top = Frame1.Top + ((frameCount - 1) * 80) ' Adjusted positioning
        .Left = Frame1.Left
        .Width = Frame1.Width
        .Height = Frame1.Height
        .Visible = True
        .Font.Name = Frame1.Font.Name
        .Font.Size = Frame1.Font.Size
        .Font.Bold = False
        .ForeColor = Frame1.ForeColor
    End With
    
    ' Create "From Location" label
    Set lblFrom = newFrame.Controls.Add("Forms.Label.1", "lblFrom" & frameCount, True)
    With lblFrom
        .Caption = "From Location"
        .Top = Label19.Top
        .Left = Label19.Left
        .Width = Label19.Width
        .Height = Label19.Height
        .Visible = True
        .Font.Name = Label19.Font.Name
        .Font.Size = Label19.Font.Size
        .Font.Bold = False
        .ForeColor = Label19.ForeColor
    End With
    
    ' Create "From Location" text box
    Set newTextBoxFrom = newFrame.Controls.Add("Forms.TextBox.1", "txtFrom" & frameCount, True)
    With newTextBoxFrom
        .Top = TextBox21.Top
        .Left = TextBox21.Left
        .Width = TextBox21.Width
        .Height = TextBox21.Height
        .Visible = True
        .Font.Name = TextBox21.Font.Name
        .Font.Size = TextBox21.Font.Size
    End With
    
    ' Create "From Location" combo box
    Set newComboBoxFrom = newFrame.Controls.Add("Forms.ComboBox.1", "cmbFrom" & frameCount, True)
    With newComboBoxFrom
        .Top = ComboBox9.Top
        .Left = ComboBox9.Left
        .Width = ComboBox9.Width
        .Height = ComboBox9.Height
        .Visible = True
        .Font.Name = ComboBox9.Font.Name
        .Font.Size = ComboBox9.Font.Size
        .RowSource = "RoadandSuburb"
        .MatchRequired = False
    End With
    
    ' Create "To Location" label
    Set lblTo = newFrame.Controls.Add("Forms.Label.1", "lblTo" & frameCount, True)
    With lblTo
        .Caption = "To Location"
        .Top = Label20.Top
        .Left = Label20.Left
        .Width = Label20.Width
        .Height = Label20.Height
        .Visible = True
        .Font.Name = Label20.Font.Name
        .Font.Size = Label20.Font.Size
        .Font.Bold = False
        .ForeColor = Label20.ForeColor
    End With
    
    ' Create "To Location" text box
    Set newTextBoxTo = newFrame.Controls.Add("Forms.TextBox.1", "txtTo" & frameCount, True)
    With newTextBoxTo
        .Top = TextBox22.Top
        .Left = TextBox22.Left
        .Width = TextBox22.Width
        .Height = TextBox22.Height
        .Visible = True
        .Font.Name = TextBox22.Font.Name
        .Font.Size = TextBox22.Font.Size
    End With
    
    ' Create "To Location" combo box
    Set newComboBoxTo = newFrame.Controls.Add("Forms.ComboBox.1", "cmbTo" & frameCount, True)
    With newComboBoxTo
        .Top = ComboBox10.Top
        .Left = ComboBox10.Left
        .Width = ComboBox10.Width
        .Height = ComboBox10.Height
        .Visible = True
        .Font.Name = ComboBox10.Font.Name
        .Font.Size = ComboBox10.Font.Size
        .RowSource = "RoadandSuburb"
        .MatchRequired = False
    End With
    
    ' Create instances of the class and assign controls
    Set TextBoxHandler = New ClsTextBoxHandler
    Set TextBoxHandler.TextBoxControl = newTextBoxFrom
    Set TextBoxHandler.ComboBoxControl = newComboBoxFrom
    TextBoxHandlers.Add TextBoxHandler
    
    Set TextBoxHandler = New ClsTextBoxHandler
    Set TextBoxHandler.TextBoxControl = newTextBoxTo
    Set TextBoxHandler.ComboBoxControl = newComboBoxTo
    TextBoxHandlers.Add TextBoxHandler
    
    ' Position the command buttons below the new frame
    CommandButton3.Top = newFrame.Top + newFrame.Height + 10
    CommandButton4.Top = newFrame.Top + newFrame.Height + 10
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Runtime Error"
End Sub

Private Function GetFrameCount() As Integer
    Dim ctrl As Control
    Dim count As Integer
    count = 0
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "Frame" Then
            count = count + 1
        End If
    Next ctrl
    
    GetFrameCount = count
End Function

Private Sub TextBox21_Enter()
    ' Code for Text Box selected for From Location
    lblBoldText.Caption = "ENTER VALID STREET NUMBER"
    lblBoldText.Visible = True
    lblInputMessage.Caption = "Designated Field is optional" & vbCrLf & _
                              " " & vbCrLf & _
                              "Please leave it blank if nothing to enter."
    lblInputMessage.Visible = True
End Sub

Private Sub TextBox21_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Hides the input message when the text box loses focus
    lblBoldText.Visible = False
    lblInputMessage.Visible = False
End Sub

Private Sub TextBox21_Change()
    Dim inputText As String
    Dim isValid As Boolean
    Dim i As Integer
    Dim alphabetCount As Integer
    
    ' Get the input text from TextBox21
    inputText = TextBox21.Text
    
    ' Initialize the validation flag and alphabet count
    isValid = True
    alphabetCount = 0
    
    ' Check if the text length is more than 6 characters
    If Len(inputText) > 6 Then
        isValid = False
    Else
        ' Loop through each character in the input text
        For i = 1 To Len(inputText)
            If Not IsNumeric(Mid(inputText, i, 1)) Then
                alphabetCount = alphabetCount + 1
            End If
        Next i
        
        ' Check if the input is valid (only two characters allowed)
        If alphabetCount > 2 Then
            isValid = False
        End If
    End If
    
    ' If the input is not valid, display the error message and clear the text box
    If Not isValid Then
        MsgBox "Please enter a valid street number." & vbCrLf & "You must enter a number in the designated field, no other characters allowed", vbCritical + vbOKOnly, "Invalid Input"
        TextBox21.Text = ""
    End If
End Sub

Private Sub ComboBox9_Enter()
    ' Code for Combo Box selected for From Location
    lblBoldText.Caption = "PLEASE SELECT OR WRITE VALID ADDRESS"
    lblBoldText.Visible = True
    lblInputMessage.Caption = "Enter the address in right format-" & vbCrLf & _
                              "ABBREVIATIONS ARE NOT ALLOWED" & vbCrLf & _
                              "Example: 20 Viaduct Harbour Avenue"
    lblInputMessage.Visible = True
End Sub

Private Sub ComboBox9_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Hides the input message when the combo box loses focus
    lblBoldText.Visible = False
    lblInputMessage.Visible = False
End Sub

Private Sub ComboBox9_Change()
    Dim isValid As Boolean
    Dim i As Integer

    isValid = False

    ' Checks the value if it matches in the dropdown list
    For i = 0 To ComboBox9.ListCount - 1
        If ComboBox9.Value = ComboBox9.List(i) Then
            isValid = True
            Exit For
        End If
    Next i

    ' If the result is false, then it will display an error message.
    If Not isValid Then
        MsgBox "Only Use the drop-down arrow to select an option.", vbCritical + vbOKOnly, "Select from Drop Down List"
        ComboBox9.Value = ""
    End If
End Sub

Private Sub CommandButton4_Click()
    ' Code for removing the last added frame
    Dim frameCount As Integer
    frameCount = GetFrameCount()
    
    If frameCount > 1 Then
        Me.Controls.Remove "Frame" & frameCount
    End If
    
    ' Reposition the command buttons
    If frameCount > 1 Then
        CommandButton3.Top = Me.Controls("Frame" & (frameCount - 1)).Top + Me.Controls("Frame" & (frameCount - 1)).Height + 10
        CommandButton4.Top = Me.Controls("Frame" & (frameCount - 1)).Top + Me.Controls("Frame" & (frameCount - 1)).Height + 10
    Else
        CommandButton3.Top = Frame1.Top + Frame1.Height + 10
        CommandButton4.Top = Frame1.Top + Frame1.Height + 10
    End If
End Sub

Private Sub TextBox22_Enter()
' Code for Text Box selected for From Location
    lblBoldText.Caption = "ENTER VALID STREET NUMBER"
    lblBoldText.Visible = True
    lblInputMessage.Caption = "Designated Feild is optional" & vbCrLf & _
                              " " & vbCrLf & _
                              "Please leave it blank if nothing to enter."
    lblInputMessage.Visible = True

End Sub

Private Sub TextBox22_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Hides the input message when the text box loses focus
    lblBoldText.Visible = False
    lblInputMessage.Visible = False
End Sub


Private Sub TextBox22_Change()
    Dim inputText As String
    Dim isNumber As Boolean
    Dim i As Integer
    Dim alphabetCount As Integer
    
    ' Get the input text from TextBox22
    inputText = TextBox22.Text
    
    ' Initialize the validation flag and alphabet count
    isValid = True
    alphabetCount = 0
    
    ' Check if the text length is more than 6 characters
    If Len(inputText) > 6 Then
        isValid = False
    Else
    
    ' Loop through each character in the input text
    For i = 1 To Len(inputText)
        If Not IsNumeric(Mid(inputText, i, 1)) Then
            alphabetCount = alphabetCount + 1
        End If
    Next i
    
    ' Check if the input is valid (only two characters allowed)
    If alphabetCount > 2 Then
        isValid = False
    End If
    End If
    ' If the input is not valid, display the error message and clear the text box
    If Not isValid Then
        MsgBox "Please enter a valid street number." & vbCrLf & "You must enter a number in designated field, no other charecters allowed", vbCritical + vbOKOnly, "Invalid Input"
        TextBox22.Text = ""
    End If
End Sub

Private Sub ComboBox10_Enter()
    ' Code for Text Box selected for From Location
    lblBoldText.Caption = "PLEASE SELECT OR WRITE VALID ADDRESS"
    lblBoldText.Visible = True
    lblInputMessage.Caption = "Enter the address in right format-" & vbCrLf & _
                              "ABBREVIATIONS ARE NOT ALLOWED" & vbCrLf & _
                              "Example: 20 Viaduct Harbour Avenue"
    lblInputMessage.Visible = True
End Sub

Private Sub ComboBox10_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Hides the input message when the text box loses focus
    lblBoldText.Visible = False
    lblInputMessage.Visible = False
End Sub

Private Sub ComboBox10_Change()
    Dim isValid As Boolean
    Dim i As Integer

    isValid = False

    ' Checks the value if it matches in the dropdown list
    For i = 0 To ComboBox10.ListCount - 1
        If ComboBox10.Value = ComboBox10.List(i) Then
            isValid = True
            Exit For
        End If
    Next i

    ' if the result is false, then it will display an error message.
    If Not isValid Then
        MsgBox "Only Use the drop-down arrow to select an option.", vbCritical + vbOKOnly, "Select from Drop Down List"
        ComboBox10.Value = ""
    End If
End Sub

Private Sub TextBox3_Enter()
' Code for Text Box selected for Start Date
    lblBoldText.Caption = "PLEASE ENTER VALID DATE"
    lblBoldText.Visible = True
    lblInputMessage.Caption = "Enter the date in right format-" & vbCrLf & _
                              "must be in the right format dd/mm/yyyy" & vbCrLf & _
                              " " & vbCrLf & _
                              "Example: 20/12/2025"
    lblInputMessage.Visible = True
End Sub

Private Sub TextBox3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Hides the input message when the text box loses focus
    lblBoldText.Visible = False
    lblInputMessage.Visible = False
End Sub
Private Sub TextBox3_AfterUpdate()
    Dim inputDate As Date
    Dim formattedDate As String
    
    On Error GoTo InvalidDate
    ' Convert the input text to a date
    inputDate = CDate(TextBox3.Text)
    
    ' Format the date as dd/mm/yyyy
    formattedDate = Format(inputDate, "dd/mm/yyyy")
    
    ' Set the formatted date back to TextBox3
    TextBox3.Text = formattedDate
    Exit Sub

InvalidDate:
    MsgBox "Please enter a valid date in the format dd/mm/yyyy.", vbCritical + vbOKOnly, "Invalid Date"
    TextBox3.Text = ""
End Sub

Private Sub TextBox4_Enter()
' Code for Text Box selected for Expected Finish Date
    lblBoldText.Caption = "PLEASE ENTER VALID DATE"
    lblBoldText.Visible = True
    lblInputMessage.Caption = "Enter the date in right format-" & vbCrLf & _
                              "must be in the right format dd/mm/yyyy" & vbCrLf & _
                              " " & vbCrLf & _
                              "Should be on or after the start date" & vbCrLf & _
                              " " & vbCrLf & _
                              "Example: 20/12/2025"
    lblInputMessage.Visible = True
End Sub

Private Sub TextBox4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Hides the input message when the text box loses focus
    lblBoldText.Visible = False
    lblInputMessage.Visible = False
End Sub

Private Sub TextBox4_AfterUpdate()
    Dim inputDate3 As Date
    Dim inputDate4 As Date
    Dim formattedDate As String
    
    ' Check if TextBox3 is empty
    If TextBox3.Text = "" Then
        MsgBox "Please enter a date in Start Date box first.", vbCritical + vbOKOnly, "Missing Date"
        TextBox4.Text = ""
        Exit Sub
    End If
    
    On Error GoTo InvalidDate
    ' Convert the input text to dates
    inputDate3 = CDate(TextBox3.Text)
    inputDate4 = CDate(TextBox4.Text)
    
    ' Check if TextBox4 date is less than TextBox3 date
    If inputDate4 < inputDate3 Then
        MsgBox "Please enter a future date in Expected Finish Date box", vbCritical + vbOKOnly, "Invalid Date"
        TextBox4.Text = ""
        Exit Sub
    End If
    
    ' Format the date as dd/mm/yyyy
    formattedDate = Format(inputDate4, "dd/mm/yyyy")
    
    ' Set the formatted date back to TextBox4
    TextBox4.Text = formattedDate
    Exit Sub

InvalidDate:
    MsgBox "Please enter a valid date in the format dd/mm/yyyy.", vbCritical + vbOKOnly, "Invalid Date"
    TextBox4.Text = ""
End Sub

Private Sub ComboBox4_Enter()
' Code for List/Combo Box displayed for Impact.
    lblBoldText.Caption = "CHOOSE FROM DROP DOWN LIST"
    lblBoldText.Visible = True
    lblInputMessage.Caption = ""
    lblInputMessage.Visible = False
End Sub

Private Sub ComboBox4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Hides the input message when the text box loses focus
    lblBoldText.Visible = False
    lblInputMessage.Visible = False
End Sub
Private Sub ComboBox4_Change()
    Dim isValid As Boolean
    Dim i As Integer

    isValid = False

    ' Checks the value if it matches in the dropdown list
    For i = 0 To ComboBox4.ListCount - 1
        If ComboBox4.Value = ComboBox4.List(i) Then
            isValid = True
            Exit For
        End If
    Next i

    ' if the result is false, then it will display an error message.
    If Not isValid Then
        MsgBox "Only Use the drop-down arrow to select an option.", vbCritical + vbOKOnly, "Select from Drop Down List"
        ComboBox4.Value = ""
    End If
End Sub
Private Sub ComboBox5_Enter()
' Code for List/Combo Box displayed for Start Time.
    lblBoldText2.Caption = "CHOOSE FROM DROP DOWN LIST"
    lblBoldText2.Visible = True
    lblInputMessage.Caption = ""
    lblInputMessage.Visible = False
End Sub

Private Sub ComboBox5_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Hides the input message when the text box loses focus
    lblBoldText2.Visible = False
    lblInputMessage.Visible = False
End Sub

Private Sub ComboBox5_Change()
    Dim isValid As Boolean
    Dim i As Integer

    isValid = False

    ' Checks the value if it matches in the dropdown list
    For i = 0 To ComboBox5.ListCount - 1
        If ComboBox5.Value = ComboBox5.List(i) Then
            isValid = True
            Exit For
        End If
    Next i

    ' if the result is false, then it will display an error message.
    If Not isValid Then
        MsgBox "Only Use the drop-down arrow to select an option.", vbCritical + vbOKOnly, "Select from Drop Down List"
        ComboBox5.Value = ""
    End If
End Sub

Private Sub ComboBox6_Enter()
' Code for List/Combo Box displayed for Impact.
    lblBoldText2.Caption = "CHOOSE FROM DROP DOWN LIST"
    lblBoldText2.Visible = True
    lblInputMessage.Caption = ""
    lblInputMessage.Visible = False
End Sub

Private Sub ComboBox6_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Hides the input message when the text box loses focus
    lblBoldText2.Visible = False
    lblInputMessage.Visible = False
End Sub

Private Sub ComboBox6_Change()
    Dim isValid As Boolean
    Dim i As Integer

    isValid = False

    ' Checks the value if it matches in the dropdown list
    For i = 0 To ComboBox6.ListCount - 1
        If ComboBox6.Value = ComboBox6.List(i) Then
            isValid = True
            Exit For
        End If
    Next i

    ' if the result is false, then it will display an error message.
    If Not isValid Then
        MsgBox "Only Use the drop-down arrow to select an option.", vbCritical + vbOKOnly, "Select from Drop Down List"
        ComboBox6.Value = ""
    End If
End Sub

Private Sub TextBox5_Enter()
' Code for Text Box displayed for Principal Project Contact.
    lblBoldText2.Caption = "ENTER CONTACT PERSON FULL NAME"
    lblBoldText2.Visible = True
End Sub

Private Sub TextBox5_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Hides the input message when the text box loses focus
    lblBoldText2.Visible = False
End Sub


Private Sub TextBox6_Enter()
' Code for Text Box displayed for Principal Contact Number.
    lblBoldText2.Caption = "ENTER CONTACT PERSON NUMBER"
    lblBoldText2.Visible = True
End Sub

Private Sub TextBox6_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Hides the input message when the text box loses focus
    lblBoldText2.Visible = False
End Sub

Private Sub TextBox6_AfterUpdate()
    Dim inputText As String
    Dim isValid As Boolean
    Dim i As Integer
    Dim emailPattern As String
    Dim regex As Object
    
    inputText = TextBox6.Text
    isValid = True
    
    ' Check if the input is an email address
    emailPattern = "^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$"
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = emailPattern
    regex.IgnoreCase = True
    regex.Global = True
    
    If regex.Test(inputText) Then
        ' Input is a valid email address
        isValid = True
    Else
        ' Input is not an email address, check for phone number
        ' Check if the input contains only numbers
        For i = 1 To Len(inputText)
            If Not IsNumeric(Mid(inputText, i, 1)) Then
                isValid = False
                Exit For
            End If
        Next i
        
        ' Check if the input is exactly 10 digits long
        If Len(inputText) <> 10 Then
            isValid = False
        End If
        
        ' Check if the input starts with 64 or +64
        If Left(inputText, 2) = "64" Or Left(inputText, 3) = "+64" Then
            isValid = False
        End If
    End If
    
    ' If the input is not valid, display the error message and clear the text box
    If Not isValid Then
        MsgBox "Please enter a valid 10-digit phone number (not starting with 64 or +64) or a valid email address.", vbCritical + vbOKOnly, "Invalid Input"
        TextBox6.Text = ""
    End If
End Sub

Private Sub TextBox7_Enter()
' Code for Text Box displayed for On-Site Contact.
    lblBoldText2.Caption = "ENTER CONTACT PERSON FULL NAME"
    lblBoldText2.Visible = True
End Sub

Private Sub TextBox7_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Hides the input message when the text box loses focus
    lblBoldText2.Visible = False
End Sub

Private Sub TextBox8_Enter()
' Code for Text Box displayed for On-Site Contact Number.
    lblBoldText2.Caption = "ENTER CONTACT PERSON NUMBER"
    lblBoldText2.Visible = True
End Sub

Private Sub TextBox8_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Hides the input message when the text box loses focus
    lblBoldText2.Visible = False
End Sub
Private Sub TextBox8_AfterUpdate()
    Dim inputText As String
    Dim isValid As Boolean
    Dim i As Integer
    Dim emailPattern As String
    Dim regex As Object
    
    inputText = TextBox8.Text
    isValid = True
    
    ' Check if the input is an email address
    emailPattern = "^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$"
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = emailPattern
    regex.IgnoreCase = True
    regex.Global = True
    
    If regex.Test(inputText) Then
        ' Input is a valid email address
        isValid = True
    Else
        ' Input is not an email address, check for phone number
        ' Check if the input contains only numbers
        For i = 1 To Len(inputText)
            If Not IsNumeric(Mid(inputText, i, 1)) Then
                isValid = False
                Exit For
            End If
        Next i
        
        ' Check if the input is exactly 10 digits long
        If Len(inputText) <> 10 Then
            isValid = False
        End If
        
        ' Check if the input starts with 64 or +64
        If Left(inputText, 2) = "64" Or Left(inputText, 3) = "+64" Then
            isValid = False
        End If
    End If
    
    ' If the input is not valid, display the error message and clear the text box
    If Not isValid Then
        MsgBox "Please enter a valid 10-digit phone number (not starting with 64 or +64) or a valid email address.", vbCritical + vbOKOnly, "Invalid Input"
        TextBox8.Text = ""
    End If
End Sub

Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Road Closure Form")
    
    ' Find the next available row
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row + 1
    
    ' Check if all mandatory fields are filled
    If TextBox1.Value = "" Or TextBox2.Value = "" Or ComboBox1.Value = "" Or ComboBox2.Value = "" Or ComboBox3.Value = "" Or ComboBox4.Value = "" Or TextBox3.Value = "" Or TextBox4.Value = "" Or ComboBox5.Value = "" Or ComboBox6.Value = "" Or TextBox5.Value = "" Or TextBox6.Value = "" Or TextBox7.Value = "" Or TextBox8.Value = "" Then
        MsgBox "Please fill in all mandatory fields.", vbExclamation
        Exit Sub
    End If
    
    ' Scenario 1: OptionButtonNo is selected
    If OptionButtonNo.Value = True Then
        ws.Cells(nextRow, 1).Value = TextBox1.Value
        ws.Cells(nextRow, 2).Value = TextBox2.Value
        ws.Cells(nextRow, 3).Value = ComboBox1.Value
        ws.Cells(nextRow, 4).Value = ComboBox2.Value
        ws.Cells(nextRow, 5).Value = ComboBox3.Value
        
        If TextBox19.Value <> "" Then
            ws.Cells(nextRow, 6).Value = Replace(TextBox19.Value & " " & ComboBox7.Value, "(", "")
            ws.Cells(nextRow, 6).Value = Replace(ws.Cells(nextRow, 6).Value, ")", "")
        End If
        
        If TextBox20.Value <> "" Then
            ws.Cells(nextRow, 7).Value = Replace(TextBox20.Value & " " & ComboBox8.Value, "(", "")
            ws.Cells(nextRow, 7).Value = Replace(ws.Cells(nextRow, 7).Value, ")", "")
        End If
        
        ws.Cells(nextRow, 8).Value = TextBox3.Value
        ws.Cells(nextRow, 9).Value = TextBox4.Value
        ws.Cells(nextRow, 10).Value = ComboBox4.Value
        ws.Cells(nextRow, 11).Value = ComboBox5.Value
        ws.Cells(nextRow, 12).Value = ComboBox6.Value
        ws.Cells(nextRow, 13).Value = TextBox5.Value
        ws.Cells(nextRow, 14).Value = TextBox6.Value
        ws.Cells(nextRow, 15).Value = TextBox7.Value
        ws.Cells(nextRow, 16).Value = TextBox8.Value
        
    ' Scenario 2: OptionButtonYes is selected without additional frames
    ElseIf OptionButtonYes.Value = True And GetFrameCount() = 1 Then
        ' First set of values
        ws.Cells(nextRow, 1).Value = TextBox1.Value
        ws.Cells(nextRow, 2).Value = TextBox2.Value
        ws.Cells(nextRow, 3).Value = ComboBox1.Value
        ws.Cells(nextRow, 4).Value = ComboBox2.Value
        ws.Cells(nextRow, 5).Value = ComboBox3.Value
        
        If TextBox19.Value <> "" Then
            ws.Cells(nextRow, 6).Value = Replace(TextBox19.Value & " " & ComboBox7.Value, "(", "")
            ws.Cells(nextRow, 6).Value = Replace(ws.Cells(nextRow, 6).Value, ")", "")
        End If
        
        If TextBox20.Value <> "" Then
            ws.Cells(nextRow, 7).Value = Replace(TextBox20.Value & " " & ComboBox8.Value, "(", "")
            ws.Cells(nextRow, 7).Value = Replace(ws.Cells(nextRow, 7).Value, ")", "")
        End If
        
        ws.Cells(nextRow, 8).Value = TextBox3.Value
        ws.Cells(nextRow, 9).Value = TextBox4.Value
        ws.Cells(nextRow, 10).Value = ComboBox4.Value
        ws.Cells(nextRow, 11).Value = ComboBox5.Value
        ws.Cells(nextRow, 12).Value = ComboBox6.Value
        ws.Cells(nextRow, 13).Value = TextBox5.Value
        ws.Cells(nextRow, 14).Value = TextBox6.Value
        ws.Cells(nextRow, 15).Value = TextBox7.Value
        ws.Cells(nextRow, 16).Value = TextBox8.Value
        
        ' Second set of values
        nextRow = nextRow + 1
        ws.Cells(nextRow, 1).Value = TextBox1.Value
        ws.Cells(nextRow, 2).Value = TextBox2.Value
        ws.Cells(nextRow, 3).Value = ComboBox1.Value
        ws.Cells(nextRow, 4).Value = ComboBox2.Value
        ws.Cells(nextRow, 5).Value = ComboBox3.Value
        
        If TextBox21.Value <> "" Then
            ws.Cells(nextRow, 6).Value = Replace(TextBox21.Value & " " & ComboBox9.Value, "(", "")
            ws.Cells(nextRow, 6).Value = Replace(ws.Cells(nextRow, 6).Value, ")", "")
        End If
        
        If TextBox22.Value <> "" Then
            ws.Cells(nextRow, 7).Value = Replace(TextBox22.Value & " " & ComboBox10.Value, "(", "")
            ws.Cells(nextRow, 7).Value = Replace(ws.Cells(nextRow, 7).Value, ")", "")
        End If
        
        ws.Cells(nextRow, 8).Value = TextBox3.Value
        ws.Cells(nextRow, 9).Value = TextBox4.Value
        ws.Cells(nextRow, 10).Value = ComboBox4.Value
        ws.Cells(nextRow, 11).Value = ComboBox5.Value
        ws.Cells(nextRow, 12).Value = ComboBox6.Value
        ws.Cells(nextRow, 13).Value = TextBox5.Value
        ws.Cells(nextRow, 14).Value = TextBox6.Value
        ws.Cells(nextRow, 15).Value = TextBox7.Value
        ws.Cells(nextRow, 16).Value = TextBox8.Value
        
    ' Scenario 3: OptionButtonYes is selected with additional frames
    ElseIf OptionButtonYes.Value = True And GetFrameCount() > 1 Then
        ' Loop through the dynamically added frames and copy values to the worksheet
        For i = 2 To GetFrameCount()  ' Start from 2 because the first frame is already processed
    Row = nextRow + i - 1
    
    ' Combine additional From and To Location values and remove brackets
    If Controls("txtFrom" & i).Value <> "" Then
        ws.Cells(Row, 6).Value = Replace(Controls("txtFrom" & i).Value & " " & Controls("cmbFrom" & i).Value, "(", "")
        ws.Cells(Row, 6).Value = Replace(ws.Cells(Row, 6).Value, ")", "")
    End If
    
    If Controls("txtTo" & i).Value <> "" Then
        ws.Cells(Row, 7).Value = Replace(Controls("txtTo" & i).Value & " " & Controls("cmbTo" & i).Value, "(", "")
        ws.Cells(Row, 7).Value = Replace(ws.Cells(Row, 7).Value, ")", "")
    End If
    
    ' Copy other values to the worksheet (keeping them same as before)
    ws.Cells(Row, 1).Value = TextBox1.Value
    ws.Cells(Row, 2).Value = TextBox2.Value
    ws.Cells(Row, 3).Value = ComboBox1.Value
    ws.Cells(Row, 4).Value = ComboBox2.Value
    ws.Cells(Row, 5).Value = ComboBox3.Value
    
    ws.Cells(Row, 8).Value = TextBox3.Value
    ws.Cells(Row, 9).Value = TextBox4.Value
    ws.Cells(Row, 10).Value = ComboBox4.Value
    ws.Cells(Row, 11).Value = ComboBox5.Value
    ws.Cells(Row, 12).Value = ComboBox6.Value
    ws.Cells(Row, 13).Value = TextBox5.Value
    ws.Cells(Row, 14).Value = TextBox6.Value
    ws.Cells(Row, 15).Value = TextBox7.Value
    ws.Cells(Row, 16).Value = TextBox8.Value
    
    nextRow = nextRow + 1
      Next i
    End If
End Sub

Module 1 Code:

Sub Button1_Click()
    Road_Closure_Form.Show
End Sub

Class Module code:
' clsTextBoxHandler class module
Public WithEvents TextBoxControl As MSForms.TextBox
Public WithEvents ComboBoxControl As MSForms.ComboBox

' Event handler for TextBox Enter event
Private Sub TextBoxControl_Enter()
    lblBoldText.Caption = "ENTER VALID STREET NUMBER"
    lblBoldText.Visible = True
    lblInputMessage.Caption = "Designated Field is optional" & vbCrLf & _
                              " " & vbCrLf & _
                              "Please leave it blank if nothing to enter."
    lblInputMessage.Visible = True
End Sub

' Event handler for TextBox Exit event
Private Sub TextBoxControl_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    lblBoldText.Visible = False
    lblInputMessage.Visible = False
End Sub

' Event handler for TextBox Change event
Private Sub TextBoxControl_Change()
    Dim inputText As String
    Dim isValid As Boolean
    Dim i As Integer
    Dim alphabetCount As Integer
    
    inputText = TextBoxControl.Text
    isValid = True
    alphabetCount = 0
    
    If Len(inputText) > 6 Then
        isValid = False
    Else
        For i = 1 To Len(inputText)
            If Not IsNumeric(Mid(inputText, i, 1)) Then
                alphabetCount = alphabetCount + 1
            End If
        Next i
        
        If alphabetCount > 2 Then
            isValid = False
        End If
    End If
    
    If Not isValid Then
        MsgBox "Please enter a valid street number." & vbCrLf & "You must enter a number in the designated field, no other characters allowed", vbCritical + vbOKOnly, "Invalid Input"
        TextBoxControl.Text = ""
    End If
    
End Sub


' Event handler for ComboBox Enter event
Private Sub ComboBoxControl_Enter()
    lblBoldText.Caption = "PLEASE SELECT OR WRITE VALID ADDRESS"
    lblBoldText.Visible = True
    lblInputMessage.Caption = "Enter the address in right format-" & vbCrLf & _
                              "ABBREVIATIONS ARE NOT ALLOWED" & vbCrLf & _
                              "Example: 20 Viaduct Harbour Avenue"
    lblInputMessage.Visible = True
End Sub

' Event handler for ComboBox Exit event
Private Sub ComboBoxControl_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    lblBoldText.Visible = False
    lblInputMessage.Visible = False
End Sub

' Event handler for ComboBox Change event
Private Sub ComboBoxControl_Change()
    Dim isValid As Boolean
    Dim i As Integer

    isValid = False

    For i = 0 To ComboBoxControl.ListCount - 1
        If ComboBoxControl.Value = ComboBoxControl.List(i) Then
            isValid = True
            Exit For
        End If
    Next i

    If Not isValid Then
        MsgBox "Only Use the drop-down arrow to select an option.", vbCritical + vbOKOnly, "Select from Drop Down List"
        ComboBoxControl.Value = ""
    End If
End Sub

