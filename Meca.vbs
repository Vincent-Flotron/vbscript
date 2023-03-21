' write me a vbscript with a gui that takes 3 inputs.
' The first input is "test bench name", the second is "task" and the third is "comment".
' There is a button named "Save". When this button is clicked, if the first and second input are filled, it save the result by adding a line in a csv file named "Meca" followed by the "test bench name" and finishing by ".csv".
' The first column will be the "date and time" in this format "yyyy.MM.dd hh:mm:ss:fff".
' So the format of an inserted line will be: "date and time";"test bench name";"task";"comment". 
' When the program starts, it display the already saved inputs.

' Do you understand what I want?

'Create GUI
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objForm = objShell.CreateDialog("Test Bench Data Entry", , 260, 240)
objForm.SetIcon "shell32.dll", 24

'Setup labels and input fields
Set objLabel1 = objForm.CreateLabel("Test Bench Name:", 10, 10, 100, 20)
Set objInput1 = objForm.CreateTextbox("", 120, 10, 120, 20)
Set objLabel2 = objForm.CreateLabel("Task:", 10, 40, 100, 20)
Set objInput2 = objForm.CreateTextbox("", 120, 40, 120, 20)
Set objLabel3 = objForm.CreateLabel("Comment:", 10, 70, 100, 20)
Set objInput3 = objForm.CreateTextbox("", 120, 70, 120, 20)

'Load previously saved data
If objFSO.FileExists("Meca.csv") Then
    Set objFile = objFSO.OpenTextFile("Meca.csv", 1)
    Do Until objFile.AtEndOfStream
        strLine = objFile.ReadLine
        arrLine = Split(strLine, ";")
        If UBound(arrLine) = 3 Then
            objInput1.Text = arrLine(1)
            objInput2.Text = arrLine(2)
            objInput3.Text = arrLine(3)
        End If
    Loop
    objFile.Close
End If

'Setup Save button
Set objButton = objForm.CreateButton("Save", 100, 120, 60, 25)
objButton.OnClick = GetRef("SaveData")

Sub SaveData
    'Check if required fields are filled
    If Trim(objInput1.Text) = "" Or Trim(objInput2.Text) = "" Then
        MsgBox "Please fill in the required fields.", vbExclamation, "Error"
        Exit Sub
    End If
    
    'Create or open CSV file
    strFileName = "Meca" & objInput1.Text & ".csv"
    If Not objFSO.FileExists(strFileName) Then
        Set objFile = objFSO.CreateTextFile(strFileName, True)
        objFile.WriteLine "Date and Time;Test Bench Name;Task;Comment"
    Else
        Set objFile = objFSO.OpenTextFile(strFileName, 8)
    End If
    
    'Write data to CSV file
    strDateTime = FormatDateTime(Now, 3) & " " & FormatDateTime(Now, 4)
    strLine = """" & strDateTime & """" & ";" & _
              """" & objInput1.Text & """" & ";" & _
              """" & objInput2.Text & """" & ";" & _
              """" & objInput3.Text & """"
    objFile.WriteLine strLine
    objFile.Close
    
    MsgBox "Data saved successfully.", vbInformation, "Success"
End Sub

'Run GUI
objForm.Show
