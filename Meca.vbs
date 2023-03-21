'Create GUI
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Load previously saved data
If objFSO.FileExists("Meca.csv") Then
    Set objFile = objFSO.OpenTextFile("Meca.csv", 1)
    Do Until objFile.AtEndOfStream
        strLine = objFile.ReadLine
        arrLine = Split(strLine, ";")
        If UBound(arrLine) = 3 Then
            strSavedInput1 = arrLine(1)
            strSavedInput2 = arrLine(2)
            strSavedInput3 = arrLine(3)
        End If
    Loop
    objFile.Close
End If

'Get user inputs
strInput1 = InputBox("Test Bench Name:", "Test Bench Data Entry", strSavedInput1)
If strInput1 = "" Then
    MsgBox "Please fill in the required fields.", vbExclamation, "Error"
    WScript.Quit
End If

strInput2 = InputBox("Task:", "Test Bench Data Entry", strSavedInput2)
If strInput2 = "" Then
    MsgBox "Please fill in the required fields.", vbExclamation, "Error"
    WScript.Quit
End If

strInput3 = InputBox("Comment:", "Test Bench Data Entry", strSavedInput3)

'Save data to CSV file
strDateTime = FormatDateTime(Now, 3) & " " & FormatDateTime(Now, 4)
strFileName = "Meca_" & strInput1 & ".csv"
If Not objFSO.FileExists(strFileName) Then
    Set objFile = objFSO.CreateTextFile(strFileName, True)
    objFile.WriteLine "Date and Time;Test Bench Name;Task;Comment"
Else
    Set objFile = objFSO.OpenTextFile(strFileName, 8)
End If
strLine = """" & strDateTime & """" & ";" & _
          """" & strInput1 & """" & ";" & _
          """" & strInput2 & """" & ";" & _
          """" & strInput3 & """"
objFile.WriteLine strLine
objFile.Close

MsgBox "Data saved successfully.", vbInformation, "Success"
