
Sub import_all()
	Call import()
End Sub

Sub import(ByVal folder_path, ByVal columnsletters_, ByVal import_file_pattern, ByVal sep, byVal sheet_name, ByVal nb_lines_to_skip)
	' Deal with arguments
	If folder_path = "" Then
        folder_path = "C:\ProgramData\Virtual Unit\VA5_DISC\csv\test_export"
    End If
	
	If columnsletters_ = "" Then
		Dim columnsletters
		ReDim columnsletters(10)
		columnsletters(1) = "A"
		columnsletters(2) = "B"
		columnsletters(3) = "C"
		columnsletters(4) = "D"
		columnsletters(5) = "E"
		columnsletters(6) = "F"
		columnsletters(7) = "G"
		columnsletters(8) = "H"
		columnsletters(9) = "I"
		columnsletters(10) = "J"
	Else
		columnsletters = columnsletters_
	End If
	
	If import_file_pattern = "" Then
		import_file_pattern = "Kennwert"
	End If
	
	If sep = "" Then
		sep = ";"
	End If
	
	If sheet_name = "" Then
		sheet_name = "Import_from_csv"
	End If
	
	If nb_lines_to_skip = "" Then
		nb_lines_to_skip = 1
	End If
	
	' Open the files to import from
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folder_path)
    Set Files = folder.Files
	
	' Import all the files into the sheet "sheet_name"
    Dim Filename, line
    Dim i: i = nb_lines_to_skip + 1
    Dim column: column = 1	' column number of the sheet "sheet_name"
    Sheets(sheet_name).Select	' select the sheet
    For Each Item In Files
		' Import column of the matching file. Its name  must contain the "import_file_pattern"
        If InStr(Item.Name, import_file_pattern) Then
			' Write the filename as header of the column
            Range(columnsletters(column) & 1).Select
            ActiveCell.FormulaR1C1 = Item.Name
            
			' Open a file to import
            Filename = folder_path & "\" & Item.Name
            Set f = fso.OpenTextFile(Filename)
			
			' Skipt the lines to skip
			For i = 1 To 1 + nb_lines_to_skip
				f.ReadLine
			Next
            
			' Import 2nd column of the file to the sheet "sheet_name"
            Do Until f.AtEndOfStream
				' Read one line and take the 2nd part after the "sep"
                line = Split(f.ReadLine, sep)(1)
                'Debug.Print line
                Range(columnsletters(column) & i).Select
                ActiveCell.FormulaR1C1 = line
                i = i + 1
            Loop
			' Close the file
            f.Close
			f = Nothing
			column = column + 1
        End If
    i = nb_lines_to_skip + 1
    Next
	
	' Clean all
	Set Files = Nothing
    Set folder = Nothing
	Set fso = Nothing

End Sub
