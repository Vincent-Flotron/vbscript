'============================================================================
'  __      ______   _____           _       _     _______            
'  \ \    / /  _ \ / ____|         (_)     | |   |__   __|           
'   \ \  / /| |_) | (___   ___ _ __ _ _ __ | |_     | |_ __ ___  ___ 
'    \ \/ / |  _ < \___ \ / __| '__| | '_ \| __|    | | '__/ _ \/ _ \
'     \  /  | |_) |____) | (__| |  | | |_) | |_     | | | |  __/  __/
'      \/   |____/|_____/ \___|_|  |_| .__/ \__|    |_|_|  \___|\___|
'                                    | |                             
'                                    |_|

'============================================================================
' How to use:
'	cscript //nologo tree.vbs "C:\MyFolder"
'----------------------------------------------------------------------------

' Title
Call DisplayTitle

' Get the root folder from command line argument or use the current directory
If WScript.Arguments.Count > 0 Then
    rootPath = WScript.Arguments.Item(0)
Else
    Set fso = CreateObject("Scripting.FileSystemObject")
    rootPath = fso.GetAbsolutePathName(".")
End If

Set fso = CreateObject("Scripting.FileSystemObject")
Set rootFolder = fso.GetFolder(rootPath)

' Display the root folder path
WScript.Echo "Directory tree for " & rootFolder.Path & vbCrLf

' Recursively display folders and files
DisplayFolder rootFolder, 0

' Subroutine to display the contents of a folder and its subfolders
Sub DisplayFolder(folder, indent)
    ' Display the name of the folder
    WScript.Echo Space(indent) & "+--" & folder.Name

    ' Recursively display files and subfolders
    For Each subFolder In folder.SubFolders
        DisplayFolder subFolder, indent + 3
    Next

    For Each file In folder.Files
        WScript.Echo Space(indent + 3) & "|  " & file.Name
    Next
End Sub

Sub DisplayTitle
	WScript.Echo " __      ______   _____           _       _     _______            "
	WScript.Echo " \ \    / /  _ \ / ____|         (_)     | |   |__   __|           "
	WScript.Echo "  \ \  / /| |_) | (___   ___ _ __ _ _ __ | |_     | |_ __ ___  ___ "
	WScript.Echo "   \ \/ / |  _ < \___ \ / __| '__| | '_ \| __|    | | '__/ _ \/ _ \"
	WScript.Echo "    \  /  | |_) |____) | (__| |  | | |_) | |_     | | | |  __/  __/"
	WScript.Echo "     \/   |____/|_____/ \___|_|  |_| .__/ \__|    |_|_|  \___|\___|"
	WScript.Echo "                                   | |                              "
	WScript.Echo "                                   |_|                              "
	WScript.Echo vbCrLf
End Sub
