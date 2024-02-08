Option Explicit

Dim objFSO, objShell, objFolder, objOneDriveFolder
Dim sourceFolder, destinationFolder

' Define source and destination paths
sourceFolder = "P:\" ' Change this to the network drive letter
destinationFolder = CreateObject("WScript.Shell").SpecialFolders("OneDrive") & "\P drive\"
'Add note for K
' Create objects for file system operations and shell
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")

' Check if source folder exists
If objFSO.FolderExists(sourceFolder) Then
    ' Create destination folder if it doesn't exist
    If Not objFSO.FolderExists(destinationFolder) Then
        objFSO.CreateFolder(destinationFolder)
    End If
    
    ' Get the source folder object
    Set objFolder = objFSO.GetFolder(sourceFolder)
    
    ' Copy files from source to destination
    For Each objFile In objFolder.Files
        objFSO.CopyFile objFile.Path, destinationFolder & objFile.Name, True
    Next
    
    ' Show completion message
    MsgBox "Files copied successfully!", vbInformation, "Copy Complete"
Else
    ' If source folder doesn't exist, show error message
    MsgBox "Source folder does not exist!", vbExclamation, "Error"
End If

' Clean up objects
Set objFSO = Nothing
Set objShell = Nothing
