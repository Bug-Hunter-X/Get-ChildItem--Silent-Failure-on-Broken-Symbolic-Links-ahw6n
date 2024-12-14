Function GetChildItemSafe(path)
  Dim fso, file, err
  Set fso = CreateObject("Scripting.FileSystemObject")
  On Error Resume Next

  'Check if target file/directory exists
  If fso.FileExists(path) Or fso.FolderExists(path) Then
    Set file = fso.GetFile(path) 
    'Process the file
    'Example: WScript.Echo file.Name, file.Size
    GetChildItemSafe = file
  Else
    'Handle the error, broken link etc.
    Err.Raise vbError, , "File or directory not found: " & path
  End If
  On Error Goto 0
  Set fso = Nothing
  Set file = Nothing
End Function

' Example usage
Dim fileInfo
fileInfo = GetChildItemSafe("brokenLink.txt")
If IsObject(fileInfo) Then
  WScript.Echo "File exists, Name: " & fileInfo.Name & ", Size: " & fileInfo.Size
Else
  WScript.Echo "Error: File does not exist or is inaccessible."
End If