' Initialize/Set the variables

Dim WshShell, strSourceFile, strZipFile, strYear, strMonth, strDay

Set WshShell = WScript.CreateObject("WScript.Shell")
'Source

strSourceFile = WScript.Arguments.Item(0)
'Target

strZipFile = WScript.Arguments.Item(1)
fZip strFile, strZipFile

Function fZip(strFileToZip,sTargetZIPFile)

'This function will add all of the files in a source folder to a ZIP file
'using Windows' native folder ZIP capability.

Dim oShellApp, oFSO, iErr, sErrSource, sErrDescription

Set oShellApp = CreateObject("Shell.Application")

Set oFSO = CreateObject("Scripting.FileSystemObject")
'The source folder needs to have a \ on the End
' If Right(strFileToZip,1) <> "\" Then strFileToZip = strFileToZip & "\"

On Error Resume Next
'If a target ZIP exists already, delete it

If oFSO.FileExists(sTargetZIPFile) Then oFSO.DeleteFile sTargetZIPFile,True
iErr = Err.Number
sErrSource = Err.Source
sErrDescription = Err.Description

On Error GoTo 0
If iErr <> 0 Then
fZip = Array(iErr,sErrSource,sErrDescription)
Exit Function

End If



On Error Resume Next

'Write the fileheader for a blank zipfile.

oFSO.OpenTextFile(sTargetZIPFile, 2, True).Write "PK" & Chr(5) & Chr(6) & String(18, Chr(0))
iErr = Err.Number
sErrSource = Err.Source
sErrDescription = Err.Description

On Error GoTo 0
If iErr <> 0 Then
fZip = Array(iErr,sErrSource,sErrDescription)
Exit Function

End If
On Error Resume Next
'Start copying files into the zip from the source folder.
' oShellApp.NameSpace(sTargetZIPFile).CopyHere oShellApp.NameSpace(strFileToZip).Items
'Copy only one file

oShellApp.NameSpace(sTargetZIPFile).CopyHere strFileToZip
iErr = Err.Number
sErrSource = Err.Source
sErrDescription = Err.Description

On Error GoTo 0
If iErr <> 0 Then
fZip = Array(iErr,sErrSource,sErrDescription)
Exit Function

End If


WScript.Sleep 500

'Because the copying occurs in a separate process, the script will just continue. Run a DO...LOOP to prevent the function

'from exiting until the file is finished zipping.

'Do Until oShellApp.NameSpace(sTargetZIPFile).Items.Count = oShellApp.NameSpace(strFileToZip).Items.Count

'WScript.Sleep 500

'Loop

fZip = Array(0,"","")

End Function 