Option Explicit
Dim objFSO, objFolder, objFile, inputFolder, outputFolder, scriptFile, scriptPath, batchFile, batchPath, dxfFilePath

' Create File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Ask for the input folder
inputFolder = InputBox("Enter the folder path containing DWG files:", "Select Input Folder")

' Validate input folder
If inputFolder = "" Or Not objFSO.FolderExists(inputFolder) Then
    MsgBox "Invalid folder path!", vbExclamation, "Error"
    WScript.Quit
End If

' Define output folder
outputFolder = objFSO.BuildPath(inputFolder, "DXF_Converted")

' Create output folder if it doesn't exist
If Not objFSO.FolderExists(outputFolder) Then
    objFSO.CreateFolder(outputFolder)
End If

' Define script file path
scriptPath = objFSO.BuildPath(inputFolder, "convert.scr")

' Create and write to script file
Set scriptFile = objFSO.CreateTextFile(scriptPath, True)

' Suppress dialog boxes and prevent sysvar change prompts in AutoCAD
scriptFile.WriteLine "_FILEDIA 0"
scriptFile.WriteLine "_CMDDIA 0"
scriptFile.WriteLine "_SETVAR QAFLAGS 2"

' Loop through DWG files and generate script commands
Set objFolder = objFSO.GetFolder(inputFolder)
For Each objFile In objFolder.Files
    If LCase(objFSO.GetExtensionName(objFile.Name)) = "dwg" Then
        ' Define the corresponding DXF file path
        dxfFilePath = objFSO.BuildPath(outputFolder, objFSO.GetBaseName(objFile.Name) & ".dxf")

        ' Check if DXF file already exists
        If Not objFSO.FileExists(dxfFilePath) Then
            ' Open the DWG file
            scriptFile.WriteLine "_OPEN " & Chr(34) & objFile.Path & Chr(34)
            
            ' Convert the DWG file to DXF and save it in the output folder
            scriptFile.WriteLine "_SAVEAS DXF 16 " & Chr(34) & dxfFilePath & Chr(34)
            
            ' Add a delay to ensure AutoCAD processes the file properly
            scriptFile.WriteLine "_DELAY 2000"
        End If
    End If
Next

' Restore dialog settings and QAFLAGS
scriptFile.WriteLine "_SETVAR QAFLAGS 0"
scriptFile.WriteLine "_FILEDIA 1"
scriptFile.WriteLine "_CMDDIA 1"

' Close script file
scriptFile.Close

' Create batch file to kill AutoCAD if it's running
batchPath = objFSO.BuildPath(inputFolder, "TaskKillAutoCad.bat")
Set batchFile = objFSO.CreateTextFile(batchPath, True)
batchFile.WriteLine "@echo off"
batchFile.WriteLine "taskkill /IM acad.exe /F"
batchFile.Close

' Notify user
MsgBox "Script file 'convert.scr' and batch file 'TaskKillAutoCad.bat' created successfully in " & inputFolder, vbInformation, "Done"
