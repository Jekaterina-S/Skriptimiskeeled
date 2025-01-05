Option Explicit

' Check arguments
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript sort.vbs <SourceFolder> <TargetFolder>"
    WScript.Quit
End If

Dim sourceFolder, targetFolder, fso, filesMoved, foldersCreated
sourceFolder = WScript.Arguments(0)
targetFolder = WScript.Arguments(1)

Set fso = CreateObject("Scripting.FileSystemObject")
filesMoved = 0
foldersCreated = 0

SortPictures sourceFolder, targetFolder

WScript.Echo filesMoved & " picture" & Pluralize(filesMoved, "s") & " were sorted into " & foldersCreated & " folder" & Pluralize(foldersCreated, "s") & "."

' Sort pictures
Sub SortPictures(currentFolder, targetFolder)
    Dim folder, file, subFolder, fileDate, formattedDate,newFolderPath

    ' Create target folder if it doesn't exist
    If Not fso.FolderExists(targetFolder) Then
        fso.CreateFolder targetFolder
        WScript.Echo "Created folder: " & targetFolder
    End If

    For Each file In fso.GetFolder(currentFolder).Files
        If LCase(fso.GetExtensionName(file.Name)) = "jpg" Or LCase(fso.GetExtensionName(file.Name)) = "jpeg" Then
            formattedDate = Year(file.DateCreated) & "-" & Month(file.DateCreated) & "-" & Day(file.DateCreated)
            newFolderPath = fso.BuildPath(targetFolder, formattedDate)
            If Not fso.FolderExists(newFolderPath) Then
                fso.CreateFolder newFolderPath
                foldersCreated = foldersCreated + 1
            End If
            fso.MoveFile file.Path, fso.BuildPath(newFolderPath, file.Name)
            filesMoved = filesMoved + 1
        End If
    Next

    For Each subFolder In fso.GetFolder(currentFolder).SubFolders
        SortPictures subFolder.Path, targetFolder
    Next
End Sub

' Pluralize helper function
Function Pluralize(count, suffix)
    If count = 1 Then
        Pluralize = ""
    Else
        Pluralize = suffix
    End If
End Function
