Option Explicit

Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim srcDir, destDir
srcDir = WScript.Arguments(0) & "\"
destDir = WScript.Arguments(1) & "\"

Dim extJPEG, extJPG
extJPEG = "jpeg"
extJPG = "jpg"

Dim countFiles, countDirs
countFiles = 0
countDirs = 0

If Not objFSO.FolderExists(srcDir) Or Not objFSO.FolderExists(destDir) Then
    WScript.Echo "Allikas puudub."
    WScript.Quit
End If

Function FormatNumber(num, digits)
    FormatNumber = Right(String(digits, "0") & num, digits)
End Function

Function ExploreFolder(path)
    Dim fld, item, f
    Set fld = objFSO.GetFolder(path)
    
    For Each f In fld.Files
        SortFile f
    Next
    
    For Each item In fld.SubFolders
        ExploreFolder item.Path
    Next
End Function

Function SortFile(ByRef f)
    If LCase(objFSO.GetExtensionName(f.Path)) = extJPEG Or LCase(objFSO.GetExtensionName(f.Path)) = extJPG Then
        Dim yearModified, formattedDate, yearDir, dateDir
        yearModified = Year(f.DateLastModified)
        formattedDate = yearModified & "-" & FormatNumber(Month(f.DateLastModified), 2) & "-" & FormatNumber(Day(f.DateLastModified), 2)
        yearDir = destDir & yearModified & "\"
        dateDir = yearDir & formattedDate & "\"
        
        If Not objFSO.FolderExists(yearDir) Then
            objFSO.CreateFolder(yearDir)
            countDirs = countDirs + 1
        End If
        
        If Not objFSO.FolderExists(dateDir) Then
            objFSO.CreateFolder(dateDir)
            countDirs = countDirs + 1
        End If
        
        f.Copy dateDir
        countFiles = countFiles + 1
    End If
End Function

Function LogResults(path)
    Dim dir, subDir, fl, fileList, numberOfFiles
    Set dir = objFSO.GetFolder(path)
    Set fileList = CreateObject("System.Collections.ArrayList")
    numberOfFiles = 0
    
    For Each subDir In dir.SubFolders
        For Each fl In subDir.Files
            fileList.Add(fl.Name)
            numberOfFiles = numberOfFiles + 1
        Next
        
        If numberOfFiles > 0 Then
            WScript.Echo "--------"
            WScript.Echo numberOfFiles & " fail(id)"
            WScript.Echo Join(fileList.ToArray(), ", ")
            WScript.Echo "teisaldati " & subDir.Path
        End If
        
        fileList.Clear()
        LogResults subDir.Path
    Next
End Function

ExploreFolder srcDir
LogResults(destDir)

If countFiles = 1 Then
    If countDirs = 1 Then
        WScript.Echo countFiles & " pilt sorteeritud " & countDirs & " kataloogi."
    Else
        WScript.Echo countFiles & " pildid sorteeritud " & countDirs & " kataloogidesse."
    End If
Else
    If countDirs = 1 Then
        WScript.Echo countFiles & " pilt sorteeritud " & countDirs & " kataloogi."
    Else
        WScript.Echo countFiles & " pildid sorteeritud " & countDirs & " kataloogidesse."
    End If
End If
