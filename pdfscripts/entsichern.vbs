Dim objFSO
Dim objFSOEntsichert
Dim objFolderEntsichert
Dim strDirectoryEntsichert
Dim strPDFfile
Dim strPDFfileEntsichert
dim filesys
dim newfolder
dim fso
dim cmd
Dim objEntsichernFolder
Dim objEntsichertFolder
 
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFSOEntsichert = CreateObject("Scripting.FileSystemObject")
set filesys = CreateObject("Scripting.FileSystemObject")
set fso = createobject("scripting.filesystemobject")
Set wshShell = WScript.CreateObject ("WSCript.shell")
 
objEntsichernFolder = "..\entsichern\"
objEntsichertFolder = "..\entsichert\"
 
 
 
 
 
Set objFolder = objFSO.GetFolder(objEntsichernFolder)
'Wscript.Echo objFolder.Name
Set colFiles = objFolder.Files
For Each objFile in colFiles
    If fso.GetExtensionName(objFile.Path) = "pdf" Then
        strPDFfile = Chr(34) & objEntsichernFolder & objFile.Name & Chr(34)
        strPDFfileEntsichert = Chr(34) & objEntsichertFolder & fso.getbasename(objFile.Path) & "_entsichert.pdf" & Chr(34)
        cmd = ".\mupdf\pdfclean.exe " & strPDFfile & " " & strPDFfileEntsichert
        wshshell.run cmd
'        Wscript.Echo objFile.Name
'        Wscript.Echo fso.getbasename(objFile.Path)
    End If
Next
 
 
ShowSubfolders objFSO.GetFolder(objEntsichernFolder)
 
 
 
Sub ShowSubFolders(Folder)
    For Each Subfolder in Folder.SubFolders
'        Wscript.Echo Subfolder.Name
        strDirectoryEntsichert = objEntsichertFolder & Subfolder.Name
        If  Not filesys.FolderExists(strDirectoryEntsichert) Then
             Set objFolderEntsichert = objFSOEntsichert.CreateFolder(strDirectoryEntsichert)
        End If
 
 
        Set objFolder = objFSO.GetFolder(Subfolder.Path)
        Set colFiles = objFolder.Files
        For Each objFile in colFiles
            If fso.GetExtensionName(objFile.Path) = "pdf" Then
'                Wscript.Echo objFile.Name
                strPDFfile = Chr(34) & objEntsichernFolder & Subfolder.Name & "\" & objFile.Name & Chr(34)
                strPDFfileEntsichert = Chr(34) & objEntsichertFolder & Subfolder.Name & "\" & fso.getbasename(objFile.Path) & "_entsichert.pdf" & Chr(34)
'                Wscript.Echo strPDFfile
'                Wscript.Echo strPDFfileEntsichert
                cmd = ".\mupdf\pdfclean.exe " & strPDFfile & " " & strPDFfileEntsichert
'                wshshell.run ".\mupdf\pdfclean.exe " & strPDFfile & " " & strPDFfileEntsichert, 6, True
'                Wscript.Echo cmd
                wshshell.run cmd
            End If
        Next
        ShowSubFolders Subfolder
    Next
End Sub
