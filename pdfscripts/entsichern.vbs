Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objFSOEntsichert, objFolderEntsichert, strDirectoryEntsichert
Dim strPDFfile, strPDFfileEntsichert
Set objFSOEntsichert = CreateObject("Scripting.FileSystemObject")
objStartFolder = "..\entsichern"


dim filesys
dim newfolder
set filesys=CreateObject("Scripting.FileSystemObject")



dim fso
set fso = createobject("scripting.filesystemobject")

dim cmd

Set wshShell = WScript.CreateObject ("WSCript.shell")


Set objFolder = objFSO.GetFolder(objStartFolder)
Wscript.Echo objFolder.Name
Set colFiles = objFolder.Files
For Each objFile in colFiles
    If fso.GetExtensionName(objFile.Path) = "pdf" Then
        strPDFfile = Chr(34) & "..\entsichern\" & objFile.Name & Chr(34)
        strPDFfileEntsichert = Chr(34) & "..\entsichert\" & fso.getbasename(objFile.Path) & "_entsichert.pdf" & Chr(34)
        cmd = ".\mupdf\pdfclean.exe " & strPDFfile & " " & strPDFfileEntsichert
        wshshell.run cmd
        Wscript.Echo objFile.Name
        Wscript.Echo fso.getbasename(objFile.Path)
    End If
Next

ShowSubfolders objFSO.GetFolder(objStartFolder)

Sub ShowSubFolders(Folder)
    For Each Subfolder in Folder.SubFolders
        Wscript.Echo Subfolder.Name
        strDirectoryEntsichert = "..\entsichert\" &Subfolder.Name
        If  Not filesys.FolderExists(strDirectoryEntsichert) Then
             Set objFolderEntsichert = objFSOEntsichert.CreateFolder(strDirectoryEntsichert)
        End If

        Set objFolder = objFSO.GetFolder(Subfolder.Path)
        Set colFiles = objFolder.Files
        For Each objFile in colFiles
            If fso.GetExtensionName(objFile.Path) = "pdf" Then
                Wscript.Echo objFile.Name
                strPDFfile = Chr(34) & "..\entsichern\" & Subfolder.Name & "\" & objFile.Name & Chr(34)
                strPDFfileEntsichert = Chr(34) & "..\entsichert\" & Subfolder.Name & "\" & fso.getbasename(objFile.Path) & "_entsichert.pdf" & Chr(34)
                Wscript.Echo strPDFfile
                Wscript.Echo strPDFfileEntsichert
                cmd = ".\mupdf\pdfclean.exe " & strPDFfile & " " & strPDFfileEntsichert
'                wshshell.run ".\mupdf\pdfclean.exe " & strPDFfile & " " & strPDFfileEntsichert, 6, True
                Wscript.Echo cmd
                wshshell.run cmd
            End If
        Next
        ShowSubFolders Subfolder
    Next
End Sub
