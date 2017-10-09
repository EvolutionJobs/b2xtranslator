Dim objFSO, objFile
Dim strSvnFolder, strEntriesFile, strRevNr, strConfigFolder, strConfigFile
Const ForReading = 1
Const ForWriting = 2

strSvnFolder = "..\..\.svn"
strEntriesFile = "\entries"
strConfigFolder = "..\.."
strConfigFile = "\revision.txt"

Wscript.Echo("Start")

' Create the File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Check that the .svn folder exists
If objFSO.FolderExists(strSvnFolder) Then
   
    'check if the entries file exists
    If objFSO.FileExists(strSvnFolder & strEntriesFile) Then
        
        Wscript.Echo("Reading svn revision number")
        
        'read revision number
        Set objFile = objFSO.OpenTextFile(strSvnFolder & strEntriesFile, ForReading)
        'revision number is saved in the fourth line
        strRevNr = objFile.ReadLine()
        strRevNr = objFile.ReadLine()
        strRevNr = objFile.ReadLine()
        strRevNr = objFile.ReadLine()
        objFile.Close()
        
        Wscript.Echo(strRevNr)
        
        Wscript.Echo("Writing revision.txt")
        
        'write revision number
        Set objFile = objFSO.OpenTextFile(strConfigFolder & strConfigFile, ForWriting, true)
        objFile.Write(strRevNr)
        objFile.Close()
            
        End If

End If