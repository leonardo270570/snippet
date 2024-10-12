Sub GetAllDirectories()
    Dim objFso As Object
    Dim objFolder As Object
    Dim objSubFolder As Object
    Dim i As Integer
    
    'Create an instance of the FileSystemObject
    Set objFso = CreateObject("Scripting.FileSystemObject")
    
    'Get the folder object
    m_mainfolder = "H:\Project\r\prolife-actuarial\gpvreserve\assumption\"
    Set objFolder = objFso.GetFolder(m_mainfolder)
    
    i = 1
    'loops through each file in the directory and prints their names and path
    For Each objSubFolder In objFolder.subfolders
    
        Set objChildFolder = objFso.GetFolder(m_mainfolder & objSubFolder.Name & "\")
        
        For Each objFile In objChildFolder.Files
            'print folder name
            Debug.Print objFile.Name
            'print folder path
            Debug.Print objFile.Path
            i = i + 1
        Next objFile

        i = i + 1
    Next objSubFolder
    
    For Each objSubFolder In objFolder.Files
        'print folder name
        Debug.Print objSubFolder.Name
        'print folder path
        Debug.Print objSubFolder.Path
        i = i + 1
    Next objSubFolder

End Sub
