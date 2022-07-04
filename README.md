# VBA-HTTP

ספרייה ליצירת בקשות HTTPS מיישום MS Access.

```vba
Sub SaveToFile()

    Dim http1 As New http
    Dim file_name As String
    
    
    http1.url = "http://forum.enativ.com/filebase.php?d=1&id=816&f=816&what=c&c_old=6&page=1"

    
    http1.send False
    
    http1.WaitForResponse
    
    file_name = "E:\User\Downloads\" 'Application.CurrentProject.path & "\"
    
    file_name = file_name & http1.get_header("Content-Disposition")("filename")
    

    http1.save_response_to_file file_name, adSaveCreateOverWrite

End Sub
```
