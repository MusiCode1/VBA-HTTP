# VBA-HTTP

ספרייה ליצירת בקשות HTTPS מיישום MS Access.

```vba
Sub SaveToFile()

    Dim http1 As New http
    Dim file_name As String
    
    
    http1.url = "https://file-examples.com/storage/fee9d866f762c2ea69ae2a5/2017/04/file_example_MP4_1280_10MG.mp4"

    
    http1.send False
    
    http1.WaitForResponse
    
    file_name = "C:\User\Downloads\"
    
    file_name = file_name & http1.get_header("Content-Disposition")("filename")
    

    http1.save_response_to_file file_name, adSaveCreateOverWrite

End Sub
```
