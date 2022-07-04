Option Compare Database
Option Explicit

'
' |~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~|
' |  מחלקת תקשורת                                                          |
' |  @MusiCode נבנה ע"י אבי המכונה                                         |
' |  מותר להשתמש, אך אין להעתיק ללא ציון קרדיט                           |
' |  MusiCode3@gmail.com                                                    |
' |  גרסה 3.0.0                                                             |
' |  המחלקה מאפשרת להעלות קבצים ב"פרום דטה", ואפ' קבצים בינאריים        |
' |  כמו כן, בעת הורדת הקובץ, יש פס זמן-ריצה, בשורת המצב                 |
' |                                                                         |
' |  ניתן לצרף באדי, בקלות ובנוחות                                        |
' |                                                                         |
' |~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~|
'

' HTTP Var 3.0.0
Private Const HTTP_VAR As Double = 3

'==== Consts ==============================================
' קבועים הקשורים לאובייקט
'  HTTP
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_SERVICE_HTTP = 3

Public Enum Ports
    INTERNET_DEFAULT_HTTP_PORT = 80 ' HTTP
    INTERNET_DEFAULT_HTTPS_PORT = 443 ' HTTPS
End Enum

' פלאגים
Public Enum Internet_Flags
    INTERNET_FLAG_KEEP_CONNECTION = &H400000
    INTERNET_FLAG_PRAGMA_NOCACHE = &H100
    INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
    
    INTERNET_FLAG_NO_AUTO_REDIRECT = &H200000 ' ללא הפניות מחדש 301/302
    INTERNET_FLAG_NO_COOKIES = &H80000 ' ללא שליחת קוקיס
End Enum
Private Const INTERNET_FLAG_RELOAD = &H80000000 ' HTTP
Private Const INTERNET_FLAG_SECURE = &H800000 ' HTTPS

' קבועם הקשורים לבקשת סטטוס או כותרות של
' HTTP
Public Enum HTTP_Query
    HTTP_QUERY_STATUS_CODE = 19
    HTTP_QUERY_STATUS_TEXT = 20
    HTTP_QUERY_RAW_HEADERS = 21
    HTTP_QUERY_RAW_HEADERS_CRLF = 22
    
    HTTP_QUERY_CONTENT_LENGTH = 5 'גידל המשאב בבתים
End Enum

Private Const bufSize = 128

' קבועים הקשורים לאוביקט הזרימה
Private Const adReadAll = -1
Private Const adReadLine = -2
Private Const adSaveCreateOverWrite = 2
Private Const adSaveCreateNotExist = 1
Private Const adTypeBinary = 1
Private Const adTypeText = 2
Private Const adWriteChar = 0
Private Const adWriteLine = 1

'==== Propertys ============================================
' מאפייני המחלקה
'
' ידיות חיבור
Private hSession As Long
Private hConnection As Long
Private hOpenRequest As Long

' האם התבצע חיבור?
Private connecting As Boolean

' פרטי הבקשה
Private qUrl As String
Private qHttps As Boolean ' האם השרת מאובטח
Private qServer As String ' שם השרת
Private qPath As String ' נתיב בתוך השרת
Private qHttp_method As String ' המתודה
Private qFlags As Long 'פלאגים לחיבור
Private qFlags_object As New Collection
Private qHttp_port As Long 'פורט החיבור
Private qUri_object As New Dictionary ' מערך פרמטרי שורת הכתובת
Private qUri_txt_final As String ' שורת הכתובת הסופית

' מערך הכותרות
Private qRequest_headers As New Dictionary
Private qRequest_headers_txt_final As String

' פרטי גוף הבקשה
Private bBody_object As New Dictionary ' אובייקט הגוף
Private bBody_txt As String ' הגוף, בטקסט
Private bBody_stream_final As Object ' אובייקט סטרים עם הגוף שיישלח למעשה

' פרטי התשובה
Private rResponse_headers_arr As Dictionary
Private rResponse_headers_str As String
Private rResponse_stream As Object

' הגדרות האובייקט
Private intWithe_Events As Integer
Private oShow_meter As Boolean

'==== WinAPI ================================================
#If VBA7 Then
    Private Declare PtrSafe Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" ( _
        ByVal lpszAgent As String, _
        ByVal dwAccessType As Long, _
        ByVal lpszProxyName As String, _
        ByVal lpszProxyBypass As String, _
        ByVal dwFlags As Long) As Long
        
    Private Declare PtrSafe Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" ( _
        ByVal hInternetSession As Long, _
        ByVal lpszServerName As String, _
        ByVal nServerPort As Long, _
        ByVal lpszUserName As String, _
        ByVal lpszPassword As String, _
        ByVal dwService As Long, _
        ByVal dwFlags As Long, _
        ByVal dwContext As Long) As Long
    
    Private Declare PtrSafe Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" ( _
        ByVal hHttpSession As Long, _
        ByVal lpszVerb As String, _
        ByVal lpszObjectName As String, _
        ByVal lpszVersion As String, _
        ByVal lpszReferer As String, _
        ByVal lpszAcceptTypes As String, _
        ByVal dwFlags As Long, _
        ByVal dwContext As Long) As Long
    
    Private Declare PtrSafe Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" ( _
        ByVal hHttpRequest As Long, _
        ByVal lpszHeaders As String, _
        ByVal dwHeadersLength As Long, _
        ByVal lpOptional As LongPtr, _
        ByVal dwOptionalLength As Long) As Boolean
    
    Private Declare PtrSafe Function InternetReadBinaryFile Lib "wininet.dll" Alias "InternetReadFile" ( _
        ByVal hFile As Long, _
        ByRef bytearray_firstelement As Byte, _
        ByVal lNumBytesToRead As Long, _
        ByRef lNumberOfBytesRead As Long) As Integer
        
    Private Declare PtrSafe Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, _
        ByVal lInfoLevel As Long, _
        ByRef sBuffer As Any, _
        ByRef lBufferLength As Long, _
        ByRef lIndex As Long) As Boolean
        
    Private Declare PtrSafe Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
    
    Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" ( _
        lpdwFlags As LongPtr, _
        ByVal dwReserved As Long) As Boolean
        
    Private Declare PtrSafe Function GetLastError Lib "Kernel32" () As Long
    
    Private Declare PtrSafe Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" _
        (MemoryDestination As Any, MemorySource As Any, ByVal ByteLength As Long)
    
    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As Long
    
    Private Declare PtrSafe Function SendMessageTimeout Lib "user32" _
        Alias "SendMessageTimeoutA" (ByVal hwnd As Long, _
        ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, _
        ByVal fuFlags As Long, ByVal uTimeout As Long, _
        pdwResult As Long) As Long
    
    Private Declare PtrSafe Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
  
#Else

    Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" ( _
        ByVal lpszAgent As String, _
        ByVal dwAccessType As Long, _
        ByVal lpszProxyName As String, _
        ByVal lpszProxyBypass As String, _
        ByVal dwFlags As Long) As Long
        
    Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" ( _
        ByVal hInternetSession As Long, _
        ByVal lpszServerName As String, _
        ByVal nServerPort As Integer, _
        ByVal lpszUserName As String, _
        ByVal lpszPassword As String, _
        ByVal dwService As Long, _
        ByVal dwFlags As Long, _
        ByVal dwContext As Long) As Long
    
    Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" ( _
        ByVal hHttpSession As Long, _
        ByVal lpszVerb As String, _
        ByVal lpszObjectName As String, _
        ByVal lpszVersion As String, _
        ByVal lpszReferer As String, _
        ByVal lpszAcceptTypes As String, _
        ByVal dwFlags As Long, _
        ByVal dwContext As Long) As Long
    
    Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" ( _
        ByVal hHttpRequest As Long, _
        ByVal lpszHeaders As String, _
        ByVal dwHeadersLength As Long, _
        ByVal lpOptional As LongPtr, _
        ByVal dwOptionalLength As Long) As Boolean
    
    Private Declare Function InternetReadBinaryFile Lib "wininet.dll" Alias "InternetReadFile" ( _
        ByVal hFile As Long, _
        ByRef bytearray_firstelement As Byte, _
        ByVal lNumBytesToRead As Long, _
        ByRef lNumberOfBytesRead As Long) As Integer
        
    Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, _
        ByVal lInfoLevel As Long, _
        ByRef sBuffer As Any, _
        ByRef lBufferLength As Long, _
        ByRef lIndex As Long) As Boolean
        
    Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
    
    Private Declare Function InternetGetConnectedState Lib "wininet.dll" ( _
        lpdwFlags As Long, _
        ByVal dwReserved As Long) As Boolean
        
    Private Declare Function GetLastError Lib "Kernel32" () As Long
    
    Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" _
        (MemoryDestination As Any, MemorySource As Any, ByVal ByteLength As Long)
    
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
    
    Private Declare Function SendMessageTimeout Lib "user32" _
        Alias "SendMessageTimeoutA" (ByVal hwnd As Long, _
        ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, _
        ByVal fuFlags As Long, ByVal uTimeout As Long, _
        pdwResult As Long) As Long
    
    Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
        
#End If

Private Sub Class_Initialize()

    If InternetGetConnectedState(0, 0) = False Then Err.Raise 30, , "אין חיבור לאינטרנט!"
    Withe_Events = 100
    connecting = False

End Sub

Private Sub Class_Terminate()

    Set rResponse_headers_arr = Nothing
    Set rResponse_stream = Nothing
    
End Sub

'========================================================================================
' גישה (קביעת/ קבלת) מאפיינים
'
Public Property Get get_response_binary() As String
        
    private_is_connecting
    
    rResponse_stream.Position = 0
    rResponse_stream.Type = adTypeBinary
    get_response_binary = rResponse_stream.Read

End Property

Public Property Get is_connecting()
    is_connecting = connecting
End Property

Public Property Get get_response(Optional charset As String = "utf-8") As String
    
    Dim BufferPosition As Long: BufferPosition = 0
    Dim BufferLength As Long: BufferLength = 0
    Dim length As Long
    
    private_is_connecting
        
    rResponse_stream.Position = 0
    rResponse_stream.Type = adTypeText
    rResponse_stream.charset = charset
    
    length = rResponse_stream.Size
    
    get_response = ""
    
    If oShow_meter Then SysCmd acSysCmdInitMeter, "העתקת הקובץ מתבצעת כעת...", length * 2
    
    Do
        
        BufferAppend get_response, rResponse_stream.ReadText(50), BufferPosition, BufferLength
        
        If rResponse_stream.Position Mod Withe_Events = 0 Then
            If oShow_meter Then SysCmd acSysCmdUpdateMeter, length + rResponse_stream.Position
            DoEvents
        End If

    Loop While Not rResponse_stream.EOS
    
    BufferToString get_response, BufferPosition, BufferLength
    
    If oShow_meter Then SysCmd acSysCmdRemoveMeter
    
End Property

Public Property Get Withe_Events() As Integer
    Withe_Events = intWithe_Events
End Property

Public Property Let Withe_Events(value As Integer)
    
    If value = 0 Then
        Err.Raise 33, , "Withe_Events צריך להיות גדול מ 1"
    ElseIf value > 500 Then
        Err.Raise 34, , "Withe_Events לא יכול לההות גדול מ ?500"
    End If
    
    intWithe_Events = value
End Property

Public Property Let Show_meter(value As Boolean)
    oShow_meter = value
End Property

Public Property Let url(value As String)
    
    qUrl = value
    
End Property

Public Property Let method(value As String)
    qHttp_method = value
End Property

Public Property Let set_header(key As String, value As String)
    qRequest_headers(key) = value
End Property

Public Property Let Set_Flags(value As Internet_Flags)
    qFlags_object.Add value
End Property

Public Property Let set_str_body(value As String)
    bBody_txt = value
End Property

Public Property Let set_obj_body(value As Dictionary)
    Set bBody_object = value
End Property

Public Property Let set_content_type(value As String)
    qRequest_headers("Content-Type") = value
End Property

Public Function get_header(header_name As String, Optional get_object As Boolean = False, _
                                                Optional sub_object As String = "") As Variant
    Dim zz
    
    private_is_connecting
    
    If Not rResponse_headers_arr.Exists(header_name) Then
        get_header = False
        Exit Function
    End If
    
    If VarType(rResponse_headers_arr(header_name)) <> vbString Then
    
        If Not sub_object = "" Then
            Err.Clear
            On Error Resume Next
            zz = rResponse_headers_arr(header_name)(sub_object)
            If Err.Number = 0 Then
                get_header = rResponse_headers_arr(header_name)(sub_object)
            Else
                get_header = False
            End If
            On Error GoTo 0
            Exit Function
        End If
    
        If get_object Then
            Set get_header = rResponse_headers_arr(header_name)
        Else
            get_header = rResponse_headers_arr(header_name)(0)
        End If
    Else
        get_header = rResponse_headers_arr(header_name)
    End If
    
End Function

Public Function get_all_headers() As String
    
    private_is_connecting
    get_all_headers = rResponse_headers_str
End Function

Public Sub save_response_to_file(file_name As String, Optional Overwrite_existing_file As Boolean = False)

    Select Case Overwrite_existing_file
        
        Case False
            rResponse_stream.SaveToFile file_name, adSaveCreateNotExist
        Case True
            rResponse_stream.SaveToFile file_name, adSaveCreateOverWrite
    End Select
    
End Sub

Public Sub add_body(key As String, value As String)
    bBody_object(key) = value
End Sub

Public Sub add_uri(key As String, value As String)
    qUri_object(key) = value
End Sub

'==========================================================================================================
'
' הוצאה לפועל של הבקשה
Public Sub send(Optional wait_after_execution As Boolean = True)
    
    Dim str_headers As String
    Dim bRes As Boolean
    
    ready_request
        
    computer_connected_to_Internet
    
    hSession = InternetOpen("HttpClass By AYbSystems", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, qFlags)
    hConnection = InternetConnect(hSession, qServer, qHttp_port, vbNullString, vbNullString, INTERNET_SERVICE_HTTP, 0, 0)
    hOpenRequest = HttpOpenRequest(hConnection, qHttp_method, qPath, "HTTP/1.0", vbNullString, vbNullString, qFlags, 0)
    
    If TypeName(bBody_stream_final) = "Stream" Then
    
        bRes = HttpSendRequest(hOpenRequest, qRequest_headers_txt_final, Len(qRequest_headers_txt_final), _
            StrPtr(bBody_stream_final.Read) + 3, bBody_stream_final.Size - 3) ' הפחתת שלושת הבתים הראשונים, זה דילוג על ה BOM
            
        DoEvents
        bBody_stream_final.Close
        Set bBody_stream_final = Nothing
    Else
        bRes = HttpSendRequest(hOpenRequest, qRequest_headers_txt_final, Len(qRequest_headers_txt_final), 0, 0)
    End If
    
    If bRes = False Then
        Exit Sub
    End If
    
    connecting = True
    get_headers
    
    If wait_after_execution Then
    
        WaitForResponse
       
    End If
  
End Sub

Public Sub WaitForResponse()
    
    Dim sBuffer() As Byte
    Dim lngDataReturned As Long
    Dim totalRead As Long
    Dim length As Variant
    
    ReDim sBuffer(bufSize)
    
    Set rResponse_stream = CreateObject("ADODB.Stream")
    
    rResponse_stream.Type = adTypeBinary
    rResponse_stream.Open
    
    length = rResponse_headers_arr("Content-Length")
    
    If length <> False And oShow_meter Then SysCmd acSysCmdInitMeter, "הורדת הקובץ מתבצעת כעת...", length * 2
    
    Do
        InternetReadBinaryFile hOpenRequest, sBuffer(0), UBound(sBuffer) - LBound(sBuffer), lngDataReturned
        If lngDataReturned = 0 Then Exit Do
            
        ReDim Preserve sBuffer(lngDataReturned - 1)
            
        rResponse_stream.Write sBuffer
        
        ReDim sBuffer(bufSize)
        
        totalRead = totalRead + lngDataReturned
                
        If totalRead Mod Withe_Events = 0 Then
            If length <> False And oShow_meter Then SysCmd acSysCmdUpdateMeter, totalRead
            DoEvents
        End If

    Loop While lngDataReturned <> 0
    
    If length <> False And oShow_meter Then SysCmd acSysCmdRemoveMeter
    
    close_con
    
End Sub

Private Sub close_con()
    
    qFlags = 0
    
    Call InternetCloseHandle(hConnection)
    Call InternetCloseHandle(hSession)
    Call InternetCloseHandle(hOpenRequest)
    
    Set bBody_object = Nothing
    Set qRequest_headers = Nothing
    Set qUri_object = Nothing
    
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Private Sub get_headers()

    Dim headers As String
    
    private_is_connecting
    
    headers = GetQueryInfo(hOpenRequest, HTTP_Query.HTTP_QUERY_RAW_HEADERS_CRLF)
    
    rResponse_headers_str = headers
    
    Set rResponse_headers_arr = parseHeaders(headers)
    
End Sub

Private Function GetQueryInfo(ByVal hHttpRequest As Long, ByVal iInfoLevel As HTTP_Query) As String
    Dim sBuffer As String * 1024
    Dim lBufferLength As Long
    Dim ErrNumber As Long

    lBufferLength = Len(sBuffer)
    HttpQueryInfo hHttpRequest, iInfoLevel, ByVal sBuffer, lBufferLength, 0
    ErrNumber = GetLastError
    GetQueryInfo = Left$(sBuffer, lBufferLength)
End Function

Private Function parseHeaders(headers As String) As Dictionary
    
    'On Error Resume Next
    
    Dim head As New Dictionary
    Dim a As Variant
    Dim b() As String
    Dim c() As String
    Dim d() As String
    
    Dim headers_arr As Variant
    
    headers = Replace(headers, """", "")
    
    headers_arr = Split(headers, vbCrLf)
    
    For Each a In headers_arr
    
        If a = "" Then GoTo continue
    
        b = Split(a, ":", 2)
        
        If UBound(b) > 0 Then
            
            c = Split(b(1), ";")
            d = Split(b(1), ",")
            
            If UBound(c) > 0 Then
                head.Add Trim(b(0)), explode_str_to_obj(b(1), ";", "=")
            ElseIf UBound(d) > 0 Then
                head.Add Trim(b(0)), explode_str_to_obj(b(1), ",", "=")
            Else
                head.Add Trim(b(0)), Trim(b(1))
            End If
        Else
        
            head.Add "Status", Split(a, " ")(1)
            head.Add "StatusText", Split(a, " ", 3)(2)
            
        End If
continue:
    Next
    
    Set parseHeaders = head
    
End Function

Private Sub private_is_connecting()
    If Not connecting Then
        Err.Raise 4554, , "לא התבצע עדיין חיבור!"
    End If
End Sub


'===============================================================================================
' הוספה למשתנה מהירה
Private Sub BufferAppend(ByRef buffer As String, _
                              ByRef Append As Variant, _
                              ByRef BufferPosition As Long, _
                              ByRef BufferLength As Long)
                              
    ' VBA can be slow to append strings due to allocating a new string for each append
    ' Instead of using the traditional append, allocate a large empty string and then copy string at append position
    '
    ' Example:
    ' Buffer: "abc  "
    ' Append: "def"
    ' Buffer Position: 3
    ' Buffer Length: 5
    '
    ' Buffer position + Append length > Buffer length -> Append chunk of blank space to buffer
    ' Buffer: "abc       "
    ' Buffer Length: 10
    '
    ' Copy memory for "def" into buffer at position 3 (0-based)
    ' Buffer: "abcdef    "
    '
    ' Approach based on cStringBuilder from vbAccelerator
    ' http://www.vbaccelerator.com/home/VB/Code/Techniques/RunTime_Debug_Tracing/VB6_Tracer_Utility_zip_cStringBuilder_cls.asp

    Dim AppendLength As Long
    Dim LengthPlusPosition As Long

    AppendLength = VBA.LenB(Append)
    LengthPlusPosition = AppendLength + BufferPosition

    If LengthPlusPosition > BufferLength Then
        ' Appending would overflow buffer, add chunks until buffer is long enough
        Dim TemporaryLength As Long

        TemporaryLength = BufferLength
        Do While TemporaryLength < LengthPlusPosition
            ' Initially, initialize string with 255 characters,
            ' then add large chunks (8192) after that
            '
            ' Size: # Characters x 2 bytes / character
            If TemporaryLength = 0 Then
                TemporaryLength = TemporaryLength + 510
            Else
                TemporaryLength = TemporaryLength + 16384
            End If
        Loop

        buffer = buffer & VBA.Space$((TemporaryLength - BufferLength) \ 2)
        BufferLength = TemporaryLength
    End If

    ' Copy memory from append to buffer at buffer position
    CopyMemory ByVal UnsignedAdd(StrPtr(buffer), _
                    BufferPosition), _
                    ByVal StrPtr(Append), _
                    AppendLength

    BufferPosition = BufferPosition + AppendLength

End Sub

Private Function BufferToString(ByRef buffer As String, ByVal BufferPosition As Long, ByVal BufferLength As Long) As String

    If BufferPosition > 0 Then
        buffer = VBA.Left$(buffer, BufferPosition \ 2)
    End If

End Function

#If VBA7 Then
Private Function UnsignedAdd(start As LongPtr, Increment As Long) As LongPtr
#Else
Private Function UnsignedAdd(start As Long, Increment As Long) As Long
#End If

    If start And &H80000000 Then
        UnsignedAdd = start + Increment
    ElseIf (start Or &H80000000) < -Increment Then
        UnsignedAdd = start + Increment
    Else
        UnsignedAdd = (start + &H80000000) + (Increment + &H80000000)
    End If
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


Private Function computer_connected_to_Internet() As Boolean

    Dim objPing As Object
    Dim objStatus As Object
    Dim Result As Integer
    Dim object_str As String
    Dim query_str As String
    Dim start As Double
    
    object_str = "winmgmts:{impersonationLevel=impersonate}"
    query_str = "Select * from Win32_PingStatus Where Address = '" & qServer & "'"
    
    Set objPing = GetObject(object_str).ExecQuery(query_str)
    
    For Each objStatus In objPing
        Result = objStatus.StatusCode
    Next
    
    If Result = 0 Then
        computer_connected_to_Internet = True
    Else
        Err.Raise 30, , "אין חיבור לאינטרנט!"
        computer_connected_to_Internet = False
    End If
    
End Function

'==========================================================================================================
'
' הכנת הבקשה

Private Sub ready_request()
    ready_url
    ready_uri
    ready_data
    ready_headers
    ready_flags
End Sub

Private Sub ready_flags()
    Dim flag
    For Each flag In qFlags_object
        qFlags = qFlags + flag
    Next
End Sub

Private Sub ready_txt_multipart_formData(data As String, Name, boundary As String)
    
    bBody_stream_final.WriteText String(2, "-") & boundary & vbCrLf
    bBody_stream_final.WriteText "Content-Disposition: form-data; name=""" & Name & """;"
    bBody_stream_final.WriteText vbCrLf & vbCrLf
    bBody_stream_final.WriteText data
    bBody_stream_final.WriteText vbCrLf

End Sub

Private Sub ready_file_multipart_formData(data As file, Name, boundary As String)
    
    bBody_stream_final.WriteText String(2, "-") & boundary & vbCrLf
    bBody_stream_final.WriteText "Content-Disposition: form-data;"
    bBody_stream_final.WriteText " name=" & Chr(34) & "fileUpload" & Chr(34) & ";"
    bBody_stream_final.WriteText " filename=" & Chr(34) & Name & Chr(34)
    bBody_stream_final.WriteText "Content-Type: application/upload" & vbCrLf & vbCrLf
    
    If TypeName(data.data) <> "String" Then
    
        bBody_stream_final.Position = 0
        bBody_stream_final.Type = adTypeBinary
        bBody_stream_final.Position = bBody_stream_final.Size
        
        bBody_stream_final.Write data.data
        
        bBody_stream_final.Position = 0
        bBody_stream_final.Type = adTypeText
        bBody_stream_final.Position = bBody_stream_final.Size
    Else
        bBody_stream_final.WriteText data.data
    End If
    
    bBody_stream_final.WriteText vbCrLf

End Sub

Private Sub ready_uri()

    If qUri_object.count > 0 Then
    
        qUri_txt_final = imlode_obj_to_str(qUri_object, "&", "=")
        
        qPath = qPath & "?" & qUri_txt_final
    End If
    
End Sub

Private Sub ready_data_application_x_www_form_urlencoded()
    
    bBody_stream_final.WriteText imlode_obj_to_str(bBody_object, "&", "=")
    
End Sub

Private Sub ready_data_application_json()
    bBody_stream_final.WriteText JsonConverter.ConvertToJson(bBody_object)
End Sub

Private Sub ready_data_multipart_formData()
    
    Dim sName As Variant
    Dim boundary As String
    
    boundary = String(4, "-") & StringToMD5Hex(Now & "Salt")
    
    qRequest_headers("Content-Type") = "multipart/form-data; boundary=" & boundary
    
    For Each sName In bBody_object.Keys
    
        If TypeName(bBody_object(sName)) = "file" Then
            
            ready_file_multipart_formData bBody_object(sName), sName, boundary
            
        Else
            
            ready_txt_multipart_formData bBody_object(sName), sName, boundary
            
        End If
        
    Next
    
    bBody_stream_final.WriteText String(2, "-") & boundary & String(2, "-")
    
End Sub

Private Sub ready_data()
     
    If qHttp_method = "" And bBody_object.count > 0 Or bBody_txt <> "" Then
        qHttp_method = "POST"
    ElseIf qHttp_method = "" Then
        qHttp_method = "GET"
    End If
    
    If qHttp_method = "GET" Or bBody_object.count < 1 Then Exit Sub
    
    If bBody_txt = "" Then
    
        If Not qRequest_headers.Exists("Content-Type") Then _
            qRequest_headers("Content-Type") = "application/x-www-form-urlencoded"
            
        Set bBody_stream_final = CreateObject("ADODB.Stream")
        
        bBody_stream_final.Type = adTypeText
        bBody_stream_final.charset = "UTF-8" ' "Windows-1255"
        bBody_stream_final.Open

        Select Case qRequest_headers("Content-Type")
        
        Case "multipart/form-data"
            ready_data_multipart_formData
                
        Case "application/json"
            ready_data_application_json
            
        Case "application/x-www-form-urlencoded"
                ready_data_application_x_www_form_urlencoded
                
        End Select
            
    ElseIf bBody_txt <> "" Then
        bBody_stream_final.WriteText bBody_txt
    End If
    
    bBody_stream_final.Position = 0
    bBody_stream_final.Type = adTypeBinary
    
End Sub

Private Function ready_url() As Boolean
    
    Dim http As String
    Dim custom_port() As String
    Dim value_arr() As String
    Dim value As String
    
    value = qUrl
    
    http = Split(value, ":")(0)
    
    http = UCase(http)
    
    If http = "HTTP" Then
    
        qHttps = False
        qFlags = qFlags + INTERNET_FLAG_RELOAD
        qHttp_port = Ports.INTERNET_DEFAULT_HTTP_PORT
        
    ElseIf http = "HTTPS" Then
    
        qHttps = True
        qFlags = qFlags + INTERNET_FLAG_SECURE
        qHttp_port = Ports.INTERNET_DEFAULT_HTTPS_PORT
        
    Else
        Err.Raise 1, , "URL ???? ????"
    End If
    
    value = Split(value, "//", 2)(1)
    
    value_arr = Split(value, "/", 2)
    
    qServer = value_arr(0)
    qPath = value_arr(1)
    
    custom_port = Split(qServer, ":", 2)
    
    If UBound(custom_port) > 0 Then
        qHttp_port = custom_port(1)
        qServer = custom_port(0)
    End If
End Function

Private Sub ready_headers()
    
    Dim str_headers As String
    Dim i As Long
    Dim key As Variant
    
    str_headers = ""
    
    i = 1
    
    For Each key In qRequest_headers.Keys
    
        str_headers = str_headers & key & ": " _
            & qRequest_headers(key)
            
        If i > qRequest_headers.count Then str_headers = str_headers & vbCrLf
        i = i + 1
            
    Next
    
    qRequest_headers_txt_final = str_headers
    
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'==========================================================================================================
'
' פונקציות עזר
Private Function explode_str_to_obj(ByVal str_txt As String, Optional delimiter_val_and_key As String = "&", Optional delimiter_key_and_val As String = "=") As Dictionary
    Dim a() As String
    Dim b, c() As String
    Dim i As Integer
    Dim rows As New Dictionary
    a = Split(str_txt, delimiter_val_and_key)
    i = 0
    For Each b In a
        c = Split(b, delimiter_key_and_val)
        If UBound(c) > 0 Then
            rows(Trim(c(0))) = Trim(c(1))
        Else
            rows(i) = Trim(b)
            i = i + 1
        End If
    Next
    Set explode_str_to_obj = rows
End Function

Private Function imlode_obj_to_str(body As Dictionary, Optional delimiter_val_and_key As String = "&", Optional delimiter_key_and_val As String = "=") As String
    
    Dim i As Long
    Dim b As Variant
    
    If body.count = 0 Then Exit Function
    
    ReDim a(body.count - 1) As String
    i = 0
    For Each b In body.Keys
        a(i) = b & delimiter_key_and_val & body(b)
        i = i + 1
    Next
    imlode_obj_to_str = Join(a, delimiter_val_and_key)
End Function

Private Function StringToMD5Hex(ByVal s As String) As String

    Dim enc As Object
    Dim Bytes() As Byte
    Dim Pos As Long
    Dim outstr As String
    
    Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    
    Bytes = StrConv(s, vbFromUnicode)
    Bytes = enc.ComputeHash_2(Bytes)
    
    For Pos = 1 To UBound(Bytes) + 1
       outstr = outstr & LCase( _
                        Right("0" & Hex( _
                            AscB( _
                            MidB(Bytes, Pos, 1))), 2))
    Next Pos
    
    StringToMD5Hex = outstr
    Set enc = Nothing
    
End Function

'Private Function Program_is_responding() As Boolean
'
'    Const SMTO_BLOCK = &H1
'    Const SMTO_ABORTIFHUNG = &H2
'    Const WM_NULL = &H0
'    Const WM_CLOSE = &H10
'    Const PROCESS_ALL_ACCESS = &H1F0FFF
'
'    Dim lReturn As Long
'    Dim lResult As Long
'    Dim lHwnd As Long
'
'    lHwnd = GetForegroundWindow
'
'
'    lReturn = SendMessageTimeout(lHwnd, WM_NULL, 0&, 0&, SMTO_ABORTIFHUNG And SMTO_BLOCK, 1, lResult)
'
'    Call CloseHandle(lHwnd)
'
'    If lReturn Then
'        Program_is_responding = True
'    Else
'        Program_is_responding = False
'    End If
'
'End Function


