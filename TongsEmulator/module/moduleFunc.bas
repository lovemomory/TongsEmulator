Attribute VB_Name = "moduleFunc"
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

' 현재 경로 구하기
Public Function AP() As String
    AP = App.Path
    If Not Right(AP, 1) = "\" Then AP = AP & "\"
End Function

Public Sub Log(str As String, Optional Level As Long = 2)
    formMain.textLog.SelStart = Len(formMain.textLog.Text)
    formMain.textLog.SelText = vbCrLf & "[" & Now & "] " & str & vbCrLf
    formMain.textLog.SelStart = Len(formMain.textLog.Text)
End Sub

Public Function HexToStr(data() As Byte)
    Dim str As String
    Dim i As Long
    For i = 0 To UBound(data)
         str = str & Right("00" & Hex(data(i)), 2) & " "
    Next i
    
    HexToStr = str
End Function

Public Sub WriteINI(section As String, Key As String, value As String, filename As String)
    Call WritePrivateProfileString(section, Key, value, filename)
End Sub

Public Function ReadINI(section As String, Key As String, filename As String, defvalue As String) As String
    Dim StringBuffer As String
    Dim StringBufferSize As Integer

    StringBuffer = Space$(255)
    StringBufferSize = Len(StringBuffer)

    StringBufferSize = GetPrivateProfileString(section, Key, defvalue, StringBuffer, StringBufferSize, filename)

    If StringBufferSize > 0 Then
        ReadINI = Left$(StringBuffer, StringBufferSize)
    Else
        ReadINI = defvalue
    End If
    
    If InStr(ReadINI, vbNullChar) Then ReadINI = Left$(ReadINI, InStr(ReadINI, vbNullChar) - 1)
End Function

Public Function RemoveNullStr(str As String) As String
    If InStr(str, vbNullChar) Then str = Left$(str, InStr(str, vbNullChar) - 1)
    RemoveNullStr = str
End Function
