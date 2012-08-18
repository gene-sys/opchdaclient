Attribute VB_Name = "INI"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpapplicationname As String, ByVal lpkeyname As Any, ByVal lpdefault As String, _
        ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpapplicationname As Any, ByVal lpkeyname As Any, _
        ByVal lpString As Any, ByVal lpfilename As String) As Long

Public Sub WriteINI(sINIFile As String, sSection As String, sKey As String, sValue As String)
    Dim N As Integer
    Dim sTemp  As String
    sTemp = sValue
    For N = 1 To Len(sValue)
        If Mid$(sValue, N, 1) = vbCr Or Mid$(sValue, N, 1) = vbLf Then Mid$(sValue, N) = " "
    Next N
    N = WritePrivateProfileString(sSection, sKey, sTemp, sINIFile)
End Sub
Public Function GetINI(sINIFile As String, sSection As String, sKey As String, sdefault As String) As String
    Dim sTemp  As String * 256
    Dim nLength As Integer
    sTemp = Space$(256)
    nLength = GetPrivateProfileString(sSection, sKey, sdefault, sTemp, 255, sINIFile)
    GetINI = Left$(sTemp, nLength)
End Function


