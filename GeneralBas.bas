Attribute VB_Name = "GeneralBas"
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public Function ReadIni(Seção As String, Chave As String, NomeArq As String)

    Dim LpAppName As String, lpKeyName As String
    Dim lpDefault As String, lpReturnedString As String
    Dim nSize As Integer, lpFileName As String, X As Integer
    
    lpFileName = NomeArq
    LpAppName = Seção
    lpKeyName = Chave
    lpDefault = ""
    lpReturnedString = Space$(512)
    nSize = 512
    X = GetPrivateProfileString(LpAppName, lpKeyName, lpDefault, lpReturnedString, nSize, lpFileName)
    ReadIni = Left$(lpReturnedString, X)

End Function
Public Sub WriteIni(Seção As String, Chave As String, Valor As String, NomeArq As String)

    Dim lpFileName As String, X As Integer
    lpFileName = NomeArq
    X = WritePrivateProfileString(Seção, Chave, Valor, lpFileName)

End Sub


Public Function GetTempDir() As String
   
   GetTempDir = String$(145, Chr$(0))
   GetTempDir = Left$(GetTempDir, GetTempPath(Len(GetTempDir), GetTempDir))

End Function


