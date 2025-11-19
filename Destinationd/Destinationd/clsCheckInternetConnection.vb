
Public Class clsCheckInternetConnection

#Region "InternetGetConnectedStateEx function"
    Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" Alias "InternetGetConnectedStateExA" _
      (ByRef lpdwFlags As Long, _
       ByVal lpszConnectionName As String, _
       ByVal dwNameLen As Long, ByVal dwReserved As Long) As Boolean

    ''''''''''''CONSTANT's'''''''''''''''
    'Internet connection VIA Proxy server.
    Public Const ProxyConnection As Integer = &H4S
    'Modem is busy.
    Public Const ModemConnectionIsBusy As Integer = &H8S
    'Internet connection is currently Offline
    Public Const InternetIsOffline As Integer = &H20S
    'Internet connection is currently configured
    Public Const InternetConnectionIsConfigured As Integer = &H40S
    'Internet connection VIA Modem.
    Public Const ModemConnection As Integer = &H1S
    'Remote Access Server is installed.
    Public Const RasInstalled As Integer = &H10S
    'Internet connection VIA LAN.
    Public Const LanConnection As Integer = &H2S

    ''''''''''''''VARIABLES'''''''''''''
    Public CntType As Long, CntName As String = Space(50), CntNameLen As Long = 50
    Dim Status As Boolean
    '''''''''''''''FUNCTIONS'''''''''''''''
    Public Function IsConnected() As Boolean
        'Returns true if there is any internet connection.
        IsConnected = InternetGetConnectedStateEx(CntType, CntName, CntNameLen, 0)
    End Function

    Public Function IsLanConnection() As Boolean
        Dim dwflags As Integer
        'return True if LAN connection
        Status = InternetGetConnectedStateEx(CntType, CntName, CntNameLen, 0)
        IsLanConnection = CntType And LanConnection And Status
    End Function

    Public Function IsModemConnection() As Boolean
        Dim dwflags As Integer

        'return True if modem connection.
        Status = InternetGetConnectedStateEx(CntType, CntName, CntNameLen, 0)
        IsModemConnection = CntType And ModemConnection And Status
    End Function

    Public Function IsProxyConnection() As Boolean
        Dim dwflags As Integer
        'return True if connected through a proxy.
        Call InternetGetConnectedStateEx(CntType, CntName, CntNameLen, 0)
        IsProxyConnection = CntType And ProxyConnection
    End Function

    Public Function IsRasInstalled() As Boolean
        Dim dwflags As Integer
        'return True if RAS installed.
        Call InternetGetConnectedStateEx(CntType, CntName, CntNameLen, 0)
        IsRasInstalled = CntType And RasInstalled
    End Function
#End Region

End Class