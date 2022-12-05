Imports System.Net
Module Computer
    Public Username As String = Environment.UserName
    Public CurrentDirectory As String = Environment.CurrentDirectory
    Public MachineName As String = Environment.MachineName
    Public Function GetIPAddress() As String
        GetIPAddress = String.Empty
        Dim strHostName As String = Dns.GetHostName()
        Dim iphe As IPHostEntry = Dns.GetHostEntry(strHostName)
        For Each ipheal As IPAddress In iphe.AddressList
            If ipheal.AddressFamily = Sockets.AddressFamily.InterNetwork Then
                GetIPAddress = ipheal.ToString
            End If
        Next
    End Function

End Module
