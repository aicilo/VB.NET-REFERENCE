Imports System.DirectoryServices
Imports System.DirectoryServices.AccountManagement
Public Class Domain

    Public Function ValidateLogin(ByVal path As String, ByVal username As String, ByVal password As String) As Boolean
        Dim Success As Boolean
        Dim Entry As New DirectoryEntry(path, username, password)
        'Dim Entry As New DirectoryEntry("LDAP://pki.com.ph", username, password)

        Dim Searcher As New DirectorySearcher(Entry)
        Searcher.SearchScope = DirectoryServices.SearchScope.OneLevel
        Try
            Dim Results As SearchResult = Searcher.FindOne
            Success = Not (Results Is Nothing)
        Catch ex As Exception
            Success = False
        End Try

        Entry.Dispose()
        Searcher.Dispose()

        Return Success

    End Function
End Class
