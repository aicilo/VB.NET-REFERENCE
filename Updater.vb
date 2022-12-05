Imports System.IO
Public Class Updater

    Private _strSourcePath As String
    Public Property StrSourcePath() As String
        Get
            Return _strSourcePath
        End Get
        Set(ByVal value As String)
            _strSourcePath = value
        End Set
    End Property

    Private _strDestinationPath As String
    Public Property StrDestinationPath() As String
        Get
            Return _strDestinationPath
        End Get
        Set(ByVal value As String)
            _strDestinationPath = value
        End Set
    End Property

    Private _strApplicationName As String
    Public Property StrApplicationName() As String
        Get
            Return _strApplicationName
        End Get
        Set(ByVal value As String)
            _strApplicationName = value
        End Set
    End Property

    Private Sub NewUpdates(ByVal strSourcePath As String, ByVal strDestinationPath As String)
        Try
            Dim sourceDirectoryInfo As New DirectoryInfo(strSourcePath)
            If Not Directory.Exists(strDestinationPath) Then
                Directory.CreateDirectory(strDestinationPath)
            End If
            Dim fileSystemInfo As FileSystemInfo
            For Each fileSystemInfo In sourceDirectoryInfo.GetFileSystemInfos
                Dim destinationFilename As String = Path.Combine(strDestinationPath, fileSystemInfo.Name)
                If TypeOf fileSystemInfo Is FileInfo Then
                    File.Copy(fileSystemInfo.FullName, destinationFilename, True)
                Else
                    NewUpdates(fileSystemInfo.FullName, destinationFilename)
                End If
            Next
            MessageBox.Show("Updated Successfully.", "Software Updates", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.ToString, "Error detected", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    ''' <summary>
    ''' Requirement to fill
    ''' StrSourcePath > Location for updated file
    ''' StrDestinationPath > Location of current file
    ''' StrApplicationName > Exe file name
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CheckForUpdates()
        Dim updatedFile As Date = File.GetLastWriteTime(StrSourcePath & "\" & StrApplicationName & ".exe")
        Dim currentFile As Date = File.GetLastWriteTime(StrDestinationPath & "\" & StrApplicationName & ".exe")
        If updatedFile <> currentFile Then
            If MessageBox.Show("New updates is now available. Would you like to update?", "Software updates", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                NewUpdates(StrSourcePath, StrDestinationPath)
            End If
        Else
            MessageBox.Show("Up to date.", "Software Updates", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

End Class
