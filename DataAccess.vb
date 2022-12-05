Imports System.Data.SqlClient
Public Class DataAccess
    Private con As SqlConnection
    Private cmd As SqlCommand
    Private da As SqlDataAdapter
    Private dr As SqlDataReader
    Private tran As SqlTransaction
    Private dtable As DataTable

    Private _strQuery As String
    Public Property StrQuery() As String
        Get
            Return _strQuery
        End Get
        Set(ByVal value As String)
            _strQuery = value
        End Set
    End Property

    Private _strArrQuery As New ArrayList
    Public Property StrArrQuery() As ArrayList
        Get
            Return _strArrQuery
        End Get
        Set(ByVal value As ArrayList)
            _strArrQuery = value
        End Set
    End Property

    Private Function Connect(ByVal strConnection As String) As Boolean
        Try
            con = New SqlConnection(strConnection)
            con.Open()
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Connection failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Private Sub Disconnect()
        Try
            con.Close()
            con = Nothing
            cmd = Nothing
            dr = Nothing
            da = Nothing
            tran = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Disconnection failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Function ReadData(ByVal strConnection As String) As DataTable
        dtable = New DataTable
        If Connect(strConnection) Then
            Try
                If Not String.IsNullOrEmpty(StrQuery) Then
                    cmd = New SqlCommand(StrQuery, con)
                    da = New SqlDataAdapter(cmd)
                    da.Fill(dtable)
                Else
                    MessageBox.Show("StrQuery is empty.", "Error 404", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error 404", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
        Return dtable
        Disconnect()
    End Function

    Public Function TransactionData(ByVal strConnection As String) As Boolean
        If Connect(strConnection) Then
            tran = con.BeginTransaction
            Try
                If Not IsNothing(StrArrQuery) Then
                    For Each query As String In StrArrQuery
                        cmd = New SqlCommand(query, con, tran)
                        cmd.ExecuteNonQuery()
                    Next
                    tran.Commit()
                    Return True
                Else
                    MessageBox.Show("StrArrQuery is empty.", "Error 404", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Return False
                End If
            Catch ex As Exception
                tran.Rollback()
                MessageBox.Show(ex.Message, "Transaction failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
            Disconnect()
        Else
            Return False
        End If
    End Function

    Public Function CheckData(ByVal strConnection As String) As Boolean
        If Connect(strConnection) Then
            Try
                If Not String.IsNullOrEmpty(StrQuery) Then
                    cmd = New SqlCommand(StrQuery, con)
                    dr = cmd.ExecuteReader
                    If dr.HasRows Then
                        Return True
                    Else
                        Return False
                    End If
                Else
                    MessageBox.Show("StrQuery is empty.", "Error 404", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Return False
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error 404", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
            Disconnect()
        Else
            Return False
        End If

    End Function

    Public Function Merge(ByVal ParamArray dtable() As DataTable) As DataTable
        Try
            Dim table As New DataTable
            For Each table In dtable
                table.Merge(table)
            Next
            Return table
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error 404", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return New DataTable
        End Try
    End Function

    Public Function CountData(ByVal strConnection As String) As Integer
        If Connect(strConnection) Then
            Try
                If Not String.IsNullOrEmpty(StrQuery) Then
                    cmd = New SqlCommand(StrQuery, con)
                    CountData = DirectCast(cmd.ExecuteScalar(), Int32)
                Else
                    MessageBox.Show("StrQuery is empty.", "Error 404", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error 404", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            Disconnect()
        End If
        Return CountData
    End Function

End Class
