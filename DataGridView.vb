Public Class DataGridView
    ''' <summary>
    ''' Note: This types of function can only filter the string data type.
    ''' </summary>
    ''' <param name="dgv">Name of your datagridview (object).</param>
    ''' <param name="strLookup">String that you want to filter.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Search(ByVal dgv As System.Windows.Forms.DataGridView, ByVal strLookup As String) As System.Windows.Forms.DataGridView
        Dim st As New System.Text.StringBuilder()
        Dim tmp As DataTable = DirectCast(dgv.DataSource, DataTable)
        For Each dgCol As DataGridViewColumn In dgv.Columns
            If dgCol.ValueType IsNot Nothing Then
                If dgCol.ValueType.FullName.ToString() = "System.String" Then
                    Dim b As Boolean = (st.ToString() = "")

                    If b Then
                        st.Append(String.Format("[{0}] LIKE '%{1}%' ", dgCol.Name, strLookup.Replace("'", "")))
                    Else
                        st.Append(String.Format(" or [{0}] LIKE '%{1}%' ", dgCol.Name, strLookup.Replace("'", "")))
                    End If
                End If
            End If
        Next
        tmp.DefaultView.RowFilter = st.ToString()
        dgv.ClearSelection()

        Return dgv
    End Function
End Class
