Imports System.Runtime.InteropServices
Public Class TextBox
    Private collectData As AutoCompleteStringCollection

    Public Sub SuggestData(ByVal dtable As DataTable, ByVal txtData As System.Windows.Forms.TextBox, ByVal ParamArray columnArg() As Object)
        collectData = New AutoCompleteStringCollection
        For i = 0 To dtable.Rows.Count - 1
            Application.DoEvents()
            For Each param As Object In columnArg
                collectData.Add(dtable.Rows(i).Item(param).ToString.Trim)
            Next
        Next
        With txtData
            .AutoCompleteSource = AutoCompleteSource.CustomSource
            .AutoCompleteCustomSource = collectData
            .AutoCompleteMode = AutoCompleteMode.Suggest
        End With
    End Sub

    <DllImport("user32.dll", EntryPoint:="SendMessage")> _
    Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As Integer, <MarshalAs(UnmanagedType.LPWStr)> IParam As String) As Int32
    End Function
    Private Const EM_SETCUEBANNER As Integer = &H1501 'for TextBox
    Private Const CB_SETCUEBANNER As Integer = &H1703 ''for ComboBox
    Public Sub WaterMark(ByVal strMark As String, ByVal txtBox As System.Windows.Forms.TextBox)
        SendMessage(txtBox.Handle, EM_SETCUEBANNER, 1, strMark)
    End Sub

End Class
