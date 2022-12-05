Imports Microsoft.Office.Interop.Excel

Public Class Excel

    Private _xlsApplication As Application
    Private _xlsWorkBook As Workbook
    Private _xlsWorkSheet As Worksheet

    Public Sub ExportDatatableToExcel(ByVal dtable As System.Data.DataTable)
        Try
            _xlsApplication = New Application

            _xlsWorkBook = _xlsApplication.Workbooks.Add()
            _xlsWorkSheet = _xlsWorkBook.ActiveSheet()

            Dim dtTemp As System.Data.DataTable = dtable
            Dim dcTemp As System.Data.DataColumn
            Dim drTemp As System.Data.DataRow
            Dim colIndex As Integer = 0
            Dim rowIndex As Integer = 0

            For Each dcTemp In dtTemp.Columns
                colIndex = colIndex + 1
                _xlsApplication.Cells(1, colIndex) = dcTemp.ColumnName
            Next
            For Each drTemp In dtTemp.Rows
                rowIndex = rowIndex + 1
                colIndex = 0
                For Each dcTemp In dtTemp.Columns
                    colIndex = colIndex + 1
                    _xlsApplication.Cells(rowIndex + 1, colIndex) = drTemp(dcTemp.ColumnName)
                Next
            Next
            _xlsWorkSheet.Columns.AutoFit()

            Dim sfd As New SaveFileDialog
            sfd.Title = "Export Data"
            sfd.Filter = "Excel Workbook|*.xlsx"
            sfd.FileName = "C:\" & "data_" & DateTime.Now.ToString("yyyyMMdd")
            If sfd.ShowDialog = DialogResult.OK Then
                _xlsWorkBook.SaveAs(sfd.FileName)
                _xlsWorkBook.Close()
            Else
                _xlsWorkBook.Close(SaveChanges:=False)
            End If
            
        Catch ex As Exception
            MessageBox.Show(ex.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            _xlsWorkBook.Close()
        End Try
        _xlsApplication.Quit()
    End Sub


    Public Function ImportExcelToDatable(Optional ByVal strPath As String = "") As System.Data.DataTable
        ImportExcelToDatable = New System.Data.DataTable
        If strPath = String.Empty Then
            Dim opf As New OpenFileDialog
            opf.Filter = "Excel Workbook| *.xlsx"
            If opf.ShowDialog = DialogResult.OK Then
                strPath = opf.FileName
            End If
        End If
        Try
            _xlsWorkBook = _xlsApplication.Workbooks.Add(strPath)
            _xlsWorkSheet = _xlsWorkBook.ActiveSheet()
            Dim rows As Integer = _xlsWorkSheet.UsedRange.Rows.Count
            Dim cols As Integer = _xlsWorkSheet.UsedRange.Columns.Count

            Dim noOfRow As Integer = 1

            For colIndex As Integer = 1 To cols
                Dim colName As String = _xlsWorkSheet.Cells(1, colIndex).Text
                ImportExcelToDatable.Columns.Add(colName)
                noOfRow = 2
            Next

            For rowIndex As Integer = noOfRow To rows
                Dim drr As DataRow = ImportExcelToDatable.NewRow()
                For column As Integer = 1 To cols
                    If column = 11 Then
                        Dim rng As Range = _xlsWorkSheet.Cells(rowIndex, column)
                        drr(column - 1) = rng.Text
                    Else
                        drr(column - 1) = _xlsWorkSheet.Cells(rowIndex, column).Value
                    End If
                Next
                ImportExcelToDatable.Rows.Add(drr)
            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        _xlsWorkBook.Close()
        _xlsApplication.Quit()
    End Function

End Class
