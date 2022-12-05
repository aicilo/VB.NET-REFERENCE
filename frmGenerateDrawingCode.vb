Imports Microsoft.Office.Interop

Public Class frmGenerateDrawingCode


    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Dispose()
    End Sub



    Private _dtBarcode As New DataTable("Selected")
    Private Sub GenerateDrawingCode_Load(sender As Object, e As EventArgs) Handles Me.Load
        _dtBarcode.Columns.Add("Drawing Code", GetType(String))
        _dtBarcode.Columns.Add("Product Number", GetType(String))
        _dtBarcode.Columns.Add("Settsu Number", GetType(String))
        _dtBarcode.Columns.Add("Frame Number", GetType(String))

        Dim drawing As New ControlOfCustomerDrawing
        drawing.DGV = dgvDrawingInfo
        drawing.LoadDrawingInfoForGeneratingBarcode()

        Dim dataAccess As New DataAccess
        dataAccess.StrQuery = "SELECT [Product_Number] FROM [tbl_Drawing];"
        Dim productList As DataTable = dataAccess.ReadData(My.Settings.MainDatabase)
        Dim dataCollection As New QAManagementSystem.TextBox
        dataCollection.SuggestData(productList, txtSearch, "Product_Number")

        Dim waterMark As New QAManagementSystem.TextBox
        waterMark.WaterMark("Search Product Number", txtSearch)

    End Sub

    Structure Data
        Dim barcode As String
        Dim productno As String
        Dim settsu As String
        Dim frameno As String
    End Structure

    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click

        Dim list As New List(Of Data)
        Dim count As Integer = 0

        For Each dr As DataRow In _dtBarcode.Rows
            Dim data As New Data
            data.barcode = dr.Item("Drawing Code").ToString
            data.productno = dr.Item("Product Number").ToString
            data.settsu = dr.Item("Settsu Number").ToString
            data.frameno = dr.Item("Frame Number").ToString
            list.Add(data)
        Next

        Dim xlApp As Excel.Application = New Excel.Application
        Dim xlwb As Excel.Workbook = xlApp.Workbooks.Add()
        Dim xlws As Excel.Worksheet = xlwb.Worksheets(1)
        xlws.Select()
        If _dtBarcode.Rows.Count >= 1 Then
            If MessageBox.Show("Barcode generate to excel file.", "Manual Printing", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = Windows.Forms.DialogResult.OK Then
                xlApp.Visible = True
            End If
        Else
            MessageBox.Show("Please select the product number that you want to generate. Please try again.", " Generate Barcode", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
        xlws.PageSetup.TopMargin = 0.5
        xlws.PageSetup.LeftMargin = 0.5
        xlws.PageSetup.RightMargin = 0
        xlws.PageSetup.BottomMargin = 0
        xlws.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4
        xlws.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape

        xlApp.ActiveWindow.View = Excel.XlWindowView.xlPageLayoutView

        Dim y As Integer = 1
        Dim x As Integer = 1

        For Each lst As Data In list
            With xlws
                .Cells(y, x) = lst.productno : .Range(.Cells(y, x), .Cells(y, x)).Font.Name = "Consolas" : .Range(.Cells(y, x), .Cells(y, x)).Font.Size = 20 : .Range(.Cells(y, x), .Cells(y, x)).Font.FontStyle = "Bold" : .Range(.Cells(y, x), .Cells(y, x)).ColumnWidth = 35 : .Range(.Cells(y, x), .Cells(y, x)).HorizontalAlignment = Excel.Constants.xlCenter
                .Cells(y + 1, x) = "*" & lst.barcode & "*" : .Range(.Cells(y + 1, x), .Cells(y + 1, x)).Font.Name = "3 of 9 Barcode" : .Range(.Cells(y + 1, x), .Cells(y + 1, x)).Font.Size = 24 : .Range(.Cells(y, x), .Cells(y, x)).ColumnWidth = 35 : .Range(.Cells(y + 1, x), .Cells(y + 1, x)).HorizontalAlignment = Excel.Constants.xlCenter
                .Cells(y + 2, x) = lst.settsu & "-" & lst.frameno : .Range(.Cells(y + 2, x), .Cells(y + 2, x)).Font.Name = "Consolas" : .Range(.Cells(y + 2, x), .Cells(y + 2, x)).Font.Size = 11 : .Range(.Cells(y, x), .Cells(y, x)).ColumnWidth = 35 : .Range(.Cells(y + 2, x), .Cells(y + 2, x)).HorizontalAlignment = Excel.Constants.xlCenter

                .Range(xlws.Cells(y, x), xlws.Cells(y + 2, x)).BorderAround2(LineStyle:=Excel.XlLineStyle.xlDash, Weight:=Excel.XlBorderWeight.xlThick, ColorIndex:=0)

            End With
            If x = 4 Then
                x = 1
                y += 4
            Else
                x += 1
            End If

            Application.DoEvents()

        Next



    End Sub

    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged
        Dim search As New Search
        search.GridSearch(dgvDrawingInfo, txtSearch.Text)

        For Each row As DataGridViewRow In dgvDrawingInfo.Rows
            Dim chk As DataGridViewCheckBoxCell = row.Cells(0)
            For Each dr As DataRow In _dtBarcode.Rows
                If dgvDrawingInfo.Rows(chk.RowIndex).Cells(1).Value = dr.Item(0).ToString Then
                    dgvDrawingInfo.Rows(chk.RowIndex).Cells(0).Value = True
                End If
            Next

        Next
    End Sub


    Private Sub dgvDrawingInfo_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvDrawingInfo.CellMouseClick
        If e.ColumnIndex <> 0 And e.RowIndex >= 0 Then
            If dgvDrawingInfo.Rows(e.RowIndex).Cells(0).Value = False Then
                dgvDrawingInfo.Rows(e.RowIndex).Cells(0).Value = True
            Else
                dgvDrawingInfo.Rows(e.RowIndex).Cells(0).Value = False
            End If

            If (dgvDrawingInfo.Rows.Count > 0) Then
                If (dgvDrawingInfo.Rows(e.RowIndex).Cells(0).Value = True) Then
                    Dim dr = _dtBarcode.Select("`Drawing Code` = '" & dgvDrawingInfo.Rows(e.RowIndex).Cells("Drawing Code").Value.ToString() & "'")
                    If (Not IsNothing(dr)) Then
                        If (dr.Length = 0) Then
                            _dtBarcode.Rows.Add(dgvDrawingInfo.Rows(e.RowIndex).Cells("Drawing Code").Value.ToString(), dgvDrawingInfo.Rows(e.RowIndex).Cells("Product No.").Value.ToString(), dgvDrawingInfo.Rows(e.RowIndex).Cells("Settsu No.").Value.ToString(), dgvDrawingInfo.Rows(e.RowIndex).Cells("Frame No.").Value.ToString())
                        End If
                    Else
                        _dtBarcode.Rows.Add(dgvDrawingInfo.Rows(e.RowIndex).Cells("Drawing Code").Value.ToString(), dgvDrawingInfo.Rows(e.RowIndex).Cells("Product No.").Value.ToString(), dgvDrawingInfo.Rows(e.RowIndex).Cells("Settsu No.").Value.ToString(), dgvDrawingInfo.Rows(e.RowIndex).Cells("Frame No.").Value.ToString())
                    End If
                Else
                    Dim dr = _dtBarcode.Select("`Drawing Code` = '" & dgvDrawingInfo.Rows(e.RowIndex).Cells("Drawing Code").Value.ToString() & "'")
                    If (Not IsNothing(dr)) Then
                        If (dr.Length > 0) Then
                            _dtBarcode.Rows.Remove(dr(0))
                            _dtBarcode.AcceptChanges()
                        End If
                    End If

                End If
            End If
        End If

        Dim i As Integer = 0
        For Each row As DataRow In _dtBarcode.Rows
            i += 1
        Next
        lblTotalSelectedRow.Text = "Selected Row(s): " & i.ToString
    End Sub
End Class