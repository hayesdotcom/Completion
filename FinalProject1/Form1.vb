Imports System.Drawing.Printing
Imports System.DataTable
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office
Imports Microsoft.Office.Interop
Imports System.IO

Public Class Form1
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        MeExit()
    End Sub

    Private Sub MeExit()
        Dim iExit As DialogResult

        iExit = MessageBox.Show("Confirm if you want to exit", "Datagridview System", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        If iExit = DialogResult.Yes Then
            Application.Exit()

        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        For Each row As DataGridViewRow In DataGridView1.SelectedRows
            DataGridView1.Rows.Remove(row)
        Next
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        txtFirstname.Text = ""
        txtLastname.Text = ""
        cmbGender.Text = ""
        txtDOB.Text = ""
        txtAge.Text = ""
        txtAddress.Text = ""
        txtStud.Text = ""
        txtDept.Text = ""
        txtCourse.Text = ""
        txtYear.Text = ""
        txtSection.Text = ""
        txtMobile.Text = ""
        txtEmail.Text = ""
        cmbVaccine.Text = ""


    End Sub

    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        DataGridView1.Rows.Add(txtFirstname.Text, txtLastname.Text, cmbGender.Text, txtDOB.Text, txtAge.Text, txtAddress.Text,
                               txtStud.Text, txtDept.Text, txtCourse.Text, txtYear.Text,
                               txtSection.Text, txtMobile.Text, txtEmail.Text, cmbVaccine.Text)
    End Sub

    Private bitmap As Bitmap

    Private Sub iPrint()
        Dim height As Integer = DataGridView1.Height
        DataGridView1.Height = DataGridView1.RowCount * DataGridView1.RowTemplate.Height
        bitmap = New Bitmap(Me.DataGridView1.Width, Me.DataGridView1.Height)
        DataGridView1.DrawToBitmap(bitmap, New Rectangle(0, 0, Me.DataGridView1.Width, Me.DataGridView1.Height))
        PrintPreviewDialog1.Document = PrintDocument1
        PrintPreviewDialog1.PrintPreviewControl.Zoom = 1
        PrintPreviewDialog1.ShowDialog()
        DataGridView1.Height = height
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        iPrint()
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        e.Graphics.DrawImage(bitmap, 0, 0)

    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        Using sfd As SaveFileDialog = New SaveFileDialog() With {.Filter = "Excel Workbook| *.xlsx"}
            If sfd.ShowDialog() = DialogResult.OK Then
                Try
                    Using workbook As XLWorkbook = New XLWorkbook()
                        workbook.Worksheet.Add(Me.DataGridView1.Products.CopyTodataTable(), "Products")
                        workbook.SaveAs(sfd.FileName)

                    End Using
                    MessageBox.Show("Succesfuly added in Excel!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            End If
        End Using
    End Sub

    Private Sub SaveToExcel()

        Dim excel As Microsoft.Office.Interop.Excel._Application = New Microsoft.Office.Interop.Excel.Application()
        Dim workbook As Microsoft.Office.Interop.Excel._Workbook = excel.Workbooks.Add(Type.Missing)
        Dim worksheet As Microsoft.Office.Interop.Excel._Worksheet = Nothing

        Try
            worksheet = workbook.ActiveSheet

            worksheet.Name = "ExportedFromDataGrid"

            Dim cellRowIndex As Integer = 1
            Dim cellColumnIndex As Integer = 1

            For j As Integer = 0 To DataGridView1.Columns.Count - 1
                worksheet.Cells(cellRowIndex, cellColumnIndex) = DataGridView1.Columns(j).HeaderText
                cellColumnIndex += 1
            Next

            cellColumnIndex = 1
            cellRowIndex += 1

            For i As Integer = 0 To DataGridView1.Rows.Count - 2
                For j As Integer = 0 To DataGridView1.Columns.Count - 1
                    worksheet.Cells(cellRowIndex, cellColumnIndex) = DataGridView1.Rows(i).Cells(j).Value.ToString()
                    cellColumnIndex += 1
                Next

                cellColumnIndex = 1
                cellRowIndex += 1

            Next

            Dim saveDialog As New SaveFileDialog()
            saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            saveDialog.FilterIndex = 2

            If saveDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                workbook.SaveAs(saveDialog.FileName)
                MessageBox.Show("Export Successful!")
            End If


        Catch ex As System.Exception
            MessageBox.Show(ex.Message)
        Finally
            excel.Quit()
            workbook = Nothing
            excel = Nothing

        End Try
    End Sub

    Private Sub Label16_Click(sender As Object, e As EventArgs) Handles Label16.Click

    End Sub
End Class
