Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Public Class frmWorkExcel

    'Private xlApp As Excel.Application
    'Private xlLlibre As Excel.Workbook
    'Private xlFulla As Excel.Worksheet


    Private Sub frmWorkExcel_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        btnExplorarExcel.Enabled = False
    End Sub

    Private Sub btnExaminar_Click(sender As Object, e As EventArgs) Handles btnExaminar.Click
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.Title = "Selecciona un fitxer"
        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            tbRuta.Text = openFileDialog1.FileName.ToString()
            btnExplorarExcel.Enabled = True
        End If
    End Sub

    Private Sub btnExplorarExcel_Click(sender As Object, e As EventArgs) Handles btnExplorarExcel.Click
        Dim xlApp = New Excel.Application
        Dim xlLlibre As Excel.Workbook
        Dim xlFulla As Excel.Worksheet

        Dim numColumnes As Integer
        Dim numFiles As Integer
        xlLlibre = xlApp.Workbooks.Open(tbRuta.Text, True, True, , "")
        xlFulla = xlApp.Worksheets("Hoja1")
        numColumnes = xlFulla.Columns.Count
        numFiles = xlFulla.Rows.Count

        Dim contingutCela As String

        For i As Integer = 1 To numColumnes
            For j As Integer = 1 To numFiles
                contingutCela = xlFulla.Range(i, j).Text.ToString()
                lblContingutCela.Text = contingutCela
            Next
        Next

        xlApp.Quit()
        xlLlibre.Close()

        xlApp = Nothing
        xlLlibre = Nothing
        xlFulla = Nothing


    End Sub
End Class
