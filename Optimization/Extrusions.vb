Imports Excel = Microsoft.Office.Interop.Excel
Public Class Extrusions
    Private Sub Extrusions_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set the caption bar text of the form.  
        Me.Text = "Easy Cut V1.0"

        ListBox1.Items.Add("Bronze")
        ListBox1.Items.Add("White")
        ListBox1.Items.Add("Silver")
        ListBox1.Items.Add("Other")
        ListBox2.Items.Add("New Part")

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ListBox2.SelectedItem = "New Part" Then
            Dim frmNewExtr As NewExtrusions = New NewExtrusions()
            frmNewExtr.ShowDialog()
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim appXL As Excel.Application
        Dim wbXl As Excel.Workbook
        Dim shXL As Excel.Worksheet
        Dim raXL As Excel.Range
        ' Start Excel and get Application object.
        appXL = CreateObject("Excel.Application")
        appXL.Visible = True
        ' Add a new workbook.
        wbXl = appXL.Workbooks.Add
        shXL = wbXl.ActiveSheet
        ' Add table headers going cell by cell.
        shXL.Cells(1, 1).Value = "ID1"
        shXL.Cells(1, 2).Value = "ID2"
        shXL.Cells(1, 3).Value = "ID3"
        shXL.Cells(1, 4).Value = "Description"
        shXL.Cells(1, 5).Value = "Color"
        shXL.Cells(1, 6).Value = "Size"
        shXL.Cells(1, 7).Value = "Count"
        ' Format A1:D1 as bold, vertical alignment = center.
        With shXL.Range("A1", "G1")
            .Font.Bold = True
            .ColumnWidth = 20
            .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        End With
        ' Create an array to set multiple values at once.
        Dim students(5, 2) As String
        students(0, 0) = "Zara"
        students(0, 1) = "Ali"
        students(1, 0) = "Nuha"
        students(1, 1) = "Ali"
        students(2, 0) = "Arilia"
        students(2, 1) = "RamKumar"
        students(3, 0) = "Rita"
        students(3, 1) = "Jones"
        students(4, 0) = "Umme"
        students(4, 1) = "Ayman"
        ' Fill A2:B6 with an array of values (First and Last Names).
        shXL.Range("A2", "B6").Value = students
        ' Fill C2:C6 with a relative formula (=A2 & " " & B2).
        raXL = shXL.Range("C2", "C6")
        raXL.Formula = "=A2 & "" "" & B2"
        ' Fill D2:D6 values.
        With shXL
            .Cells(2, 4).Value = "Biology"
            .Cells(3, 4).Value = "Mathmematics"
            .Cells(4, 4).Value = "Physics"
            .Cells(5, 4).Value = "Mathmematics"
            .Cells(6, 4).Value = "Arabic"
        End With
        ' AutoFit columns A:D.
        raXL = shXL.Range("A1", "G1")
        raXL.EntireColumn.AutoFit()
        ' Make sure Excel is visible and give the user control
        ' of Excel's lifetime.
        appXL.Visible = True
        appXL.UserControl = True
        ' Release object references.
        raXL = Nothing
        shXL = Nothing
        wbXl = Nothing
        appXL.Quit()
        appXL = Nothing
    End Sub

    Private Sub ListBox2_SelectedValueChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedValueChanged
        Dim pString As String
        pString = "Temporary Description" + vbCrLf + "Temporary color"
        RichTextBox1.Text = pString
    End Sub
End Class