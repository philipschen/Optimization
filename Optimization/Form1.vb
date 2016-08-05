Imports System.Data.SqlClient
Public Class Form1
    Dim connectionstring As Class1 = New Class1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim con1 As New SqlConnection
        Dim cmd1 As New SqlCommand
        con1.ConnectionString = "Data Source=(Local)\SQLEXPRESS;Initial Catalog=TestDB;Integrated Security=True"

        '
        ' Initial Setup for DataBase and Tables
        '
        Try
            con1.Open()
            If (con1.State = ConnectionState.Open) Then
                Dim con As New SqlConnection
                Dim cmd As New SqlCommand
                con.ConnectionString = "Data Source=(Local)\SQLEXPRESS;Initial Catalog=TestDB;Integrated Security=True"
                con.Open()
                cmd.Connection = con

                Try
                    cmd.CommandText = "SELECT internalID, shopNumber  FROM parts"
                    cmd.ExecuteNonQuery()
                Catch
                    Try
                        cmd.CommandText = "CREATE TABLE dbo.parts" + "(partID nvarchar(50) NOT NULL, description nvarchar(50), color nvarchar(50), size FLOAT NOT NULL, count int NOT NULL, internalID int NOT NULL, shopNumber nvarchar(50) NOT NULL, itemNumber float, itemQuantity int, context1 nvarchar(50), context2 float, context3 nvarchar(50))"
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "CREATE TABLE dbo.stockNew" + "(stockID1 nvarchar(50), stockID2 nvarchar(50) NOT NULL, stockID3 nvarchar(50), description nvarchar(MAX), color nvarchar(50), size FLOAT NOT NULL, count int NOT NULL, internalID int NOT NULL, context1 nvarchar(50), context2 float, context3 nvarchar(50))"
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "CREATE TABLE dbo.stockUsed" + "(stockID1 nvarchar(50), stockID2 nvarchar(50) NOT NULL, stockID3 nvarchar(50), description nvarchar(MAX), color nvarchar(50), size FLOAT NOT NULL, count int NOT NULL, internalID int NOT NULL, location nvarchar(50), context1 nvarchar(50), context2 float, context3 nvarchar(50))"
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "CREATE TABLE dbo.activationCode" + "(CutGLib nvarchar(50))"
                        cmd.ExecuteNonQuery()
                        MessageBox.Show("Tables have been created successfully", "MyProgram", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Catch ex As Exception
                        MessageBox.Show(ex.ToString())
                    End Try
                End Try
            End If
        Catch
        Finally
            If (con1.State = ConnectionState.Open) Then
                con1.Close()
            Else
                Dim ext1 As InitialSetup = New InitialSetup()
                ext1.ShowDialog()
                If ext1.DialogResult = DialogResult.No Then
                    Me.Close()
                End If
                MessageBox.Show("Re-Launch EasyCut", "MyProgram", MessageBoxButtons.OK)
                Me.Close()
            End If
        End Try

        ' Set the caption bar text of the form.  
        Me.Text = connectionstring.version
        Button3.Visible = True

        Try
            PictureBox1.Image = Image.FromFile("temp.jpg")
            PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
        Catch

        End Try
    End Sub

    Private Sub BStock_Click(sender As Object, e As EventArgs) Handles BStock.Click
        Dim ext1 As Extrusions = New Extrusions()
        ext1.Show()
    End Sub

    Private Sub BPart_Click(sender As Object, e As EventArgs) Handles BPart.Click
        Dim cutman1 As CutManagement = New CutManagement()
        cutman1.Show()
    End Sub

    Private Sub set_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim frm1 As Confirm1 = New Confirm1
        If frm1.ShowDialog() = DialogResult.OK Then
            Dim con As New SqlConnection
            Dim cmd As New SqlCommand
            con.ConnectionString = connectionstring.connect1
            con.Open()
            cmd.Connection = con
            cmd.CommandText = "DELETE FROM stockNew"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "DELETE FROM stockUsed"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "DELETE FROM parts"
            cmd.ExecuteNonQuery()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim tables1 As DirectlyEdit = New DirectlyEdit()
        tables1.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim excellin1 As StockInputExcell = New StockInputExcell()
        excellin1.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim excellin1 As PartInputExcell = New PartInputExcell()
        excellin1.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim excellin1 As UsedStockInputExcel = New UsedStockInputExcel()
        excellin1.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim excellin1 As NavInput = New NavInput()
        excellin1.Show()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = "Data Source=(Local)\SQLEXPRESS;Initial Catalog=TestDB;Integrated Security=True"
        con.Open()
        cmd.Connection = con
        cmd.CommandText = "CREATE TABLE dbo.parts" + "(partID nvarchar(50) NOT NULL, description nvarchar(50), color nvarchar(50), size FLOAT NOT NULL, count int NOT NULL, internalID int NOT NULL, shopNumber nvarchar(50) NOT NULL, itemNumber float, itemQuantity int, context1 nvarchar(50), context2 float, context3 nvarchar(50))"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "CREATE TABLE dbo.stockNew" + "(stockID1 nvarchar(50), stockID2 nvarchar(50) NOT NULL, stockID3 nvarchar(50), description nvarchar(MAX), color nvarchar(50), size FLOAT NOT NULL, count int NOT NULL, internalID int NOT NULL, context1 nvarchar(50), context2 float, context3 nvarchar(50))"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "CREATE TABLE dbo.stockUsed" + "(stockID1 nvarchar(50), stockID2 nvarchar(50) NOT NULL, stockID3 nvarchar(50), description nvarchar(MAX), color nvarchar(50), size FLOAT NOT NULL, count int NOT NULL, internalID int NOT NULL, location nvarchar(50), context1 nvarchar(50), context2 float, context3 nvarchar(50))"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "CREATE TABLE dbo.activationCode" + "(CutGLib nvarchar(50))"
        cmd.ExecuteNonQuery()
        MessageBox.Show("Tables have been created successfully", "MyProgram", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
End Class

