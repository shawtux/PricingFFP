Imports System.IO
Imports System.Text.RegularExpressions

Public Class Form1

    Private appPath = System.AppDomain.CurrentDomain.BaseDirectory


    Private Function getCheckedItems(element As DataGridView, ByVal ColumnName As String) As List(Of DataGridViewRow)
        Return _
            (
                From SubRows In
                    (
                        From Rows In element.Rows.Cast(Of DataGridViewRow)()
                        Where Not Rows.IsNewRow
                    ).ToList
                Where CBool(SubRows.Cells(ColumnName).Value) = True
            ).ToList
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim directoryPath As String = ""
        ToolStripStatusLabel2.Text = "Ejecutando"
        If DataGridView1.Rows.Count > 0 Then
            If DataGridView1.Rows.Count = 1 And (DataGridView1.Rows(0).Cells("fileNameColumn").Value = "ningun archivo encontrado") Then
                MsgBox("No se puede iniciar, no hay archivos a procesar")
            Else

                For Each item As DataGridViewRow In getCheckedItems(DataGridView1, "chkColumn")
                    Dim filename As String = item.Cells("fileNameColumn").Value.Substring(0, item.Cells("fileNameColumn").Value.IndexOf("."))

                    Dim options As RegexOptions = RegexOptions.Singleline Or RegexOptions.Multiline
                    Dim fareArray As New List(Of faredata)
                    Dim ruta As String = appPath & "\" & item.Cells("fileNameColumn").Value
                    'Dim ruta As String = "C:\\Users\\Marcos Orrego\\Documents\\Visual Studio 2015\\Projects\\PricingFFP\\output_test_script_PRICING_CustomizedInputs_RISCL JACQUELINE P.csv"
                    directoryPath = Path.GetDirectoryName(ruta)
                    Dim salida As StreamWriter
                    salida = My.Computer.FileSystem.OpenTextFileWriter(directoryPath & "\\salida.txt", False)
                    If File.Exists(ruta) Then
                        Dim readText As String = File.ReadAllText(ruta)
                        Dim responsePattern As String = "Response: \[(.*?)\].*?(?=\"")"

                        For Each match As Match In Regex.Matches(readText, responsePattern, options)

                            salida.WriteLine(match.Groups(1).ToString)

                        Next

                    End If
                    salida.Close()
                    Dim util As Utilities = New Utilities()
                    'fareArray = util.parseFareTable(File.ReadAllText(directoryPath & "\\salida.txt"))
                    util.parseFQN(File.ReadAllText(directoryPath & "\\salida.txt"), appPath & "\" & filename & ".xlsx")
                    'Console.WriteLine(util.getCleanFQN(File.ReadAllText(directoryPath & "\\fqn_test.txt")))


                Next
                'File.Delete(directoryPath & "\\salida.txt")

            End If
        End If
        ToolStripStatusLabel2.Text = "Terminado"
        MsgBox("Proceso Terminado", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            txtBoxPath.Text = FolderBrowserDialog1.SelectedPath
            appPath = FolderBrowserDialog1.SelectedPath
            populateFileList()
        End If
    End Sub

    Private Sub datagridview_addColumns()
        If DataGridView1.Columns.Count = 0 Then
            Dim col As New DataGridViewCheckBoxColumn
            With col
                .Name = "chkColumn"
                .HeaderText = ""
                .ReadOnly = False
            End With
            DataGridView1.Columns.Add(col)
            Dim col1 As New DataGridViewTextBoxColumn
            With col1
                .Name = "fileNameColumn"
                .HeaderText = "Archivo"
            End With
            DataGridView1.Columns.Add(col1)
            Dim col2 As New DataGridViewTextBoxColumn
            With col2
                .Name = "modifDateColumn"
                .HeaderText = "Ultima fecha modificacon"

            End With
            DataGridView1.Columns.Add(col2)

            DataGridView1.Sort(DataGridView1.Columns("modifDateColumn"), System.ComponentModel.ListSortDirection.Descending)
        End If
    End Sub

    Private Sub Form1_onload(sender As Object, e As EventArgs) Handles MyBase.Load

        datagridview_addColumns()



        'txtBoxPath.Text = appPath
        'populateFileList()

    End Sub

    Private Sub Form1_onshow(sender As Object, e As EventArgs) Handles MyBase.Shown
        DataGridView1.AutoGenerateColumns = True

        txtBoxPath.Text = appPath

        populateFileList()
    End Sub

    Private Sub txtBoxPath_TextChanged(sender As Object, e As EventArgs) Handles txtBoxPath.TextChanged
        If (txtBoxPath.Text <> Nothing) Then
            appPath = txtBoxPath.Text
            populateFileList()
        End If
    End Sub

    Public Sub populateFileList()


        DataGridView1.Rows.Clear()
        If DataGridView1.Columns.Count() = 0 Then
            datagridview_addColumns()
        End If

        Try
            For Each filefound As String In My.Computer.FileSystem.GetFiles(appPath)
                Dim extension As String = System.IO.Path.GetExtension(filefound)

                If extension.ToLower = ".csv" Then

                    DataGridView1.Rows.Add(New Object() {False, Path.GetFileName(filefound), File.GetLastWriteTime(filefound)})

                End If
            Next
        Catch
            MessageBox.Show("Ruta no existe", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            If DataGridView1.Rows.Count = 0 Then

                DataGridView1.Rows.Add(New Object() {False, "ningun archivo encontrado", ""})
            End If
        End Try

        If (DataGridView1.Rows.Count = 0) Then

            DataGridView1.Rows.Add(New Object() {False, "ningun archivo encontrado", ""})
        End If
        DataGridView1.Sort(DataGridView1.Columns("modifDateColumn"), System.ComponentModel.ListSortDirection.Descending)
    End Sub

End Class
