Imports System.IO
Imports System.Text.RegularExpressions

Public Class Form1

    Private appPath = System.AppDomain.CurrentDomain.BaseDirectory
    Sub populateFileList()

        CheckedListBox1.Items.Clear()
        DataGridView1.Rows.Clear()

        Try
        For Each filefound As String In My.Computer.FileSystem.GetFiles(appPath)
                Dim extension As String = System.IO.Path.GetExtension(filefound)

                If extension.ToLower = ".csv" Then
                    CheckedListBox1.Enabled = True
                    CheckedListBox1.Items.Add(System.IO.Path.GetFileName(filefound))

                    DataGridView1.Rows.Add(False, Path.GetFileName(filefound), File.GetLastWriteTime(filefound))

                End If
            Next
        Catch
            MessageBox.Show("Ruta no existe", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            If CheckedListBox1.Items.Count = 0 Then
                CheckedListBox1.Items.Add("ningun archivo encontrado")
                DataGridView1.Rows.Add(False, "ningun archivo encontrado", "")
            End If
        End Try

        If (CheckedListBox1.Items.Count = 0 Or DataGridView1.Rows.Count = 0) Then
            CheckedListBox1.Items.Add("ningun archivo encontrado")
            'DataGridView1.Rows.Add(False, "ningun archivo encontrado", "")
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim directoryPath As String = ""
        ToolStripStatusLabel2.Text = "Ejecutando"
        If CheckedListBox1.CheckedItems.Count > 0 Then
            If CheckedListBox1.Items.Count = 1 And (CheckedListBox1.Items(1).ToString = "ningun archivo encontrado" Or DataGridView1.Rows(1).Cells("fileNameColumn").Value = "ningun archivo encontrado") Then
                MsgBox("No se puede iniciar, no hay archivos a procesar")
            Else

                For Each item In CheckedListBox1.CheckedItems
                    Dim filename As String = item.ToString.Substring(0, item.ToString.IndexOf("."))

                    Dim options As RegexOptions = RegexOptions.Singleline Or RegexOptions.Multiline
                    Dim fareArray As New List(Of faredata)
                    Dim ruta As String = appPath & "\" & item.ToString
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
        MsgBox("Proceso Terminado", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            txtBoxPath.Text = FolderBrowserDialog1.SelectedPath
            appPath = FolderBrowserDialog1.SelectedPath
            populateFileList()
        End If
    End Sub

    Private Sub Form1_onload(sender As Object, e As EventArgs) Handles MyBase.Load

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


End Class
