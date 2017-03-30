Imports System.IO
Public Class Form2
    Public print_data() As String
    Public Shared Function Create_and_Display_Form(display() As String)
        Dim i As Integer
        ReDim Form2.print_data(UBound(display))
        display.CopyTo(Form2.print_data, 0)
        For i = 0 To UBound(display)
            If display(i) = "Using bisection method " Then
                Form2.DataGridView1.Rows.Add(New String() {display(i), display(i + 1), " = ", display(i + 2) & " " & display(i + 3)})
                i = i + 3
            Else
                Form2.DataGridView1.Rows.Add(New String() {display(i), display(i + 1), " = ", display(i + 2)})
                Form2.DataGridView1.Rows.Add(New String() {"", "", " = ", display(i + 3)})
                Form2.DataGridView1.Rows.Add(New String() {"", "", " = ", display(i + 4) & " " & display(i + 5)})
                Form2.DataGridView1.Rows.Add(New String() {"", "", " ", " "})
                i = i + 5
            End If
        Next i
        Form2.DataGridView1.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Form2.DataGridView1.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Form2.DataGridView1.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Form2.DataGridView1.Columns(3).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Form2.Button2.Select()
        Form2.Show()
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call Form1.update_input_variables()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.Title = "Print to Text File"
        saveFileDialog1.OverwritePrompt = True
        saveFileDialog1.AddExtension = True
        Dim DR As DialogResult = saveFileDialog1.ShowDialog
        If saveFileDialog1.FileName <> "" Then
            saveFileDialog1.Filter = "TXT Files (*.txt*)|*.txt"
            If DR = Windows.Forms.DialogResult.OK Then
                Dim i As Integer
                Dim new_file As System.IO.StreamWriter
                If System.IO.File.Exists(saveFileDialog1.FileName) = True Then
                    File.Delete(saveFileDialog1.FileName)
                    saveFileDialog1.FileName = Replace(saveFileDialog1.FileName, ".txt", "")
                End If
                new_file = My.Computer.FileSystem.OpenTextFileWriter(saveFileDialog1.FileName & ".txt", True)
                For i = 0 To UBound(print_data)
                    If print_data(i) = "Using bisection method " Then
                        new_file.Write(print_data(i)) : new_file.Write(" ") : new_file.Write(print_data(i + 1)) : new_file.Write(" = ") : new_file.Write(print_data(i + 2)) : new_file.Write(" ") : new_file.WriteLine(print_data(i + 3))
                        new_file.WriteLine("  ")
                        i = i + 3
                    Else
                        new_file.Write(print_data(i)) : new_file.Write(" ") : new_file.Write(print_data(i + 1)) : new_file.Write(" = ") : new_file.WriteLine(print_data(i + 2))
                        new_file.Write(" = ") : new_file.WriteLine(print_data(i + 3))
                        new_file.Write(" = ") : new_file.Write(print_data(i + 4)) : new_file.Write(" ") : new_file.WriteLine(print_data(i + 5))
                        new_file.WriteLine("  ")
                        i = i + 5
                    End If
                Next i
                new_file.Close()
            End If
        End If
    End Sub
End Class