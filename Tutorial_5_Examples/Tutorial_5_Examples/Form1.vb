' ----- Form1.vb -----

Imports System
Imports System.IO
Imports System.Data.SqlClient

Public Class Form1
    Private Sub BtnUpdate_Click(sender As Object, e As EventArgs) Handles BtnUpdate.Click

        Dim txtOutput As String = ""

        ' Create an object
        Dim spot As New Animal With {
            .Name = "Spot"
        }

        ' And another using the Shared method
        Dim fluffy = Animal.MakeAnimal("Fluffy")

        ' Access the shared property
        txtOutput &= "Num of Animals " &
            fluffy.NumOfAnimals &
            Environment.NewLine

        ' Create a string array with lines of text
        Dim lines() As String = {"This is some random text",
            "saved to a file."}

        ' Target the My Documents path
        Dim mydocpath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

        ' Write the strings to the file
        ' Using acquires the system resource StreamWriter
        Using outputFile As New StreamWriter(mydocpath & Convert.ToString("\randomtext.txt"))
            For Each line As String In lines
                outputFile.WriteLine(line)
            Next
        End Using

        ' Append text to the file
        Using outputFile As New StreamWriter(mydocpath & Convert.ToString("\randomtext.txt"), True)
            outputFile.WriteLine("Here is more info")
        End Using

        ' Read text from a file and output it
        Try
            ' Open the file using StreamReader
            Using sr As New StreamReader(mydocpath & Convert.ToString("\randomtext.txt"))
                Dim line As String
                ' Read the stream and save it to a string
                ' which is written in the text box
                line = sr.ReadToEnd()
                txtOutput &= line &
                    Environment.NewLine
            End Using
        Catch ex As Exception
            Console.WriteLine("Couldn't Read the File")
            Console.WriteLine(ex.Message)
        End Try

        ' MessageBoxs provide an easy way to pass information
        ' to the user and then pass info back to you
        ' There are multiple icons you can use Asterisk,
        ' Error, Exclamation, Hand, Information, Question,
        ' Stop, Warning, None
        ' There are multiple buttons AbortRetryIgnore,
        ' OK, OKCancel, RetryCancel, YesNo, YesNoCancel
        ' You can set the default highlighted button with
        ' Button1, Button2, or Button3

        If MessageBox.Show("Message to Show", "Title",
                        MessageBoxButtons.AbortRetryIgnore,
                        MessageBoxIcon.Exclamation,
                        MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.Abort Then
            txtOutput &= "Abort Clicked" & Environment.NewLine
        Else
            txtOutput &= "Retry or Ignore Clicked" & Environment.NewLine

        End If

        ' Use OpenFileDialog to open files
        Dim OpenFileDialogEx As New OpenFileDialog() With {
            .Filter = "Text Documents (*.txt)|*.txt|All Files (*.*)|*.*",
            .FilterIndex = 2,
            .Title = "Open Important File"
        }

        ' Get the file selected
        Dim fileSelected As String

        If OpenFileDialogEx.ShowDialog =
                System.Windows.Forms.DialogResult.OK Then
            Try
                fileSelected = OpenFileDialogEx.FileName

                txtOutput &= "File Selected : " &
                    fileSelected &
                    Environment.NewLine

            Catch ex As Exception
                MessageBox.Show("Error Getting File", "Error")
            End Try
        End If

        ' SaveFileDialog allows you to define what file to save
        Dim fileToSave As String = ""

        ' OverwritePrompt protects you if the file already exists
        Dim SaveFileDialogEx As New SaveFileDialog() With {
            .Filter = "Text Documents (*.txt)|*.txt|All Files (*.*)|*.*",
            .FilterIndex = 2,
            .Title = "Save Important File",
            .DefaultExt = "txt",
            .FileName = fileToSave,
            .OverwritePrompt = True
        }

        If SaveFileDialogEx.ShowDialog =
                System.Windows.Forms.DialogResult.OK Then
            Try
                fileSelected = SaveFileDialogEx.FileName

                txtOutput &= "File Saved : " &
                    fileSelected &
                    Environment.NewLine

            Catch ex As Exception
                MessageBox.Show("Error Saving File", "Error")
            End Try

        End If

        ' FontDialog allows the user to select the font,
        ' font style, size, color, strikeout and underline
        Dim FontDialogEx As New FontDialog() With {
            .ShowColor = True
        }

        If FontDialogEx.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            txtOutput &= FontDialogEx.Font.ToString &
                Environment.NewLine

            txtOutput &= FontDialogEx.Color.ToString &
                Environment.NewLine
        End If

        TextBox1.Text = txtOutput
    End Sub
End Class

' ----- Animal.vb -----

Public Class Animal

    Public Name As String

    ' Some times we want to define properties that have
    ' the same value for every object. These are
    ' called Shared Properties
    Public Shared NumOfAnimals As Integer = 0

    Public Sub New(Optional n As String = "Unknown")
        Name = n

        ' Every time an object is created we'll keep
        ' track of it
        NumOfAnimals += 1
    End Sub

    ' You can only access other shared functions and 
    ' shared properties with shared methods / functions
    ' This 
    Public Shared Function MakeAnimal(n As String)
        Dim newAnimal As New Animal() With {
            .Name = n
        }

        Return newAnimal
    End Function

End Class
