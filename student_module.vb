Public Class Form1
    Dim Currword As String = ""
    Dim data As String


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click

        Dim result As DialogResult = OFD_Cloze.ShowDialog()
        Dim file As String = OFD_Cloze.FileName
        Dim FileNum As Integer = FreeFile()
        Dim CarOv As String = "", TempLine As String
        Dim data As String = ""
        Dim Done As Boolean = False

       
        If result = Windows.Forms.DialogResult.OK Then   'if the dialog box was successful
            FileOpen(FileNum, file, OpenMode.Input)     ' open file
            Do Until EOF(FileNum)                       ' read line by line to end of file
                TempLine = LineInput(FileNum)
                CarOv += TempLine
                CarOv += Environment.NewLine
            Loop
            FileClose(FileNum)          ' close the file
            data = CarOv       ' dumps in variable
        End If

        ' show cloze passage in text box
        Dim i As Integer = 0
        While i < data.Length
            If data(i) = "|" Then
                i = i + 1
                If data(i) = "|" Then
                    i = i + 1
                    If data(i) = "|" Then
                        Exit While
                    End If
                End If
            Else
                TB_Cloze.Text = TB_Cloze.Text + data(i)
            End If
            i = i + 1
        End While

        ' Puts Word Bank into the Combo Box
        i = 78
        While i < data.Length
            If data(i) = "|" Then
                i = i + 1
                If data(i) = "|" Then
                    i = i + 1
                    If data(i) = "|" Then
                        Exit While
                    End If
                End If
                ComboBox1.Text = ComboBox1.Items.Add(Currword)
                Currword = ""
            Else
                Currword = Currword + data(i)
            End If
            i = i + 1
            Done = True
        End While

    End Sub

    Private Sub ComboBox1_TextChanged(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged
       
    End Sub

    Private Sub BTN_Replace_Click(sender As Object, e As EventArgs) Handles BTN_Replace.Click
        'Dim o As Integer = 1
        Dim INDEX As Integer = 1
        For i to data.Length
            'Do Until INDEX > "(" & INDEX & ")______"
            'While INDEX < Currword
            'If TB_Cloze.Text = ("(" & INDEX & ")______") Then
            'Exit While
            'Else
            TB_Cloze.Text = Replace(TB_Cloze.Text, "(" & INDEX & ")______", ComboBox1.Text)
            INDEX = INDEX + 1
            'End If
            'End While
            'Loop
        Next i
    End Sub

    Private Sub BTN_Submit_Click(sender As Object, e As EventArgs) Handles BTN_Submit.Click
        'While TB_Cloze.Text(i) = Test.text(i)
    End Sub
End Class
