Public Class Student_Module
    Dim Currword As String = ""
    Dim data As String

    Private Sub Student_Module_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CutToolStripMenuItem.Enabled = False
        PrintPreviewToolStripMenuItem.Enabled = False
        CopyToolStripMenuItem.Enabled = False
        PasteToolStripMenuItem.Enabled = False
        SelectAllToolStripMenuItem.Enabled = False
        CustomizeToolStripMenuItem.Enabled = False
        OptionsToolStripMenuItem.Enabled = False
        ContentsToolStripMenuItem.Enabled = False
        'IndexToolStripMenuItem.Enabled = False
        'SearchToolStripMenuItem.Enabled = False
        AboutToolStripMenuItem.Enabled = False



    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        Dim result As DialogResult = OFD_Cloze.ShowDialog()
        Dim file As String = OFD_Cloze.FileName
        Dim FileNum As Integer = FreeFile()
        Dim CarOv As String = "", TempLine As String
        Dim data As String = ""
        Dim Done As Boolean = False

        If result = Windows.Forms.DialogResult.OK Then  ' if the dialog box was successful
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
    Private Sub BTN_Replace_Click(sender As Object, e As EventArgs) Handles BTN_Replace.Click
        Dim o As Integer = 1
        Dim INDEX As Integer = 1
        Dim example As Integer = 0

        Do While example <= 10
            While INDEX < Currword
                'If TB_Cloze.Text = ("(" & INDEX & ")______") Then
                'Else
                TB_Cloze.Text = Replace(TB_Cloze.Text, "(" & INDEX & ")______", ComboBox1.Text)
                INDEX = INDEX + 1
                'End If
            End While
                example += 1
        Loop
    End Sub

    Private Sub BTN_Submit_Click(sender As Object, e As EventArgs)
        If TB_Cloze.Text = ClozePassage.CarOv Then
            MsgBox("Correct")
        Else
            MsgBox("THIS IS INCORRECT")
            End
        End If
    End Sub
    Private Sub Student_Module_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Application.Exit()
    End Sub
End Class
