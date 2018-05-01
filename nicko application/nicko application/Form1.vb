Public Class Form1

    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click
        'Check if there's text added to the textbox
        If TextBox1.Modified Then
            'If the text of notepad changed, the program will ask the user if they want to save the changes
            Dim ask As MsgBoxResult
            ask = MsgBox("Do you want to save the changes", MsgBoxStyle.YesNoCancel, "New Document")
            If ask = MsgBoxResult.No Then
                TextBox1.Clear()
            ElseIf ask = MsgBoxResult.Cancel Then
            ElseIf ask = MsgBoxResult.Yes Then
                OpenFileDialog1.ShowDialog()
                My.Computer.FileSystem.WriteAllText(OpenFileDialog1.FileName, TextBox1.Text, False)
                TextBox1.Clear()
            End If
            'If textbox's text is still the same, notepad will open a new page:
        Else
            TextBox1.Clear()
        End If
    End Sub
    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        'Check if there's text added to the textbox
        If TextBox1.Modified Then
            'If the text of notepad changed, the program will ask the user if they want to save the changes
            Dim ask As MsgBoxResult
            ask = MsgBox("Do you want to save the changes", MsgBoxStyle.YesNoCancel, "Open Document")
            If ask = MsgBoxResult.No Then
                OpenFileDialog1.ShowDialog()
                TextBox1.Text = My.Computer.FileSystem.ReadAllText(OpenFileDialog1.FileName)
            ElseIf ask = MsgBoxResult.Cancel Then
            ElseIf ask = MsgBoxResult.Yes Then
                OpenFileDialog1.ShowDialog()
                My.Computer.FileSystem.WriteAllText(OpenFileDialog1.FileName, TextBox1.Text, False)
                TextBox1.Clear()
            End If
        Else
            'If textbox's text is still the same, notepad will show the OpenFileDialog
            OpenFileDialog1.ShowDialog()
            Try
                TextBox1.Text = My.Computer.FileSystem.ReadAllText(OpenFileDialog1.FileName)
            Catch ex As Exception
            End Try
        End If
    End Sub
    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        'Check if there's text added to the textbox
        If TextBox1.Modified Then
            'If the text of notepad changed, the program will ask the user if they want to save the changes
            Dim ask As MsgBoxResult
            ask = MsgBox("Do you want to save the changes", MsgBoxStyle.YesNoCancel, "Open Document")
            If ask = MsgBoxResult.No Then
                OpenFileDialog1.ShowDialog()
                TextBox1.Text = My.Computer.FileSystem.ReadAllText(OpenFileDialog1.FileName)
            ElseIf ask = MsgBoxResult.Cancel Then
            ElseIf ask = MsgBoxResult.Yes Then
                OpenFileDialog1.ShowDialog()
                My.Computer.FileSystem.WriteAllText(OpenFileDialog1.FileName, TextBox1.Text, False)
                TextBox1.Clear()
            End If
        Else
            'If textbox's text is still the same, notepad will show the OpenFileDialog
            OpenFileDialog1.ShowDialog()
            Try
                TextBox1.Text = My.Computer.FileSystem.ReadAllText(OpenFileDialog1.FileName)
            Catch ex As Exception
            End Try
        End If
    End Sub
    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        'exit the program
        End

    End Sub
    Private Sub UndoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UndoToolStripMenuItem.Click
        'check if textbox can undo
        If TextBox1.CanUndo Then
            TextBox1.Undo()
        Else
        End If
    End Sub
    Private Sub CopyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripMenuItem.Click

        My.Computer.Clipboard.Clear()
        If TextBox1.SelectionLength > 0 Then
            My.Computer.Clipboard.SetText(TextBox1.SelectedText)

        End If
    End Sub
    Private Sub PasteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteToolStripMenuItem.Click
        If My.Computer.Clipboard.ContainsText Then
            TextBox1.Paste()
        End If
    End Sub
    Private Sub SelectAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectAllToolStripMenuItem.Click
        TextBox1.SelectAll()
    End Sub
    Private Sub FindToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindToolStripMenuItem.Click
        Dim a As String
        Dim b As String
        a = InputBox("Enter text to be found")
        b = InStr(TextBox1.Text, a)
        If b Then
            TextBox1.Focus()
            TextBox1.SelectionStart = b - 1
            TextBox1.SelectionLength = Len(a)
        Else
            MsgBox("Text not found.")
        End If

    End Sub
    Private Sub FontToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FontToolStripMenuItem.Click
        FontDialog1.ShowDialog()
        TextBox1.Font = FontDialog1.Font
    End Sub
    Private Sub FontColorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FontColorToolStripMenuItem.Click
        ColorDialog1.ShowDialog()
        TextBox1.ForeColor = ColorDialog1.Color
    End Sub
    Private Sub LeftToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LeftToolStripMenuItem.Click
        TextBox1.TextAlign = HorizontalAlignment.Left
        LeftToolStripMenuItem.Checked = True
        CenterToolStripMenuItem.Checked = False
        RightToolStripMenuItem.Checked = False

    End Sub
    Private Sub CenterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CenterToolStripMenuItem.Click
        TextBox1.TextAlign = HorizontalAlignment.Center
        LeftToolStripMenuItem.Checked = False
        CenterToolStripMenuItem.Checked = True
        RightToolStripMenuItem.Checked = False

    End Sub
    Private Sub RightToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RightToolStripMenuItem.Click
        TextBox1.TextAlign = HorizontalAlignment.Right
        LeftToolStripMenuItem.Checked = False
        CenterToolStripMenuItem.Checked = False
        RightToolStripMenuItem.Checked = True
    End Sub
    Private Sub BackColorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackColorToolStripMenuItem.Click
        ColorDialog1.ShowDialog()
        TextBox1.BackColor = ColorDialog1.Color
    End Sub
End Class