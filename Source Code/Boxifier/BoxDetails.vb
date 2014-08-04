Public Class BoxDetails
    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal message As String = "")
        Try
            MsgBox("The system has encountered the following error and will attempt to recover from it accordingly:" & vbCrLf & message & vbCrLf & vbCrLf & ex.Message.ToLower, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error Encountered")
        Catch exc As Exception
            MsgBox("The system has encountered a critical, unhandled exception." & vbCrLf & vbCrLf & exc.Message.ToLower, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Critical Error")
        End Try
    End Sub

    Public Sub ClearInfo()
        Try
            TextBox1.Text = ""
            TextBox2.Text = ""
            RichTextBox1.Text = ""
        Catch ex As Exception
            Error_Handler(ex, "[ClearInfo]")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text.Length < 1 Then
            MsgBox("The box details you are trying to submit is not allowed as the box ID is required to have a valid value.", MsgBoxStyle.Information, "Error in Input")
            Button1.DialogResult = Windows.Forms.DialogResult.None
        Else
            Button1.DialogResult = Windows.Forms.DialogResult.OK
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()
        End If
    End Sub

    Private Sub BoxDetails_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button1.DialogResult = Windows.Forms.DialogResult.None
        TextBox2.Select()
        TextBox2.Focus()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Try
            If IsNumeric(TextBox1.Text) = False Then
                If TextBox1.Text.Length > 1 Then
                    TextBox1.Text = TextBox1.Text.Remove(TextBox1.Text.Length - 1, 1)
                Else
                    TextBox1.Text = ""
                End If
                If TextBox1.Text.Length > 0 Then
                    TextBox1.Select(TextBox1.Text.Length - 1, 0)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
End Class