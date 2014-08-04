Public Class UserPrompt

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        TextBox1.Text = TextBox1.Text.Replace("\", "")
        TextBox1.Text = TextBox1.Text.Replace("/", "")
        TextBox1.Text = TextBox1.Text.Replace(":", "")
        TextBox1.Text = TextBox1.Text.Replace("*", "")
        TextBox1.Text = TextBox1.Text.Replace("?", "")
        TextBox1.Text = TextBox1.Text.Replace("""", "")
        TextBox1.Text = TextBox1.Text.Replace("<", "")
        TextBox1.Text = TextBox1.Text.Replace(">", "")
        TextBox1.Text = TextBox1.Text.Replace("|", "")
        If TextBox1.Text.Length > 0 Then
            TextBox1.Select(TextBox1.Text.Length, 0)
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
    End Sub

    Private Sub UserPrompt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Select()
        TextBox1.Focus()
    End Sub
End Class