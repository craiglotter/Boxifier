Imports System.IO

Imports System.Data.OleDb
Public Class Main_Screen

    Private basedatabase As String = (Application.StartupPath & "\boxifier.mdb").Replace("\\", "\")
    Private adapter As OleDbDataAdapter
    Private boxes As DataSet
    Private boxdetails1 As BoxDetails
    Private results1 As Results

    Private ds As DataSet
    Private da As OleDbDataAdapter
    Private conn As OleDbConnection
    Private cmd As OleDbCommand

    Private lastselectedindex As Integer = -2

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal message As String = "")
        Try
            If CheckBox1.Checked = False Then
                MsgBox("The system has encountered the following error and will attempt to recover from it accordingly:" & vbCrLf & message & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error Encountered")
            Else
                MsgBox("The system has encountered the following error and will attempt to recover from it accordingly:" & vbCrLf & message & vbCrLf & vbCrLf & ex.ToString, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error Encountered")
            End If

        Catch exc As Exception
            MsgBox("The system has encountered a critical, unhandled exception." & vbCrLf & vbCrLf & exc.Message.ToLower, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Critical Error")
        End Try
    End Sub

    Private Sub LoadComboBox()
        Try
            ComboBox1.Items.Clear()
            Dim dinfo As DirectoryInfo = New DirectoryInfo(Application.StartupPath)
            For Each finfo As FileInfo In dinfo.GetFiles
                If finfo.Extension.ToLower = ".mdb" Then
                    If Not finfo.Name = "Boxifier.mdb" Then
                        ComboBox1.Items.Add(finfo.Name)
                    End If
                End If
            Next
            If ComboBox1.Items.Count > 0 Then
                'If lastselectedindex = -2 Then
                '    ComboBox1.SelectedIndex = 0
                'Else
                '    If ComboBox1.Items.Count >= lastselectedindex + 1 Then
                '        ComboBox1.SelectedIndex = lastselectedindex
                '    Else
                ComboBox1.SelectedIndex = 0
                'End If
                '    End If
            End If
        Catch ex As Exception
            Error_Handler(ex, "[Loading existing box sets]")
        End Try
    End Sub

    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Text = "Boxifier " & Format(My.Application.Info.Version.Major, "00") & Format(My.Application.Info.Version.Minor, "00") & Format(My.Application.Info.Version.Build, "00") & "." & My.Application.Info.Version.Revision
            boxdetails1 = New BoxDetails()

            LoadComboBox()
        


        Catch ex As Exception
            Error_Handler(ex, "[Application Load]")
        End Try
    End Sub

    Private Sub main_screen_close(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Try
            If ComboBox1.SelectedIndex <> -1 Then
                'If MsgBox("Do you wish to save any changes made to your box set?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Update Database?") = MsgBoxResult.Yes Then
                '    da.Update(ds)
                'End If
            End If
            If conn IsNot Nothing Then
                conn.Close()
                da.Dispose()
                ds.Dispose()
            End If
        Catch ex As Exception
            Error_Handler(ex, "[Application Closing]")
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            If (lastselectedindex <> ComboBox1.SelectedIndex) And (lastselectedindex <> -2) Then
                If ComboBox1.SelectedIndex <> -1 Then
                    'If MsgBox("Do you wish to save any changes made to your box set?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Update Database?") = MsgBoxResult.Yes Then
                    '    da.Update(ds)
                    'End If
                End If
            End If
            If lastselectedindex <> ComboBox1.SelectedIndex Then
                InitialiseDataGrid()
            End If
            lastselectedindex = ComboBox1.SelectedIndex
        Catch ex As Exception
            Error_Handler(ex, "[Change DB Selection]")
        End Try
    End Sub
    Private Sub InitialiseDataGrid()
        Try
            DataGridView1.Enabled = True
            GroupBox1.Text = ComboBox1.Items(ComboBox1.SelectedIndex).ToString.Replace(".mdb", "")
            If conn IsNot Nothing Then
                conn.Close()
                da.Dispose()
                ds.Dispose()
            End If

            Dim connect As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (Application.StartupPath & "\" & ComboBox1.Items(ComboBox1.SelectedIndex).ToString).Replace("\\", "\") & ";Persist Security Info=False"
            conn = New OleDbConnection(connect)
            conn.Open()
            da = CreateCustomerAdapter(conn)
            ds = New DataSet
            da.Fill(ds)
            DataGridView1.DataSource = ds.Tables(0)
            DataGridView1.Columns(0).HeaderText = "ID"
            DataGridView1.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            DataGridView1.Columns(1).HeaderText = "Box Name"
            DataGridView1.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            DataGridView1.Columns(2).HeaderText = "Contents"
            DataGridView1.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            DataGridView1.Columns(3).HeaderText = " "
            DataGridView1.Columns(3).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            DataGridView1.Columns(0).ReadOnly = True
            DataGridView1.Columns(1).ReadOnly = True
            DataGridView1.Columns(2).ReadOnly = True
            DataGridView1.Columns(3).ReadOnly = True

            Dim dataGridViewCellStyle2 As DataGridViewCellStyle = New DataGridViewCellStyle
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.SteelBlue
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.Color.Navy
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True
            DataGridView1.Columns(3).DefaultCellStyle = dataGridViewCellStyle2
            ' DataGridView1.EnableHeadersVisualStyles = False

            Button2.Enabled = True
        Catch ex As Exception
            Error_Handler(ex, "[Initialise DataGrid]")
        End Try
    End Sub

    Private Sub DataGridView1_Rowsremoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DataGridView1.RowsRemoved

        Try
            da.Update(ds)
        Catch ex As Exception
            Error_Handler(ex, "[Edit Row Initialize]")
        End Try
    End Sub


    Private Sub DataGridView1_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Try
            If e.ColumnIndex = 3 And e.RowIndex <> -1 Then
                boxdetails1.ClearInfo()
                boxdetails1.TextBox1.Text = DataGridView1.Item(0, e.RowIndex).Value
                boxdetails1.TextBox2.Text = DataGridView1.Item(1, e.RowIndex).Value
                Dim results() As String = DataGridView1.Item(2, e.RowIndex).Value.ToString.Split(";")
                boxdetails1.RichTextBox1.Clear()
                For Each str As String In results
                    If str.Length > 0 Then
                        boxdetails1.RichTextBox1.Text = boxdetails1.RichTextBox1.Text & str & vbCrLf
                    End If
                Next
                If boxdetails1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Dim contentsres As String = ""
                    Dim linetoread() As String = boxdetails1.RichTextBox1.Text.Split(Chr(10))
                    For Each str As String In linetoread
                        str = str.Trim
                        If str.Length > 0 Then
                            contentsres = contentsres & str & ";"
                        End If
                    Next
                    If contentsres.EndsWith(";") Then
                        contentsres = contentsres.Remove(contentsres.Length - 1, 1)
                    End If

                    'MsgBox("Update Boxes set box_ID =" & boxdetails1.TextBox1.Text & ", box_title = '" & boxdetails1.TextBox2.Text & ", ' box_contents = '" & contentsres & "', box_edit = 'Edit' where Box_ID = " & DataGridView1.Item(0, e.RowIndex).Value)
                    cmd = New OleDbCommand("Update Boxes set box_ID =" & boxdetails1.TextBox1.Text & ", box_title = '" & boxdetails1.TextBox2.Text & "',  box_contents = '" & contentsres & "', box_edit = 'Edit' where Box_ID = " & DataGridView1.Item(0, e.RowIndex).Value)


                    Dim connect As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (Application.StartupPath & "\" & ComboBox1.Items(ComboBox1.SelectedIndex).ToString).Replace("\\", "\") & ";Persist Security Info=False"
                    conn = New OleDbConnection(connect)
                    cmd.Connection = conn
                    cmd.Connection.Open()
                    cmd.ExecuteNonQuery()
                    cmd.Connection.Close()
                    InitialiseDataGrid()
                    cmd.Dispose()

                    'DataGridView1.CurrentCell = DataGridView1.Item(0, e.RowIndex)
                    'DataGridView1.CurrentCell.Value = boxdetails1.TextBox1.Text
                    'DataGridView1.CurrentCell = DataGridView1.Item(1, e.RowIndex)
                    'DataGridView1.CurrentCell.Value = boxdetails1.TextBox2.Text
                    'DataGridView1.CurrentCell = DataGridView1.Item(2, e.RowIndex)
                    'DataGridView1.CurrentCell.Value = contentsres
                    
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex, "[Edit Row Initialize]")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            boxdetails1.ClearInfo()

            Dim maxresult As String = 0
            Try
                cmd = New OleDbCommand("Select max(Box_ID) from Boxes")
                Dim connect As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (Application.StartupPath & "\" & ComboBox1.Items(ComboBox1.SelectedIndex).ToString).Replace("\\", "\") & ";Persist Security Info=False"
                conn = New OleDbConnection(connect)
                cmd.Connection = conn
                cmd.Connection.Open()
                maxresult = cmd.ExecuteScalar().ToString()
                cmd.Connection.Close()
                cmd.Dispose()
            Catch ex As Exception
                Error_Handler(ex, "[Learning Maximum ID]")
                maxresult = "0"
            End Try
            Dim max As Integer = 0
            If IsNumeric(maxresult) Then
                max = Integer.Parse(maxresult) + 1
            End If
            boxdetails1.TextBox1.Text = max


            If boxdetails1.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim contentsres As String = ""
                Dim linetoread() As String = boxdetails1.RichTextBox1.Text.Split(Chr(10))
                For Each str As String In linetoread
                    str = str.Trim
                    If str.Length > 0 Then
                        contentsres = contentsres & str & ";"
                    End If
                Next
                If contentsres.EndsWith(";") Then
                    contentsres = contentsres.Remove(contentsres.Length - 1, 1)
                End If

                'DataGridView1.Rows.Add(boxdetails1.TextBox1.Text, boxdetails1.TextBox2.Text, contentsres, "Edit")

                cmd = New OleDbCommand("INSERT INTO [Boxes] (box_ID, box_Title,box_contents) VALUES (" & boxdetails1.TextBox1.Text & ", '" & boxdetails1.TextBox2.Text & "','" & contentsres & "')")


                Dim connect As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (Application.StartupPath & "\" & ComboBox1.Items(ComboBox1.SelectedIndex).ToString).Replace("\\", "\") & ";Persist Security Info=False"
                conn = New OleDbConnection(connect)
                cmd.Connection = conn
                cmd.Connection.Open()
                cmd.ExecuteNonQuery()
                cmd.Connection.Close()
                InitialiseDataGrid()
                cmd.Dispose()
            End If
        Catch ex As Exception
            Error_Handler(ex, "[Add Box Details]")
        End Try
    End Sub


    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            If My.Computer.FileSystem.FileExists(basedatabase) = True Then
                Dim userprompt1 As UserPrompt = New UserPrompt
                If userprompt1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    If My.Computer.FileSystem.FileExists((Application.StartupPath & "\" & userprompt1.TextBox1.Text & ".mdb").Replace("\\", "\").Replace(".mdb.mdb", ".mdb").Replace(".MDB.mdb", ".mdb")) = False Then
                        If userprompt1.TextBox1.Text.Length > 0 Then
                            My.Computer.FileSystem.CopyFile(basedatabase, (Application.StartupPath & "\" & userprompt1.TextBox1.Text & ".mdb").Replace("\\", "\").Replace(".mdb.mdb", ".mdb").Replace(".MDB.mdb", ".mdb"), False)
                            LoadComboBox()
                        End If
                    Else
                        MsgBox("The box set you're trying to create already exists. Please select a unique set name when creating a new box set.", MsgBoxStyle.Information, "Set Name Already Exists")
                    End If

                End If
                userprompt1.Dispose()
            End If
        Catch ex As Exception
            Error_Handler(ex, "[Creating new box set]")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            If ComboBox1.SelectedIndex <> -1 Then


                If conn IsNot Nothing Then
                    conn.Close()
                    da.Dispose()
                    ds.Dispose()
                End If
                Dim filetodelete As String = (Application.StartupPath & "\" & ComboBox1.Items(ComboBox1.SelectedIndex).ToString).Replace("\\", "\")
                If My.Computer.FileSystem.FileExists(filetodelete) = True Then
                    If ComboBox1.Items.Count > 1 Then
                        ComboBox1.Items.RemoveAt(ComboBox1.SelectedIndex)
                        lastselectedindex = -2
                        ComboBox1.SelectedIndex = 0
                    Else
                        ComboBox1.Text = ""
                        ComboBox1.Items.Clear()
                        lastselectedindex = -2
                        DataGridView1.DataSource = Nothing
                        Button2.Enabled = False
                        GroupBox1.Text = "No Box Set Located"
                    End If
                    My.Computer.FileSystem.DeleteFile(filetodelete, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex, "[Deleting box set]")
        End Try
    End Sub

    Private Sub LinkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Try
            LoadComboBox()
        Catch ex As Exception
            Error_Handler(ex, "[Refresh Combobox Contents]")
        End Try
    End Sub

    Public Function CreateCustomerAdapter(ByVal connection As OleDbConnection) As OleDbDataAdapter

        Dim adapter As OleDbDataAdapter = New OleDbDataAdapter()

        ' Create the SelectCommand.
        Dim command As OleDbCommand = New OleDbCommand( _
            "SELECT * FROM boxes ", connection)

        adapter.SelectCommand = command

        ' Create the InsertCommand.
        command = New OleDbCommand( _
            "INSERT INTO Boxes (Box_ID, Box_Title,Box_Contents,Box_Edit) " & _
            "VALUES (@BoxID, @BoxTitle, @BoxContents,@Box_Edit)", connection)

        ' Add the parameters for the InsertCommand.
        command.Parameters.Add("@BoxID", OleDbType.Integer, 10, "Box_ID")
        command.Parameters.Add("@BoxTitle", OleDbType.WChar, 255, "Box_Title")
        command.Parameters.Add("@BoxContents", OleDbType.WChar, 65535, "Box_Contents")
        command.Parameters.Add("@BoxEdit", OleDbType.WChar, 4, "Box_Edit")
        adapter.InsertCommand = command

        ' Create the UpdateCommand.
        command = New OleDbCommand( _
            "UPDATE Boxes SET Box_ID = @BoxID, Box_Title = @BoxTitle, Box_Contents = @BoxContents, Box_Edit = @BoxEdit " & _
            "WHERE Box_ID = @oldBoxID", connection)

        ' Add the parameters for the UpdateCommand.
        command.Parameters.Add("@BoxID", OleDbType.Integer, 10, "Box_ID")
        command.Parameters.Add("@BoxTitle", OleDbType.WChar, 255, "Box_Title")
        command.Parameters.Add("@BoxContents", OleDbType.WChar, 65535, "Box_Contents")
        command.Parameters.Add("@BoxEdit", OleDbType.WChar, 4, "Box_Edit")
        command.Parameters.Add("@CustomerID", SqlDbType.NChar, 5, "CustomerID")
        command.Parameters.Add("@CompanyName", SqlDbType.NVarChar, 40, "CompanyName")

        Dim parameter As OleDbParameter = command.Parameters.Add( _
            "@oldBoxID", OleDbType.Integer, 10, "Box_ID")
        parameter.SourceVersion = DataRowVersion.Original
        command.Parameters.Add("@oldBoxID", OleDbType.Integer, 10, "Box_ID")

        adapter.UpdateCommand = command

        ' Create the DeleteCommand.
        command = New OleDbCommand( _
            "DELETE FROM Boxes WHERE Box_ID = @BoxID", connection)

        ' Add the parameters for the DeleteCommand.
        command.Parameters.Add( _
            "@BoxID", OleDbType.Integer, 10, "Box_ID")
        parameter.SourceVersion = DataRowVersion.Original

        adapter.DeleteCommand = command

        Return adapter
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Dim reader As OleDbDataReader
            cmd = New OleDbCommand("Select * from Boxes where Box_Title like '%" & TextBox1.Text.Replace(" ", "%") & "%' union select * from Boxes where Box_Contents like '%" & TextBox1.Text.Replace(" ", "%") & "%' order by box_id")
            Dim connect As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (Application.StartupPath & "\" & ComboBox1.Items(ComboBox1.SelectedIndex).ToString).Replace("\\", "\") & ";Persist Security Info=False"
            conn = New OleDbConnection(connect)
            cmd.Connection = conn
            cmd.Connection.Open()
            reader = cmd.ExecuteReader()
            results1 = New Results
            Dim top As Integer = 10
            Dim left As Integer = 10
            Dim counter As Integer = 0
            If reader.HasRows = True Then

                While (reader.Read())
                    Dim lbl As Label = New Label
                    lbl.AutoSize = True
                    lbl.Text = reader("Box_ID").ToString() & ": " & reader("Box_Title").ToString()
                    lbl.Top = top
                    top = top + lbl.Height
                    lbl.Left = left
                    results1.Panel1.Controls.Add(lbl)
                    Dim lbl2 As Label = New Label
                    lbl2.AutoSize = True
                    lbl2.Text = reader("Box_Contents").ToString()
                    lbl2.Top = top
                    top = top + lbl2.Height
                    lbl2.ForeColor = Color.Gray
                    lbl2.Left = left + 20
                    results1.Panel1.Controls.Add(lbl2)
                    counter = counter + 1
                End While
            End If
            reader.Close()
            

            If counter <> 1 Then
                results1.Label1.Text = counter & " results matched your search terms"
            Else
                results1.Label1.Text = counter & " result matched your search terms"
            End If


            cmd.Connection.Close()
            cmd.Dispose()
            results1.Show()
            conn.Dispose()
        Catch ex As Exception
            Error_Handler(ex, "[Search Function]")
        End Try
    End Sub
End Class
