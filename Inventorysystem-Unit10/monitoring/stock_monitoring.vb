Imports MySql.Data.MySqlClient
Public Class stock_monitoring
    Private Sub stock_monitoring_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        refreshgrid()
    End Sub
    Private Sub refreshgrid()
        Try
            con.Close()
            con.Open()
            Dim cmdrefreshgrid As New MySqlCommand("SELECT `partcode`, `partname`,  `section`, `stockU10` FROM `unit10_tblmaster` ORDER BY `partcode` ASC", con)

            Dim da As New MySqlDataAdapter(cmdrefreshgrid)
            Dim dt As New DataTable
            da.Fill(dt)
            datagrid1.DataSource = dt
            datagrid1.AutoResizeColumns()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally

            con.Close()
        End Try
    End Sub

    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles Guna2Button1.Click
        refreshgrid()

    End Sub

    Private Sub cmbsearch_TextChanged(sender As Object, e As EventArgs) Handles cmbsearch.TextChanged
        Try
            con.Close()
            con.Open()
            Dim cmdrefreshgrid As New MySqlCommand("SELECT `partcode`, `partname`,  `section`, `stockU10` FROM `unit10_tblmaster` WHERE `partcode` REGEXP '" & cmbsearch.Text & "' or `partname` REGEXP '" & cmbsearch.Text & "'ORDER BY `partcode` ASC", con)

            Dim da As New MySqlDataAdapter(cmdrefreshgrid)
            Dim dt As New DataTable
            da.Fill(dt)
            datagrid1.DataSource = dt
            datagrid1.AutoResizeColumns()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally

            con.Close()
        End Try
    End Sub
End Class