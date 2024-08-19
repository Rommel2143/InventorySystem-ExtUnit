Imports MySql.Data.MySqlClient
Public Class results_OUT
    Dim itempartcode As String

    Private Sub scan_results_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtpicker.Value = Date.Now.ToString("yyyy-MM-dd")
    End Sub


    Private Sub dtpicker_ValueChanged(sender As Object, e As EventArgs) Handles dtpicker.ValueChanged
        Try
            con.Close()
            con.Open()
            Dim cmdselect As New MySqlCommand("Select distinct `batchout` FROM `unit10_tblscan`
                                                WHERE `dateout`='" & dtpicker.Value.ToString("yyyy-MM-dd") & "' and `located`='" & PClocation & "' and `userout`='" & idno & "' ", con)
            dr = cmdselect.ExecuteReader
            cmbbatchin.Items.Clear()
            While (dr.Read())
                cmbbatchin.Items.Add(dr.GetString("batchout"))
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub


    Private Sub cmbbatchout_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbbatchin.SelectedIndexChanged
        Try
            con.Close()
            con.Open()
            Dim cmdrefreshgrid As New MySqlCommand("SELECT `qrcode`,`partcode`,`lotnumber`, `remarks`, `qty` FROM `unit10_tblscan`
                                                     WHERE `dateout`='" & dtpicker.Value.ToString("yyyy-MM-dd") & "' and `located`='" & PClocation & "' and `userout`='" & idno & "' and `batchout`='" & cmbbatchin.Text & "' ", con)

            Dim da As New MySqlDataAdapter(cmdrefreshgrid)
            Dim dt As New DataTable
            da.Fill(dt)
            datagrid2.DataSource = dt


            con.Close()
            con.Open()
            Dim cmdrefreshgrid2 As New MySqlCommand("SELECT `partcode` AS Partcode, SUM(`qty`) AS TOTAL FROM `unit10_tblscan`
                                                  WHERE `dateout`='" & dtpicker.Value.ToString("yyyy-MM-dd") & "' and `located`='" & PClocation & "' and `userout`='" & idno & "' and `batchout`='" & cmbbatchin.Text & "'              
                                                  GROUP BY partcode", con)

            Dim da2 As New MySqlDataAdapter(cmdrefreshgrid2)
            Dim dt2 As New DataTable
            da2.Fill(dt2)
            datagrid1.DataSource = dt2
            datagrid1.AutoResizeColumns()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub datagrid1_CellContextMenuStripChanged(sender As Object, e As DataGridViewCellEventArgs) Handles datagrid1.CellContextMenuStripChanged

    End Sub

    Private Sub datagrid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles datagrid1.CellClick
        Try
            With datagrid1
                itempartcode = .SelectedCells(0).Value.ToString()
            End With
            con.Close()
            con.Open()
            Dim cmdrefreshgrid As New MySqlCommand("SELECT `qrcode`,`partcode`,`lotnumber`, `remarks`, `qty` FROM `unit10_tblscan`
                                                     WHERE `partcode`='" & itempartcode & "' and `dateout`='" & dtpicker.Value.ToString("yyyy-MM-dd") & "' and `located`='" & PClocation & "' and `userout`='" & idno & "' and `batchout`='" & cmbbatchin.Text & "' ", con)

            Dim da As New MySqlDataAdapter(cmdrefreshgrid)
            Dim dt As New DataTable
            da.Fill(dt)
            datagrid2.DataSource = dt
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
End Class