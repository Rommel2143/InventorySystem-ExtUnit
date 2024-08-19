Imports MySql.Data.MySqlClient
Public Class results_IN
    Dim itempartcode As String

    Private Sub scan_results_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtpicker.Value = Date.Now.ToString("yyyy-MM-dd")
    End Sub



    Private Sub datagrid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles datagrid1.CellClick
        Try
            With datagrid1
                itempartcode = .SelectedCells(0).Value.ToString()
            End With

            con.Close()
            con.Open()
            Dim cmdrefreshgrid As New MySqlCommand("SELECT ts.`qrcode`,ts.`partcode`,ts.`lotnumber`, ts.`remarks`, ts.`qty`
                                                    FROM `unit10_tblscan` ts
                                                    LEFT JOIN scanoperator_is so ON ts.userin = so.IDno
                                                    WHERE       `datein`='" & dtpicker.Value.ToString("yyyy-MM-dd") & "' 
                                                    and `partcode`='" & itempartcode & "'
                                                            and `located`='" & PClocation & "' 
                                                            and `Fullname`='" & cmbuser.Text & "'  
                                                            and `batch`='" & cmbbatchin.Text & "' ", con)

            Dim da As New MySqlDataAdapter(cmdrefreshgrid)
            Dim dt As New DataTable
            da.Fill(dt)
            datagrid2.DataSource = dt
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub cmbbatchout_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles cmbbatchin.SelectedIndexChanged
        Try


            con.Close()
            con.Open()
            Dim cmdrefreshgrid As New MySqlCommand("SELECT ts.`qrcode`,ts.`partcode`,ts.`lotnumber`, ts.`remarks`, ts.`qty`
                                                    FROM `unit10_tblscan` ts
                                                    LEFT JOIN scanoperator_is so ON ts.userin = so.IDno
                                                    WHERE       `datein`='" & dtpicker.Value.ToString("yyyy-MM-dd") & "' 
                                                            and `located`='" & PClocation & "' 
                                                            and `Fullname`='" & cmbuser.Text & "'  
                                                            and `batch`='" & cmbbatchin.Text & "' ", con)

            Dim da As New MySqlDataAdapter(cmdrefreshgrid)
            Dim dt As New DataTable
            da.Fill(dt)
            datagrid2.DataSource = dt


            con.Close()
            con.Open()
            Dim cmdrefreshgrid2 As New MySqlCommand("SELECT ts.`partcode` AS Partcode, SUM(`qty`) AS TOTAL 
                                                  FROM `unit10_tblscan` ts
                                                    LEFT JOIN scanoperator_is so ON ts.userin = so.IDno
                                                    WHERE       `datein`='" & dtpicker.Value.ToString("yyyy-MM-dd") & "' 
                                                            and `located`='" & PClocation & "' 
                                                            and `Fullname`='" & cmbuser.Text & "'  
                                                            and `batch`='" & cmbbatchin.Text & "'          
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

    Private Sub dtpicker_ValueChanged_1(sender As Object, e As EventArgs) Handles dtpicker.ValueChanged
        Try
            con.Close()
            con.Open()
            Dim cmdselect As New MySqlCommand("Select distinct so.fullname FROM `unit10_tblscan` uts
                                                INNER JOIN `scanoperator_is` so ON uts.userin = so.IDno
                                                WHERE uts.located='" & PClocation & "' and uts.datein='" & dtpicker.Value.ToString("yyyy-MM-dd") & "'", con)
            dr = cmdselect.ExecuteReader
            cmbuser.Items.Clear()
            While (dr.Read())
                cmbuser.Items.Add(dr.GetString("fullname"))
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub cmbuser_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbuser.SelectedIndexChanged
        Try
            con.Close()
            con.Open()
            Dim cmdselect As New MySqlCommand("Select distinct ts.`batch` FROM `unit10_tblscan` ts
                                              Left Join scanoperator_is tsoout ON ts.userin = tsoout.IDno
                                               
                                                WHERE `located`='" & PClocation & "' and `datein`='" & dtpicker.Value.ToString("yyyy-MM-dd") & "' and `fullname`='" & cmbuser.Text & "'", con)
            dr = cmdselect.ExecuteReader
            cmbbatchin.Items.Clear()
            While (dr.Read())
                cmbbatchin.Items.Add(dr.GetString("batch"))
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

End Class