Imports System
Public Class sub_FRAME

    Private Sub Scan_Frame_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        userstrip.Text = fname
    End Sub

    Private Sub PARTSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PARTSToolStripMenuItem.Click
        display_formscan(Scan_In)
    End Sub
    Private Sub FGToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FGToolStripMenuItem.Click
        display_formscan(Scan_out)
    End Sub
    Private Sub display_formscan(form As Form)
        With form
            .Refresh()
            .TopLevel = False
            Panel1.Controls.Add(form)
            .BringToFront()
            .Show()

        End With
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub unitdisplay_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        con.Close()
        Application.Exit()

    End Sub

    Private Sub LogoutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogoutToolStripMenuItem.Click
        display_form(Login)
        Login.txtbarcode.Clear()
    End Sub

    Private Sub DeviceInfoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeviceInfoToolStripMenuItem.Click
        MessageBox.Show("Device loc:" & PClocation & "   /  Mac:" & PCmac & "   /  Device:" & PCname & "", "This Device is Registered", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub SuggestToImproveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SuggestToImproveToolStripMenuItem.Click
        suggest_improvent.Show()
        suggest_improvent.BringToFront()
    End Sub

    Private Sub RETURNToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RETURNToolStripMenuItem.Click
        display_formscan(Scan_return)
    End Sub

    Private Sub QRCheckerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles QRCheckerToolStripMenuItem.Click

    End Sub

    Private Sub QRCheckerToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles QRCheckerToolStripMenuItem1.Click
        display_formscan(QR_checker)
    End Sub

    Private Sub StockMonitoringToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StockMonitoringToolStripMenuItem.Click
        display_formscan(stock_monitoring)
    End Sub

    Private Sub ManageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ManageToolStripMenuItem.Click
        display_formscan(Add_item)
    End Sub
End Class