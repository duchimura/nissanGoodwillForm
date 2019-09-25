'set tab index
'msgbox syntax Is msgbox.show("")
'change titlebar icon to nissan
'make type frame font red on start

Public Class gwForm
    Public Sub gwForm_load(sender As Object, e As EventArgs) Handles MyBase.Load
        'include these items in initForm
        Me.Height = 118
        Me.Width = 490
        Me.CenterToScreen()
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        'selectTypeFrame.RectangleToClient = RectangleToScreen()
        'initializeform
    End Sub

    Private Sub RepairOnlyOB_CheckedChanged(sender As Object, e As EventArgs) Handles repairOnlyOB.CheckedChanged
        Me.Height = 583
        Me.Width = 980
        Me.CenterToScreen()
        selectTypeFrame.ForeColor = Color.Black
        If repairOnlyOB.Checked = True Then
            repairOnlyOB.Font = New Font(repairOnlyOB.Font, FontStyle.Bold)
        Else
            repairOnlyOB.Font = New Font(repairOnlyOB.Font, FontStyle.Regular)
        End If
    End Sub

    Private Sub RentalOnlyOB_CheckedChanged(sender As Object, e As EventArgs) Handles rentalOnlyOB.CheckedChanged
        Me.Height = 583
        Me.Width = 980
        Me.CenterToScreen()
        selectTypeFrame.ForeColor = Color.Black
        If rentalOnlyOB.Checked = True Then
            rentalOnlyOB.Font = New Font(rentalOnlyOB.Font, FontStyle.Bold)
        Else
            rentalOnlyOB.Font = New Font(rentalOnlyOB.Font, FontStyle.Regular)
        End If
    End Sub

    Private Sub RepairAndRentalOB_CheckedChanged(sender As Object, e As EventArgs) Handles repairAndRentalOB.CheckedChanged
        Me.Height = 583
        Me.Width = 980
        Me.CenterToScreen()
        selectTypeFrame.ForeColor = Color.Black
        If repairAndRentalOB.CheckAlign = True Then
            repairAndRentalOB.Font = New Font(repairAndRentalOB.Font, FontStyle.Bold)
        Else
            repairAndRentalOB.Font = New Font(repairAndRentalOB.Font, FontStyle.Regular)
        End If
    End Sub

    Private Sub ResetBtn_Click(sender As Object, e As EventArgs) Handles resetBtn.Click
        'initForm goes here
        Me.Height = 118
        Me.Width = 490
        Me.CenterToScreen()
        Me.repairOnlyOB.Checked = False
        Me.rentalOnlyOB.Checked = False
        Me.repairAndRentalOB.Checked = False
        repairAndRentalOB.Font = New Font(repairAndRentalOB.Font, FontStyle.Regular)
        rentalOnlyOB.Font = New Font(rentalOnlyOB.Font, FontStyle.Regular)
        repairOnlyOB.Font = New Font(repairOnlyOB.Font, FontStyle.Regular)
    End Sub

    Private Sub emailBtn_Click(sender As Object, e As EventArgs) Handles emailBtn.Click

    End Sub
End Class
