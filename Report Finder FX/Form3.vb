'--- GPL COPYRIGHT 2024 - CODE-CATS https://orcid.org/0009-0006-7849-1462 --- 
Public Class Form3
    Private Sub BT_BACK_Click(sender As Object, e As EventArgs) Handles BT_BACK.Click
        Call INFO_PANEL_OFF()
    End Sub
    Private Sub BT_GPL_Click(sender As Object, e As EventArgs) Handles BT_GPL.Click
        Call OPEN_GPLN()
    End Sub
End Class