'--- GPL COPYRIGHT 2024 - CODE-CATS https://orcid.org/0009-0006-7849-1462 --- 
Public Class Form1
    Private Sub BT_SEARCH_Click(sender As Object, e As EventArgs) Handles BT_SEARCH.Click
        If unlocked = True Then
            Call START_SEARCH()
        End If
    End Sub
    Private Sub BT_EXIT_Click(sender As Object, e As EventArgs) Handles BT_EXIT.Click
        Call SHUT_DOWN()
    End Sub
    Private Sub LOAD_BUTTON_Click(sender As Object, e As EventArgs) Handles LOAD_BUTTON.Click
        If unlocked = False Then
            Call CATALOG_LOAD()
            Call FUNCTIONS_LOAD()
        End If
    End Sub
    Private Sub BT_FG_SEL_Click(sender As Object, e As EventArgs) Handles BT_FG_SEL.Click
        Call FG_SELECT_ALL()
    End Sub
    Private Sub BT_FG_DS_Click(sender As Object, e As EventArgs) Handles BT_FG_DS.Click
        Call FG_DESELECT_ALL()
    End Sub
    Private Sub BT_SG_SEL_Click(sender As Object, e As EventArgs) Handles BT_SG_SEL.Click
        Call SG_SELECT_ALL()
    End Sub
    Private Sub BT_SG_DS_Click(sender As Object, e As EventArgs) Handles BT_SG_DS.Click
        Call SG_DESELECT_ALL()
    End Sub
    Private Sub BT_INFO_Click(sender As Object, e As EventArgs) Handles BT_INFO.Click
        Call INFO_PANEL_ON()
    End Sub
    Private Sub BT_SET_Click(sender As Object, e As EventArgs) Handles BT_SET.Click
        Call SET_PANEL_ON()
    End Sub
End Class
