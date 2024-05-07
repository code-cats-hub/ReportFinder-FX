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
        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()
        GC.WaitForPendingFinalizers()
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
    Private Sub BT_FG01_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG01.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG02_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG02.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG03_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG03.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG04_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG04.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG05_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG05.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG06_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG06.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG07_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG07.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG08_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG08.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG09_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG09.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG10_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG10.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG11_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG11.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG12_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG12.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG13_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG13.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG14_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG14.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG15_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG15.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG16_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG16.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG17_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG17.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG18_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG18.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG19_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG19.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG20_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG20.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG21_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG21.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG22_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG22.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG23_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG23.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG24_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG24.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_FG25_CheckedChanged(sender As Object, e As EventArgs) Handles BT_FG25.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_SG01_CheckedChanged(sender As Object, e As EventArgs) Handles BT_SG01.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_SG02_CheckedChanged(sender As Object, e As EventArgs) Handles BT_SG02.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_SG03_CheckedChanged(sender As Object, e As EventArgs) Handles BT_SG03.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_SG04_CheckedChanged(sender As Object, e As EventArgs) Handles BT_SG04.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_SG05_CheckedChanged(sender As Object, e As EventArgs) Handles BT_SG05.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_SG06_CheckedChanged(sender As Object, e As EventArgs) Handles BT_SG06.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_SG07_CheckedChanged(sender As Object, e As EventArgs) Handles BT_SG07.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_SG08_CheckedChanged(sender As Object, e As EventArgs) Handles BT_SG08.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_SG09_CheckedChanged(sender As Object, e As EventArgs) Handles BT_SG09.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_SG10_CheckedChanged(sender As Object, e As EventArgs) Handles BT_SG10.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_SG11_CheckedChanged(sender As Object, e As EventArgs) Handles BT_SG11.CheckedChanged
        ActiveControl = Nothing
    End Sub
    Private Sub BT_SG12_CheckedChanged(sender As Object, e As EventArgs) Handles BT_SG12.CheckedChanged
        ActiveControl = Nothing
    End Sub

End Class
