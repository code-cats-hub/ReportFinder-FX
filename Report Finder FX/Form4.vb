'--- GPL COPYRIGHT 2024 - CODE-CATS https://orcid.org/0009-0006-7849-1462 --- 
Public Class Form4
    Private Sub BT_BACK_Click(sender As Object, e As EventArgs) Handles BT_BACK.Click
        Call SET_PANEL_OFF()
    End Sub
    Private Sub BT_EXIT_Click(sender As Object, e As EventArgs) Handles BT_EXIT.Click
        Call SHUT_DOWN()
    End Sub
    Private Sub BT_INFO_Click(sender As Object, e As EventArgs) Handles BT_INFO.Click
        Call INFO_PANEL_ON()
    End Sub
    Private Sub BT_PREVIEW_Click(sender As Object, e As EventArgs) Handles BT_PREVIEW.Click
        Call PREVIEW_FILLER()
    End Sub
    Private Sub BT_RELOAD_Click(sender As Object, e As EventArgs) Handles BT_RELOAD.Click
        Call CATALOG_LOAD()
    End Sub
    Private Sub BT_SOURCE_Click(sender As Object, e As EventArgs) Handles BT_SOURCE.Click
        Call SOURCE_LOAD()
    End Sub
    Private Sub BT_CROMEFIND_Click(sender As Object, e As EventArgs) Handles BT_CROMEFIND.Click
        CHROME_FINDER()
    End Sub
    Private Sub BT_LINKS_CheckedChanged(sender As Object, e As EventArgs) Handles BT_LINKS.CheckedChanged
        Call LINK_PANEL_MODE()
    End Sub
    Private Sub BT_FOX_OFF_Click(sender As Object, e As EventArgs) Handles BT_FOX_OFF.Click
        Call FIREFOX_FINDER()
    End Sub
    Private Sub BT_FOX_ON_Click(sender As Object, e As EventArgs) Handles BT_FOX_ON.Click
        Call FIREFOX_REMOVER()
    End Sub
    Private Sub BT_LINK_1_Click(sender As Object, e As EventArgs) Handles BT_LINK_1.Click
        Call LINK2BROWSER("https://www.google.com/")
    End Sub
    Private Sub BT_LINK_2_Click(sender As Object, e As EventArgs) Handles BT_LINK_2.Click
        Call LINK2BROWSER("https://www.google.com/")
    End Sub
End Class