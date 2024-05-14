'--- GPL COPYRIGHT 2024 - CODE-CATS https://orcid.org/0009-0006-7849-1462 --- 
Imports Excel = Microsoft.Office.Interop.Excel
Module Module4
    Public Sub PREVIEW_FILLER()

        Call PREV_FILL_COLUMNS()
        Call PREV_FILL_ROWS()

    End Sub
    Private Sub PREV_FILL_COLUMNS()

        Dim i As Integer
        Dim nc As New DataGridViewTextBoxColumn

        '--- TABLE ST1 ---
        Form4.ST_TABLE1.Columns.Clear()

        nc = New DataGridViewTextBoxColumn
        With nc
            .Name = "ID_R1"
            .HeaderText = "ID-R"
        End With
        Form4.ST_TABLE1.Columns.Add(nc)

        For i = 1 To dim_FG_option
            Dim nc2 As New DataGridViewTextBoxColumn
            With nc2
                .Name = arr_SEARCH_FG(i, 1)
                .HeaderText = arr_SEARCH_FG(i, 1)
            End With
            Form4.ST_TABLE1.Columns.Add(nc2)
        Next

        '--- TABLE ST2 ---
        Form4.ST_TABLE2.Columns.Clear()

        nc = New DataGridViewTextBoxColumn
        With nc
            .Name = "ID_R2"
            .HeaderText = "ID-R"
        End With
        Form4.ST_TABLE2.Columns.Add(nc)

        For i = 1 To dim_SG_option
            Dim nc2 As New DataGridViewTextBoxColumn
            With nc2
                .Name = arr_SEARCH_SG(i, 1)
                .HeaderText = arr_SEARCH_SG(i, 1)
            End With
            Form4.ST_TABLE2.Columns.Add(nc2)
        Next

        '--- TABLE SB2 ---
        Form4.SB_TABLE2.Columns.Clear()

        nc = New DataGridViewTextBoxColumn
        With nc
            .Name = "ID_FG"
            .HeaderText = "ID-FG"
        End With
        Form4.SB_TABLE2.Columns.Add(nc)

        nc = New DataGridViewTextBoxColumn
        With nc
            .Name = "SF"
            .HeaderText = "SEARCH FIELDS"
        End With
        Form4.SB_TABLE2.Columns.Add(nc)

        '--- TABLE SB3 ---
        Form4.SB_TABLE3.Columns.Clear()

        nc = New DataGridViewTextBoxColumn
        With nc
            .Name = "ID_SG"
            .HeaderText = "ID-SG"
        End With
        Form4.SB_TABLE3.Columns.Add(nc)

        nc = New DataGridViewTextBoxColumn
        With nc
            .Name = "RT"
            .HeaderText = "RESULT TYPES"
        End With
        Form4.SB_TABLE3.Columns.Add(nc)

        '--- TABLE SB1 ---
        Form4.SB_TABLE1.Columns.Clear()

        nc = New DataGridViewTextBoxColumn
        With nc
            .Name = "ID_R3"
            .HeaderText = "ID-R"
        End With
        Form4.SB_TABLE1.Columns.Add(nc)

        nc = New DataGridViewTextBoxColumn
        With nc
            .Name = "M01"
            .HeaderText = "NAME"
        End With
        Form4.SB_TABLE1.Columns.Add(nc)

        nc = New DataGridViewTextBoxColumn
        With nc
            .Name = "M02"
            .HeaderText = "SEARCH FIELDS"
        End With
        Form4.SB_TABLE1.Columns.Add(nc)

        nc = New DataGridViewTextBoxColumn
        With nc
            .Name = "M03"
            .HeaderText = "NOTES"
        End With
        Form4.SB_TABLE1.Columns.Add(nc)

        nc = New DataGridViewTextBoxColumn
        With nc
            .Name = "M04"
            .HeaderText = "COGNOS PATH"
        End With
        Form4.SB_TABLE1.Columns.Add(nc)

        nc = New DataGridViewTextBoxColumn
        With nc
            .Name = "BTC"
            .HeaderText = "BROWSER LINK"
        End With
        Form4.SB_TABLE1.Columns.Add(nc)

    End Sub
    Private Sub PREV_FILL_ROWS()

        Dim x, y As Integer
        Dim list_ROW As New List(Of String)

        '--- TABLE ST1 ---
        Form4.ST_TABLE1.Rows.Clear()

        For x = 1 To dim_RKEY
            For y = 1 To dim_FG_cap
                list_ROW.Add(arr_MAP_FG(x, y))
            Next
            Form4.ST_TABLE1.Rows.Add(list_ROW.ToArray)
            list_ROW.Clear()
        Next

        '--- TABLE ST2 ---
        Form4.ST_TABLE2.Rows.Clear()

        For x = 1 To dim_RKEY
            For y = 1 To dim_SG_cap
                list_ROW.Add(arr_MAP_SG(x, y))
            Next
            Form4.ST_TABLE2.Rows.Add(list_ROW.ToArray)
            list_ROW.Clear()
        Next

        '--- TABLE SB1 ---
        Form4.SB_TABLE1.Rows.Clear()

        For x = 1 To dim_RKEY
            For y = 1 To 6
                list_ROW.Add(arr_MAP_RES(x, y))
            Next
            Form4.SB_TABLE1.Rows.Add(list_ROW.ToArray)
            list_ROW.Clear()
        Next

        '--- TABLE SB2 ---
        Form4.SB_TABLE2.Rows.Clear()

        For x = 1 To dim_FG_option
            For y = 1 To 2
                list_ROW.Add(arr_SEARCH_FG(x, y))
            Next
            Form4.SB_TABLE2.Rows.Add(list_ROW.ToArray)
            list_ROW.Clear()
        Next

        '--- TABLE SB3 ---
        Form4.SB_TABLE3.Rows.Clear()

        For x = 1 To dim_SG_option
            For y = 1 To 2
                list_ROW.Add(arr_SEARCH_SG(x, y))
            Next
            Form4.SB_TABLE3.Rows.Add(list_ROW.ToArray)
            list_ROW.Clear()
        Next

    End Sub
    Public Sub LINK_PANEL_MODE()

        If Form4.BT_LINKS.Checked = True Then
            Form4.LINK_PANEL.Visible = True
            Form4.LINK_PANEL_PIPE.Visible = True
            Form4.LINK_PANEL_SHADOW1.Visible = True
            Form4.LINK_PANEL_SHADOW2.Visible = True
        Else
            Form4.LINK_PANEL.Visible = False
            Form4.LINK_PANEL_PIPE.Visible = False
            Form4.LINK_PANEL_SHADOW1.Visible = False
            Form4.LINK_PANEL_SHADOW2.Visible = False
        End If

    End Sub
    Public Sub SOURCE_LOAD()
        Dim CALL_EXCEL As New Excel.Application

        CALL_EXCEL.Workbooks.Open(path_user & "Data\REPORTMATRIX.xlsx")
        CALL_EXCEL.Visible = True

    End Sub
End Module
