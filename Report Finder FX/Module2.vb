'--- GPL COPYRIGHT 2024 - CODE-CATS https://orcid.org/0009-0006-7849-1462 --- 
Imports Excel = Microsoft.Office.Interop.Excel
Module Module2

    '--- EXTERNAL TABLES ---
    Public arr_SEARCH_FG(100, 2) As String
    Public arr_SEARCH_SG(100, 2) As String
    Public arr_MAP_FG(100, 100) As String
    Public arr_MAP_SG(100, 100) As String
    Public arr_MAP_RES(100, 6) As String
    '--- DIMENSION WRITE ---
    Public dim_RKEY As Integer
    Public dim_FG_option As Integer
    Public dim_SG_option As Integer
    Public dim_FG_cap As Integer
    Public dim_SG_cap As Integer
    Public unlocked As Boolean = False

    Public Sub CATALOG_LOAD()

        Form1.LOAD_BAR.Value = 0
        Form4.LOAD_BAR3.Value = 0

        Dim x, y, z As Integer
        Dim filled As Integer
        Dim CALL_EXCEL As New Excel.Application

        CALL_EXCEL.Workbooks.Open(path_user & "Data\REPORT MATRIX.xlsx")
        filled = 0

        Dim CALL_WS1 As Excel.Worksheet = CALL_EXCEL.Sheets(1)
        Dim CALL_WS2 As Excel.Worksheet = CALL_EXCEL.Sheets(2)
        Dim CALL_WS3 As Excel.Worksheet = CALL_EXCEL.Sheets(3)
        Dim CALL_WS4 As Excel.Worksheet = CALL_EXCEL.Sheets(4)
        Dim CALL_WS5 As Excel.Worksheet = CALL_EXCEL.Sheets(5)

        dim_RKEY = CALL_WS1.Range("A" & CALL_WS1.Rows.Count).End(Excel.XlDirection.xlUp).Row - 2
        ReDim arr_MAP_RES(dim_RKEY, 6)

        '--- LOAD SEARCH FG options ---'

        z = CALL_WS4.Range("A" & CALL_WS4.Rows.Count).End(Excel.XlDirection.xlUp).Row

        For x = 1 To z
            If CALL_WS4.Cells(x + 1, 2).Value <> "" Then filled = filled + 1
        Next

        dim_FG_option = filled
        dim_FG_cap = filled + 1
        ReDim arr_SEARCH_FG(dim_FG_option, 2)
        ReDim arr_MAP_FG(dim_RKEY, dim_FG_cap)

        For x = 1 To filled
            For y = 1 To 2
                arr_SEARCH_FG(x, y) = CALL_WS4.Cells(x + 1, y).Value
            Next
        Next

        z = 0
        filled = 0

        Call PAUSE()
        Form1.LOAD_BAR.Value = 20
        Form4.LOAD_BAR3.Value = 20

        '--- LOAD SEARCH SG options ---'

        z = CALL_WS5.Range("A" & CALL_WS5.Rows.Count).End(Excel.XlDirection.xlUp).Row

        For x = 1 To z
            If CALL_WS5.Cells(x + 1, 2).Value <> "" Then filled = filled + 1
        Next

        dim_SG_option = filled
        dim_SG_cap = filled + 1
        ReDim arr_SEARCH_SG(dim_SG_option, 2)
        ReDim arr_MAP_SG(dim_RKEY, dim_SG_cap)

        For x = 1 To filled
            For y = 1 To 2
                arr_SEARCH_SG(x, y) = CALL_WS5.Cells(x + 1, y).Value
            Next
        Next

        z = 0
        filled = 0

        Call PAUSE()
        Form1.LOAD_BAR.Value = 40
        Form4.LOAD_BAR3.Value = 40

        '--- LOAD SEARCH FG MAP table ---'

        For x = 1 To dim_RKEY
            For y = 1 To dim_FG_cap
                arr_MAP_FG(x, y) = CALL_WS1.Cells(x + 2, y).Value
            Next
        Next

        Form1.LOAD_BAR.Value = 60
        Form4.LOAD_BAR3.Value = 60

        '--- LOAD SEARCH SG MAP table ---'

        For x = 1 To dim_RKEY
            For y = 1 To dim_SG_cap
                arr_MAP_SG(x, y) = CALL_WS2.Cells(x + 2, y).Value
            Next
        Next

        Call PAUSE()
        Form1.LOAD_BAR.Value = 80
        Form4.LOAD_BAR3.Value = 80

        '--- LOAD RESULT MAP table ---'

        For x = 1 To dim_RKEY
            For y = 1 To 6
                arr_MAP_RES(x, y) = CALL_WS3.Cells(x + 2, y).Value
            Next
        Next

        Call PAUSE()
        Form1.LOAD_BAR.Value = 100
        Form4.LOAD_BAR3.Value = 100

        CALL_EXCEL.Workbooks.Close()
        CALL_EXCEL.Quit()

        GC.Collect()
        GC.WaitForPendingFinalizers()

    End Sub

    Public Sub FUNCTIONS_LOAD()

        Form1.LOAD_BAR2.Value = 0

        Dim i As Integer

        '--- FILL FG BUTTONS ---

        For Each ctrl As Control In Form1.Controls
            If TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                Dim cb As System.Windows.Forms.CheckBox = ctrl
                If Mid(cb.Name, 4, 2) = "FG" Then
                    If CInt(Right(cb.Name, 2)) <= dim_FG_option Then
                        cb.BackColor = Color.FromArgb(66, 120, 190)
                        With cb.FlatAppearance
                            .BorderSize = 0
                            .CheckedBackColor = Color.FromArgb(70, 115, 50)
                            .MouseOverBackColor = Color.FromArgb(45, 80, 130)
                        End With
                        For i = 1 To dim_FG_option
                            If Right(cb.Name, 4) = arr_SEARCH_FG(i, 1) Then cb.Text = arr_SEARCH_FG(i, 2)
                        Next
                    Else
                        cb.BackColor = Color.White
                        With cb.FlatAppearance
                            .BorderSize = 14
                            .CheckedBackColor = Color.White
                            .MouseOverBackColor = Color.RosyBrown
                        End With
                    End If
                End If
            End If
        Next

        Call PAUSE()
        Form1.LOAD_BAR2.Value = 25

        '--- FILL FG BUTTONS ---

        For Each ctrl As Control In Form1.Controls
            If TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                Dim cb As System.Windows.Forms.CheckBox = ctrl
                If Mid(cb.Name, 4, 2) = "SG" Then
                    If CInt(Right(cb.Name, 2)) <= dim_SG_option Then
                        cb.BackColor = Color.FromArgb(66, 120, 190)
                        With cb.FlatAppearance
                            .BorderSize = 0
                            .CheckedBackColor = Color.FromArgb(70, 115, 50)
                            .MouseOverBackColor = Color.FromArgb(45, 80, 130)
                        End With
                        For i = 1 To dim_SG_option
                            If Right(cb.Name, 4) = arr_SEARCH_SG(i, 1) Then cb.Text = arr_SEARCH_SG(i, 2)
                        Next
                    Else
                        cb.BackColor = Color.White
                        With cb.FlatAppearance
                            .BorderSize = 14
                            .CheckedBackColor = Color.White
                            .MouseOverBackColor = Color.RosyBrown
                        End With
                    End If
                End If
            End If
        Next

        Call PAUSE()
        Form1.LOAD_BAR2.Value = 50

        '--- UNLOCK UI BUTTONS ---
        'moved below doe to timing
        Call PAUSE()
        Form1.LOAD_BAR2.Value = 75

        '--- FINALIZE AND HIDE LOAD PANEL ---

        Call PAUSE()
        Form1.LOAD_BAR2.Value = 100

        Call PAUSE()
        Call PAUSE()
        With Form1.LOAD_BUTTON
            .Text = "SUCCESS"
            .BackColor = Color.FromArgb(6, 176, 37)
            .FlatAppearance.MouseOverBackColor = Color.FromArgb(6, 176, 37)
            .FlatAppearance.MouseDownBackColor = Color.FromArgb(6, 176, 37)
        End With

        Call PAUSE2()

        unlocked = True

        With Form1.BT_SET
            .ForeColor = Color.FromArgb(215, 230, 250)
            .FlatAppearance.MouseDownBackColor = Color.FromArgb(45, 70, 110)
            .FlatAppearance.MouseOverBackColor = Color.FromArgb(45, 70, 110)
        End With
        With Form1.BT_SEARCH
            .ForeColor = Color.White
            .FlatAppearance.MouseDownBackColor = Color.FromArgb(70, 130, 50)
            .FlatAppearance.MouseOverBackColor = Color.FromArgb(50, 95, 30)
        End With

        Form1.LOAD_PANEL.Visible = False

    End Sub

End Module
