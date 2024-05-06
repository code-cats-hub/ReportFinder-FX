'--- GPL COPYRIGHT 2024 - CODE-CATS https://orcid.org/0009-0006-7849-1462 --- 
Module Module3

    '--- INPUT TABLES ---
    Public arr_IN_FG(100) As String
    Public arr_IN_SG(100) As String
    '--- CROSSOVER TABLES ---
    Public arr_HIT(100) As String
    Public arr_PASS(100) As String
    '--- DIMENSION WRITE ---
    Public dim_IN_FG As Integer
    Public dim_IN_SG As Integer
    Public dim_HIT As Integer
    Public dim_PASS As Integer

    Public Sub SEARCH_PERFORMER()

        Call FILL_arr_IN()
        Call SEARCH_PHASE1()
        Call SEARCH_PHASE2()
        Call FORMAT_SEARCH_UI()
        Call FILL_SEARCH_UI()

    End Sub
    Private Sub FILL_arr_IN()

        Dim x As Integer
        x = 0

        '--- FG SECTION ---

        For Each ctrl As Control In Form1.Controls
            If TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                Dim cb As System.Windows.Forms.CheckBox = ctrl
                If Left(cb.Name, 5) = "BT_FG" And cb.FlatAppearance.BorderSize = 0 And cb.Checked = True Then
                    x = x + 1
                End If
            End If
        Next

        dim_IN_FG = x
        ReDim arr_IN_FG(dim_IN_FG)
        x = 0

        For Each ctrl As Control In Form1.Controls
            If TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                Dim cb As System.Windows.Forms.CheckBox = ctrl
                If Left(cb.Name, 5) = "BT_FG" And cb.FlatAppearance.BorderSize = 0 And cb.Checked = True Then
                    x = x + 1
                    arr_IN_FG(x) = Right(cb.Name, 4)
                End If
            End If
        Next

        x = 0

        '--- SG SECTION ---

        For Each ctrl As Control In Form1.Controls
            If TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                Dim cb As System.Windows.Forms.CheckBox = ctrl
                If Left(cb.Name, 5) = "BT_SG" And cb.FlatAppearance.BorderSize = 0 And cb.Checked = True Then
                    x = x + 1
                End If
            End If
        Next

        dim_IN_SG = x
        ReDim arr_IN_SG(dim_IN_SG)
        x = 0

        For Each ctrl As Control In Form1.Controls
            If TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                Dim cb As System.Windows.Forms.CheckBox = ctrl
                If Left(cb.Name, 5) = "BT_SG" And cb.FlatAppearance.BorderSize = 0 And cb.Checked = True Then
                    x = x + 1
                    arr_IN_SG(x) = Right(cb.Name, 4)
                End If
            End If
        Next

        x = 0

    End Sub
    Private Sub SEARCH_PHASE1()

        Dim x, y, z, h As Integer
        Dim hitcount As Integer
        Dim flaghit As Boolean

        hitcount = 0
        flaghit = False

        '>>> cycle through input1
        For z = 1 To dim_IN_FG
            '>>> look for input1 through search1 (FG) sub-array
            For x = 1 To dim_RKEY
                For y = 2 To dim_FG_cap
                    '>>> input1 was found at x,y - check if hit is already in array, if not then add
                    If arr_IN_FG(z) = arr_MAP_FG(x, y) Then
                        For h = 1 To 100
                            If arr_HIT(h) = arr_MAP_FG(x, 1) Then flaghit = True
                        Next
                        If flaghit = False Then
                            hitcount = hitcount + 1
                            arr_HIT(hitcount) = arr_MAP_FG(x, 1)
                        Else
                            flaghit = False
                        End If
                    End If
                    '<<<
                Next
            Next
            '<<<
        Next
        '<<<

        '>>> write dimension
        'ReDim arr_HIT(hitcount)
        dim_HIT = hitcount
        '<<<

    End Sub
    Private Sub SEARCH_PHASE2()

        Dim x, y, z, h, p As Integer
        Dim passcount As Integer
        Dim flagpass As Boolean

        passcount = 0
        flagpass = False

        '>>> cycle through input2
        For z = 1 To dim_IN_SG
            '>>> look for input2 through search2 (SG) sub-array
            For x = 1 To dim_RKEY
                For y = 2 To dim_SG_cap
                    '>>> input2 was found at x,y
                    If arr_IN_SG(z) = arr_MAP_SG(x, y) Then
                        '>>> check if new hit is on hitlist, if so check if this hit was passed to pass array, if not then pass
                        For h = 1 To dim_HIT
                            If arr_MAP_SG(x, 1) = arr_HIT(h) Then
                                For p = 1 To 100
                                    If arr_PASS(p) = arr_HIT(h) Then flagpass = True
                                Next
                                If flagpass = False Then
                                    passcount = passcount + 1
                                    arr_PASS(passcount) = arr_HIT(h)
                                Else
                                    flagpass = False
                                End If
                            End If
                        Next
                        '<<<
                    End If
                    '<<<
                Next
            Next
            '<<<
        Next
        '<<<

        '>>> write dimension
        'ReDim arr_PASS(passcount)
        dim_PASS = passcount
        '<<<
    End Sub
    Private Sub FORMAT_SEARCH_UI()

        For Each ctrl As Control In Form2.PANEL_PARENT.Controls
            If TypeOf ctrl Is System.Windows.Forms.Panel Then
                Dim rpanel As System.Windows.Forms.Panel = ctrl
                If Right(rpanel.Name, 5) = "CHILD" Then
                    If CInt(Mid(rpanel.Name, 2, 2)) > dim_PASS Then
                        rpanel.Visible = False
                    End If
                End If
            End If
        Next

    End Sub
    Private Sub FILL_SEARCH_UI()

        Dim i As Integer

        For Each ctrl As Control In Form2.PANEL_PARENT.Controls
            If TypeOf ctrl Is System.Windows.Forms.Panel Then
                Dim rpanel As System.Windows.Forms.Panel = ctrl
                For Each ctrl2 As Control In rpanel.Controls
                    If TypeOf ctrl2 Is System.Windows.Forms.Label Then
                        Dim rlabel As System.Windows.Forms.Label = ctrl2
                        Select Case Mid(rlabel.Name, 5, 3)
                            Case "M01"
                                For i = 1 To dim_RKEY
                                    If arr_MAP_RES(i, 1) = arr_PASS(CInt(Mid(rlabel.Name, 2, 2))) Then
                                        rlabel.Text = arr_MAP_RES(i, 2)
                                    End If
                                Next
                            Case "M02"
                                For i = 1 To dim_RKEY
                                    If arr_MAP_RES(i, 1) = arr_PASS(CInt(Mid(rlabel.Name, 2, 2))) Then
                                        rlabel.Text = arr_MAP_RES(i, 3)
                                    End If
                                Next
                            Case "M03"
                                For i = 1 To dim_RKEY
                                    If arr_MAP_RES(i, 1) = arr_PASS(CInt(Mid(rlabel.Name, 2, 2))) Then
                                        rlabel.Text = arr_MAP_RES(i, 4)
                                    End If
                                Next
                            Case "M04"
                                For i = 1 To dim_RKEY
                                    If arr_MAP_RES(i, 1) = arr_PASS(CInt(Mid(rlabel.Name, 2, 2))) Then
                                        rlabel.Text = arr_MAP_RES(i, 5)
                                    End If
                                Next
                        End Select
                    End If
                Next
            End If
        Next

    End Sub
End Module
