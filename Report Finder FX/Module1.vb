'--- GPL COPYRIGHT 2024 - CODE-CATS https://orcid.org/0009-0006-7849-1462 --- 
Imports System.Collections.ObjectModel
Imports System.IO
Module Module1
    Public path_chrome As String = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    Public path_fox As String = ""
    Public path_browser As String = path_chrome
    Public path_user As String = Application.StartupPath()
    Public chrome_lock As Boolean = False
    Public browser_error As Boolean = False
    Public browser_error2 As Boolean = False

    Public Sub PAUSE()
        Dim wait As Date
        wait = Now.AddMilliseconds(250)
        Do Until Now > wait
            System.Windows.Forms.Application.DoEvents()
        Loop
    End Sub
    Public Sub PAUSE2()
        Dim wait As Date
        wait = Now.AddMilliseconds(1500)
        Do Until Now > wait
            System.Windows.Forms.Application.DoEvents()
        Loop
    End Sub
    Public Sub START_SEARCH()
        Dim a, b As Integer
        a = 0
        b = 0

loopback:
        For Each ctrl As Control In Form1.Controls
            If TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                Dim cb As System.Windows.Forms.CheckBox = ctrl
                If Left(cb.Name, 5) = "BT_FG" And cb.FlatAppearance.BorderSize = 0 And cb.Checked = True Then
                    a = a + 1
                End If
            End If
        Next
        For Each ctrl As Control In Form1.Controls
            If TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                Dim cb As System.Windows.Forms.CheckBox = ctrl
                If Left(cb.Name, 5) = "BT_SG" And cb.FlatAppearance.BorderSize = 0 And cb.Checked = True Then
                    b = b + 1
                End If
            End If
        Next

        If a = 0 Then
            Dim msg As Integer = MsgBox("No search fields were selected" & vbCrLf & "Did you want to select all fields?", vbYesNo, "Report Finder FX")
            If msg = vbYes Then
                Call FG_SELECT_ALL()
                GoTo loopback
            End If
            If msg = vbNo Then GoTo finish
        End If
        If b = 0 Then
            Dim msg As Integer = MsgBox("No result types were selected" & vbCrLf & "Did you want to select all types?", vbYesNo, "Report Finder FX")
            If msg = vbYes Then
                Call SG_SELECT_ALL()
                GoTo loopback
            End If
            If msg = vbNo Then GoTo finish
        End If


        On Error GoTo fault
        Call SEARCH_PERFORMER()
        Form2.Show()
        Form1.Hide()
        GoTo finish

fault:
        MsgBox("search failed due to unmatched forces (at least by this app)", vbOKOnly, "Oh no...")
finish:
    End Sub
    Public Sub RETURN_N_RESET()

        Form1.Show()
        Form2.Hide()

        '--- INPUT TABLES RESET ---
        ReDim arr_IN_FG(100)
        ReDim arr_IN_SG(100)
        '--- CROSSOVER TABLES RESET ---
        ReDim arr_HIT(100)
        ReDim arr_PASS(100)
        '--- DIMENSION RESET ---
        dim_IN_FG = 0
        dim_IN_SG = 0
        dim_HIT = 0
        dim_PASS = 0

        For Each ctrl As Control In Form1.Controls
            If TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                Dim cb As System.Windows.Forms.CheckBox = ctrl
                cb.Checked = False
            End If
        Next

        For Each ctrl As Control In Form2.PANEL_PARENT.Controls
            If TypeOf ctrl Is System.Windows.Forms.Panel Then
                Dim rpanel As System.Windows.Forms.Panel = ctrl
                If Right(rpanel.Name, 5) = "CHILD" Then
                    rpanel.Visible = True
                End If
            End If
        Next

    End Sub
    Public Sub SHUT_DOWN()

        Form1.Hide()
        Form2.Hide()
        Form3.Hide()
        Form4.Hide()

        Form1.Close()

    End Sub
    Public Sub INFO_PANEL_ON()

        Form3.Show()

    End Sub
    Public Sub INFO_PANEL_OFF()

        Form3.Hide()
        Form1.Focus()
        Form2.Focus()
        Form4.Focus()

    End Sub
    Public Sub SET_PANEL_ON()

        Form4.Show()

    End Sub
    Public Sub SET_PANEL_OFF()

        Form4.Hide()
        Form1.Focus()
        Form2.Focus()

    End Sub
    Public Sub FG_SELECT_ALL()

        For Each ctrl As Control In Form1.Controls
            If TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                Dim cb As System.Windows.Forms.CheckBox = ctrl
                If Mid(cb.Name, 4, 2) = "FG" Then
                    cb.Checked = True
                End If
            End If
        Next

    End Sub
    Public Sub FG_DESELECT_ALL()

        For Each ctrl As Control In Form1.Controls
            If TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                Dim cb As System.Windows.Forms.CheckBox = ctrl
                If Mid(cb.Name, 4, 2) = "FG" Then
                    cb.Checked = False
                End If
            End If
        Next

    End Sub
    Public Sub SG_SELECT_ALL()

        For Each ctrl As Control In Form1.Controls
            If TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                Dim cb As System.Windows.Forms.CheckBox = ctrl
                If Mid(cb.Name, 4, 2) = "SG" Then
                    cb.Checked = True
                End If
            End If
        Next

    End Sub
    Public Sub SG_DESELECT_ALL()

        For Each ctrl As Control In Form1.Controls
            If TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                Dim cb As System.Windows.Forms.CheckBox = ctrl
                If Mid(cb.Name, 4, 2) = "SG" Then
                    cb.Checked = False
                End If
            End If
        Next

    End Sub
    Public Sub BTC_PUSH2CHROME(panel_nr As String)

        Dim i As Integer
        Dim name_read As String
        Dim direct_link As String

        name_read = TryCast(Form2.Controls.Find(panel_nr & "_M01", True).First, Label).Text
        direct_link = "https://www.google.com/"

        For i = 1 To dim_RKEY
            If arr_MAP_RES(i, 2) = name_read Then direct_link = arr_MAP_RES(i, 6)
        Next

        MsgBox("Sending to selected browser >>>", vbOKOnly, "Report Finder FX")

rollback:
        On Error GoTo error1
        Process.Start(path_browser, direct_link)
        GoTo finish
error1:
        If browser_error = False Then
            Call CHROME_FINDER()
            browser_error = True
            GoTo rollback
        Else
            Dim msg As Integer = MsgBox("Chrome failed. Do you want to try Firefox instead?", vbYesNo, "It's not a good day for Chrome it seems")
            If msg = vbYes Then Call FIREFOX_FINDER()
        End If
finish:
    End Sub
    Public Sub CHROME_FINDER()
        Dim files As ReadOnlyCollection(Of String)
        Dim xpathx As String
        xpathx = ""

        If chrome_lock = True Then GoTo finish

        On Error GoTo next1
        files = My.Computer.FileSystem.GetFiles(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Google\Chrome"), FileIO.SearchOption.SearchAllSubDirectories, "chrome.exe")
        If files.Count > 0 Then
            xpathx = files(0)
        End If
        If xpathx <> "" Then GoTo rundown
next1:
        On Error GoTo next2
        files = My.Computer.FileSystem.GetFiles("C:\Program Files\Google\Chrome\", FileIO.SearchOption.SearchAllSubDirectories, "chrome.exe")
        If files.Count > 0 Then
            xpathx = files(0)
        End If
        If xpathx <> "" Then GoTo rundown
next2:
        On Error GoTo rundown
        files = My.Computer.FileSystem.GetFiles("C:\Program Files (x86)\Google\Chrome\", FileIO.SearchOption.SearchAllSubDirectories, "chrome.exe")
        If files.Count > 0 Then
            xpathx = files(0)
        End If
rundown:
        If xpathx <> "" Then
            path_chrome = xpathx
            With Form4.BT_CROMEFIND
                .Text = "CHROME LOCATED"
                .BackColor = Color.FromArgb(50, 95, 30)
                .FlatAppearance.MouseDownBackColor = Color.FromArgb(50, 95, 30)
                .FlatAppearance.MouseOverBackColor = Color.FromArgb(50, 95, 30)
                .Cursor = Cursors.Default
            End With
            path_browser = path_chrome
            chrome_lock = True
        Else
            MsgBox("path to chrome no found amigo", vbOKOnly, "Oh no...")
        End If
finish:
    End Sub
    Public Sub FIREFOX_FINDER()
        Dim files As ReadOnlyCollection(Of String)
        Dim xpathx As String
        xpathx = ""

        If path_fox <> "" Then GoTo finishdown

        On Error GoTo next1
        files = My.Computer.FileSystem.GetFiles("C:\Program Files\Mozilla Firefox\", FileIO.SearchOption.SearchAllSubDirectories, "firefox.exe")
        If files.Count > 0 Then
            xpathx = files(0)
        End If
        If xpathx <> "" Then GoTo rundown
next1:
        On Error GoTo rundown
        files = My.Computer.FileSystem.GetFiles("C:\Program Files (x86)\Mozilla Firefox\", FileIO.SearchOption.SearchAllSubDirectories, "firefox.exe")
        If files.Count > 0 Then
            xpathx = files(0)
        End If
rundown:
        If xpathx <> "" Then
            path_fox = xpathx
            GoTo finishdown
        Else
            MsgBox("path to fox no found amigo", vbOKOnly, "Oh no...")
            GoTo finish
        End If
finishdown:
        path_browser = path_fox
        Form4.BT_FOX_OFF.Visible = False
        Form4.LINK_LABEL_FOX_OFF.Visible = False
        Form4.LOCK_CHROMEFIND.Visible = True
finish:
    End Sub
    Public Sub FIREFOX_REMOVER()
        path_browser = path_chrome
        Form4.BT_FOX_OFF.Visible = True
        Form4.LINK_LABEL_FOX_OFF.Visible = True
        Form4.LOCK_CHROMEFIND.Visible = False
    End Sub
    Public Sub LINK2BROWSER(push_lnk As String)

rollback:
        On Error GoTo error1
        Process.Start(path_browser, push_lnk)
        GoTo finish
error1:
        If browser_error2 = False Then
            Call CHROME_FINDER()
            browser_error2 = True
            GoTo rollback
        Else
            Dim msg As Integer = MsgBox("Chrome failed. Do you want to try Firefox instead?", vbYesNo, "It's not a good day for Chrome it seems")
            If msg = vbYes Then Call FIREFOX_FINDER()
        End If
finish:
    End Sub
    Public Sub OPEN_GPLN()
        On Error Resume Next
        Process.Start("C:\Windows\notepad.exe", path_user & "Data\LICENSE NOTICE.txt")
    End Sub
End Module
