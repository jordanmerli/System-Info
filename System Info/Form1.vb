Imports System.Management
Imports System
Imports System.Diagnostics
Public Class Form1
    Dim dt As New DataTable
    Dim X, Y As Integer
    Dim newpoint As System.Drawing.Point
    Dim cpu As New Diagnostics.PerformanceCounter("Processor", "% Processor Time", "_Total")
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Timer3.Interval = 10
        Timer2.Start()
        Timer3.Enabled = True
        Timer1.Enabled = True
        Timer4.Interval = 500
        Timer4.Start()
        dt.Columns.Add("Processname")
        dt.Columns.Add("Session ID")
        dt.Columns.Add("Process ID")
        dt.Columns.Add("Respond?")
        dt.Columns.Add("Allocated Virtual Memory")
        dt.Columns.Add("Allocated Phisic Memory")
        DataGridView1.DataSource = dt
        DataGridView1.DefaultCellStyle.Font = New Font("Microsoft Sans Serif", 7.0F, FontStyle.Regular)
    End Sub
    Private Function GetIPv4Address() As String
        GetIPv4Address = String.Empty
        Dim strHostName As String = System.Net.Dns.GetHostName()
        Dim iphe As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(strHostName)
        For Each ipheal As System.Net.IPAddress In iphe.AddressList
            If ipheal.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                GetIPv4Address = ipheal.ToString()
                Label32.Text = Label32.Text + "   " + GetIPv4Address
                Label29.Text = Label29.Text + "   " + strHostName
            End If
        Next
    End Function
    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Label3.Text = DateTime.Now
    End Sub
    Private dtBootTime As New DateTime()
    Declare Function GetKeyState Lib "user32" Alias "GetKeyState" (ByVal ByValnVirtKey As Int32) As Int16
    Private Const VK_CAPSLOCK = &H14
    Private Const VK_NUMLOCK = &H90
    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        Timer3.Interval = 10
        If GetKeyState(VK_CAPSLOCK) = 1 And GetKeyState(VK_NUMLOCK) = 1 Then
            Label55.Text = "CAPS & NUM ON"
        ElseIf GetKeyState(VK_CAPSLOCK) = 1 And GetKeyState(VK_NUMLOCK) <> 1 Then
            Label55.Text = "CAPS ON"
        ElseIf GetKeyState(VK_CAPSLOCK) <> 1 And GetKeyState(VK_NUMLOCK) = 1 Then
            Label55.Text = "NUM ON"
        Else
            Label55.Text = ""
        End If
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim psBattery As PowerStatus = SystemInformation.PowerStatus
        Dim perFull As Single = psBattery.BatteryLifePercent
        If perFull * 100 > 100 Then
            ProgressBar1.Value = 100
        ElseIf perFull * 100 < 100 Then
            ProgressBar1.Value = perFull * 100
        End If
        If psBattery.PowerLineStatus = PowerLineStatus.Online Then
            Label54.Text = "Battery Info - " & perFull * 100 & "%" & " charging!"
        ElseIf psBattery.PowerLineStatus = PowerLineStatus.Offline Then
            If (perFull * 100) < 50 Then
                Label54.Text = "Battery Info - " & perFull * 100 & "%" & " Connect Charger!"
            Else
                Label54.Text = "Battery Info - " & perFull * 100 & "%" & "Disconnect charger!"
            End If
        End If
        Timer1.Interval = 1000
    End Sub
    Private Sub SI_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        X = Control.MousePosition.X - Me.Location.X
        Y = Control.MousePosition.Y - Me.Location.Y
    End Sub
    Private Sub SI_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove
        If e.Button = Windows.Forms.MouseButtons.Left Then
            newpoint = Control.MousePosition
            newpoint.X -= (X)
            newpoint.Y -= (Y)
            Me.Location = newpoint
        End If
    End Sub
    Private Sub Timer4_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer4.Tick
        Label1.Text = "CPU: " & cpu.NextValue() & " %"
        GetIPv4Address()
        Label17.Text = My.Computer.Info.OSFullName
        Label43.Text = My.User.Name
        Label23.Text = My.Computer.Info.AvailablePhysicalMemory & " bytes" & " / " & My.Computer.Info.TotalPhysicalMemory & " bytes"
        Label30.Text = My.Computer.Info.AvailableVirtualMemory & " bytes" & " / " & My.Computer.Info.TotalVirtualMemory & " bytes"
        Label31.Text = My.Computer.Info.OSPlatform
        Label39.Text = My.Computer.Info.OSVersion
    End Sub
    Private Sub TrackBar1_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.Scroll
        If TrackBar1.Value = 1 Then
            Timer5.Start()
        Else
            Timer5.Stop()
            Me.Width = 711
            Me.Height = 253
            Label2.Text = "Process(es): 0"
        End If
    End Sub
    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start("mailto:joe_nanoteck@hotmail.it")
    End Sub
    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Process.Start("http://nanoteck.altervista.org")
    End Sub
    Private Sub Timer5_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer5.Tick
        Me.Width = 711
        Me.Height = 538
        Dim processes As System.Diagnostics.Process()
        processes = System.Diagnostics.Process.GetProcesses()
        Dim process As System.Diagnostics.Process
        For Each process In processes
            Dim selRows As DataRow() = dt.Select("Processname='" & process.ProcessName & ".exe'")
            If selRows.Length > 0 Then
                selRows(0).Item("Processname") = process.ProcessName & ".exe"
                selRows(0).Item("Session ID") = process.SessionId
                selRows(0).Item("Process ID") = process.Id
                selRows(0).Item("Respond?") = process.Responding
                selRows(0).Item("Allocated Virtual Memory") = process.VirtualMemorySize64 & " bytes"
                selRows(0).Item("Allocated Phisic Memory") = process.WorkingSet64 & " bytes"
            Else
                Dim dr As DataRow = dt.NewRow
                dr.Item("Processname") = process.ProcessName & ".exe"
                dr.Item("Session ID") = process.SessionId
                dr.Item("Process ID") = process.Id
                dr.Item("Respond?") = process.Responding
                dr.Item("Allocated Virtual Memory") = process.VirtualMemorySize64 & " bytes"
                dr.Item("Allocated Phisic Memory") = process.WorkingSet64 & " bytes"
                dt.Rows.Add(dr)
            End If
        Next
        Label2.Text = "Process(es): " & DataGridView1.Rows.Count()
    End Sub
End Class
