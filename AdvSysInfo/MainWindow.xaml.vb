Imports System.Management
Imports System.Data
Imports System.ComponentModel
Imports System.ServiceProcess
Imports Microsoft.Win32

Class MainWindow

    Private mstrCurrentTargetServer As String
    Private mstrPreviousTargetServer As String

    Private mbrshAlternatingRowColor As Brush = Brushes.GhostWhite

    Private mbwSysInfo As New BackgroundWorker
    Private mobjSysInfo As clsSysInfo

    Private mbwCPULoad As New BackgroundWorker
    Private mintCPULogicalProcessors As Integer
    Private mlstCPULoad As List(Of clsCPU)
    Private mblnCPULoadError As Boolean

    Private mbwMemLoad As New BackgroundWorker
    Private mobjMemLoad As clsMem
    Private mblnMemLoadError As Boolean

    Private mbwDiskLoad As New BackgroundWorker
    Private mlstDiskLoad As List(Of clsDisk)
    Private mblnDiskLoadError As Boolean

    Private mbwApplications As New BackgroundWorker
    Private mdvApplications As DataView
    Private mstrApplicationsFilter As String

    Private mbwProcesses As New BackgroundWorker
    Private mdvProcesses As DataView
    Private mstrProcessesFilter As String

    Private mbwServices As New BackgroundWorker
    Private mdvServices As DataView
    Private mblnDeviceDriver As Boolean
    Private mstrServicesFilter As String

    Private mbwEventLog As New BackgroundWorker
    Private mobjLogEntries As IEnumerable(Of Object)
    Private mstrLogType As String
    Private mstrLogEntryTypes As String
    Private mstrLogEntryFromDate As String
    Private mstrLogEntryToDate As String

    Private mbwOpenFiles As New BackgroundWorker
    Private mdvOpenFiles As DataView
    Private mstrOpenFilesFilter As String
    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        Try

            Reset_Application()

            AddHandler mbwSysInfo.DoWork, AddressOf mbwSysInfo_DoWork
            AddHandler mbwSysInfo.RunWorkerCompleted, AddressOf mbwSysInfo_WorkCompleted

            mbwSysInfo.WorkerReportsProgress = True
            mbwSysInfo.WorkerSupportsCancellation = True

            AddHandler mbwCPULoad.DoWork, AddressOf mbwCPULoad_DoWork
            AddHandler mbwCPULoad.RunWorkerCompleted, AddressOf mbwCPULoad_WorkCompleted

            mbwCPULoad.WorkerReportsProgress = True
            mbwCPULoad.WorkerSupportsCancellation = True

            AddHandler mbwMemLoad.DoWork, AddressOf mbwMemLoad_DoWork
            AddHandler mbwMemLoad.RunWorkerCompleted, AddressOf mbwMemLoad_WorkCompleted

            mbwMemLoad.WorkerReportsProgress = True
            mbwMemLoad.WorkerSupportsCancellation = True

            AddHandler mbwDiskLoad.DoWork, AddressOf mbwDiskLoad_DoWork
            AddHandler mbwDiskLoad.RunWorkerCompleted, AddressOf mbwDiskLoad_WorkCompleted

            mbwDiskLoad.WorkerReportsProgress = True
            mbwDiskLoad.WorkerSupportsCancellation = True

            AddHandler mbwApplications.DoWork, AddressOf mbwApplications_DoWork
            AddHandler mbwApplications.ProgressChanged, AddressOf mbwApplications_ProgressChanged
            AddHandler mbwApplications.RunWorkerCompleted, AddressOf mbwApplications_WorkCompleted

            mbwApplications.WorkerReportsProgress = True
            mbwApplications.WorkerSupportsCancellation = True

            AddHandler mbwProcesses.DoWork, AddressOf mbwProcesses_DoWork
            AddHandler mbwProcesses.ProgressChanged, AddressOf mbwProcesses_ProgressChanged
            AddHandler mbwProcesses.RunWorkerCompleted, AddressOf mbwProcesses_WorkCompleted

            mbwProcesses.WorkerReportsProgress = True
            mbwProcesses.WorkerSupportsCancellation = True

            AddHandler mbwServices.DoWork, AddressOf mbwServices_DoWork
            AddHandler mbwServices.ProgressChanged, AddressOf mbwServices_ProgressChanged
            AddHandler mbwServices.RunWorkerCompleted, AddressOf mbwServices_WorkCompleted

            mbwServices.WorkerReportsProgress = True
            mbwServices.WorkerSupportsCancellation = True

            AddHandler mbwEventLog.DoWork, AddressOf mbwEventLog_DoWork
            AddHandler mbwEventLog.ProgressChanged, AddressOf mbwEventLog_ProgressChanged
            AddHandler mbwEventLog.RunWorkerCompleted, AddressOf mbwEventLog_WorkCompleted

            mbwEventLog.WorkerReportsProgress = True
            mbwEventLog.WorkerSupportsCancellation = True

            AddHandler mbwOpenFiles.DoWork, AddressOf mbwOpenFiles_DoWork
            AddHandler mbwOpenFiles.ProgressChanged, AddressOf mbwOpenFiles_ProgressChanged
            AddHandler mbwOpenFiles.RunWorkerCompleted, AddressOf mbwOpenFiles_WorkCompleted

            mbwOpenFiles.WorkerReportsProgress = True
            mbwOpenFiles.WorkerSupportsCancellation = True

        Catch ex As Exception

            MsgBox("Error initializing application." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Sub
    Private Sub cmdScan_Click(sender As Object, e As RoutedEventArgs) Handles cmdScan.Click

        Try

            If String.IsNullOrWhiteSpace(txtServerNameIP.Text) Then Exit Sub

            mstrPreviousTargetServer = mstrCurrentTargetServer
            mstrCurrentTargetServer = txtServerNameIP.Text.Trim

            If mstrPreviousTargetServer = mstrCurrentTargetServer Then Exit Sub

            Reset_Application()

            lblServerNameIP.Content = Get_HostName_Or_IPAddress(mstrCurrentTargetServer, True)

            tabctrl.IsEnabled = True

            If mbwSysInfo.IsBusy Then
                mbwSysInfo.CancelAsync()
            ElseIf Not mbwSysInfo.CancellationPending Then
                mbwSysInfo.RunWorkerAsync()
            End If

            If mbwCPULoad.IsBusy Then
                mbwCPULoad.CancelAsync()
            ElseIf Not mbwCPULoad.CancellationPending Then
                mbwCPULoad.RunWorkerAsync()
            End If

            If mbwMemLoad.IsBusy Then
                mbwMemLoad.CancelAsync()
            ElseIf Not mbwMemLoad.CancellationPending Then
                mbwMemLoad.RunWorkerAsync()
            End If

            If mbwDiskLoad.IsBusy Then
                mbwDiskLoad.CancelAsync()
            ElseIf Not mbwDiskLoad.CancellationPending Then
                mbwDiskLoad.RunWorkerAsync()
            End If

            If mbwApplications.IsBusy Then mbwApplications.CancelAsync()
            If mbwProcesses.IsBusy Then mbwProcesses.CancelAsync()
            If mbwServices.IsBusy Then mbwServices.CancelAsync()
            If mbwEventLog.IsBusy Then mbwEventLog.CancelAsync()
            If mbwOpenFiles.IsBusy Then mbwOpenFiles.CancelAsync()

        Catch ex As Exception

            MsgBox("Error initializing metrics at cmdScan.Click." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Sub
    Private Sub Reset_Application()

        Try

            mintCPULogicalProcessors = -1

            scrlCPU.ScrollToTop()
            scrlDisks.ScrollToTop()

            tabctrl.IsEnabled = False

            lbxDetails.Items.Clear()
            grdProc.Children.Clear()
            grdMem.Children.Clear()
            grdDisks.Children.Clear()

            dtgrdApplications.DataContext = Nothing
            dtgrdApplications.ItemsSource = Nothing
            dtgrdProcesses.DataContext = Nothing
            dtgrdProcesses.ItemsSource = Nothing
            dtgrdServices.DataContext = Nothing
            dtgrdServices.ItemsSource = Nothing
            dtgrdEventLogs.DataContext = Nothing
            dtgrdEventLogs.ItemsSource = Nothing
            dtgrdOpenFiles.DataContext = Nothing
            dtgrdOpenFiles.ItemsSource = Nothing

            txtApplicationName.Text = ""
            txtProcessName.Text = ""
            txtAccessedBy.Text = ""
            txtOpenFile.Text = ""
            txtMessage.Text = ""
            txtAccessedBy.IsEnabled = False
            txtOpenFile.IsEnabled = False

            rbGeneral.IsChecked = True
            rbApplication.IsChecked = True
            rbAnd.IsChecked = True
            rbAnd.IsEnabled = False
            rbOr.IsEnabled = False

            chk32Bit.IsChecked = False
            chk64Bit.IsChecked = False
            chkRunning.IsChecked = False
            chkStopped.IsChecked = False
            chkAutomatic.IsChecked = False
            chkManual.IsChecked = False
            chkDisabled.IsChecked = False

            chkCritical.IsChecked = False
            chkWarning.IsChecked = False
            chkVerbose.IsChecked = False
            chkError.IsChecked = False
            chkInformation.IsChecked = False
            chkAccessedBy.IsChecked = False
            chkOpenFile.IsChecked = False

            dtpFromDate.SelectedDate = Nothing
            dtpToDate.SelectedDate = Nothing

            lblServerNameIP.Content = ""
            lblApplicationsCount.Content = ""
            lblProcessesCount.Content = ""
            lblServicesCount.Content = ""
            lblEventLogCount.Content = ""
            lblOpenFilesCount.Content = ""
            lblEventId.Content = ""
            lblEventType.Content = ""
            lblGenerated.Content = ""
            lblSource.Content = ""
            lblMachineName.Content = ""

        Catch ex As Exception

            MsgBox("Error resetting UI." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Sub
    Private Function Retrieve_SysInfo() As clsSysInfo

        Try

            Dim sysInfo As New clsSysInfo
            Dim strOSName As String
            Dim strCPUName As String
            Dim strDiskModel As String
            Dim lngDiskSize As Long
            Dim i As Integer
            Dim strScope As String = "\\" & mstrCurrentTargetServer & "\root\cimv2"
            Dim objOS As New ManagementObjectSearcher(strScope, "SELECT * FROM Win32_OperatingSystem")
            Dim objCS As New ManagementObjectSearcher(strScope, "SELECT * FROM Win32_ComputerSystem")
            Dim objCPU As New ManagementObjectSearcher(strScope, "SELECT * FROM Win32_Processor")
            Dim objDisk As New ManagementObjectSearcher(strScope, "SELECT * FROM Win32_DiskDrive")
            Dim objMgmt As ManagementObject

            For Each objMgmt In objOS.Get
                sysInfo.OSManufacturer = IIf(IsNothing(objMgmt("manufacturer")), "", objMgmt("manufacturer").ToString)
                strOSName = IIf(IsNothing(objMgmt("name")), "", objMgmt("name").ToString)
                sysInfo.OSVersion = IIf(IsNothing(objMgmt("version")), "", objMgmt("version").ToString)
                sysInfo.OSArchitecture = IIf(IsNothing(objMgmt("osarchitecture")), "", objMgmt("osarchitecture").ToString)
                sysInfo.OSInstallDate = IIf(IsNothing(objMgmt("installdate")), "", objMgmt("installdate").ToString)
                sysInfo.OSLastBootUpTime = IIf(IsNothing(objMgmt("lastbootuptime")), "", objMgmt("lastbootuptime").ToString)
                sysInfo.WindowsDir = IIf(IsNothing(objMgmt("windowsdirectory")), "", objMgmt("windowsdirectory").ToString)
                sysInfo.ComputerName = IIf(IsNothing(objMgmt("csname")), "", objMgmt("csname").ToString)
            Next

            If InStr(strOSName, "|") > 0 Then
                Dim vArray() As String
                vArray = Split(strOSName, "|")
                sysInfo.OSName = vArray(0)
            End If

            For Each objMgmt In objCS.Get
                sysInfo.PhysicalMemoryTotal = IIf(IsNothing(objMgmt("totalphysicalmemory")), 0, (Convert.ToInt64(objMgmt("totalphysicalmemory")))) ' / 1024000000)
                sysInfo.Manufacturer = IIf(IsNothing(objMgmt("manufacturer")), "", objMgmt("manufacturer").ToString)
                sysInfo.Model = IIf(IsNothing(objMgmt("model")), "", objMgmt("model").ToString)
                sysInfo.SystemType = IIf(IsNothing(objMgmt("systemtype")), "", objMgmt("systemtype").ToString)
                sysInfo.ProcessorsPhysical = IIf(IsNothing(objMgmt("numberofprocessors")), "", objMgmt("numberofprocessors").ToString)
            Next

            For Each objMgmt In objCPU.Get
                sysInfo.CPUManufacturer = IIf(IsNothing(objMgmt("manufacturer")), "", objMgmt("manufacturer").ToString)
                strCPUName = IIf(IsNothing(objMgmt("name")), "", objMgmt("name").ToString)
                sysInfo.CPUCaption = IIf(IsNothing(objMgmt("caption")), "", objMgmt("caption").ToString)
                sysInfo.CPUClockSpeed = IIf(IsNothing(objMgmt("maxclockspeed")), "", objMgmt("maxclockspeed").ToString)
                sysInfo.CPUCores = IIf(IsNothing(objMgmt("numberofcores")), "", objMgmt("numberofcores").ToString)
                mintCPULogicalProcessors = IIf(IsNothing(objMgmt("numberoflogicalprocessors")), 0, Convert.ToInt32(objMgmt("numberoflogicalprocessors")))
                sysInfo.CPUProcessorsLogical = mintCPULogicalProcessors.ToString
            Next

            sysInfo.CPUName = Replace(Replace(Replace(Replace(strCPUName, "(R)", ""), "(TM)", ""), "CPU", ""), "@ ", "@")

            i = 0

            For Each objMgmt In objDisk.Get
                strDiskModel = IIf(IsNothing(objMgmt("model")), "", objMgmt("model").ToString)
                lngDiskSize = IIf(IsNothing(objMgmt("size")), "", Convert.ToInt64(objMgmt("size")))
                sysInfo.LogicalDisks.Add("Physical Disk (" & i & "): " & strDiskModel & " [" & CalculateBestByteSize(lngDiskSize, True) & "]")
                i = i + 1
            Next

            Retrieve_SysInfo = sysInfo

        Catch ex As Exception

            MsgBox("Error retrieving system information." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Function
    Private Sub Display_SysInfo()

        Try

            If mobjSysInfo Is Nothing Then Exit Sub

            lbxDetails.Items.Add(mobjSysInfo.Manufacturer & " " & mobjSysInfo.Model & " (" & mobjSysInfo.SystemType & ")")
            lbxDetails.Items.Add(mobjSysInfo.CPUName & " | " & "Processors: " & mobjSysInfo.CPUProcessorsLogical & " (Cores: " & mobjSysInfo.CPUCores & ") @" & mobjSysInfo.CPUClockSpeed & "MHz")
            'lbxDetails.Items.Add(mobjSysInfo.CPUCaption)
            lbxDetails.Items.Add(CalculateBestByteSize(mobjSysInfo.PhysicalMemoryTotal) & " RAM")

            For i = 0 To (mobjSysInfo.LogicalDisks.Count - 1)
                lbxDetails.Items.Add(mobjSysInfo.LogicalDisks(i))
            Next

            lbxDetails.Items.Add(mobjSysInfo.OSName & " - " & mobjSysInfo.OSArchitecture & " (" & mobjSysInfo.OSVersion & ")")
            lbxDetails.Items.Add("  - Installed: " & MakeFriendlyDateString(mobjSysInfo.OSInstallDate))
            lbxDetails.Items.Add("Uptime: " & CalculateSystemUpTime(mobjSysInfo.OSLastBootUpTime))
            lbxDetails.Items.Add("  - Last Boot: " & MakeFriendlyDateString(mobjSysInfo.OSLastBootUpTime))
            lbxDetails.Items.Add(" ")

        Catch ex As Exception

            MsgBox("Error displaying system information." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Sub
    Private Function Retrieve_CPULoad() As List(Of clsCPU)

        Try
            mblnCPULoadError = False

            Dim lstPerfCntrCPU As New List(Of PerformanceCounter)
            Dim lstCPULoad As New List(Of clsCPU)

            If mintCPULogicalProcessors = -1 Then
                Dim strScope As String = "\\" & mstrCurrentTargetServer & "\root\cimv2"
                Dim objCPU As New ManagementObjectSearcher(strScope, "SELECT * FROM Win32_Processor")
                Dim objMgmt As ManagementObject
                For Each objMgmt In objCPU.Get
                    mintCPULogicalProcessors = IIf(IsNothing(objMgmt("numberoflogicalprocessors")), 0, Convert.ToInt32(objMgmt("numberoflogicalprocessors").ToString))
                Next
            End If

            lstPerfCntrCPU.Add(New PerformanceCounter("Processor", "% Processor Time", "_Total", mstrCurrentTargetServer))

            'Environment.ProcessorCount
            Dim perfCounter As PerformanceCounter
            For i = 0 To Get_TotalCPUs(mstrCurrentTargetServer) - 1
                perfCounter = New PerformanceCounter("Processor", "% Processor Time", i.ToString, mstrCurrentTargetServer)
                lstPerfCntrCPU.Add(perfCounter)
            Next

            For Each proccesor In lstPerfCntrCPU
                Dim intCoreLoad As Integer = CInt(proccesor.NextValue())
                Dim cpu As New clsCPU
                System.Threading.Thread.Sleep(150)
                intCoreLoad = CInt(proccesor.NextValue())
                cpu.Id = proccesor.InstanceName
                cpu.Load = intCoreLoad
                lstCPULoad.Add(cpu)
            Next

            Retrieve_CPULoad = lstCPULoad

        Catch ex As Exception

            mblnCPULoadError = True

            If mbwCPULoad.IsBusy Then mbwCPULoad.CancelAsync()

            MsgBox("Error retrieving CPU load values." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Function
    Private Sub Display_CPULoad()

        Try

            mblnCPULoadError = False

            If mlstCPULoad Is Nothing Then Exit Sub

            Dim intCnt As Integer = 0
            Dim stkpnlAllCPULoad As New StackPanel

            With stkpnlAllCPULoad
                .Orientation = Orientation.Vertical
                .HorizontalAlignment = HorizontalAlignment.Left
                .VerticalAlignment = VerticalAlignment.Top
            End With

            For Each cpu In mlstCPULoad
                Dim stkpnlLoad As New StackPanel
                Dim lblCPU As New Label
                Dim prgbarLoad As New ProgressBar
                Dim lblLoad As New Label
                Dim lblSpacer As New Label
                With stkpnlLoad
                    .Orientation = Orientation.Horizontal
                    .Margin = New Thickness(0, 0, 0, 0.5)
                End With
                With lblCPU
                    .Content = IIf(intCnt = 0, "cpuT ", "cpu" & cpu.Id & " ")
                    .FontSize = 9
                    .Height = 22
                End With
                With prgbarLoad
                    .Foreground = IIf(intCnt = 0, ProgressBarForeColor(cpu.Load, False), ProgressBarForeColor(cpu.Load, True))
                    .Value = cpu.Load
                    .Height = IIf(intCnt = 0, 10, 5)
                    .Width = 210
                End With
                With lblLoad
                    .Content = " " & cpu.Load & "%"
                    .FontSize = 9
                    .Height = 22
                End With
                stkpnlLoad.Children.Add(lblCPU)
                stkpnlLoad.Children.Add(prgbarLoad)
                stkpnlLoad.Children.Add(lblLoad)
                stkpnlAllCPULoad.Children.Add(stkpnlLoad)
                intCnt = intCnt + 1
            Next

            With grdProc
                .Children.Clear()
                .Children.Add(stkpnlAllCPULoad)
            End With

            scrlCPU.IsEnabled = (mlstCPULoad.Count >= 7)

        Catch ex As Exception

            mblnCPULoadError = True

            If mbwCPULoad.IsBusy Then mbwCPULoad.CancelAsync()

            MsgBox("Error displaying CPU load metrics." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Sub
    Private Function Retrieve_MemLoad() As clsMem

        Try

            mblnMemLoadError = False

            Dim mem As New clsMem
            Dim strScope As String = "\\" & mstrCurrentTargetServer & "\root\cimv2"
            Dim objOS As New ManagementObjectSearcher(strScope, "SELECT * FROM Win32_OperatingSystem")
            Dim objCS As New ManagementObjectSearcher(strScope, "SELECT * FROM Win32_ComputerSystem")
            Dim sngPhysicalMemoryFree As Single
            Dim sngPhysicalMemoryTotal As Single
            Dim sngPhysicalMemoryUsage As Single
            Dim sngPhysicalMemoryLoad As Single
            Dim objMgmt As ManagementObject

            For Each objMgmt In objOS.Get
                sngPhysicalMemoryFree = IIf(IsNothing(objMgmt("freephysicalmemory")), 0, (Convert.ToInt64(objMgmt("freephysicalmemory")))) '/ 1024000)
            Next

            For Each objMgmt In objCS.Get
                sngPhysicalMemoryTotal = IIf(IsNothing(objMgmt("totalphysicalmemory")), 0, (Convert.ToInt64(objMgmt("totalphysicalmemory")))) ' / 1024000000)
            Next

            sngPhysicalMemoryUsage = sngPhysicalMemoryTotal - (sngPhysicalMemoryFree * 1024)
            sngPhysicalMemoryLoad = (sngPhysicalMemoryUsage / sngPhysicalMemoryTotal) * 100

            mem.PhysicalMemoryFree = sngPhysicalMemoryFree
            mem.PhysicalMemoryTotal = sngPhysicalMemoryTotal
            mem.PhysicalMemoryUsage = sngPhysicalMemoryUsage
            mem.PhysicalMemoryLoad = sngPhysicalMemoryLoad

            Retrieve_MemLoad = mem

        Catch ex As Exception

            mblnMemLoadError = True

            If mbwMemLoad.IsBusy Then mbwMemLoad.CancelAsync()

            MsgBox("Error retrieving memory load values." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Function
    Private Sub Display_MemLoad()

        Try

            mblnMemLoadError = False

            If mobjMemLoad Is Nothing Then Exit Sub

            Dim stkpnlMemoryLoad As New StackPanel

            With stkpnlMemoryLoad
                .Orientation = Orientation.Vertical
                .HorizontalAlignment = HorizontalAlignment.Left
                .VerticalAlignment = VerticalAlignment.Center
            End With

            Dim stkpnlTotMemoryLoad As New StackPanel
            Dim lblTotMemory As New Label
            Dim prgbarTotMemory As New ProgressBar
            Dim lblTotMemoryLoad As New Label
            Dim lblTotUsageRatio As New Label

            stkpnlTotMemoryLoad.Orientation = Orientation.Horizontal

            With lblTotMemory
                .Content = "Mem "
                .FontSize = 9
                .Height = 22
            End With

            With prgbarTotMemory
                .Foreground = ProgressBarForeColor(mobjMemLoad.PhysicalMemoryLoad)
                .Value = mobjMemLoad.PhysicalMemoryLoad
                .Height = 10
                .Width = 210
            End With

            With lblTotMemoryLoad
                .Content = " " & CInt(mobjMemLoad.PhysicalMemoryLoad) & "%"
                .FontSize = 9
                .Height = 22
            End With

            stkpnlTotMemoryLoad.Children.Add(lblTotMemory)
            stkpnlTotMemoryLoad.Children.Add(prgbarTotMemory)
            stkpnlTotMemoryLoad.Children.Add(lblTotMemoryLoad)

            With lblTotUsageRatio
                .Content = CalculateBestByteSize(mobjMemLoad.PhysicalMemoryUsage) & "/" & CalculateBestByteSize(mobjMemLoad.PhysicalMemoryTotal)
                .FontSize = 9
                .Height = 22
                .HorizontalAlignment = HorizontalAlignment.Center
                .VerticalAlignment = VerticalAlignment.Top
            End With

            stkpnlMemoryLoad.Children.Add(stkpnlTotMemoryLoad)
            stkpnlMemoryLoad.Children.Add(lblTotUsageRatio)

            With grdMem
                .Children.Clear()
                .Children.Add(stkpnlMemoryLoad)
            End With

        Catch ex As Exception

            mblnMemLoadError = True

            If mbwMemLoad.IsBusy Then mbwMemLoad.CancelAsync()

            MsgBox("Error displaying memory load metrics." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Sub
    Private Function Retrieve_DiskLoad() As List(Of clsDisk)

        Try

            mblnDiskLoadError = False

            Dim lstDiskLoad As New List(Of clsDisk)
            Dim objMgmt As ManagementObject
            Dim strScope As String = "\\" & mstrCurrentTargetServer & "\root\cimv2"
            Dim objDisk As New ManagementObjectSearcher(strScope, "SELECT * FROM Win32_DiskDrive")
            Dim objLogicalDisk As New ManagementObjectSearcher(strScope, "SELECT * FROM Win32_LogicalDisk")
            Dim sngLogicalDiskSize As Single
            Dim strLogicalDiskDeviceId As String
            Dim strLogicalDiskFS As String
            Dim strLogicalDiskName As String
            Dim sngLogicalDiskFree As Single
            Dim sngLogicalDiskUsed As Single
            Dim sngLogicalDiskLoad As Single
            Dim strLogicalDiskVolumeName As String

            For Each objMgmt In objLogicalDisk.Get
                Dim disk As New clsDisk
                sngLogicalDiskSize = IIf(IsNothing(objMgmt("size")), 0, Convert.ToInt64(objMgmt("size")))
                If sngLogicalDiskSize = 0 Then Continue For
                strLogicalDiskDeviceId = IIf(IsNothing(objMgmt("deviceid")), "", objMgmt("deviceid").ToString)
                strLogicalDiskFS = IIf(IsNothing(objMgmt("filesystem")), "", objMgmt("filesystem").ToString)
                strLogicalDiskName = IIf(IsNothing(objMgmt("name")), "", objMgmt("name").ToString)
                sngLogicalDiskFree = IIf(IsNothing(objMgmt("freespace")), 0, Convert.ToInt64(objMgmt("freespace")))
                sngLogicalDiskUsed = sngLogicalDiskSize - sngLogicalDiskFree
                sngLogicalDiskLoad = (sngLogicalDiskUsed / sngLogicalDiskSize) * 100
                strLogicalDiskVolumeName = IIf(IsNothing(objMgmt("volumename")), "", objMgmt("volumename").ToString)
                disk.LogicalDiskSize = sngLogicalDiskSize
                disk.LogicalDiskDeviceId = strLogicalDiskDeviceId
                disk.LogicalDiskFS = strLogicalDiskFS
                disk.LogicalDiskName = strLogicalDiskName
                disk.LogicalDiskFree = sngLogicalDiskFree
                disk.LogicalDiskUsed = sngLogicalDiskUsed
                disk.LogicalDiskLoad = sngLogicalDiskLoad
                disk.LogicalDiskVolumeName = strLogicalDiskVolumeName
                lstDiskLoad.Add(disk)
            Next

            Retrieve_DiskLoad = lstDiskLoad

        Catch ex As Exception

            mblnDiskLoadError = True

            If mbwDiskLoad.IsBusy Then mbwDiskLoad.CancelAsync()

            MsgBox("Error retrieving disk load values." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Function
    Private Sub Display_DiskLoad()

        Try

            mblnDiskLoadError = False

            If mlstDiskLoad Is Nothing Then Exit Sub

            Dim i As Integer = 0
            Dim intCurrRow As Integer = 0
            Dim intCurrColumn As Integer = 0
            Dim intTotRows As Integer = grdDisks.RowDefinitions.Count - 1
            Dim intTotColumns As Integer = grdDisks.ColumnDefinitions.Count - 1

            grdDisks.Children.Clear()

            For Each disk In mlstDiskLoad
                Dim stkpnlDiskLoad As New StackPanel
                Dim stkpnlTotDiskLoad As New StackPanel
                Dim lblDiskDriveLetter As New Label
                Dim prgbarDiskUsage As New ProgressBar
                Dim lblTotDiskUsage As New Label
                Dim lblDiskInfo As New Label
                With stkpnlDiskLoad
                    .Orientation = Orientation.Vertical
                    .HorizontalAlignment = HorizontalAlignment.Left
                    .VerticalAlignment = VerticalAlignment.Center
                End With
                stkpnlTotDiskLoad.Orientation = Orientation.Horizontal
                With lblDiskDriveLetter
                    .Content = disk.LogicalDiskDeviceId
                    .FontSize = 9
                    .Height = 22
                End With
                With prgbarDiskUsage
                    .Foreground = ProgressBarForeColor(disk.LogicalDiskLoad)
                    .Value = disk.LogicalDiskLoad
                    .Height = 15
                    .Width = 190
                End With
                With lblTotDiskUsage
                    .Content = " " & CInt(disk.LogicalDiskLoad) & "%"
                    .FontSize = 9
                    .Height = 22
                End With
                stkpnlTotDiskLoad.Children.Add(lblDiskDriveLetter)
                stkpnlTotDiskLoad.Children.Add(prgbarDiskUsage)
                stkpnlTotDiskLoad.Children.Add(lblTotDiskUsage)
                With lblDiskInfo
                    .Content = CalculateBestByteSize(disk.LogicalDiskUsed, True) & "/" & CalculateBestByteSize(disk.LogicalDiskSize, True) & " [" & disk.LogicalDiskVolumeName & "] - " & disk.LogicalDiskFS
                    .FontSize = 9
                    .Height = 22
                    .HorizontalAlignment = HorizontalAlignment.Center
                    .VerticalAlignment = VerticalAlignment.Top
                End With
                stkpnlDiskLoad.Children.Add(stkpnlTotDiskLoad)
                stkpnlDiskLoad.Children.Add(lblDiskInfo)
                With grdDisks
                    .SetRow(stkpnlDiskLoad, intCurrRow)
                    .SetColumn(stkpnlDiskLoad, intCurrColumn)
                    .Children.Add(stkpnlDiskLoad)
                End With
                If intCurrColumn = intTotColumns Then
                    intCurrColumn = 0
                    intCurrRow = intCurrRow + 1
                Else
                    intCurrColumn = intCurrColumn + 1
                End If
                i = i + 1
            Next

            scrlDisks.IsEnabled = (intCurrRow >= 3)

        Catch ex As Exception

            mblnDiskLoadError = True

            If mbwDiskLoad.IsBusy Then mbwDiskLoad.CancelAsync()

            MsgBox("Error displaying Disk load metrics." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Sub
    Private Sub cmdSearchApplications_Click(sender As Object, e As RoutedEventArgs) Handles cmdSearchApplications.Click

        Try

            If mstrCurrentTargetServer = "" Then Exit Sub

            dtgrdApplications.DataContext = Nothing
            dtgrdApplications.ItemsSource = Nothing

            Dim strApplicationsFilter As String = IIf(String.IsNullOrWhiteSpace(txtApplicationName.Text), "", "Name LIKE '%" & txtApplicationName.Text.Trim & "%'")
            Dim strEnvironment As String
            Dim strEnvironmentFilter As String

            If chk32Bit.IsChecked Then strEnvironment = strEnvironment & " Environment = '32-bit' OR"
            If chk64Bit.IsChecked Then strEnvironment = strEnvironment & " Environment = '64-bit' OR"

            If Len(strEnvironment) > 0 Then strEnvironmentFilter = " (" & Strings.Left(strEnvironment, Len(strEnvironment) - 2) & ") "

            If Len(strApplicationsFilter) > 0 Then
                If Len(strEnvironmentFilter) > 0 Then
                    mstrApplicationsFilter = strApplicationsFilter & " AND " & strEnvironmentFilter
                Else
                    mstrApplicationsFilter = strApplicationsFilter
                End If
            Else
                If Len(strEnvironmentFilter) > 0 Then
                    mstrApplicationsFilter = strEnvironmentFilter
                Else
                    mstrApplicationsFilter = ""
                End If
            End If

            mbwApplications.RunWorkerAsync()

        Catch ex As Exception

            MsgBox("Error initializing load of installed applications at cmdSearchApplications.Click." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Sub
    Private Function Retrieve_Applications() As DataView

        Try

            Dim dtApplications As DataTable = New DataTable("apps")
            Dim col1 As DataColumn = New DataColumn("Name", System.Type.GetType("System.String"))
            Dim col2 As DataColumn = New DataColumn("Version", System.Type.GetType("System.String"))
            Dim col3 As DataColumn = New DataColumn("Size", System.Type.GetType("System.String"))
            Dim col4 As DataColumn = New DataColumn("Installed", System.Type.GetType("System.String"))
            Dim col5 As DataColumn = New DataColumn("Location", System.Type.GetType("System.String"))
            Dim col6 As DataColumn = New DataColumn("Publisher", System.Type.GetType("System.String"))
            Dim col7 As DataColumn = New DataColumn("Environment", System.Type.GetType("System.String"))
            Dim strHostName As String = IIf(IsNumeric(Replace(mstrCurrentTargetServer, ".", "")), Get_HostName_Or_IPAddress(mstrCurrentTargetServer), mstrCurrentTargetServer)
            Dim regKey64 As RegistryKey
            Dim regKey32 As RegistryKey
            Dim strKeyPath64 As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            Dim strKeyPath32 As String = "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            Dim strSubKeyName As String
            Dim objApplications(6) As Object

            With dtApplications.Columns
                .Add(col1)
                .Add(col2)
                .Add(col3)
                .Add(col4)
                .Add(col5)
                .Add(col6)
                .Add(col7)
            End With

            'regKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
            regKey64 = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, strHostName).OpenSubKey(strKeyPath64)

            For Each strSubKeyName In regKey64.GetSubKeyNames
                Dim subRegKey As RegistryKey = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, strHostName).OpenSubKey(strKeyPath64 & "\" & strSubKeyName)
                If IsNothing(subRegKey.GetValue("DisplayName")) Then
                    objApplications(0) = ""
                Else
                    objApplications(0) = subRegKey.GetValue("DisplayName").ToString
                End If
                If IsNothing(subRegKey.GetValue("DisplayVersion")) Then
                    objApplications(1) = ""
                Else
                    objApplications(1) = subRegKey.GetValue("DisplayVersion").ToString
                End If
                If IsNothing(subRegKey.GetValue("EstimatedSize")) Then
                    objApplications(2) = ""
                Else
                    objApplications(2) = CalculateBestByteSize(Convert.ToInt64(subRegKey.GetValue("EstimatedSize")) * 1000)
                End If
                If IsNothing(subRegKey.GetValue("InstallDate")) Then
                    objApplications(3) = ""
                Else
                    If InStr(subRegKey.GetValue("InstallDate").ToString, "/") = 0 Then
                        objApplications(3) = MakeFriendlyDateString(subRegKey.GetValue("InstallDate").ToString)
                    Else
                        objApplications(3) = subRegKey.GetValue("InstallDate").ToString
                    End If
                End If
                If IsNothing(subRegKey.GetValue("InstallLocation")) Then
                    objApplications(4) = ""
                Else
                    objApplications(4) = subRegKey.GetValue("InstallLocation").ToString
                End If
                If IsNothing(subRegKey.GetValue("Publisher")) Then
                    objApplications(5) = ""
                Else
                    objApplications(5) = subRegKey.GetValue("Publisher").ToString
                End If
                If objApplications(0) <> "" Or objApplications(1) <> "" Or objApplications(2) <> "" Then
                    objApplications(6) = "64-bit"
                    dtApplications.LoadDataRow(objApplications, True)
                End If
            Next

            regKey32 = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, strHostName).OpenSubKey(strKeyPath32)

            For Each strSubKeyName In regKey32.GetSubKeyNames
                Dim subRegKey As RegistryKey = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, strHostName).OpenSubKey(strKeyPath32 & "\" & strSubKeyName)
                If IsNothing(subRegKey.GetValue("DisplayName")) Then
                    objApplications(0) = ""
                Else
                    objApplications(0) = subRegKey.GetValue("DisplayName").ToString
                End If
                If IsNothing(subRegKey.GetValue("DisplayVersion")) Then
                    objApplications(1) = ""
                Else
                    objApplications(1) = subRegKey.GetValue("DisplayVersion").ToString
                End If
                If IsNothing(subRegKey.GetValue("EstimatedSize")) Then
                    objApplications(2) = ""
                Else
                    objApplications(2) = CalculateBestByteSize(Convert.ToInt64(subRegKey.GetValue("EstimatedSize")) * 1000)
                End If
                If IsNothing(subRegKey.GetValue("InstallDate")) Then
                    objApplications(3) = ""
                Else
                    If InStr(subRegKey.GetValue("InstallDate").ToString, "/") = 0 Then
                        objApplications(3) = MakeFriendlyDateString(subRegKey.GetValue("InstallDate").ToString)
                    Else
                        objApplications(3) = subRegKey.GetValue("InstallDate").ToString
                    End If
                End If
                If IsNothing(subRegKey.GetValue("InstallLocation")) Then
                    objApplications(4) = ""
                Else
                    objApplications(4) = subRegKey.GetValue("InstallLocation").ToString
                End If
                If IsNothing(subRegKey.GetValue("Publisher")) Then
                    objApplications(5) = ""
                Else
                    objApplications(5) = subRegKey.GetValue("Publisher").ToString
                End If
                If objApplications(0) <> "" Or objApplications(1) <> "" Or objApplications(2) <> "" Then
                    objApplications(6) = "32-bit"
                    dtApplications.LoadDataRow(objApplications, True)
                End If
            Next

            If Len(mstrApplicationsFilter) > 0 Then
                Dim dvApplications As New DataView(dtApplications)
                dvApplications.RowFilter = mstrApplicationsFilter
                Retrieve_Applications = dvApplications
            Else
                Retrieve_Applications = dtApplications.DefaultView
            End If

        Catch ex As Exception

            MsgBox("Error retrieving installed applications." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Function
    Private Sub Display_Applications()

        Try

            With dtgrdApplications
                .IsReadOnly = True
                .RowHeight = 22
                .AlternatingRowBackground = mbrshAlternatingRowColor
                .CanUserReorderColumns = True
                .CanUserResizeColumns = True
                .CanUserResizeRows = True
                .CanUserSortColumns = True
            End With

            mdvApplications.Sort = "Name ASC"

            dtgrdApplications.ItemsSource = mdvApplications

            lblApplicationsCount.Content = "Applications: " & Format(dtgrdApplications.Items.Count, "#,##0")

        Catch ex As Exception

            MsgBox("Error displaying installed applications." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Sub
    Private Sub cmdSearchProcesses_Click(sender As Object, e As RoutedEventArgs) Handles cmdSearchProcesses.Click

        Try

            If mstrCurrentTargetServer = "" Then Exit Sub

            dtgrdProcesses.DataContext = Nothing
            dtgrdProcesses.ItemsSource = Nothing

            Dim strProcessesFilter As String = txtProcessName.Text.Trim

            mstrProcessesFilter = IIf(Len(strProcessesFilter) > 0, "Name LIKE '%" & strProcessesFilter & "%'", "")

            mbwProcesses.RunWorkerAsync()

        Catch ex As Exception

            MsgBox("Error initializing load of processes at cmdSearchProcesses.Click." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Sub
    Private Function Retrieve_Processes() As DataView

        Try

            Dim dtProcesses As DataTable = New DataTable("procs")
            Dim col1 As DataColumn = New DataColumn("Name", System.Type.GetType("System.String"))
            Dim col2 As DataColumn = New DataColumn("ProcessId", System.Type.GetType("System.String"))
            Dim col3 As DataColumn = New DataColumn("Priority", System.Type.GetType("System.String"))
            Dim col4 As DataColumn = New DataColumn("Status", System.Type.GetType("System.String"))
            Dim col5 As DataColumn = New DataColumn("ExecutablePath", System.Type.GetType("System.String"))
            Dim col6 As DataColumn = New DataColumn("CommandLine", System.Type.GetType("System.String"))
            Dim objProcesses(5) As Object
            Dim strScope As String = "\\" & mstrCurrentTargetServer & "\root\cimv2"
            Dim objProcess As New ManagementObjectSearcher(strScope, "SELECT * FROM Win32_Process")
            Dim objMgmt As ManagementObject

            With dtProcesses.Columns
                .Add(col1)
                .Add(col2)
                .Add(col3)
                .Add(col4)
                .Add(col5)
                .Add(col6)
            End With

            For Each objMgmt In objProcess.Get
                If IsNothing(objMgmt("name")) Then
                    objProcesses(0) = ""
                Else
                    objProcesses(0) = objMgmt("name").ToString
                End If
                If IsNothing(objMgmt("processid")) Then
                    objProcesses(1) = ""
                Else
                    objProcesses(1) = objMgmt("processid").ToString
                End If
                If IsNothing(objMgmt("priority")) Then
                    objProcesses(2) = ""
                Else
                    objProcesses(2) = objMgmt("priority").ToString
                End If
                If IsNothing(objMgmt("status")) Then
                    objProcesses(3) = ""
                Else
                    objProcesses(3) = objMgmt("status").ToString
                End If
                If IsNothing(objMgmt("executablepath")) Then
                    objProcesses(4) = ""
                Else
                    objProcesses(4) = objMgmt("executablepath").ToString
                End If
                If IsNothing(objMgmt("commandline")) Then
                    objProcesses(5) = ""
                Else
                    objProcesses(5) = objMgmt("commandline").ToString
                End If
                dtProcesses.LoadDataRow(objProcesses, True)
            Next

            If Len(mstrProcessesFilter) > 0 Then
                Dim dvProcesses As New DataView(dtProcesses)
                dvProcesses.RowFilter = mstrProcessesFilter
                Retrieve_Processes = dvProcesses
            Else
                Retrieve_Processes = dtProcesses.DefaultView
            End If

        Catch ex As Exception

            MsgBox("Error retrieving process information." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Function
    Private Sub Display_Processes()

        Try

            With dtgrdProcesses
                .IsReadOnly = True
                .RowHeight = 22
                .AlternatingRowBackground = mbrshAlternatingRowColor
                .CanUserReorderColumns = True
                .CanUserResizeColumns = True
                .CanUserResizeRows = True
                .CanUserSortColumns = True
            End With

            mdvProcesses.Sort = "Name ASC"

            dtgrdProcesses.ItemsSource = mdvProcesses

            lblProcessesCount.Content = "Processes: " & Format(dtgrdProcesses.Items.Count, "#,##0")

        Catch ex As Exception

            MsgBox("Error displaying processes." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Sub
    Private Sub cmdSearchServices_Click(sender As Object, e As RoutedEventArgs) Handles cmdSearchServices.Click

        Try

            If mstrCurrentTargetServer = "" Then Exit Sub

            dtgrdServices.DataContext = Nothing
            dtgrdServices.ItemsSource = Nothing

            Dim strStatus As String
            Dim strStatusFilter As String
            Dim strStartup As String
            Dim strStartupFilter As String

            mblnDeviceDriver = (rbDeviceDriver.IsChecked)

            If chkRunning.IsChecked Then strStatus = strStatus & " Status = 'Running' OR"
            If chkStopped.IsChecked Then strStatus = strStatus & " Status = 'Stopped' OR"

            If Len(strStatus) > 0 Then strStatusFilter = " (" & Strings.Left(strStatus, Len(strStatus) - 2) & ") "

            If chkAutomatic.IsChecked Then strStartup = strStartup & " Startup = 'Automatic' OR"
            If chkManual.IsChecked Then strStartup = strStartup & " Startup = 'Manual' OR"
            If chkDisabled.IsChecked Then strStartup = strStartup & " Startup = 'Disabled' OR"

            If Len(strStartup) > 0 Then strStartupFilter = " (" & Strings.Left(strStartup, Len(strStartup) - 2) & ") "

            If Len(strStatusFilter) > 0 Then
                If Len(strStartupFilter) > 0 Then
                    mstrServicesFilter = strStatusFilter & " AND " & strStartupFilter
                Else
                    mstrServicesFilter = strStatusFilter
                End If
            Else
                If Len(strStartupFilter) > 0 Then
                    mstrServicesFilter = strStartupFilter
                Else
                    mstrServicesFilter = ""
                End If
            End If

            mbwServices.RunWorkerAsync()

        Catch ex As Exception

            MsgBox("Error initializing load of services at cmdSearchServices.Click." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Sub
    Private Function Retrieve_Services(Optional ByVal blnDeviceDriver As Boolean = False) As DataView

        Try

            Dim sc As New ServiceController()
            Dim dtServices As DataTable = New DataTable("services")
            Dim col1 As DataColumn = New DataColumn("Name", System.Type.GetType("System.String"))
            Dim col2 As DataColumn = New DataColumn("DisplayName", System.Type.GetType("System.String"))
            Dim col3 As DataColumn = New DataColumn("Type", System.Type.GetType("System.String"))
            Dim col4 As DataColumn = New DataColumn("Status", System.Type.GetType("System.String"))
            Dim col5 As DataColumn = New DataColumn("Startup", System.Type.GetType("System.String"))
            Dim objServicesFields(4) As Object

            With dtServices.Columns
                .Add(col1)
                .Add(col2)
                .Add(col3)
                .Add(col4)
                .Add(col5)
            End With

            If blnDeviceDriver Then
                For Each service As ServiceController In ServiceController.GetDevices(mstrCurrentTargetServer)
                    objServicesFields(0) = service.ServiceName
                    objServicesFields(1) = service.DisplayName
                    objServicesFields(2) = service.ServiceType.ToString
                    objServicesFields(3) = service.Status.ToString
                    objServicesFields(4) = service.StartType.ToString
                    dtServices.LoadDataRow(objServicesFields, True)
                Next
            Else
                For Each service As ServiceController In ServiceController.GetServices(mstrCurrentTargetServer)
                    objServicesFields(0) = service.ServiceName
                    objServicesFields(1) = service.DisplayName
                    objServicesFields(2) = service.ServiceType.ToString
                    objServicesFields(3) = service.Status.ToString
                    objServicesFields(4) = service.StartType.ToString
                    dtServices.LoadDataRow(objServicesFields, True)
                Next
            End If

            If Len(mstrServicesFilter) > 0 Then
                Dim dvServices As New DataView(dtServices)
                dvServices.RowFilter = mstrServicesFilter
                Retrieve_Services = dvServices
            Else
                Retrieve_Services = dtServices.DefaultView
            End If

        Catch ex As Exception

            MsgBox("Error retreiving services information." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Function
    Private Sub Display_Services()

        Try

            With dtgrdServices
                .IsReadOnly = True
                .RowHeight = 22
                .AlternatingRowBackground = mbrshAlternatingRowColor
                .CanUserReorderColumns = True
                .CanUserResizeColumns = True
                .CanUserResizeRows = True
                .CanUserSortColumns = True
            End With

            dtgrdServices.ItemsSource = mdvServices

            lblServicesCount.Content = "Services: " & Format(dtgrdServices.Items.Count, "#,##0")

        Catch ex As Exception

            MsgBox("Error displaying services." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Sub
    Private Sub cmdSearchEventLogs_Click(sender As Object, e As RoutedEventArgs) Handles cmdSearchEventLogs.Click

        Try

            If mstrCurrentTargetServer = "" Then Exit Sub

            dtgrdEventLogs.DataContext = Nothing

            Dim strLogEntryTypes As String
            Dim strLogType As String

            If rbApplication.IsChecked Then strLogType = "application"
            If rbSystem.IsChecked Then strLogType = "system"

            mstrLogType = strLogType

            If chkCritical.IsChecked Then strLogEntryTypes = strLogEntryTypes & "Critical|"
            If chkWarning.IsChecked Then strLogEntryTypes = strLogEntryTypes & "Warning|"
            If chkVerbose.IsChecked Then strLogEntryTypes = strLogEntryTypes & "Verbose|"
            If chkError.IsChecked Then strLogEntryTypes = strLogEntryTypes & "Error|"
            If chkInformation.IsChecked Then strLogEntryTypes = strLogEntryTypes & "Information|"

            If chkCritical.IsChecked Or chkWarning.IsChecked Or chkVerbose.IsChecked Or chkError.IsChecked Or chkInformation.IsChecked Then strLogEntryTypes = Strings.Left(strLogEntryTypes, Len(strLogEntryTypes) - 1)

            mstrLogEntryTypes = strLogEntryTypes

            mstrLogEntryFromDate = IIf(dtpFromDate.SelectedDate IsNot Nothing, "#" & dtpFromDate.SelectedDate & "#", "")
            mstrLogEntryToDate = IIf(dtpToDate.SelectedDate IsNot Nothing, "#" & dtpToDate.SelectedDate & " 23:59:59#", "")

            mbwEventLog.RunWorkerAsync()

        Catch ex As Exception

            MsgBox("Error initializing load of event logs at cmdSearchEventLogs.Click." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Sub
    Private Function Retrieve_EventLog() As IEnumerable(Of Object)

        Try

            Dim evtLog As New EventLog()
            Dim strLongEntryTypesFilter As String()
            Dim objFilteredLogEntries As IEnumerable(Of Object)

            With evtLog
                .Log = mstrLogType
                .MachineName = mstrCurrentTargetServer
            End With

            If Len(mstrLogEntryTypes) > 0 Then
                strLongEntryTypesFilter = Split(mstrLogEntryTypes, "|")
                If mstrLogEntryFromDate <> "" Then
                    If mstrLogEntryToDate <> "" Then
                        objFilteredLogEntries = (From evtLogEntry In evtLog.Entries Where strLongEntryTypesFilter.Contains(evtLogEntry.EntryType.ToString) And Convert.ToDateTime(evtLogEntry.TimeGenerated) >= mstrLogEntryFromDate And Convert.ToDateTime(evtLogEntry.TimeGenerated) <= mstrLogEntryToDate Select evtLogEntry.EntryType, evtLogEntry.TimeGenerated, evtLogEntry.Source, evtLogEntry.EventID, evtLogEntry.MachineName, evtLogEntry.Message).ToList
                    Else
                        objFilteredLogEntries = (From evtLogEntry In evtLog.Entries Where strLongEntryTypesFilter.Contains(evtLogEntry.EntryType.ToString) And Convert.ToDateTime(evtLogEntry.TimeGenerated) >= mstrLogEntryFromDate Select evtLogEntry.EntryType, evtLogEntry.TimeGenerated, evtLogEntry.Source, evtLogEntry.EventID, evtLogEntry.MachineName, evtLogEntry.Message).ToList
                    End If
                Else
                    If mstrLogEntryToDate <> "" Then
                        objFilteredLogEntries = (From evtLogEntry In evtLog.Entries Where strLongEntryTypesFilter.Contains(evtLogEntry.EntryType.ToString) And Convert.ToDateTime(evtLogEntry.TimeGenerated) <= mstrLogEntryToDate Select evtLogEntry.EntryType, evtLogEntry.TimeGenerated, evtLogEntry.Source, evtLogEntry.EventID, evtLogEntry.MachineName, evtLogEntry.Message).ToList
                    Else
                        objFilteredLogEntries = (From evtLogEntry In evtLog.Entries Where strLongEntryTypesFilter.Contains(evtLogEntry.EntryType.ToString) Select evtLogEntry.EntryType, evtLogEntry.TimeGenerated, evtLogEntry.Source, evtLogEntry.EventID, evtLogEntry.MachineName, evtLogEntry.Message).ToList
                    End If
                End If
            Else
                If mstrLogEntryFromDate <> "" Then
                    If mstrLogEntryToDate <> "" Then
                        objFilteredLogEntries = (From evtLogEntry In evtLog.Entries Where Convert.ToDateTime(evtLogEntry.TimeGenerated) >= mstrLogEntryFromDate And Convert.ToDateTime(evtLogEntry.TimeGenerated) <= mstrLogEntryToDate Select evtLogEntry.EntryType, evtLogEntry.TimeGenerated, evtLogEntry.Source, evtLogEntry.EventID, evtLogEntry.MachineName, evtLogEntry.Message).ToList
                    Else
                        objFilteredLogEntries = (From evtLogEntry In evtLog.Entries Where Convert.ToDateTime(evtLogEntry.TimeGenerated) >= mstrLogEntryFromDate Select evtLogEntry.EntryType, evtLogEntry.TimeGenerated, evtLogEntry.Source, evtLogEntry.EventID, evtLogEntry.MachineName, evtLogEntry.Message).ToList
                    End If
                Else
                    If mstrLogEntryToDate <> "" Then
                        objFilteredLogEntries = (From evtLogEntry In evtLog.Entries Where Convert.ToDateTime(evtLogEntry.TimeGenerated) <= mstrLogEntryToDate Select evtLogEntry.EntryType, evtLogEntry.TimeGenerated, evtLogEntry.Source, evtLogEntry.EventID, evtLogEntry.MachineName, evtLogEntry.Message).ToList
                    Else
                        objFilteredLogEntries = (From evtLogEntry In evtLog.Entries Select evtLogEntry.EntryType, evtLogEntry.TimeGenerated, evtLogEntry.Source, evtLogEntry.EventID, evtLogEntry.MachineName, evtLogEntry.Message).ToList
                    End If
                End If
            End If

            Retrieve_EventLog = objFilteredLogEntries

        Catch ex As Exception

            MsgBox("Error retrieving event log entries." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Function
    Private Sub Display_EventLog()

        Try

            With dtgrdEventLogs
                .IsReadOnly = True
                .RowHeight = 22
                .AlternatingRowBackground = mbrshAlternatingRowColor
                .CanUserReorderColumns = True
                .CanUserResizeColumns = True
                .CanUserResizeRows = True
                .CanUserSortColumns = True
            End With

            'dtgrdEventLogs.DataContext = mobjLogEntries
            dtgrdEventLogs.ItemsSource = mobjLogEntries

            lblEventLogCount.Content = "Log Entries: " & Format(dtgrdEventLogs.Items.Count, "#,##0")

        Catch ex As Exception

            MsgBox("Error displaying event log entries." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Sub
    Private Sub dtgrdEventLogs_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles dtgrdEventLogs.SelectionChanged

        Try

            If dtgrdEventLogs.Items.Count = 0 Then Exit Sub
            If dtgrdEventLogs.SelectedIndex < 0 Then Exit Sub

            lblEventType.Content = dtgrdEventLogs.SelectedValue.EntryType.ToString
            lblGenerated.Content = dtgrdEventLogs.SelectedValue.TimeGenerated.ToString
            lblSource.Content = dtgrdEventLogs.SelectedValue.Source.ToString
            lblEventId.Content = dtgrdEventLogs.SelectedValue.EventId.ToString
            lblMachineName.Content = dtgrdEventLogs.SelectedValue.MachineName.ToString
            txtMessage.Text = dtgrdEventLogs.SelectedValue.Message.ToString

        Catch ex As Exception

            MsgBox("Error displaying event log entry details." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Sub
    Private Sub dtgrdEventLogs_LostFocus(sender As Object, e As RoutedEventArgs) Handles dtgrdEventLogs.LostFocus

        lblEventType.Content = ""
        lblGenerated.Content = ""
        lblSource.Content = ""
        lblEventId.Content = ""
        lblMachineName.Content = ""
        txtMessage.Text = ""

    End Sub
    Private Sub chkAccessedBy_Click(sender As Object, e As RoutedEventArgs) Handles chkAccessedBy.Click

        txtAccessedBy.IsEnabled = chkAccessedBy.IsChecked

        rbAnd.IsEnabled = (chkAccessedBy.IsChecked And chkOpenFile.IsChecked)
        rbOr.IsEnabled = (chkAccessedBy.IsChecked And chkOpenFile.IsChecked)

    End Sub
    Private Sub chkOpenFile_Click(sender As Object, e As RoutedEventArgs) Handles chkOpenFile.Click

        txtOpenFile.IsEnabled = chkOpenFile.IsChecked

        rbAnd.IsEnabled = (chkAccessedBy.IsChecked And chkOpenFile.IsChecked)
        rbOr.IsEnabled = (chkAccessedBy.IsChecked And chkOpenFile.IsChecked)

    End Sub
    Private Sub cmdSearchOpenFiles_Click(sender As Object, e As RoutedEventArgs) Handles cmdSearchOpenFiles.Click

        Try

            If mstrCurrentTargetServer = "" Then Exit Sub

            dtgrdOpenFiles.DataContext = Nothing
            dtgrdOpenFiles.ItemsSource = Nothing

            Dim strAccessFilter As String
            Dim strFileFilter As String
            Dim strOrAnd As String

            If txtAccessedBy.IsEnabled And Not String.IsNullOrWhiteSpace(txtAccessedBy.Text) Then strAccessFilter = "AccessedBy LIKE('%" & txtAccessedBy.Text.Trim & "%')"
            If txtOpenFile.IsEnabled And Not String.IsNullOrWhiteSpace(txtOpenFile.Text) Then strFileFilter = "OpenFile LIKE('%" & txtOpenFile.Text.Trim & "%')"

            If rbOr.IsEnabled And rbOr.IsChecked Then strOrAnd = " OR "
            If rbAnd.IsEnabled And rbAnd.IsChecked Then strOrAnd = " AND "

            'Display open files and folders
            If Len(strAccessFilter) > 0 Or Len(strFileFilter) > 0 Then mstrOpenFilesFilter = strAccessFilter & strOrAnd & strFileFilter

            mbwOpenFiles.RunWorkerAsync()

        Catch ex As Exception

            MsgBox("Error initializing load of open files at cmdSearchOpenFiles.Click." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Sub
    Private Function Retrieve_OpenFiles() As DataView

        Try

            Dim procOpenFiles As New Process
            Dim dtOpenFiles As DataTable = New DataTable("openfiles")
            Dim col1 As DataColumn = New DataColumn("HostName", System.Type.GetType("System.String"))
            Dim col2 As DataColumn = New DataColumn("Id", System.Type.GetType("System.String"))
            Dim col3 As DataColumn = New DataColumn("AccessedBy", System.Type.GetType("System.String"))
            Dim col4 As DataColumn = New DataColumn("Type", System.Type.GetType("System.String"))
            Dim col5 As DataColumn = New DataColumn("Locks", System.Type.GetType("System.String"))
            Dim col6 As DataColumn = New DataColumn("OpenMode", System.Type.GetType("System.String"))
            Dim col7 As DataColumn = New DataColumn("OpenFile", System.Type.GetType("System.String"))
            Dim strStandardError As String
            Dim strOutputLine As String
            Dim objOpenFilesFields(6) As Object

            With dtOpenFiles.Columns
                .Add(col1)
                .Add(col2)
                .Add(col3)
                .Add(col4)
                .Add(col5)
                .Add(col6)
                .Add(col7)
            End With

            With procOpenFiles.StartInfo
                .FileName = "C:\Windows\System32\openfiles.exe"
                .Arguments = "/query /s " & mstrCurrentTargetServer & " /v"
                .UseShellExecute = False
                .CreateNoWindow = True
                .RedirectStandardError = True
                .RedirectStandardOutput = True
                .Verb = "runas"
            End With

            procOpenFiles.Start()

            If procOpenFiles.StandardOutput Is Nothing Then Exit Function

            strStandardError = procOpenFiles.StandardError.ReadToEnd

            If Not String.IsNullOrEmpty(strStandardError) Then
                MsgBox(Replace(strStandardError, "ERROR: ", ""), vbExclamation, "AdvSysInfo")
                Exit Function
            End If

            'Read the first blank line and then the header line
            strOutputLine = procOpenFiles.StandardOutput.ReadLine
            strOutputLine = procOpenFiles.StandardOutput.ReadLine

            Do
                strOutputLine = procOpenFiles.StandardOutput.ReadLine
                If Len(strOutputLine) > 0 And InStr(strOutputLine, "=====") = 0 Then
                    objOpenFilesFields(0) = strOutputLine.Substring(0, 15).Trim   'Host Name
                    objOpenFilesFields(1) = strOutputLine.Substring(16, 8).Trim   'ID
                    objOpenFilesFields(2) = strOutputLine.Substring(25, 20).Trim  'Accessed By        
                    objOpenFilesFields(3) = strOutputLine.Substring(46, 10).Trim  'Type
                    objOpenFilesFields(4) = strOutputLine.Substring(57, 10).Trim  'Locks        
                    objOpenFilesFields(5) = strOutputLine.Substring(68, 15).Trim  'Open Mode
                    objOpenFilesFields(6) = strOutputLine.Substring(84).Trim      'Open File
                    dtOpenFiles.LoadDataRow(objOpenFilesFields, True)
                End If
            Loop Until strOutputLine Is Nothing

            If Len(mstrOpenFilesFilter) > 0 Then
                Dim dvOpenFiles As New DataView(dtOpenFiles)
                dvOpenFiles.RowFilter = mstrOpenFilesFilter
                Retrieve_OpenFiles = dvOpenFiles
            Else
                Retrieve_OpenFiles = dtOpenFiles.DefaultView
            End If

        Catch ex As Exception

            MsgBox("Error retrieving open file entries." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Function
    Private Sub Display_OpenFiles()

        Try

            With dtgrdOpenFiles
                .IsReadOnly = True
                .RowHeight = 22
                .AlternatingRowBackground = mbrshAlternatingRowColor
                .CanUserReorderColumns = True
                .CanUserResizeColumns = True
                .CanUserResizeRows = True
                .CanUserSortColumns = True
            End With

            dtgrdOpenFiles.ItemsSource = mdvOpenFiles

            lblOpenFilesCount.Content = "Open Files: " & Format(dtgrdOpenFiles.Items.Count, "#,##0")

        Catch ex As Exception

            MsgBox("Error displaying open file entries." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Sub
    Private Sub mbwSysInfo_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs)

        If mbwSysInfo.CancellationPending Then
            e.Cancel = True
            Exit Sub
        End If

        mobjSysInfo = Retrieve_SysInfo()

    End Sub
    Private Sub mbwSysInfo_WorkCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs)

        If Not e.Cancelled Then
            Display_SysInfo()
        End If

    End Sub
    Private Sub mbwCPULoad_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs)

        If mbwCPULoad.CancellationPending Then
            e.Cancel = True
            Exit Sub
        End If

        mlstCPULoad = Retrieve_CPULoad()

        System.Threading.Thread.Sleep(2000)

    End Sub
    Private Sub mbwCPULoad_WorkCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs)

        If mblnCPULoadError Then Exit Sub     'Do not loop and run the worker again

        If Not e.Cancelled Then
            Display_CPULoad()
            mbwCPULoad.RunWorkerAsync()
        End If

    End Sub
    Private Sub mbwMemLoad_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs)

        If mbwMemLoad.CancellationPending Then
            e.Cancel = True
            Exit Sub
        End If

        mobjMemLoad = Retrieve_MemLoad()

        System.Threading.Thread.Sleep(2000)

    End Sub
    Private Sub mbwMemLoad_WorkCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs)

        If mblnMemLoadError Then Exit Sub    'Do not loop and run the worker again

        If Not e.Cancelled Then
            Display_MemLoad()
            mbwMemLoad.RunWorkerAsync()
        End If

    End Sub
    Private Sub mbwDiskLoad_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs)

        If mbwDiskLoad.CancellationPending Then
            e.Cancel = True
            Exit Sub
        End If

        mlstDiskLoad = Retrieve_DiskLoad()

        System.Threading.Thread.Sleep(5000)

    End Sub
    Private Sub mbwDiskLoad_WorkCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs)

        If mblnDiskLoadError Then Exit Sub    'Do not loop and run the worker again

        If Not e.Cancelled Then
            Display_DiskLoad()
            mbwDiskLoad.RunWorkerAsync()
        End If

    End Sub
    Private Sub mbwApplications_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs)

        If mbwApplications.CancellationPending Then
            e.Cancel = True
            Exit Sub
        End If

        mbwApplications.ReportProgress(0)

        mdvApplications = Retrieve_Applications()

        mbwApplications.ReportProgress(100)

    End Sub
    Private Sub mbwApplications_ProgressChanged(ByVal sender As System.Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs)

        lblApplicationsCount.Content = "Retrieving Applications..."

    End Sub
    Private Sub mbwApplications_WorkCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs)

        If Not e.Cancelled Then
            Display_Applications()
        End If

    End Sub
    Private Sub mbwProcesses_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs)

        If mbwProcesses.CancellationPending Then
            e.Cancel = True
            Exit Sub
        End If

        mbwProcesses.ReportProgress(0)

        mdvProcesses = Retrieve_Processes()

        mbwProcesses.ReportProgress(100)

    End Sub
    Private Sub mbwProcesses_ProgressChanged(ByVal sender As System.Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs)

        lblProcessesCount.Content = "Retrieving processes..."

    End Sub
    Private Sub mbwProcesses_WorkCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs)

        If Not e.Cancelled Then
            Display_Processes()
        End If

    End Sub

    Private Sub mbwServices_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs)

        If mbwServices.CancellationPending Then
            e.Cancel = True
            Exit Sub
        End If

        mbwServices.ReportProgress(0)

        mdvServices = Retrieve_Services(mblnDeviceDriver)
        mbwServices.ReportProgress(100)

    End Sub
    Private Sub mbwServices_ProgressChanged(ByVal sender As System.Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs)

        lblServicesCount.Content = "Retrieving services..."

    End Sub
    Private Sub mbwServices_WorkCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs)

        If Not e.Cancelled Then
            Display_Services()
        End If

    End Sub
    Private Sub mbwEventLog_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs)

        If mbwEventLog.CancellationPending Then
            e.Cancel = True
            Exit Sub
        End If

        mbwEventLog.ReportProgress(0)

        mobjLogEntries = Retrieve_EventLog()

        mbwEventLog.ReportProgress(100)

    End Sub
    Private Sub mbwEventLog_ProgressChanged(ByVal sender As System.Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs)

        lblEventLogCount.Content = "Retrieving log entries..."

    End Sub
    Private Sub mbwEventLog_WorkCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs)

        If Not e.Cancelled Then
            Display_EventLog()
        End If

    End Sub
    Private Sub mbwOpenFiles_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs)

        If mbwOpenFiles.CancellationPending Then
            e.Cancel = True
            Exit Sub
        End If

        mbwOpenFiles.ReportProgress(0)

        mdvOpenFiles = Retrieve_OpenFiles()

        mbwOpenFiles.ReportProgress(100)

    End Sub
    Private Sub mbwOpenFiles_ProgressChanged(ByVal sender As System.Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs)

        lblOpenFilesCount.Content = "Retrieving open files..."

    End Sub
    Private Sub mbwOpenFiles_WorkCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs)

        If Not e.Cancelled Then
            Display_OpenFiles()
        End If

    End Sub

End Class
'Parking Lot

'Dim x() As EventLog = EventLog.GetEventLogs("10.20.4.37")
'Dim e As EventLogEntryCollection

'e = x(0).Entries         'Application log
'e = x(1).Entries         'Hardware Events
'e = x(2).Entries         'IntelAudioServiceLog
'e = x(3).Entries         'Internet Explorer
'e = x(4).Entries         'Key Management Service
'e = x(5).Entries         'Microsoft-Windows-Windows Defender/Operational
'e = x(6).Entries         'Microsoft Office Alerts
'e = x(7).Entries         'OneApp_IGCC
'e = x(8).Entries         'Parameters
'e = x(9).Entries         'Security
'e = x(10).Entries        'State
'e = x(11).Entries        'System
'e = x(12).Entries        'Tanium Protect - AppLocker
'e = x(13).Entries        'Tanium Protect - AppLocker - MSI and Script
'e = x(14).Entries        'Azure
'e = x(15).Entries        'Windows PowerShell