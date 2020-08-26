Private Function Retrieve_CPULoad() As List(Of clsCPU)

    Try
        '--------------------------------------------------------------------------------------------------------------------------------------------
        '  METHOD 1: USING PASSED IN CREDENTIALS
        '--------------------------------------------------------------------------------------------------------------------------------------------
        Dim connOpt As New ConnectionOptions()

        With connOpt
            .Impersonation = ImpersonationLevel.Impersonate
            .Authentication = AuthenticationLevel.Packet
            .Timeout = New TimeSpan(0, 0, 30)
            .EnablePrivileges = True
            .Username = "dgrewal"
            .Password = "my cool password"
        End With

        Dim mgmtPath As New ManagementPath()

        With mgmtPath
            .NamespacePath = "\root\cimv2"
            .Server = mstrCurrentTargetServer
        End With

        Dim mgmtScope As New ManagementScope(mgmtPath, connOpt)
        Dim objCPU As New ManagementObjectSearcher(mgmtScope, New ObjectQuery("SELECT * FROM Win32_Processor"))

        mgmtScope.Connect()
        '--------------------------------------------------------------------------------------------------------------------------------------------    
        '  METHOD 2: USING SINGLE SIGN-ON / INTEGRATED CREDENTIALS
        '--------------------------------------------------------------------------------------------------------------------------------------------
        Dim lstPerfCntrCPU As New List(Of PerformanceCounter)
        Dim lstCPULoad As New List(Of clsCPU)
        Dim strScope As String = "\\" & mstrCurrentTargetServer & "\root\cimv2"
        Dim objCPU As New ManagementObjectSearcher(strScope, "SELECT * FROM Win32_Processor")
        '-------------------------------------------------------------------------------------------------------------------------------------------- 
 
        Dim objMgmt As ManagementObject
        For Each objMgmt In objCPU.Get
            mintCPULogicalProcessors = IIf(IsNothing(objMgmt("numberoflogicalprocessors")), 0, Convert.ToInt32(objMgmt("numberoflogicalprocessors").ToString))
        Next

        Dim lstPerfCntrCPU As New List(Of PerformanceCounter)
        Dim lstCPULoad As New List(Of clsCPU)

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