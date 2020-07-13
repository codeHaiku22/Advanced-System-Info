Module CommonCode

    Public Function CalculateBestByteSize(ByVal lngBytes As Long, Optional blnMetric As Boolean = False) As String

        Try

            Dim strBestByteSize As String

            If blnMetric Then
                Select Case lngBytes
                    Case < 1000
                        strBestByteSize = lngBytes.ToString & "B"
                    Case < 1000000
                        strBestByteSize = Format(lngBytes / 1000, "#,###,###") & "KB"
                    Case < 1000000000
                        strBestByteSize = Format(lngBytes / 1000000, "#,###,###") & "MB"
                    Case < 1000000000000
                        strBestByteSize = Format(lngBytes / 1000000000, "#,###,###") & "GB"
                    Case < 1000000000000000
                        strBestByteSize = Format(lngBytes / 1000000000000, "#,###,###") & "TB"
                    Case < 1000000000000000000
                        strBestByteSize = Format(lngBytes / 1000000000000000, "#,###,###") & "PB"
                    Case Else
                        strBestByteSize = ""
                End Select
            Else
                Select Case lngBytes
                    Case < 1024
                        strBestByteSize = lngBytes.ToString & "B"
                    Case < 1048576
                        strBestByteSize = Format(lngBytes / 1024, "#,###,###") & "KB"
                    Case < 1073741824
                        strBestByteSize = Format(lngBytes / 1048576, "#,###,###") & "MB"
                    Case < 1099511627776
                        strBestByteSize = Format(lngBytes / 1073741824, "#,###,###") & "GB"
                    Case < 1125899906842624
                        strBestByteSize = Format(lngBytes / 1099511627776, "#,###,###") & "TB"
                    Case < 1152921504606846976
                        strBestByteSize = Format(lngBytes / 1125899906842624, "#,###,###") & "PB"
                    Case Else
                        strBestByteSize = ""
                End Select
            End If

            CalculateBestByteSize = strBestByteSize

        Catch ex As Exception

            MsgBox("Error calculating byte size conversion." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try


    End Function
    Public Function CalculateSystemUpTime(ByVal strLastBootUpTime As String, Optional blnFriendlyUpTime As Boolean = True) As String

        Try

            Dim dtmLastBootUpTime As DateTime = Mid(strLastBootUpTime, 5, 2) & "/" & Mid(strLastBootUpTime, 7, 2) & "/" & Strings.Left(strLastBootUpTime, 4) & " " & Mid(strLastBootUpTime, 9, 2) & ":" & Mid(strLastBootUpTime, 11, 2) & ":" & Mid(strLastBootUpTime, 13, 2)
            Dim dtmCurrent As DateTime = DateTime.Now
            Dim tsSystemUpTime As TimeSpan = dtmCurrent - dtmLastBootUpTime
            Dim strStystemUpTime As String

            If blnFriendlyUpTime Then
                Dim strDays As String = tsSystemUpTime.Days.ToString & " days "
                Dim strHours As String = tsSystemUpTime.Hours.ToString & " hours "
                Dim strMinutes As String = tsSystemUpTime.Minutes.ToString & " minutes "
                Dim strSeconds As String = tsSystemUpTime.Seconds.ToString & " seconds"
                strStystemUpTime = strDays & strHours & strMinutes & strSeconds
            Else
                strStystemUpTime = tsSystemUpTime.ToString
            End If

            CalculateSystemUpTime = strStystemUpTime

        Catch ex As Exception

            MsgBox("Error calculating system up time." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Function
    Public Function Get_HostName_Or_IPAddress(ByVal strHostNameorIPAddress As String, Optional blnIncludeDomain As Boolean = False) As String

        Try

            Dim blnIsIPAddress As Boolean = IsNumeric(Replace(strHostNameorIPAddress, ".", ""))
            Dim strHostName As String
            Dim strIPAddress As String

            If blnIsIPAddress Then
                strHostName = System.Net.Dns.GetHostEntry(strHostNameorIPAddress).HostName.ToString
                If blnIncludeDomain Then
                    Get_HostName_Or_IPAddress = strHostName
                Else
                    Get_HostName_Or_IPAddress = strHostName.Substring(0, strHostName.IndexOf("."))
                End If
            Else
                For Each ipAddress As System.Net.IPAddress In System.Net.Dns.GetHostAddresses(strHostNameorIPAddress)
                    If ipAddress.AddressFamily.ToString = "InterNetwork" Then
                        strIPAddress = strIPAddress & "/" & ipAddress.ToString
                    End If
                Next
                If strIPAddress.Length > 0 Then strIPAddress = Replace(strIPAddress, "/", "", 1, 1)
                Get_HostName_Or_IPAddress = strIPAddress
            End If

        Catch ex As Exception

            MsgBox("Error retrieving host name and/or IP address." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Function
    Public Function Get_TotalCPUs(ByVal strCurrentTargetServer As String) As Integer

        Dim i As Integer

        Try

            Dim perfCounter As PerformanceCounter
            Dim lngRawValue As Long = 0

            Do While IsNumeric(lngRawValue)
                perfCounter = New PerformanceCounter("Processor", "% Processor Time", i.ToString, strCurrentTargetServer)
                lngRawValue = perfCounter.RawValue
                i = i + 1
            Loop

        Catch ex As Exception

        Finally

            Get_TotalCPUs = i

        End Try

    End Function

    Public Function MakeFriendlyDateString(ByVal strSerializedDateTime As String) As String

        Try

            Dim strDate As String = ""
            Dim strTime As String = ""
            Dim strFriendlyDate As String = ""

            strSerializedDateTime = Trim(strSerializedDateTime)

            strDate = Mid(strSerializedDateTime, 5, 2) & "/" & Mid(strSerializedDateTime, 7, 2) & "/" & Strings.Left(strSerializedDateTime, 4)
            strFriendlyDate = Format(CDate(strDate), "M/d/yyyy")

            If Len(strSerializedDateTime) > 8 Then
                strTime = Mid(strSerializedDateTime, 9, 2) & ":" & Mid(strSerializedDateTime, 11, 2) & ":" & Mid(strSerializedDateTime, 13, 2)
            End If

            MakeFriendlyDateString = Trim(strFriendlyDate & " " & strTime)

        Catch ex As Exception

            MsgBox("Error creating date string." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Function
    Public Function ProgressBarForeColor(ByVal intLoad As Integer, Optional ByVal blnChildProgressBar As Boolean = False) As SolidColorBrush

        Try

            Select Case intLoad
                Case < 90
                    ProgressBarForeColor = IIf(blnChildProgressBar, Brushes.LightGreen, Brushes.Green)
                Case < 95
                    ProgressBarForeColor = IIf(blnChildProgressBar, Brushes.Orange, Brushes.DarkOrange)
                Case <= 100
                    ProgressBarForeColor = IIf(blnChildProgressBar, Brushes.Tomato, Brushes.Red)
            End Select

        Catch ex As Exception

            MsgBox("Error generating progress bar color." & vbCrLf & ex.Message, vbCritical, "AdvSysInfo")

        End Try

    End Function

End Module
