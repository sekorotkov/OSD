' // ***************************************************************************
' // Sergey Korotkov
' // SCCM User Group Russia: https://www.facebook.com/groups/sccm.russia/
' //
' // File:      OSDPreStartTool.vbs
' // 
' // Purpose:   Addition function for WinPE based on MDT & etc
' // 
' // Usage:     Use ZTIGather.wsf & UserExit for call this functions
' // Sub Files: ZTIGather.wsf + CustomSettings.ini, ZTIGather.xml, ZTIUtility.vbs, ZTIDataAccess.vbs
' //       and: Unattend.xml - for start ZTIGather.wsf script
' //
' // Version:   0.1 - 2016-11-03
' //            Starting, base function
' // Version:   0.2 - 2017-02-01
' //            Added/Changed functions DateTimeSync, WarnPingHosts
' // Version:   0.3 - 2017-02-09
' //            Added/Changed functions WarnAndDiskpartClearDiskX (beta)
' // Version:   0.4 - 2017-11-11
' //            A few bug fixes & formatting
' // Version:   0.5 - 2018-08-14
' //            Bug fixes. Thanks "p g" for report.
' //            Upgrade MDT lib to 6.3.8450.1000.
' // ***************************************************************************


' Default UserExit MDT function
Function UserExit(sType, sWhen, sDetail, bSkip) 
    oLogging.CreateEntry "OSDPreStartTool: Entered UserExit ", LogTypeInfo
    UserExit = Success
End Function

' // ***************************************************************************

' Warning message if Hosts do not ping
Function WarnPingHosts(sTitle, sText, iTimeOut, aHosts, bAllPositive)
    oLogging.CreateEntry "OSDPreStartTool: Entered WarnPingHosts", LogTypeInfo
    oLogging.CreateEntry "OSDPreStartTool: Hosts = " & Join(aHosts), LogTypeInfo

    Dim bEcho, sHost, oPing, oPings, oPingStatuses
    bEcho = True

    Set oPingStatuses = CreateObject("Scripting.Dictionary")
    oPingStatuses.CompareMode = vbTextCompare
    With oPingStatuses ' https://msdn.microsoft.com/en-us/library/aa394350%28v=vs.85%29.aspx?f=255&MSPPError=-2147217396
        .Add Null, 	"Could not find host"
        .Add 0, 	"Success"
        .Add 11001, "Buffer Too Small"
        .Add 11002, "Destination Net Unreachable "
        .Add 11003, "Destination Host Unreachable"
        .Add 11004, "Destination Protocol Unreachable"
        .Add 11005, "Destination Port Unreachable"
        .Add 11006, "No Resources"
        .Add 11007, "Bad Option"
        .Add 11008, "Hardware Error"
        .Add 11009, "Packet Too Big"
        .Add 11010, "Request Timed Out"
        .Add 11011, "Bad Request"
        .Add 11012, "Bad Route"
        .Add 11013, "TimeToLive Expired Transit"
        .Add 11014, "TimeToLive Expired Reassembly"
        .Add 11015, "Parameter Problem"
        .Add 11016, "Source Quench"
        .Add 11017, "Option Too Big"
        .Add 11018, "Bad Destination"
        .Add 11032, "Negotiating IPSEC"
        .Add 11050, "General Failure"
    End With
    sText = sText & vbNewLine

    For Each sHost In aHosts
        Set oPings = objWMI.ExecQuery("select * from Win32_PingStatus where Address = '" & sHost & "'")
        For Each oPing In oPings
            If oPing.StatusCode = 0 Then
                ' Positive reply
                oLogging.CreateEntry "OSDPreStartTool: Positive reply from: " & sHost, LogTypeInfo
                If bAllPositive = False then
                    WarnPingHosts = Success
                    Exit Function
            End If 
            Else
                oLogging.CreateEntry "OSDPreStartTool: No reply from: " & sHost, LogTypeInfo
                bEcho = False
                sText = sText & vbNewLine & sHost & ": " & oPingStatuses(oPing.StatusCode)
            End If ' If oPing.StatusCode = 0 Then
        Next ' For Each oPing In oPings
    Next ' For Each sHost In aHosts

    If bEcho = False Then
        WarnPingHosts = Failure
        oShell.Popup sText, iTimeOut, sTitle, vbOKOnly + vbCritical
    Else
        WarnPingHosts = Success
    End If
End Function

' Later
' Orig tool: https://gallery.technet.microsoft.com/OSD-Pre-Flight-Checks-cbb635f5
Function ShowDriveType(drvpath)
    Dim fso, d, t
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set d = fso.GetDrive(drvpath)
    Select Case d.DriveType
        Case 0: t = "Unknown"
        Case 1: t = "Removable"
        Case 2: t = "Fixed"
        Case 3: t = "Network"
        Case 4: t = "CD-ROM"
        Case 5: t = "RAM Disk"
    End Select
    ShowDriveType = "Drive " & d.DriveLetter & ": - " & t
End Function


' Warning message if USB storage is attached as disk iDisk
Function WarnUSBAsDiskX(sTitle, sText, iTimeOut, iDisk)
    oLogging.CreateEntry "OSDPreStartTool: Entered WarnUSBAttached", LogTypeInfo

    If objWMI.ExecQuery("SELECT InterfaceType, Index FROM Win32_DiskDrive WHERE InterfaceType = 'USB' AND Index = " & iDisk).Count > 0 then
        WarnUSBAsDiskX = Failure
        oShell.Popup sText, iTimeOut, sTitle, vbOKOnly + vbCritical
    Else
        WarnUSBAsDiskX = Success
    End If
End Function


' Warning message if DiskX storage is not attached or InterfaceType not IDE/SCSI
Function WarnStorNotPresentAsDiskX(sTitle, sText, iTimeOut, iDisk)
    oLogging.CreateEntry "OSDPreStartTool: Entered WarnUSBAttached", LogTypeInfo

    If objWMI.ExecQuery("SELECT InterfaceType, Index FROM Win32_DiskDrive WHERE (InterfaceType = 'IDE' OR InterfaceType = 'SCSI') AND Index = " & iDisk).Count = 0 then
        WarnStorNotPresentAsDiskX = Failure
        oShell.Popup sText, iTimeOut, sTitle, vbOKOnly + vbCritical
    Else
        WarnStorNotPresentAsDiskX = Success
    End If
End Function


' Warning message if IPAddress00X is not present by ZTIGather.wsf (Hope on ZTIGather.wsf)
Function WarnIPNotPresent(sTitle, sText, iTimeOut)
    oLogging.CreateEntry "OSDPreStartTool: Entered WarnIPNotPresent", LogTypeInfo
    oLogging.CreateEntry "OSDPreStartTool: IPAddress001 = " & oEnvironment.Item("IPAddress001"), LogTypeInfo

    If oEnvironment.Item("IPAddress001") = "" and oEnvironment.Item("IPAddress002") = "" and oEnvironment.Item("IPAddress003") = "" then
        oLogging.CreateEntry "OSDPreStartTool: IP address does not present", LogTypeInfo
        oShell.Popup sText, iTimeOut, sTitle, vbOKOnly + vbCritical
        WarnIPNotPresent = Failure
    Else
        WarnIPNotPresent = Success
    End if
End Function


' Warning message if Laptop Battery is operated
Function WarnIsOnBattery(sTitle, sText, iTimeOut)
    oLogging.CreateEntry "OSDPreStartTool: Entered WarnIsOnBattery", LogTypeInfo
    oLogging.CreateEntry "OSDPreStartTool: Env IsOnBattery = " & oEnvironment.Item("IsOnBattery"), LogTypeInfo

    ' Later: Add Warn without paused script: https://gallery.technet.microsoft.com/scriptcenter/3579394b-ac53-4ba5-9357-ea8efd59646d
    If oEnvironment.Item("IsOnBattery") = "True"  then
        oLogging.CreateEntry "OSDPreStartTool: IsOnBattery = " & oEnvironment.Item("IsOnBattery") & " - warning!", LogTypeInfo
        oShell.Popup sText, iTimeOut, sTitle, vbOKOnly + vbExclamation
    End if

    WarnIsOnBattery = Success
End Function


' Simple functions for Sleep in secconds
Function SleepSeconds(Seconds)
    oLogging.CreateEntry "OSDPreStartTool: Entered StartTimeOut", LogTypeInfo
    oLogging.CreateEntry "OSDPreStartTool: StartTimeOut on " & Seconds & " seconds...", LogTypeInfo
    oUtility.SafeSleep Seconds*1000
    SleepSeconds = Success
End Function


' Set PowerCfg schema by GUID
Function SetPowerScheme(sGUID)
    oLogging.CreateEntry "OSDPreStartTool: Entered SetPowerScheme", LogTypeInfo

    Dim sCommand, sPowerCfgPath, iRc

    iRc = oUtility.FindFile("powercfg.exe", sPowerCfgPath)

    If iRc <> Success then
        oLogging.CreateEntry "OSDPreStartTool: Unable to locate powercfg.exe, skipping", LogTypeInfo
        ' May be add MsbBox with TimeOut
        SetPowerScheme = Failure
    Else
        sCommand = sPowerCfgPath & " /SETACTIVE " & sGUID
        oLogging.CreateEntry "OSDPreStartTool: powercfg.exe located. sPowerCfgPath = " & sPowerCfgPath, LogTypeInfo
        oLogging.CreateEntry "OSDPreStartTool: sCommand = " & sCommand, LogTypeInfo

        oUtility.RunWithHeartbeat(sCommand)
        SetPowerScheme = Success
    End if	
End Function


' Set Custom "Step Name" in MDT Monitoring console
Function SetCurrentActionName(sActionName)
    oLogging.CreateEntry "OSDPreStartTool: Entered SetCurrentActionName", LogTypeInfo
    oLogging.CreateEntry "OSDPreStartTool: ActionName = " & sActionName, LogTypeInfo

    oEnvironment.Item("_SMSTSCurrentActionName") = sActionName
    SetCurrentActionName = Success
End Function


' Sync Date/Time from another Domain server over Environment variables
Function DateTimeSync()
    oLogging.CreateEntry "OSDPreStartTool: Entered DateTimeSync", LogTypeInfo		
    oLogging.CreateEntry "OSDPreStartTool: Variable: PXETimeServers = " & Join(oEnvironment.ListItem("PXETimeServers").Keys,", "), LogTypeInfo
    oLogging.CreateEntry "OSDPreStartTool: Variable: PXETimeUser = " & oEnvironment.Item("PXETimeUser"), LogTypeInfo
    oLogging.CreateEntry "OSDPreStartTool: Variable: PXETimeDomain = " & oEnvironment.Item("PXETimeDomain"), LogTypeInfo
    oLogging.CreateEntry "OSDPreStartTool: Variable: PXETimePassword = " & oEnvironment.Item("PXETimePassword"), LogTypeInfo ' Will be filtered by MDT log processing
    oLogging.CreateEntry "OSDPreStartTool: Variable: PXETimeShareName = " & oEnvironment.Item("PXETimeShareName"), LogTypeInfo
    oLogging.CreateEntry "OSDPreStartTool: DateTimeSync Before Sync - Date/Time Now: " & Now(), LogTypeInfo

    Dim sPXETimeServer, cPXETimeServers, oNetwork, sRemoteName, sFullUsername, sCommand, iNetTimeResult
    DateTimeSync = Failure

    cPXETimeServers = oEnvironment.ListItem("PXETimeServers").Keys

    ' if Domain does not present - use local user
    If oEnvironment.Item("PXETimeDomain") <> "" Then
        sFullUsername = oEnvironment.Item("PXETimeDomain") & "\" & oEnvironment.Item("PXETimeUser")
    Else
        sFullUsername = oEnvironment.Item("PXETimeUser")
    End If
    oLogging.CreateEntry "OSDPreStartTool: Variable: sFullUsername = " & sFullUsername, LogTypeInfo

    For Each sPXETimeServer In cPXETimeServers
        oLogging.CreateEntry "OSDPreStartTool: Sync time from sPXETimeServer = " & sPXETimeServer, LogTypeInfo
        Set oNetwork = WScript.CreateObject("WScript.Network")

        sRemoteName = "\\" & sPXETimeServer & "\" & oEnvironment.Item("PXETimeShareName")
        oLogging.CreateEntry "OSDPreStartTool: Variable: sRemoteName = " & sRemoteName, LogTypeInfo

        On Error Resume Next
        Err.Clear
        iRetRes = oNetwork.MapNetworkDrive("", sRemoteName, False, sFullUsername, oEnvironment.Item("PXETimePassword"))
        If Err.Number <> 0 then
            oLogging.CreateEntry "OSDPreStartTool: MapNetworkDrive error for: " & sRemoteName & ": " & Trim(Err.Description) & " (" & Err.Number & ")", LogTypeError
        End if
        ' Err.Number 0           - Good - Mapped
        ' Err.Number -2147024829 - Good - "Network resource doesn't exist"
        ' Err.Number -2147024891 - Bad  - "Access is denied"
        ' Err.Number -2147024843 - Bad  - "The network path was not found. Server does not exist"
        
        sCommand = "net.exe time \\" & sPXETimeServer & " /SET /Y"
        oLogging.CreateEntry "OSDPreStartTool: DateTimeSync command: " & sCommand, LogTypeInfo

        iNetTimeResult = oUtility.RunWithHeartbeat(sCommand)
        oLogging.CreateEntry "OSDPreStartTool: oUtility.RunWithHeartbeat exit code = " & iNetTimeResult, LogTypeInfo
        
        ' Remove network drive
        oNetwork.RemoveNetworkDrive sTimeLocalLetter, True, True
        If iNetTimeResult = Success Then
            oLogging.CreateEntry "OSDPreStartTool: DateTimeSync Success", LogTypeInfo
            oLogging.CreateEntry "OSDPreStartTool: DateTimeSync After Sync - Date/Time Now: " & Now(), LogTypeInfo
            DateTimeSync = Success
            Exit For
        End If
    Next

    Set oNetwork = Nothing
End Function


' Original sub function you can find in script: ZTISCCM.wsf / LiteTouch.wsf
' This Function is modified for always restart RemoteRecovery.exe process if exists
Function EnableDaRT

    Dim tries
    Dim oInv
    Dim oTicketNode
    Dim oIPNode
    Dim dicPortList
    Dim dicIPList

    Dim oProcesses, oProcess

    ' Remote control is only supported in Windows PE (don't use OSVersion as it isn't set yet for refresh)
    If oEnv("SystemDrive") <> "X:" then
        Exit Function
    End if


    ' Kill RemoteRecovery.exe process if already exists

    If oFSO.FileExists(oEnv("SystemRoot") & "\System32\inv32.xml") then
        oLogging.CreateEntry "RemoteRecovery file (inv32.xml) is present [Path:" & oEnv("SystemRoot") & "\System32\inv32.xml" & "], will be removed", LogTypeInfo
        oFSO.DeleteFile oEnv("SystemRoot") & "\System32\inv32.xml", True
        Set oProcesses = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'RemoteRecovery.exe'")
        'iterate all item(s)
        For Each oProcess In oProcesses
            oLogging.CreateEntry "OSDPreStartTool: Process:" & oProcess.Name & ", will be terminated", LogTypeInfo
            oProcess.Terminate()
        Next
    End if


    ' Initialize

    Set dicIPList = CreateObject("Scripting.Dictionary")
    Set dicPortList = CreateObject("Scripting.Dictionary")


    ' Make sure the executable exists

    If not oFSO.FileExists(oEnv("SystemRoot") & "\System32\RemoteRecovery.exe") then
        Exit Function
    End if


    ' Start remote recovery process

    oShell.CurrentDirectory = oEnv("SystemRoot") & "\System32"
    oShell.Run oEnv("SystemRoot") & "\System32\RemoteRecovery.exe -nomessage", 2, false


    ' Sleep until we see the inv32.xml file

    tries = 0
    Do
        WScript.Sleep 1000
        tries = tries + 1
    Loop While not oFSO.FileExists(oEnv("SystemRoot") & "\System32\inv32.xml") and tries < 10

    If not oFSO.FileExists(oEnv("SystemRoot") & "\System32\inv32.xml") then
        oLogging.CreateEntry "Unable to find the inv32.xml file, DaRT remote control is not running.", LogTypeInfo
        Exit Function
    End if


    ' Read the XML file and put the values into variables

    On Error Resume Next

    Set oInv = oUtility.CreateXMLDOMObjectEx(oEnv("SystemRoot") & "\System32\inv32.xml")
    Set oTicketNode = oInv.SelectSingleNode("//A")
    oEnvironment.Item("DartTicket") = oTicketNode.Attributes.getNamedItem("ID").value

    ' First get the IPv4 entries (skipping locally-administered ones)
    For each oIPNode in oInv.SelectNodes("//L")
        If Instr(oIPNode.Attributes.getNamedItem("N").value, ":") = 0 and Left(oIPNode.Attributes.getNamedItem("N").value, 4) <> "169." then
            dicIPList.Add oIPNode.Attributes.getNamedItem("N").value, ""
            dicPortList.Add oIPNode.Attributes.getNamedItem("P").value, ""
        End if
    Next

    ' Then add the IPv6 entries
    For each oIPNode in oInv.SelectNodes("//L")
        If Instr(oIPNode.Attributes.getNamedItem("N").value, ":") > 0 then
            dicIPList.Add oIPNode.Attributes.getNamedItem("N").value, ""
            dicPortList.Add oIPNode.Attributes.getNamedItem("P").value, ""
        End if
    Next
    oEnvironment.ListItem("DartIP") = dicIPList
    oEnvironment.ListItem("DartPort") = dicPortList

End Function


'Later
' Original function you can find in script: LiteTouch.wsf
Function GetNetworkingErrorHint(DeployRoot)
    Dim DeployHost, i,j

    GetNetworkingErrorHint = ""

    On error resume next

    DeployHost = DeployRoot

    If instr(3,DeployRoot,"\",vbTextCompare) <> 0 then

        DeployHost = mid(DeployRoot,3,instr(3,DeployRoot,"\",vbTextCompare)-3)
        
    End if

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If objWMI.ExecQuery("select * from win32_pingstatus where Address = '" & DeployHost & "' and ResponseTime <> NULL").Count > 0 then

        ' Positive reply from DeployRoot Server
        GetNetworkingErrorHint = "Connection OK. Possible cause: invalid credentials."

    ElseIf objWMI.ExecQuery("select * from win32_NetworkAdapterconfiguration where DHCPServer <> NULL").Count > 0 then

        ' DHCP OK, must be a network routing error or Virtual PC Configuraiton problem
        GetNetworkingErrorHint = "Can not reach DeployRoot. Possible cause: Network routing error or Network Configuration error."

    ElseIf objWMI.ExecQuery("select * from win32_NetworkAdapter where Installed = true and adaptertypeid = 0").Count > 0 then

        ' There is a valid Networking device present, yet the DHCP address is bad.
        GetNetworkingErrorHint = "DHCP Lease was not obtained for any Networking device! Possible Cause: Check physical connection."

    Else

        For Each i in objWMI.InstancesOf("Win32_PnPEntity")

            For Each j in i.CompatibleID
                If ucase(right(j,8)) = "\CC_0200" then
                
                    GetNetworkingErrorHint = "The following networking device did not have a driver installed."
                    GetNetworkingErrorHint = GetNetworkingErrorHint & vbNewLine & i.HardwareID(0)
                    exit function
                    
                End if
                
            Next
            
        Next

        ' Are you kidding me? THis is the 21st century, what kind of computer doesn't have a networking adatper?
        GetNetworkingErrorHint = "No networking devices were found on this machine!"

    End if

    on error goto 0

End function

' Later
' Destructive function (diskpart.exe delete function)
' Warning message And DiskPart Clean if FS is Unknown (by DiskPart) on disk iDisk
Function WarnAndDiskpartClearDiskX(sTitle, sText, iTimeOut, iDisk)
    oLogging.CreateEntry "OSDPreStartTool: Entered WarnAndDiskpartClearDiskX", LogTypeInfo

    Dim cOsByGUIDs, cPartitionStatuses, oExec, sOutput
    Set cOsByGUIDs = CreateObject("Scripting.Dictionary")
    cOsByGUIDs.CompareMode = vbTextCompare
    With cOsByGUIDs
        ' https://en.wikipedia.org/wiki/Partition_type#List_of_partition_IDs
        .Add Null, 	Array("Empty",	"")
        .Add "0", 	Array("Empty", 	"")
        .Add "07", 	Array("Windows","")
        .Add "82", 	Array("Linux",	"")
        .Add "83", 	Array("Linux", 	"")
        ' https://en.wikipedia.org/wiki/GUID_Partition_Table#Partition_type_GUIDs
        .Add "00000000-0000-0000-0000-000000000000", Array("None",          "Unused entry")
        .Add "024DEE41-33E7-11D3-9D69-0008C781F39F", Array("None",          "MBR partition scheme")
        .Add "C12A7328-F81F-11D2-BA4B-00A0C93EC93B", Array("None",          "EFI System partition")
        .Add "21686148-6449-6E6F-744E-656564454649", Array("None",          "BIOS boot partition")
        .Add "D3BFE2DE-3DAF-11DF-BA40-E3A556D89593", Array("None",          "Intel Fast Flash (iFFS) partition (for Intel Rapid Start technology)")
        .Add "F4019732-066E-4E12-8273-346C5641494F", Array("None",          "Sony boot partition[f]")
        .Add "BFBFAFE7-A34F-448A-9A5B-6213EB736C22", Array("None",          "Lenovo boot partition[f]")
        .Add "E3C9E316-0B5C-4DB8-817D-F92DF00215AE", Array("Windows",       "Microsoft Reserved Partition (MSR)")
        .Add "EBD0A0A2-B9E5-4433-87C0-68B6B72699C7", Array("Windows",       "Basic data partition")
        .Add "5808C8AA-7E8F-42E0-85D2-E1E90434CFB3", Array("Windows",       "Logical Disk Manager (LDM) metadata partition")
        .Add "AF9B60A0-1431-4F62-BC68-3311714A69AD", Array("Windows",       "Logical Disk Manager data partition")
        .Add "DE94BBA4-06D1-4D40-A16A-BFD50179D6AC", Array("Windows",       "Windows Recovery Environment")
        .Add "37AFFC90-EF7D-4E96-91C3-2D7AE055B174", Array("Windows",       "IBM General Parallel File System (GPFS) partition")
        .Add "E75CAF8F-F680-4CEE-AFA3-B001E56EFC2D", Array("Windows",       "Storage Spaces partition")
        .Add "75894C1E-3AEB-11D3-B7C1-7B03A0000000", Array("HP-UX",         "Data partition")
        .Add "E2A1E728-32E3-11D6-A682-7B03A0000000", Array("HP-UX",         "Service Partition")
        .Add "0FC63DAF-8483-4772-8E79-3D69D8477DE4", Array("Linux",         "Linux filesystem data")
        .Add "A19D880F-05FC-4D3B-A006-743F0F84911E", Array("Linux",         "RAID partition")
        .Add "44479540-F297-41B2-9AF7-D131D5F0458A", Array("Linux",         "Root partition (x86)")
        .Add "4F68BCE3-E8CD-4DB1-96E7-FBCAF984B709", Array("Linux",         "Root partition (x86-64)")
        .Add "69DAD710-2CE4-4E3C-B16C-21A1D49ABED3", Array("Linux",         "Root partition (32-bit ARM)")
        .Add "B921B045-1DF0-41C3-AF44-4C6F280D3FAE", Array("Linux",         "Root partition (64-bit ARM/AArch64)")
        .Add "0657FD6D-A4AB-43C4-84E5-0933C84B4F4F", Array("Linux",         "Swap partition")
        .Add "E6D6D379-F507-44C2-A23C-238F2A3DF928", Array("Linux",         "Logical Volume Manager (LVM) partition")
        .Add "933AC7E1-2EB4-4F13-B844-0E14E2AEF915", Array("Linux",         "/home partition")
        .Add "3B8F8425-20E0-4F3B-907F-1A25A76F98E8", Array("Linux",         "/srv (server data) partition")
        .Add "8DA63339-0007-60C0-C436-083AC8230908", Array("Linux",         "Reserved")
        .Add "83BD6B9D-7F41-11DC-BE0B-001560B84F0F", Array("FreeBSD",       "Boot partition")
        .Add "516E7CB4-6ECF-11D6-8FF8-00022D09712B", Array("FreeBSD",       "Data partition")
        .Add "516E7CB5-6ECF-11D6-8FF8-00022D09712B", Array("FreeBSD",       "Swap partition")
        .Add "516E7CB6-6ECF-11D6-8FF8-00022D09712B", Array("FreeBSD",       "Unix File System (UFS) partition")
        .Add "516E7CB8-6ECF-11D6-8FF8-00022D09712B", Array("FreeBSD",       "Vinum volume manager partition")
        .Add "516E7CBA-6ECF-11D6-8FF8-00022D09712B", Array("FreeBSD",       "ZFS partition")
        .Add "48465300-0000-11AA-AA11-00306543ECAC", Array("OS X",          "Hierarchical File System Plus (HFS+) partition")
        .Add "55465300-0000-11AA-AA11-00306543ECAC", Array("Darwin",        "Apple UFS")
        .Add "52414944-0000-11AA-AA11-00306543ECAC", Array("Darwin",        "Apple RAID partition")
        .Add "52414944-5F4F-11AA-AA11-00306543ECAC", Array("Darwin",        "Apple RAID partition, offline")
        .Add "426F6F74-0000-11AA-AA11-00306543ECAC", Array("Darwin",        "Apple Boot partition (Recovery HD)")
        .Add "4C616265-6C00-11AA-AA11-00306543ECAC", Array("Darwin",        "Apple Label")
        .Add "5265636F-7665-11AA-AA11-00306543ECAC", Array("Darwin",        "Apple TV Recovery partition")
        .Add "53746F72-6167-11AA-AA11-00306543ECAC", Array("Darwin",        "Apple Core Storage (i.e. Lion FileVault) partition")
        .Add "B6FA30DA-92D2-4A9A-96F1-871EC6486200", Array("Darwin",        "SoftRAID_Status")
        .Add "2E313465-19B9-463F-8126-8A7993773801", Array("Darwin",        "SoftRAID_Scratch")
        .Add "FA709C7E-65B1-4593-BFD5-E71D61DE9B02", Array("Darwin",        "SoftRAID_Volume")
        .Add "BBBA6DF5-F46F-4A89-8F59-8765B2727503", Array("Darwin",        "SoftRAID_Cache")
        .Add "6A82CB45-1DD2-11B2-99A6-080020736631", Array("Solaris",       "Boot partition")
        .Add "6A85CF4D-1DD2-11B2-99A6-080020736631", Array("Solaris",       "Root partition")
        .Add "6A87C46F-1DD2-11B2-99A6-080020736631", Array("Solaris",       "Swap partition")
        .Add "6A8B642B-1DD2-11B2-99A6-080020736631", Array("Solaris",       "Backup partition")
        .Add "6A898CC3-1DD2-11B2-99A6-080020736631", Array("Solaris",       "/usr partition")
        .Add "6A8EF2E9-1DD2-11B2-99A6-080020736631", Array("Solaris",       "/var partition")
        .Add "6A90BA39-1DD2-11B2-99A6-080020736631", Array("Solaris",       "/home partition")
        .Add "6A9283A5-1DD2-11B2-99A6-080020736631", Array("Solaris",       "Alternate sector")
        .Add "6A945A3B-1DD2-11B2-99A6-080020736631", Array("Solaris",       "Reserved partition")
        .Add "6A9630D1-1DD2-11B2-99A6-080020736631", Array("Solaris",       "Reserved partition")
        .Add "6A980767-1DD2-11B2-99A6-080020736631", Array("Solaris",       "Reserved partition")
        .Add "6A96237F-1DD2-11B2-99A6-080020736631", Array("Solaris",       "Reserved partition")
        .Add "6A8D2AC7-1DD2-11B2-99A6-080020736631", Array("Solaris",       "Reserved partition")
        .Add "49F48D32-B10E-11DC-B99B-0019D1879648", Array("NetBSD",        "Swap partition")
        .Add "49F48D5A-B10E-11DC-B99B-0019D1879648", Array("NetBSD",        "FFS partition")
        .Add "49F48D82-B10E-11DC-B99B-0019D1879648", Array("NetBSD",        "LFS partition")
        .Add "49F48DAA-B10E-11DC-B99B-0019D1879648", Array("NetBSD",        "RAID partition")
        .Add "2DB519C4-B10F-11DC-B99B-0019D1879648", Array("NetBSD",        "Concatenated partition")
        .Add "2DB519EC-B10F-11DC-B99B-0019D1879648", Array("NetBSD",        "Encrypted partition")
        .Add "FE3A2A5D-4F32-41A7-B725-ACCC3285A309", Array("ChromeOS",      "ChromeOS kernel")
        .Add "3CB8E202-3B7E-47DD-8A3C-7FF2A13CFCEC", Array("ChromeOS",      "ChromeOS rootfs")
        .Add "2E0A753D-9E48-43B0-8337-B15192CB1B5E", Array("ChromeOS",      "ChromeOS future use")
        .Add "42465331-3BA3-10F1-802A-4861696B7521", Array("Haiku",         "Haiku BFS")
        .Add "85D5E45E-237C-11E1-B4B3-E89A8F7FC3A7", Array("MidnightBSD",   "Boot partition")
        .Add "85D5E45A-237C-11E1-B4B3-E89A8F7FC3A7", Array("MidnightBSD",   "Data partition")
        .Add "85D5E45B-237C-11E1-B4B3-E89A8F7FC3A7", Array("MidnightBSD",   "Swap partition")
        .Add "0394EF8B-237E-11E1-B4B3-E89A8F7FC3A7", Array("MidnightBSD",   "Unix File System (UFS) partition")
        .Add "85D5E45C-237C-11E1-B4B3-E89A8F7FC3A7", Array("MidnightBSD",   "Vinum volume manager partition")
        .Add "85D5E45D-237C-11E1-B4B3-E89A8F7FC3A7", Array("MidnightBSD",   "ZFS partition")
        .Add "45B0969E-9B03-4F30-B4C6-B4B80CEFF106", Array("Ceph",          "Ceph Journal")
        .Add "45B0969E-9B03-4F30-B4C6-5EC00CEFF106", Array("Ceph",          "Ceph dm-crypt Encrypted Journal")
        .Add "4FBD7E29-9D25-41B8-AFD0-062C0CEFF05D", Array("Ceph",          "Ceph OSD")
        .Add "4FBD7E29-9D25-41B8-AFD0-5EC00CEFF05D", Array("Ceph",          "Ceph dm-crypt OSD")
        .Add "89C57F98-2FE5-4DC0-89C1-F3AD0CEFF2BE", Array("Ceph",          "Ceph disk in creation")
        .Add "89C57F98-2FE5-4DC0-89C1-5EC00CEFF2BE", Array("Ceph",          "Ceph dm-crypt disk in creation")
        .Add "824CC7A0-36A8-11E3-890A-952519AD3F61", Array("OpenBSD",       "Data partition")
        .Add "CEF5A9AD-73BC-4601-89F3-CDEEEEE321A1", Array("QNX",           "Power-safe (QNX6) file system")
        .Add "C91818F9-8025-47AF-89D2-F030D7000C2C", Array("Plan 9",        "Plan 9 partition")
        .Add "9D275380-40AD-11DB-BF97-000C2911D1B8", Array("VMware ESX",    "vmkcore (coredump partition)")
        .Add "AA31E02A-400F-11DB-9590-000C2911D1B8", Array("VMware ESX",    "VMFS filesystem partition")
        .Add "9198EFFC-31C0-11DB-8F78-000C2911D1B8", Array("VMware ESX",    "VMware Reserved")
        .Add "2568845D-2332-4675-BC39-8FA5A4748D15", Array("Android-IA",    "Bootloader")
        .Add "114EAFFE-1552-4022-B26E-9B053604CF84", Array("Android-IA",    "Bootloader2")
        .Add "49A4D17F-93A3-45C1-A0DE-F50B2EBE2599", Array("Android-IA",    "Boot")
        .Add "4177C722-9E92-4AAB-8644-43502BFD5506", Array("Android-IA",    "Recovery")
        .Add "EF32A33B-A409-486C-9141-9FFB711F6266", Array("Android-IA",    "Misc")
        .Add "20AC26BE-20B7-11E3-84C5-6CFDB94711E9", Array("Android-IA",    "Metadata")
        .Add "38F428E6-D326-425D-9140-6E0EA133647C", Array("Android-IA",    "System")
        .Add "A893EF21-E428-470A-9E55-0668FD91A2D9", Array("Android-IA",    "Cache")
        .Add "DC76DDA9-5AC1-491C-AF42-A82591580C0D", Array("Android-IA",    "Data")
        .Add "EBC597D0-2053-4B15-8B64-E0AAC75F4DB1", Array("Android-IA",    "Persistent")
        .Add "8F68CC74-C5E5-48DA-BE91-A0C8C15E9C80", Array("Android-IA",    "Factory")
        .Add "767941D0-2085-11E3-AD3B-6CFDB94711E9", Array("Android-IA",    "Fastboot / Tertiary")
        .Add "AC6D7924-EB71-4DF8-B48D-E267B27148FF", Array("Android-IA",    "OEM")
        .Add "7412F7D5-A156-4B13-81DC-867174929325", Array("ONIE",          "Boot")
        .Add "D4E6E2CD-4469-46F3-B5CB-1BFF57AFC149", Array("ONIE",          "Config")
        .Add "9E1A2D38-C612-4316-AA26-8B49521E5A8B", Array("PowerPC",       "PReP boot")
        .Add "BC13C2FF-59E6-4262-A352-B275FD6F7172", Array("Freedesktop",   "Shared boot loader configuration")
    End With

    Dim oPartition, oPartitions

    If objWMI.ExecQuery("SELECT Index FROM Win32_DiskPartition WHERE DiskIndex = " & iDisk).Count < 1 Then
        oLogging.CreateEntry "OSDPreStartTool: Partitions < 1", LogTypeInfo
        'Partition not exists
        oLogging.CreateEntry "OSDPreStartTool: Partition not exists on Disk" & iDisk & ", exit... ", LogTypeInfo
        WarnAndDiskpartClearDiskX = Success
        Exit Function
    End If

    ' Partition exists
    Dim oRegEx, oMatch, oMatches, iPartIndex, oPartitionStatus
    Set cPartitionStatuses = CreateObject("Scripting.Dictionary")

    Set oExec = oShell.Exec("diskpart.exe")

    ' Rescan the disks
    sOutput = ExecuteDiskPartCommand(oExec, "rescan")
    ' Select the disk
    sOutput = ExecuteDiskPartCommand(oExec, "select disk " & iDisk)
    ' Get all Partitions
    sOutput = ExecuteDiskPartCommand(oExec, "list partition")

    ' Parce the number of partitions
    Set oRegEx = CreateObject("VBScript.RegExp")
    oRegEx.Global = True
    oRegEx.Pattern = "  Partition (\d{1})    \S{5,}"

    Set oMatches = oRegEx.Execute(sOutput)
    oLogging.CreateEntry "OSDPreStartTool: Partition exists count = " & oMatches.Count, LogTypeInfo
    For Each oMatch In oMatches
        iPartIndex = oMatch.SubMatches(0)
        oLogging.CreateEntry vbNewLine & "OSDPreStartTool: Part Index = " & iPartIndex, LogTypeInfo
        sOutput = ExecuteDiskPartCommand(oExec, "select part " & iPartIndex)
        
        sOutput = ExecuteDiskPartCommand(oExec, "detail partition")
        
        ' Parce the Type of partition
        Dim aTemp, sGUID
        
        aTemp =  Split(sOutput, vbNewLine)
        sGUID = Trim((Split(aTemp(1), ":"))(1))
        oLogging.CreateEntry "OSDPreStartTool: Diskpart Type = " & sGUID, LogTypeInfo
        cPartitionStatuses.Add iPartIndex, sGUID
        oLogging.CreateEntry "OSDPreStartTool: OS Type by GUID = " & Join(cOsByGUIDs(sGUID), " | "), LogTypeInfo
    Next

'''''
    ' Now "cPartitionStatuses" dictionary contain all numbers and partitions GUIDs
    ' Also You may extend cOsByGUIDs Dictionary for right logic. For test Type of partition execute diskpart.exe, select disk 0, select partition N, details partition, Line "Type": GUID (for GPT) or ID for MBR.
    ' Example 1: if partition type does not equal Windows - delete part
'    For Each oPartitionStatus in cPartitionStatuses
'        oLogging.CreateEntry "OSDPreStartTool: Disk | Partition | Type | OS  = " & Join(Array(iDisk, oPartitionStatus, cPartitionStatuses(oPartitionStatus), cOsByGUIDs(cPartitionStatuses(oPartitionStatus))(0)), " | "), LogTypeInfo
'        if cOsByGUIDs(cPartitionStatuses(oPartitionStatus))(0) <> "Windows" Then
'            sOutput = ExecuteDiskPartCommand(oExec, "select partition " & oPartitionStatus)
'            sOutput = ExecuteDiskPartCommand(oExec, "delete partition override")
'        End If
'    Next
'
'    ' Example 2: if any partition type does not equal Windows - delete disk
'    Dim iUserAnswer
'    For Each oPartitionStatus in cPartitionStatuses
'    Do ' Fake loop for workaround lost-Continue. Original: https://snippets.webaware.com.au/snippets/vbscript-for-next-and-continue/	
'        oLogging.CreateEntry "OSDPreStartTool: Disk | Partition | Type | OS  = " & Join(Array(iDisk, oPartitionStatus, cPartitionStatuses(oPartitionStatus), cOsByGUIDs(cPartitionStatuses(oPartitionStatus))(0)), " | "), LogTypeInfo
'        if cOsByGUIDs(cPartitionStatuses(oPartitionStatus))(0) <> "Windows" Then
'            iUserAnswer = oShell.Popup(sText, iTimeOut, sTitle, vbYesNo + vbCritical)
'
'            Select Case iUserAnswer
'                Case vbNo
'                    oLogging.CreateEntry "Answer = No.", LogTypeInfo
'                    Exit Do
'                Case vbYes
'                    oLogging.CreateEntry "Answer = Yes. Delete disk override", LogTypeInfo
'                    sOutput = ExecuteDiskPartCommand(oExec, "select disk " & iDisk)
'                    sOutput = ExecuteDiskPartCommand(oExec, "delete disk override")		   
'                Case -1
'                    oLogging.CreateEntry "wtf???", LogTypeInfo
'            End Select
'
'        End If
'    Loop While False ' Fake loop for workaround lost-Continue
'    Next 'For Each oPartitionStatus in cPartitionStatuses
'
'''''

    ' Exit diskpart.exe
    oExec.StdIn.Write "exit" & vbNewLine
    WarnAndDiskpartClearDiskX = Success
End Function ' Main
' Diskpart routine subfunction
Function ExecuteDiskPartCommand(oExec, strCommand)
    ' Original: https://blogs.msdn.microsoft.com/alejacma/2011/04/26/how-to-automate-a-command-line-utility-like-diskpart-vbscript/

    Dim IgnoreThis
    ' Run the command we want
    oExec.StdIn.Write strCommand & vbNewLine
            
    ' If we read the output now, we will get the one from previous command (?). As we will always
    ' run a dummy command after every valid command, we can safely ignore this
    Do While True
        IgnoreThis = oExec.StdOut.ReadLine & vbNewLine              

        ' Command finishes when diskpart prompt is shown again
        If InStr(IgnoreThis, "DISKPART>") <> 0 Then Exit Do
    Loop

    ' Run a dummy command, so the next time we call this function and try to read output,
    ' we can safely ignore the result
    oExec.StdIn.Write vbNewLine
            
    ' Read command's output
    ExecuteDiskPartCommand = ""    
    Do While True
        ExecuteDiskPartCommand = ExecuteDiskPartCommand & oExec.StdOut.ReadLine & vbNewLine              

        ' Command finishes when diskpart prompt is shown again
        If InStr(ExecuteDiskPartCommand, "DISKPART>") <> 0 Then Exit Do
    Loop

End Function
