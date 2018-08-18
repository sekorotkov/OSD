# OSDPreStartTool

## This mini-Toolkit Check several prerequisites before starting a Task Sequence.

Russian Blog Post: https://skorotkov.wordpress.com/2017/02/10/osdprestarttool-ztigather-wsf-userexit-script/  
English Blog Post & description will be later, may be, sorry. Use Google Translate

### Current Functions

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;;; Sleep. Reason

http://deploymentresearch.com/Research/Post/528/Fixing-the-ldquo-Failed-to-find-a-valid-network-adapter-rdquo-error-in-ConfigMgr-Current-Branch

;; Simple functions for Sleep in secconds

;SleepSeconds = SleepSeconds=#SleepSeconds(10)#

;; MDT Build In function

;SleepMSeconds=#oUtility.SafeSleep(5*1000)#

;; No Need if using WarnValidateNetworkConnectivity

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;Windows Server for auth & Net time - command. Use One-to-One addresses i.e. One address-string must resolve to one Computer-ip.

PXETimeServers001=10.0.0.11

PXETimeServers002=lab-sccm-r2.lab.local

PXETimeServers003=lab-dc.lab.local

; Domain Or local user. PLEASE!!! restrict logon user by batch / service / logon locally / terminal services: https://technet.microsoft.com/en-us/library/bb457125.aspx

PXETimeUser=pxe-time-sync

PXETimePassword=SuperP@$$w0rd

; Domain name or comment for use local user on remote PXETimeServer

PXETimeDomain=lab.local

; Really, You don't need real share and real drive-mapping, (used for authentification only).

PXETimeShareName=FakeShare

;

PXEDateTimeSync=#DateTimeSync()#

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;;; Set Power Scheme

;;Command for check available schemas: powercfg /l

;;Background and next steps please read: https://blogs.technet.microsoft.com/deploymentguys/2015/03/26/reducing-windows-deployment-time-using-power-management/

SetPowerScheme=#SetPowerScheme("8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c")#

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;;; Set MDT Monitoring "Step Name" for visual allocate PXE-Booted device

SetCurrentActionName=#SetCurrentActionName("00: Booted into WinPE")#

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;;; Run DART Remote Connection

EnableDaRT=#EnableDaRT#

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;;; Warninig message if Laptop without Charger

;; WarnIsOnBattery(<Title Text>, <Message Text>, <Timeout in seconds>) ; If Timeout = 0 script doesn't continue before "Ok"-button press

WarnIsOnBattery=#WarnIsOnBattery("Warning! Charger not connected", "Please CONNECT CHARGER!" & vbNewLine & "(Auto-Continued after 30 sec.)", 30)#

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;;; Warninig message if USB disk is attached as Disk 0/1 ...

;; WarnUSBAsDiskX(<Title Text>, <Message Text>, <Wait TimeOut (Sec)>, <Disk Number>)

WarnUSBAsDisk0=#WarnUSBAsDiskX("Error! USB device", "USB devices Attached as Disk 0!" & vbNewLine & "Please Unplug" & vbNewLine & "(Press OK to continue.)", 0, 0)#

;WarnUSBAsDisk1=#WarnUSBAsDiskX("Warning! USB device", "USB devices Attached as Disk 1!" & vbNewLine & "Please Check or Unplug" & vbNewLine & "(AutoContinue after 15 sec.)", 15, 1)#

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;;; Warning message if DiskX storage not attached or InterfaceType not IDE/SCSI

;; WarnStorNotPresentAsDiskX(<Title Text>, <Message Text>, <Wait TimeOut (Sec)>, <Disk Number>)

WarnStorNotPresentAsDisk0=#WarnStorNotPresentAsDiskX("Error! Storage device", "Storage device not Attached as Disk 0!" & vbNewLine & "Please check HDD is Attached!" & vbNewLine & "(Press Ok to 
continue.)", 0, 0)#

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;;; Warninig message if IP addresses does not present (Hope on ZTIGather.wsf)

;; WarnIPNotPresent(<Title Text>, <Message Text>, <Wait TimeOut (Sec)>)

;WarnIPNotPresent=#WarnIPNotPresent("Warning! IP addresses does not present", "Please Check driver to network!" & vbNewLine & "(or press OK to skip)", 0)#

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;;; Ping hosts

;; WarnPingHosts(<Title Text>, <Message Text>, <Wait TimeOut (Sec)>, Array("Host1","Host2",...,"HostXX"), <All hosts must echo. True/False>)

WarnPingHosts=#WarnPingHosts("Warning! No Echo", "Please Check network OR hosts available!" & vbNewLine & "(or press OK to skip)", 0, Array("lab-sccm-r2.lab.local","lab.local","10.0.0.5"), False)#

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;;; Warning message if network subsystem do not work. Retry iRetries times with iSeconds seconds between

;; WarnValidateNetworkConnectivity(<Title Text>, <Message Text>, <Wait TimeOut (Sec)>, <Sleep on seconds>, <Retry count>)

WarnValidateNetworkConnectivity=#WarnValidateNetworkConnectivity("Warning! Network subsystem doesn't work", "Please check network subsystem (boot drivers, DHCP, etc...)!" & vbNewLine & "(press OK to 
skip)", 0, 3, 5)#

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

etc...
