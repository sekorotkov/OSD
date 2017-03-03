# OSDPreStartTool

## This mini-Toolkit Check several prerequisites before starting a Task Sequence.


Russian Blog Post: https://skorotkov.wordpress.com/2017/02/10/osdprestarttool-ztigather-wsf-userexit-script/  
English Blog Post & description will be later, sorry.  


Current Functions:  
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;  
;;; Sleep. Reason http://deploymentresearch.com/Research/Post/528/Fixing-the-ldquo-Failed-to-find-a-valid-network-adapter-rdquo-error-in-ConfigMgr-Current-Branch  
;; Simple functions for Sleep in secconds  
SleepSeconds = #SleepSeconds(10)#  
;; Or MDT Build In function  
SleepMSeconds=#oUtility.SafeSleep(10*1000)#  
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;  
;;; Sync date/time from windows servers  
PXEDateTimeSync=#DateTimeSync()#  
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;  
;;; Set Power Scheme. Background and next steps please read: https://blogs.technet.microsoft.com/deploymentguys/2015/03/26/reducing-windows-deployment-time-using-power-management/  
SetPowerScheme=#SetPowerScheme("8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c")#  
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;  
;;; Set DMT Monitoring "Step Name" for visual allocate PXE-Booted device. Example: https://skorotkov.files.wordpress.com/2017/02/clip_image003.png  
SetCurrentActionName=#SetCurrentActionName("00: I'm boot to PXE")#  
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;  
;;; Run DART Remote Connection  
EnableDaRT=#EnableDaRT#  
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;  
;;; Warninig message if Laptop without Charger  
;; WarnIsOnBattery("Title Text", "Message Text", "Timeout in seconds") ; If Timeout = 0 script doesn't continue before "Ok"-button press  
WarnIsOnBattery=#WarnIsOnBattery("Warning! Charger doesn't connect", "Please CONNECT CHARGER!" & vbNewLine & "(AutoContinue after 30 sec.)", 30)#  
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;  
;;; Warninig message if USB disk is attached as Disk 0/1 ...  
;; WarnUSBAsDiskX("Title Text", "Message Text", "Wait TimeOut (Sec)", "Disk Number")  
WarnUSBAsDisk0=#WarnUSBAsDiskX("Error! USB device", "USB devices Attached as Disk 0!" & vbNewLine & "Please Unplug" & vbNewLine & "(Press Ok for continue.)", 0, 0)#  
;WarnUSBAsDisk1=#WarnUSBAsDiskX("Warning! USB device", "USB devices Attached as Disk 1!" & vbNewLine & "Please Check or Unplug" & vbNewLine & "(AutoContinue after 15 sec.)", 15, 1)#  
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;  
;;; Warning message if DiskX storage not attached or InterfaceType not IDE/SCSI  
;; WarnStorNotPresentAsDiskX("Title Text", "Message Text", "Wait TimeOut (Sec)", "Disk Number")  
WarnStorNotPresentAsDisk0=#WarnStorNotPresentAsDiskX("Error! Storage device", "Storage device not Attached as Disk 0!" & vbNewLine & "Please check HDD is Attached!" & vbNewLine & "(Press Ok for continue.)", 0, 0)#  
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;  
;;; Warninig message if IP addresses does not present (Hope on ZTIGather.wsf)  
;; WarnIPNotPresent("Title Text", "Message Text", "Wait TimeOut (Sec)")  
;WarnIPNotPresent=#WarnIPNotPresent("Warning! IP addresses does not present", "Please Check driver for network!" & vbNewLine & "(or press OK for skip)", 0)#  
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;  
;;; ping hosts  
;; WarnPingHosts("Title Text", "Message Text", "Wait TimeOut (Sec)", Array("Host1","Host2",...,"HostXX"), "All hosts or Any host return echo. True/False")  
WarnPingHosts=#WarnPingHosts("Warning! No Echo", "Please Check network OR hosts available!" & vbNewLine & "(or press OK for skip)", 0, Array("lab-sccm-r2.lab.local","lab.local","10.0.0.5"), False)#  
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;  


etc...
