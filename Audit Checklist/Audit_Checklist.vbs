strComputer = "10.232.12.38"
Const HKLM = &H80000002
Const HKCU = &H80000001
Const HKU = &H80000003
Dim xl : Set xl = WScript.CreateObject("Excel.Application")
Dim objXL,objXL1,objXL2,objXL3,objXL5
intIndex = 1
intIndexTwo = 1
intIndexRams = 1
intIndexSoftware = 1

Set conn = CreateObject("ADODB.Connection")
strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\ctsingmrpss\Manualpush\rams.mdb;Persist Security Info=False"
conn.Open strConnect


CreateExcelFile()
'---------------------------------       Creating Excel      -------------------------------------------
Sub CreateExcelFile()
                xl.Visible = TRUE
				xl.WorkBooks.Add
                set objXL = xl.ActiveWorkbook.Worksheets(1)
				objXL.Name = "Audit CheckList"
				set objXL3 = xl.ActiveWorkbook.Worksheets(2)
				objXL3.Name = "Software List"
                set objXL2 = xl.ActiveWorkbook.Worksheets(3)
				objXL2.Name = "RAMS"
                
                objXL.Columns(1).ColumnWidth = 19
                objXL.Columns(2).ColumnWidth = 50
                objXL.Columns(3).ColumnWidth = 15
              	objXL.Columns(4).ColumnWidth = 15
              	objXL.Cells(1, 1).Value = "     "
                objXL.Cells(1, 2).Value = "Audit CheckList for the Desktop"
                objXL.Cells(1, 3).Value = "            "
				objXL.Range("A1:C1").Font.Bold = True
        	    intIndex = intIndex + 1
    			objXL.Cells(intIndex, 1).Select
    			
    			
        	    objXL2.Columns(1).ColumnWidth = 10
                objXL2.Columns(2).ColumnWidth = 20
                objXL2.Columns(3).ColumnWidth = 43
              	objXL2.Columns(4).ColumnWidth = 8
              	objXL2.Columns(5).ColumnWidth = 14
              	objXL2.Columns(6).ColumnWidth = 12
              	objXL2.Columns(7).ColumnWidth = 25
              	objXL2.Columns(8).ColumnWidth = 27
              	objXL2.Cells(1, 1).Value = "RequestID"
                objXL2.Cells(1, 2).Value = "Resource Category"
                objXL2.Cells(1, 3).Value = "ResourceItem"
				objXL2.Cells(1, 4).Value = "RaisedBy"
				objXL2.Cells(1, 5).Value = "Required From"
				objXL2.Cells(1, 6).Value = "Required Till"
				objXL2.Cells(1, 7).Value = "Associate Name"
				objXL2.Cells(1, 8).Value = "Project Name"
				objXL2.Cells(1, 9).Value = "Host Name"
				objXL2.Range("A1:H1").Font.Bold = True
        	    intIndexRams = intIndexRams + 1
    
        	    objXL3.Cells(1, 1).Value = "PC Name"
                objXL3.Cells(1, 2).Value = "Software Name"
                objXL3.Columns(1).ColumnWidth = 14
                objXL3.Columns(2).ColumnWidth = 80
                objXL3.Range("A1:B1").Font.Bold = True
        	    intIndexSoftware = intIndexSoftware + 1
End Sub


Sub AddToExcelFile(strHost, strValue, seatLoc)
    objXL.Cells(intIndex, 1).Value = strHost
    objXL.Cells(intIndex, 2).Value = strValue
    objXL.Cells(intIndex, 3).Value = seatLoc
    
    intIndex = intIndex + 1
    objXL.Cells(intIndex, 1).Select
End Sub

Sub AddToExcelFileTwo(strHost, strValue, seatLoc,fourth)
    objXL.Cells(intIndex, 1).Value = strHost
    objXL.Cells(intIndex, 2).Value = strValue
    objXL.Cells(intIndex, 3).Value = seatLoc
    objXL.Cells(intIndex, 4).Value = fourth
    
    intIndex = intIndex + 1
    objXL.Cells(intIndex, 1).Select
End Sub
'---------------------------------       End of Creating Excel      -------------------------------------------

Call AddToExcelFile(" Report Generated on  ",now,  "    ")

'------------------------------     Host Name   ----------------------------------------------
Call AddToExcelFile("   ","Project Name  ",  "    ")
Set objWMIPcName = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colPcName = objWMIPcName.ExecQuery("Select * from Win32_ComputerSystem")
For Each objComputer in colPcName
PCNAME = objComputer.Name
Call AddToExcelFile("    ","PC Name  ", objComputer.Name)
Next
'------------------------------     End Of Host Name   ----------------------------------------------

'------------------------------     IP address   ----------------------------------------------
Set objRegip = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set IPConfigSet = objRegip.ExecQuery("Select IPAddress from Win32_NetworkAdapterConfiguration ")
For Each IPConfig in IPConfigSet
    If Not IsNull(IPConfig.IPAddress) Then 
        For i=LBound(IPConfig.IPAddress) _
            to UBound(IPConfig.IPAddress)
		Call AddToExcelFile("    ","IP Address ", IPConfig.IPAddress(i))
	Next
    End If
Next
'------------------------------     End of IP Address   ----------------------------------------------

'------------------------------     Seat Location   ----------------------------------------------
Set objRegseat=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
objRegseat.GetStringValue HKLM,"SYSTEM\CurrentControlSet\Services\lanmanserver\parameters","srvcomment",seatLoc
Call AddToExcelFile("    ","Seat Location ", seatLoc)
'------------------------------     End of Seat Location   ----------------------------------------------

'------------------------------     User account   ----------------------------------------------
objRegseat.GetStringValue HKLM,"SOFTWARE\Microsoft\Windows NT\CurrentVersion\WinLogon","DefaultUserName",UserName
Call AddToExcelFile("    ","User Name  ", UserName)
'------------------------------       End of User account   ----------------------------------------------

Call AddToExcelFile("    ","Asset Number  ", "   ")

'------------------------------     Serial Number   ----------------------------------------------
Set objWMIserialnumber = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colNumber = objWMIserialnumber.ExecQuery("SELECT * FROM Win32_BIOS",,48)
For Each colNo In colNumber
    		Call AddToExcelFile("    ","Serial Number  ", colNo.SerialNumber)
Next
'------------------------------      End of Serial Number   ----------------------------------------------

Call AddToExcelFile(" Category :-  BIOS ", "", "    ")
Call AddToExcelFile("1  ", "Boot Password  ", "    ")
Call AddToExcelFile("  ", "BIOS Password  ", "    ")

'------------------------------     Floppy Drive   ----------------------------------------------
Set objWMIfloppy = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colFloppy = objWMIfloppy.ExecQuery("Select * from Win32_FloppyDrive")
FloppyStatus = 1
For Each objfloppy in colFloppy
FloppyStatus = 2
Next
If FloppyStatus = 1 Then
Call AddToExcelFile("  ", "Disable Floppy access", "Disabled")
Else
Call AddToExcelFile("  ", "Disable Floppy access", "Enabled")
End If
'------------------------------       End of Floppy Drive   ----------------------------------------------


'------------------------------     CD/DVD Drive   ----------------------------------------------
Set objWMIcdrom = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colCdrom = objWMIcdrom.ExecQuery("Select * from Win32_CDROMDrive")
CdromStatus = 1
For Each objcdrom in colCdrom
CdromStatus = 2
Next
If CdromStatus = 1 Then
Call AddToExcelFile("  ", "Disable CD/DVD ", "Disabled")
Else
Call AddToExcelFile("  ", "Disable CD/DVD ", "Enabled")
End If
'------------------------------     CD/DVD Drive   ----------------------------------------------

'-------------------------------     USB Status       -------------------------------------------
objRegseat.GetDWORDValue HKLM,"SYSTEM\CurrentControlSet\Services\usbstor","Start",UsbStatus
If UsbStatus = 4 Then
Call AddToExcelFile("  ", "Disable USB ", "Disabled")
Else
Call AddToExcelFile("  ", "Disable USB ", "Enabled")
End If
'-------------------------------     End of USB Status       -------------------------------------------

'-------------------------------     Firmware Version       -------------------------------------------
Set objWMIFirm = GetObject("winmgmts://" & strComputer & "/root\WMI")
Set objInstances = objWMIFirm.InstancesOf("MSDeviceUI_FirmwareRevision",48)
FirmVer = ""
On Error Resume Next
For Each objInstance in objInstances
    With objInstance
    FirmVer = .FirmwareRevision & ", " & FirmVer
    End With
On Error Goto 0
On Error Resume Next
Next
Call AddToExcelFile("  ", "Firmware Version / Rev  ", FirmVer)
'-------------------------------     End of Firmware Version       -------------------------------------------

Call AddToExcelFile("  Category :-  FIle System  ", "", "    ")

'-------------------------------     File System       -------------------------------------------
Call AddToExcelFile(" 1  ", "File System  ", "    ")
Set objWMIdisk = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colfilesystem = objWMIDisk.ExecQuery("Select * from Win32_LogicalDisk")
For Each objFilesystem in colfilesystem
	'If objFilesystem.FileSystem <> "" Then
    Call AddToExcelFile(" ", objFilesystem.DeviceID, objFilesystem.FileSystem)
    'End If
Next
'-------------------------------     End of File System     -------------------------------------------


'-------------------------------     Disk Size       -------------------------------------------
Call AddToExcelFileTwo(" 2  ", "Partition Information", " Total Size   ", " Free Space ")
Set colDisks = objWMIdisk.ExecQuery("Select * from Win32_LogicalDisk")
For Each objDisk in colDisks
	If objDisk.Size <> "" Then
	Call AddToExcelFileTwo(" ", objDisk.DeviceID, Left((objDisk.Size/1073741824),5) & "  GB  ", Left((objDisk.FreeSpace/1073741824),4) & "  GB  ")
	End If
Next
'-------------------------------     End of Disk Size       -------------------------------------------


'--------------------------------    Folder Shares      -----------------------------------------------
Call AddToExcelFile(" 3  ", "Folder Shares ", "    ")
Set objWMIShares = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colshares = objWMIShares.ExecQuery("SELECT * FROM Win32_Share")
For Each objshare in colshares
	If objshare.Name <> "C$" And objshare.Name <> "D$" And objshare.Name <> "E$" And objshare.Name <> "ADMIN$" And objshare.Name <> "IPC$" Then
		Call AddToExcelFile("   ", objshare.Name, objshare.Path)
	End If
Next
'--------------------------------    End of Folder Shares      -----------------------------------------------

Call AddToExcelFile(" 4  ", "Folder Permissions ", "    ")


'-------------------------------     Local User And Guest Details       -------------------------------------------
Call AddToExcelFile(" 5  ", "Local User Details ", "    ")
Set objWMILocaluser = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set collocaluser = objWMILocaluser.ExecQuery("Select * from Win32_UserAccount where LocalAccount = True")
For Each objlocaluser in collocaluser
If objlocaluser.Disabled = True Then
	If objlocaluser.Name = "Guest" Then
		GUEST_STATUS = "Disabled"
	End If
Else
	If objlocaluser.Name = "Guest" Then
		GUEST_STATUS = "Enabled"
	End If
Call AddToExcelFile(" ", " ", objlocaluser.Name)
End If
Next
Call AddToExcelFile(" 6 ", "Guest ID Disabled ", GUEST_STATUS)
'-------------------------------     End of Local User And Guest Details       -------------------------------------------



'-------------------------------     Administrative Privilages       -------------------------------------------
Call AddToExcelFile(" 7  ", "Admin Previlages ", "    ")
Set objWMIAdmin = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
strGroup = "Administrators"
Set objAdmin = GetObject("WinNT://" & strComputer & "/" & strGroup & ", group")
For Each objadmin In objAdmin.Members
Set colAdmin = objWMIAdmin.ExecQuery("Select * from Win32_Group")
If _
objAdmin.Name <> "NSS (GMR)" And _
objAdmin.Name <> "Domain Admins" And _
objAdmin.Name <> "fmsprnadmins" And _
objAdmin.Name <> "GSD (Cognizant)" And _
objAdmin.Name <> "GMRAdmins" And _
objAdmin.Name <> "fmsgmradmins" And _
objAdmin.Name <> "Aapadmin" And _
objAdmin.Name <> "NSSAdmins" And _
objAdmin.Name <> "secadmin" And _
objAdmin.Name <> "FMSGMRADMINS" And _
objAdmin.Name <> "GMRADMINS" And _
objAdmin.Name <> "AAPAdmin" And _
objAdmin.Name <> "SMS Admins" And _
objAdmin.Name <> "fmsgmr" And _
objAdmin.Name <> "secadmin" And _
objAdmin.Name <> "Enterprise Admins" And _
objAdmin.Name <> "secadmin" And _
objAdmin.Name <> "secadmin" And _
objAdmin.Name <> "unameit" And _
objAdmin.Name <> "Unameit" _
Then
Call AddToExcelFile(" ", " ", objAdmin.Name)
End If
Next 
'-------------------------------     End of Administrative previlages       -------------------------------------------

Call AddToExcelFile("  ", " ", " ")
Call AddToExcelFile(" Category :-  Software's ", "", "")




'-----------------------------    Softwares  -------------
SCCM_STATUS = 2
MCAFEE_STATUS = 2
MESSENGER_STATUS = 2
Messenger = " "
RecentUpDate = 00000000
status = 2
sccmerrstat = 2
UnauthSoftware = " "
UNAUTHSOFTWARE_STATUS = 2
					
StrSoftWare = "Select * from Table1"
Set rssoft = conn.Execute(StrSoftWare)


	Set objReg = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv")

Const strBaseKey = "Software\Microsoft\Windows\CurrentVersion\Uninstall\"

	objReg.EnumKey HKLM, strBaseKey, arrSubKeys

    
	For Each strSubKey In arrSubKeys
      status_excep = 1
		intRet = objReg.GetStringValue(HKLM, strBaseKey & strSubKey, "DisplayName", strValue)
    	intRetSecurity = objReg.GetStringValue(HKLM, strBaseKey & strSubKey, "ReleaseType", ReleaseType)
    	If intRet <> 0 Then
        	intRet = objReg.GetStringValue(HKLM, strBaseKey & strSubKey, "QuietDisplayName", strValue)
    	End If
    	If (strValue <> "") AND (intRet = 0) Then
				    	
				Do While not rssoft.EOF
					temp = rssoft("software")
					If InStr(strValue,temp) > 0 Then
					status_excep = 2
					rssoft.MoveNext
				
					Else
						rssoft.MoveNext
					End If	

				Loop
				If status_excep = 1 Then
					objXL3.Cells(intIndexSoftware, 1).Value = PCNAME
    				objXL3.Cells(intIndexSoftware, 2).Value = strValue
    				intIndexSoftware = intIndexSoftware + 1
    			Else

				End If			    	
				rssoft.MoveFirst    	
			'------------          For Latest Security Update       -----------------------------	    	
		    	If ReleaseType = "Security Update" Then
			        	SecUpDatestatus = objReg.GetStringValue(HKLM, strBaseKey & strSubKey, "InstallDate", SecUpDate)
				        	
				        	If RecentUpDate < SecUpDate Then
				        		RecentUpDate = SecUpDate
				        	End If
				End If
			'------------          End For Latest Security Update       -----------------------------	    	
		    


			If InStr(strValue,"McAfee VirusScan Enterprise") > 0 Then
				MCAFEE_STATUS = 1
			End If	
    		
    		If InStr(strValue,"Microsoft System Center") > 0 Then
				SCCM_STATUS = 1
			End If
			
    		If InStr(strValue,"Messenger") > 0 Then
    			Messenger = strValue & ", " & Messenger
				MESSENGER_STATUS = 1
			End If
			If InStr(strValue,"Communicator") > 0 Then
				Messenger = strValue & ", " & Messenger
				MESSENGER_STATUS = 1
			End If
			If InStr(strValue,"Snagit") > 0 Then
    			UnauthSoftware = strValue & ",   " & UnauthSoftware
				UNAUTHSOFTWARE_STATUS = 1
			End If
			If InStr(strValue,"Mozilla Firefox") > 0 Then
    			UnauthSoftware = strValue & ",   " & UnauthSoftware
				UNAUTHSOFTWARE_STATUS = 1
			End If
			If InStr(strValue,"Google ") > 0 Then
    			UnauthSoftware = strValue & ",    " & UnauthSoftware
				UNAUTHSOFTWARE_STATUS = 1
			End If
			If InStr(strValue,"Yahoo") > 0 Then
    			UnauthSoftware = strValue & ",    " & UnauthSoftware
				UNAUTHSOFTWARE_STATUS = 1
			End If
    	End If
	Next
		
		
		If SCCM_STATUS = 2 Then
			Set objWMISccm = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
			Set colsccm = objWMISccm.ExecQuery("Select * From Win32_Directory Where Name = 'C:\\WINDOWS\\system32\\ccmsetup'")
			If colsccm.Count < 1 Then
				Call AddToExcelFile(" 1  ", "Installation of SMS Agent" , "Not Installed")
			Else
				Call AddToExcelFile(" 1  ", "Installation of SMS Agent" , "Installed")
			End if
		Else

				Call AddToExcelFile(" 1  ", "Installation of SMS Agent" , "Installed")
		End If
		
		If MCAFEE_STATUS = 2 Then
				Call AddToExcelFile(" 2  ", "Installation of AV agent" , "Not Installed")
		Else
				Call AddToExcelFile(" 2  ", "Installation of AV agent" , "Installed")
		End If
			Set objWMIMessenger = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
			Set colMessenger = objWMIMessenger.ExecQuery("Select * From Win32_Directory Where Name = 'C:\\Program Files\\Messenger'")
			If colMessenger.Count < 1 Then

			Else
				Messenger = "Windows Messenger , " & Messenger
				MESSENGER_STATUS = 1

			End If
		If MESSENGER_STATUS = 2 Then
				Call AddToExcelFile(" 3  ", "Messenger" , "Not Installed")
		Else
				Call AddToExcelFile(" 3  ", "Messenger" , Messenger)
		End If
		If UNAUTHSOFTWARE_STATUS = 2 Then
				Call AddToExcelFile(" 4  ", "Unauthorized softwares " , "Not Installed")
		Else
				Call AddToExcelFile(" 4  ", "Unauthorized softwares " , UnauthSoftware)
		End If

'----------------------------------------------   DAT version   ----------------------------
Set objRegDAT=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
objRegDAT.GetStringValue HKLM,"SOFTWARE\Network Associates\ePolicy Orchestrator\Application Plugins\VIRUSCAN8700\","DATVersion",DATversion
If DATversion >= 5873.0000 Then
	Call AddToExcelFile(" 5  ", "AV – Up to date Virus definitions" , "Updated")
Else
	Call AddToExcelFile(" 5  ", "AV – Up to date Virus definitions" , "Not Updated")
End If
'----------------------------------------------   End of DAT version   ----------------------------

'----------------------------------------------   Logon Warning Message   ----------------------------
Set objRegwarn=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
	On Error Resume Next
objRegwarn.GetStringValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system\","legalnoticetext",LogonWarning
	If Err.Number <> 0 Then
		Call AddToExcelFile(" 6  ", "Logon Warning Message" , "No")
	End If
If InStr(LogonWarning,"You are authorized to use this system for approved ") > 0 Then 'If LogonWarning = 5857.0000 Then
	Call AddToExcelFile(" 6  ", "Logon Warning Message" , "yes")
Else
	Call AddToExcelFile(" 6  ", "Logon Warning Message" , "Yes, but text is changed.")
End If
'----------------------------------------------   End of Logon Warning Message   ----------------------------


'-------------------------------     Disable RDP       -------------------------------------------
Call AddToExcelFile(" 7  ", "Users on RDP", "    ")
count_rdp = 0
Set objWMIrdp = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
strGrouprdp = "Remote Desktop Users"
Set objrdp = GetObject("WinNT://" & strComputer & "/" & strGrouprdp & ", group")
For Each objRdp In objrdp.Members
Set colrdp = objWMIrdp.ExecQuery("Select * from Win32_Group")
Call AddToExcelFile(" ", " ", objRdp.Name)

Next 
'-------------------------------     End of Disable RDP       -------------------------------------------


'-----------------------------    Disable IIS  ------------------------------------------------------------
IIS_STATUS = 2
	Set objRegiis = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv")

Const strBaseIisKey = "SOFTWARE\Microsoft\"

	objRegiis.EnumKey HKLM, strBaseIisKey, arrSubKeysIis

    
	For Each strSubIisKey In arrSubKeysIis
		If strSubIisKey = "InetStp" Then
			IIS_STATUS = 1
		End If
	Next
		
	If IIS_STATUS = 1 Then
			Call AddToExcelFile(" 8  ", "IIS Status" , "Enabled")
	Else
			Call AddToExcelFile(" 8  ", "IIS Status" , "Disabled")
	End If
'-----------------------------    End of Disable IIS  ---------------------------------------------------


'------------------------------      Security Update     --------------------------------------------------
YearUpdate = Fix(RecentUpDate / 10000)
YearUpdate = Left(RecentUpDate,4)
temp = Right(RecentUpDate,4)
MonthUpdate = Fix(temp / 100)
DateUpdate = Right(RecentUpDate,2)
If RecentUpDate = 00000000 Then
Call AddToExcelFile(" 9   ","Security Patch Last Updated On ", " Security Patches Not updated Properly")
Else
Call AddToExcelFile(" 9   ","Security Patch Last Updated On ", DateUpDate & " / " & MonthUpDate & " / " & YearUpDate)
End If
'----------------------------    End Of Security Update    ---------------------------------------------------

'----------------------------    WallPaper    ---------------------------------------------------
ResultWallpaper = ""
ResultSCR = ""
Set objRegwallpaper=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
objRegwallpaper.EnumKey HKU, "", arrSubKeysWallpaper

For Each strSubKeyWallpaper In arrSubKeysWallpaper
	
on error resume next
Set objSID = objWMIrdp.Get("Win32_SID.SID='" & strSubKeyWallpaper & "'")
If Err.Number = 0 Then
		if objSID.AccountName = UserName Then
			objRegwallpaper.GetStringValue HKU,strSubKeyWallpaper & "\Control Panel\Desktop","ConvertedWallpaper",ConvertedWallpaper
			objRegDAT.GetStringValue HKU,strSubKeyWallpaper & "\Software\Policies\Microsoft\Windows\Control Panel\Desktop\","ScreenSaveTimeOut",screensaver
			ResultWallpaper = ResultWallpaper & ConvertedWallpaper 
			ResultSCR = ResultSCR & screensaver
		end if
end if
Next
	If InStr(ResultWallpaper,"\\CTSINGMRCFAB\NETLOGON\Operation-Clarity-Wallpaper.jpg") > 0 Then
		Call AddToExcelFile(" 10  ","Verification of Cognizant Background theme", "Enabled ")
	Else
		Call AddToExcelFile(" 10  ","Verification of Cognizant Background theme", ResultWallpaper)
	End If
'----------------------------    End Of WallPaper    ---------------------------------------------------

'------------------------------     ScreenSaver Lockout Time   ----------------------------------------------
Call AddToExcelFile(" 11   ","ScreenSaver Lockout Time  ", (ResultSCR/60) & "  Min")
'------------------------------     End of ScreenSaver Lockout Time   ----------------------------------------------



StrSQL = "Select * from rams"
Set rs = conn.Execute(StrSQL)

Do While not rs.EOF
	If UCase(rs("AssetID")) = UCase(PCNAME) Then

		objXL2.Cells(intIndexRams, 1).Value = rs("RequestID")
    	objXL2.Cells(intIndexRams, 2).Value = rs("Resource Category")
    	objXL2.Cells(intIndexRams, 3).Value = rs("ResourceItem")
    	objXL2.Cells(intIndexRams, 4).Value = rs("RaisedBy")
    	objXL2.Cells(intIndexRams, 5).Value = rs("Required From")
    	objXL2.Cells(intIndexRams, 6).Value = rs("Required Till")
    	objXL2.Cells(intIndexRams, 7).Value = rs("Associate Name")
    	objXL2.Cells(intIndexRams, 8).Value = rs("Project Name")
    	objXL2.Cells(intIndexRams, 9).Value = rs("AssetID")
    	
	    intIndexRams = intIndexRams + 1
	End If	
	rs.MoveNext
Loop
conn.Close



WScript.Echo "Completed . . ."