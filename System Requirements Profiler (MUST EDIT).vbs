''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

command_line = "product=totaltyler%environment=Production%title=Production Incode 10 Application Server%shortname=IAS%order=2%url=https://check.tylertech.com/s/XXFRW31luYA%u=XXFRW31luYA%manufacturer=*%model=*%computer_name=*%serial_number=*%processor_model=*%processor_description=*%>.processor_cores_virtual=4%>.processor_cores=6%>.ram=8%>.storage=500%-.operating_system=2000;2003;2011;2019;Essentials;Foundation%&.operating_system=Server%|.operating_system_architecture=32;64%-.domain_role=Domain Controller;Workstation%&.on_domain=yes"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set requirements = GetRequirements(command_line)
product = requirements("product")
title = requirements("title")
u = requirements("u")
url = requirements("url")
system_class = requirements("title")

If Left(u, 1) = "V" Then system_class = requirements("legacy_name")

Dim wsh, fso
Set wsh=WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

serial_number = GetSerialNumber()
CurrentDirectory = fso.GetAbsolutePathName(".")
ReportFileName = "report-" + serial_number + ".html"
ReportFile = CurrentDirectory + "/" + ReportFileName

WaitHtml = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd""><html><head>	<title>%title% - Tyler Technologies</title>	<link rel=""stylesheet"" href=""https://check.tylertech.com/assets/css/style.css"">	<meta http-equiv=""refresh"" content=""5""></head><body><div id=""container"">	<div id=""header""><img src=""https://check.tylertech.com/assets/images/header.png""></div>	<div id=""content"">		<h1>%title%</h1>		<div><p>Please wait, this may take a few minutes...</p></div>		<div id=""footer-container"">			<div id=""footer"">				<div class=""left""><img src=""https://check.tylertech.com/assets/images/footer.png"" /></div>				<div class=""right""><p>%title%</p></div>			</div>		</div>	</div></div></body></html>"
ReportHtml = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd""><html><head>	<title>%title% Hardware Report - Tyler Technologies</title>	<link rel=""stylesheet"" href=""https://check.tylertech.com/assets/css/style.css""></head><body><div id=""container"">	<div id=""header""><img src=""https://check.tylertech.com/assets/images/header.png""></div>	<div id=""content"">		<h1>%title% Hardware Report</h1>				<br><p>Review the report, then click the ""Send Report"" button below to send it to Tyler Technologies.</p><br><hr><br>		<div>%sections%</div>				<form action=""https://check.tylertech.com/assets/sysreq/submit.php"" method=""post"">			<input type=""hidden"" name=""u"" value=""%u%"">			<input type=""hidden"" name=""SystemName"" value=""%computer_name%"">			<input type=""hidden"" name=""SystemClass"" value=""%system_class%"">			<input type=""hidden"" name=""ReportHtml"" value=""%report_html%"">			<br><br>			<center>For a list of full system requirements, click <a href=""%url%"" target=""_blank"">here</a>.</center>			<br>			<table style=""width:100%; text-align:center;"">				<tr>					<td><input type=""submit"" value=""Send Report""></td>				</tr>			</table>		</form> 				<div id=""footer-container"">			<div id=""footer"">				<div class=""left""><img src=""https://check.tylertech.com/assets/images/footer.png"" /></div>					<div class=""right"">						<p>%date_time%</p>					</div>			</div>		</div>	</div></div></body></html>"

WaitHtml = Replace(WaitHtml, "%title%", title)
WaitHtml = Replace(WaitHtml, "%report_file_name%", ReportFileName)

WriteFile WaitHtml, ReportFile

OpenBrowser(ReportFile)

SystemInformationSection = GetSystemInformationSection
ProcessorSection = GetProcessorSection
MemorySection = GetMemorySection
HardDriveSection = GetHardDriveSection
NetFrameworkSection = GetNetFrameworkSection
OperatingSystemSection = GetOperatingSystemSection
OfficeSection = GetOfficeSection
SqlSection = GetSqlSection

computer_name = GetComputerName
ReportHtml = Replace(ReportHtml, "%computer_name%", computer_name)

sections = SystemInformationSection & ProcessorSection & MemorySection & HardDriveSection & OperatingSystemSection & NetFrameworkSection & OfficeSection & SqlSection
ReportHtml = Replace(ReportHtml, "%sections%", sections)

date_time = GetDateTime
ReportHtml = Replace(ReportHtml, "%date_time%", date_time)

ReportHtml = Replace(ReportHtml, "%title%", title)
ReportHtml = Replace(ReportHtml, "%system_class%", system_class)
ReportHtml = Replace(ReportHtml, "%url%", url)
ReportHtml = Replace(ReportHtml, "%u%", u)

If u = "" Then ReportHtml = Replace(ReportHtml, "<input type=""submit"" value=""Submit"">", "")

SendHtml = Mid(ReportHtml, InStr(ReportHtml, "<div id=""content"""))
SendHtml = Mid(SendHtml, 1, InStr(SendHtml, "<form") - 1)

ReportHtml = Replace(ReportHtml, "%report_html%", Escape(SendHtml))

WriteFile ReportHtml, ReportFile


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GetRequirements(command_line)
	Set product_requirements = CreateObject("scripting.dictionary") 
	requirements = Split(command_line, "%")
	For Each requirement In requirements
		pair = Split(requirement, "=")
		key = pair(0)
		value = pair(1)
		product_requirements(key) = value
	Next

	Set GetRequirements = product_requirements
	
End Function

Function GetStatus(item, ByVal data)
	informational = False
	
	Select Case item
		Case "sql"
			informational = True
		Case "security_software"
			informational = True
	End Select

	If(Contains(item, "net_framework")) Then
		informational = True
	End If
	
	
	status = ""
	
	all_of = true
	one_of = false
	none_of = true
	greater_than = false
	less_than = false
	
	requirements_plain = requirements(item)
	requirements_all_of = requirements("&." & item)
	requirements_one_of = requirements("|." & item)
	requirements_none_of = requirements("-." & item)
	requirements_greater_than = requirements(">." & item)
	requirements_less_than = requirements("<." & item)
	
	If requirements_plain = "*" Then
		status = "NONE"
	Else
		If(requirements_all_of = "") Then
			all_of = true
		Else
			requirement_list = Split(requirements_all_of, ";")
			For Each requirement In requirement_list
				If Not Contains(data, requirement) Then
					all_of = false
				End If
			Next
		End If
		
		If(requirements_one_of = "") Then
			one_of = true
		Else
			requirement_list = Split(requirements_one_of, ";")
			For Each requirement In requirement_list
				If Contains(data, requirement) Then
					one_of = true
					Exit For
				End If
			Next
		End If
		
		If(requirements_none_of = "") Then
			none_of = true
		Else
			none_of = true
			requirement_list = Split(requirements_none_of, ";")
			For Each requirement In requirement_list
				If Contains(data, requirement) Then
					none_of = false
					Exit For
				End If
			Next			
		End If		
		
		If(requirements_greater_than = "") Then
			greater_than = true
		ElseIf (CDbl(data) >= CDbl(requirements_greater_than)) Then
			greater_than = true
		End If
		
		
		If(requirements_less_than = "") Then
			less_than = true
		ElseIf (CDbl(data) <= CDbl(requirements_less_than)) Then
			less_than = true
		End If

		If all_of And one_of And none_of And greater_than And less_than Then
			status = "SUCCESS"
		Else
			status = "FAIL"
		End If	

	End If
	
	If(item = "hard_drive") Then
		If status = "SUCCESS" Then status = "DRIVE_SUCCESS"
		If status = "FAIL" Then status = "DRIVE_FAIL"
	End If
	
	If(item = "storage") Then
		If status = "SUCCESS" Then status = "DRIVE_SUCCESS"
		If status = "FAIL" Then status = "NONE"
		If data = -1 Then status = "FAIL"
	End If
	
	link = "%url%#server"
	If(InStr(1, LCase(title), "workstation")) > 0 Then link = "%url%#workstation"
	
	Select Case status
		Case "NONE"
			color = "black"
			status = ""
		Case "SUCCESS"
			color = "green"
			status = "Meets requirements"
		Case "FAIL"
			If informational Then
				color = "orange"
				status = "Does not meet requirements <a style=""color:orange; text-decoration:underline;"" href=""" + link + """>[?]</a>"
			Else
				color = "red"
				status = "Does not meet requirements <a style=""color:red; text-decoration:underline;"" href=""" + link + """>[?]</a>"
			End If
		Case "DRIVE_SUCCESS"
			color = "green"
			status = "Enough free space to run Tyler products"
		Case "DRIVE_FAIL"
			color = "orange"
			status = "Not enough free space to run Tyler products"
		Case Else
			color = "orange"
			status = "Contact support"
		End Select

	GetStatus = "<font color=" + color + "><b>" + status + "</b>"

End Function

Function AddSimpleItem(ByRef array, title, data, status)
	If(title <> "") Then
		title = "<b>" & title & ":</b> "
	End If

	item = "<tr><td>%item%</td><td align=""right"">%status%</td></tr>"
	item = Replace(item, "%item%", title & data)
	item = Replace(item, "%status%", status)
	
	Add array, item 
	
End Function

Function BuildSection(SectionTitle, items)

	section = "<h3>%section_title%</h3><table style=""width:100%"">%section_content%</table>"
	content = ""
	
	If UBound(items) = 0 Then
		BuildSection = ""
		Exit Function
	End If
	
	For Each item in items	
		content = content & item
	Next
	
	section = Replace(section, "%section_title%", SectionTitle)
	section = Replace(section, "%section_content%", content)
	
	BuildSection = section
	
End Function

Function Applicable(item)
	Applicable = true
	r = requirements(item) & requirements("|." & item) & requirements("&." & item) & requirements("-." & item) & requirements(">." & item) & requirements("<." & item)
	If(r = "") Then Applicable = false
End Function

Function GetDateTime() 

	daylight = ""

	For Each objItem in WMIQuery("Select * from Win32_TimeZone")
		For Each objItem2 in WMIQuery("Select * From Win32_ComputerSystem")
			blnDaylightInEffect = objItem2.DaylightInEffect
		Next

		If blnDaylightInEffect Then
			daylight = objItem.DaylightName
		Else
			daylight = objItem.StandardName
		End If
	Next

	GetDateTime = "Report generated on " + FormatDateTime(Now, vbShortDate) + " at " + FormatDateTime(Now, vbLongTime) + " " + daylight

End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' System Information Section

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GetSystemInformationSection()
	section = "System Information"
	Dim items()
	ReDim items(0)

	If Applicable("manufacturer") Then
		manufacturer = GetManufacturer()
		manufacturer_status = GetStatus("manufacturer", manufacturer)
		AddSimpleItem items, "Manufacturer", manufacturer, manufacturer_status
	End If

	If Applicable("model") Then
		model = GetModel()
		model_status = GetStatus("model", model)
		AddSimpleItem items, "Model", model, manufacturer_status
	End If
	
	If Applicable("serial_number") Then
		serial_number = GetSerialNumber()
		serial_number_status = GetStatus("serial_number", serial_number)
		AddSimpleItem items, "Serial Number", serial_number, serial_number_status
	End If
	
	If Applicable("computer_name") Then
		computer_name = GetComputerName()
		computer_name_status = GetStatus("computer_name", computer_name)
		AddSimpleItem items, "Computer Name", computer_name, computer_name_status 
	End If

	If Applicable("on_domain") Then
		on_domain = GetOnDomain()
		on_domain_status = GetStatus("on_domain", on_domain)
		AddSimpleItem items, "On Domain", on_domain, on_domain_status
	End If
	
	If Applicable("domain_role") Then	
		domain_role = GetDomainRole()
		domain_role_status = GetStatus("domain_role", domain_role)
		AddSimpleItem items, "Domain Role", domain_role, domain_role_status
	End If
	
	GetSystemInformationSection = BuildSection(section, items)

End Function

Function GetManufacturer()
	For Each objItem in WMIQuery("Select * from Win32_ComputerSystem")
		GetManufacturer = objItem.Manufacturer
		Exit For
	Next
End Function

Function GetModel()
	For Each objItem in WMIQuery("Select * from Win32_ComputerSystem")
		GetModel = objItem.Model
		Exit For
	Next
End Function

Function GetSerialNumber()
	For Each objItem in WMIQuery("Select * from Win32_SystemEnclosure")
		GetSerialNumber = objItem.SerialNumber
		Exit For
	Next
End Function

Function GetComputerName()
	For Each objItem in WMIQuery("Select * from Win32_ComputerSystem")
		GetComputerName = objItem.Name
		Exit For
	Next
End Function

Function GetOnDomain()
	For Each objItem in WMIQuery("Select * from Win32_ComputerSystem")
		PartOfDomain = objItem.PartOfDomain
		Exit For
	Next

	Select Case PartOfDomain
	Case "True"
		For Each objItem in WMIQuery("Select * from Win32_ComputerSystem")
			Domain = objItem.Domain
			GetOnDomain = "Yes (" + Domain + ")"
			Exit For
		Next
	Case Else
		GetOnDomain = "No"
	End Select
	
End Function

Function GetDomainRole()
	For Each objItem in WMIQuery("Select DomainRole from Win32_ComputerSystem")
		DomainRole = objItem.DomainRole
		Exit For
	Next
	
	Select Case DomainRole
		Case 0
			GetDomainRole = "Standalone Workstation"
		Case 1
			GetDomainRole = "Member Workstation"
		Case 2
			GetDomainRole = "Standalone Server"
		Case 3
			GetDomainRole = "Member Server"
		Case 4
			GetDomainRole = "Backup Domain Controller"
		Case 5
			GetDomainRole = "Primary Domain Controller"
		Case Else
			GetDomainRole = "Unknown"
	End Select
	
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Processor Section

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Function GetProcessorSection()
	section = "Processor"
	Dim items()
	ReDim items(0)

	If Applicable("processor_model") Then
		processor_model = GetProcessorModel()
		processor_model_status = GetStatus("processor_model", processor_model)
		AddSimpleItem items, "Model", processor_model, processor_model_status
	End If
	
	If Applicable("processor_description") Then
		processor_description = GetProcessorDescription()
		processor_description_status = GetStatus("processor_description", processor_description)
		AddSimpleItem items, "Description", processor_description, processor_description_status
	End If

	If IsVirtualMachine() Then
		If Applicable("processor_cores_virtual") Then
			processor_cores_virtual = GetProcessorCores()
			processor_cores_virtual_status = GetStatus("processor_cores_virtual", processor_cores_virtual)
			AddSimpleItem items, "Virtual Cores", processor_cores_virtual, processor_cores_virtual_status
		End If
		If Applicable("processor_cpus_virtual") Then
			processor_cpus_virtual = GetLogicalProcessors()
			processor_cpus_virtual_status = GetStatus("processor_cpus_virtual", processor_cores_virtual)
			AddSimpleItem items, "Logical Processors", processor_cpus_virtual, processor_cpus_virtual_status
		End If
	Else
		If Applicable("processor_cores") Then
			processor_cores = GetProcessorCores()
			processor_cores_status = GetStatus("processor_cores", processor_cores)
			AddSimpleItem items, "Cores", processor_cores, processor_cores_status
		End If

		If Applicable("processor_cpus") Then
			processor_cpus = GetPhysicalProcessors()
			processor_cpus_status = GetStatus("processor_cpus", processor_cpus)
			AddSimpleItem items, "Physical Processors", processor_cpus, processor_cpus_status
		End If
	End If
	
	GetProcessorSection = BuildSection(section, items)
End Function

Function IsVirtualMachine()
	For Each objItem in WMIQuery("Select * from Win32_ComputerSystem")
			If(Contains(objItem.Model, "virtual")) Then
				IsVirtualMachine = True	
			Else
				IsVirtualMachine = False
			End If
		Exit For
	Next
End Function

Function GetProcessorModel()
	For Each objItem in WMIQuery("Select * from Win32_Processor")
		GetProcessorModel = objItem.Name
		Exit For
	Next
End Function

Function GetProcessorDescription()
	For Each objItem in WMIQuery("Select * from Win32_Processor")
		GetProcessorDescription = objItem.Description
		Exit For
	Next
End Function

Function GetProcessorCores()
	On Error Resume Next
	Cores = 0
	For Each objItem in WMIQuery("Select * from Win32_Processor")
		Cores = Cores + objItem.NumberOfCores
	Next
	GetProcessorCores = Cores
	On Error Goto 0
End Function

Function GetLogicalProcessors()
	For Each objItem in WMIQuery("SELECT * FROM Win32_ComputerSystem")
		GetLogicalProcessors = objItem.NumberofLogicalProcessors
		Exit For
	Next
End Function

Function GetPhysicalProcessors()
	For Each objItem in WMIQuery("SELECT * FROM Win32_ComputerSystem")
		GetLogicalProcessors = objItem.NumberOfProcessors
		Exit For
	Next
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Memory Section

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GetMemorySection()
	section = "Memory"
	Dim items()
	ReDim items(0)

	If Applicable("ram") Then
		ram = GetRAM()
		ram_status = GetStatus("ram", ram)
		AddSimpleItem items, "Amount of RAM", ram & "GB", ram_status
	End If

	GetMemorySection = BuildSection(section, items)
End Function

Function GetRAM()
	InstalledRAM = 0
	For Each objItem in WMIQuery("SELECT * FROM Win32_PhysicalMemory")
		InstalledRAM = InstalledRAM + objItem.Capacity
	Next
	GetRAM = Round(InstalledRAM / 1024 / 1024 / 1024, 1)
	
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Hard Drive Section

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GetHardDriveSection()
	section = "Hard Drives"
	Dim items()
	ReDim items(0)

	Dim drives()
	ReDim drives(0)
	
	storage_drive_found = False
		
	For Each Drive in WMIQuery("SELECT * FROM Win32_DiskDrive")
		If Drive.InterfaceType <> "USB" Then
			Set Partitions = WMIQuery("ASSOCIATORS OF {Win32_DiskDrive.DeviceID="""	& Replace(Drive.DeviceID, "\", "\\") & """} WHERE AssocClass = " & "Win32_DiskDriveToDiskPartition")
			For Each Partition In Partitions
				Set LogicalDisks = WMIQuery("ASSOCIATORS OF {Win32_DiskPartition.DeviceID=""" & Partition.DeviceID & """} WHERE AssocClass = Win32_LogicalDiskToPartition")
				For Each LogicalDisk In LogicalDisks
					letter = LogicalDisk.DeviceID
					label = LogicalDisk.VolumeName
					If label = "" Then label = "Local Disk"
					
					If Not IsNull(LogicalDisk.FreeSpace) Then
						free = Round((LogicalDisk.FreeSpace - (LogicalDisk.FreeSpace * 0.1)) / 1024 / 1024 / 1024, 1)
						total = Round((LogicalDisk.Size - (LogicalDisk.Size * 0.1)) / 1024 / 1024 / 1024, 1)
					Else
						free = 0
						total = 0
					End If
					
					Add drives, label & " ;" & letter & ";" & total & ";" & free
					
					If Applicable("storage") Then
						If letter = "C:" And Applicable("hard_drive_c") Then
							'hard_drive_c was specified, so C: will get displayed below
						Else
							status = GetStatus("storage", free)
							If(Contains(status, "green")) Then
								storage_drive_found = True
							End If
						
							AddSimpleItem items, label & "(" & letter & ")", total & "GB total, " & free & "GB free", status
						End If
					
					End If
					
					If Applicable("hard_drive") Then
						AddSimpleItem items, label & "(" & letter & ")", total & "GB total, " & free & "GB free", GetStatus("hard_drive", free)
					End If					
					
				Next
			Next
		End If
	Next	
	
	If Applicable("storage") Then
		If Not (storage_drive_found) Then
			AddSimpleItem items, "Hard drive with enough free space", "Not Found", GetStatus("storage", -1)
		End If
	End If
	
	For Each drive_letter In Array("a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z")
		If Applicable("hard_drive_" & drive_letter) Then
			drive_found = false
			
			'For named drives
			For Each drive in drives
				If Contains(drive, drive_letter & ":") Then
					drive = Split(drive, ";")
					label = drive(0)
					letter = drive(1)
					total = drive(2)
					free = drive(3)
					AddSimpleItem items, label & "(" & letter & ")", total & "GB total, " & free & "GB free", GetStatus("hard_drive_" & drive_letter, total)
					drive_found = true
					Exit For
				End If
			Next
							
			If Not drive_found Then
				AddSimpleItem items, UCase(drive_letter), "Not Found", GetStatus("hard_drive_" & drive_letter, -1)
			End If
		End If
	Next
	
	
	GetHardDriveSection = BuildSection(section, items)

End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Operating System Section

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Function GetOperatingSystemSection()
	section = "Operating System"
	Dim items()
	ReDim items(0)

	If Applicable("operating_system") Then
		operating_system = GetOperatingSystem()
		operating_system_status = GetStatus("operating_system", operating_system)
		AddSimpleItem items, "Operating System", operating_system, operating_system_status
	End If

	If Applicable("operating_system_architecture") Then
		operating_system_architecture = GetArchitecture()
		operating_system_architecture_status = GetStatus("operating_system_architecture", operating_system_architecture)
		AddSimpleItem items, "Architecture", operating_system_architecture & "-bit", operating_system_architecture_status	
	End If
	
	If Applicable("security_software") Then
		security_software = GetSecuritySoftware()
		security_software_status = GetStatus("security_software", security_software)
		AddSimpleItem items, "Security Software", security_software, security_software_status
	End If
	
	GetOperatingSystemSection = BuildSection(section, items)
End Function

Function GetOperatingSystem()
	OperatingSystem = "Unknown"
	For Each objItem in WMIQuery("Select * from Win32_OperatingSystem")
		OperatingSystem = objItem.Caption
		Exit For
	Next
	
	ServicePack = "No service pack"
	
	For Each objItem in WMIQuery("Select * from Win32_OperatingSystem")
		ServicePack = objItem.ServicePackMajorVersion
		
		If ServicePack = "0" Then
			ServicePack =  "No Service Pack"
		Else
			ServicePack =  " Service Pack " + CStr(ServicePack)
		End If
		
		Exit For
	Next
	
	GetOperatingSystem = "<br>" + OperatingSystem + " (" + ServicePack + ")"
	
End Function

Function GetSecuritySoftware()
	security_software = "None detected"
	
	WSC_SECURITY_PROVIDER_NONE                   = 0
	WSC_SECURITY_PROVIDER_FIREWALL               = 1
	WSC_SECURITY_PROVIDER_AUTOUPDATE_SETTINGS    = 2
	WSC_SECURITY_PROVIDER_ANTIVIRUS              = 4
	WSC_SECURITY_PROVIDER_ANTISPYWARE            = 8
	WSC_SECURITY_PROVIDER_INTERNET_SETTINGS      = 16
	WSC_SECURITY_PROVIDER_USER_ACCOUNT_CONTROL   = 32
	WSC_SECURITY_PROVIDER_SERVICE                = 64
		
	all_security_software = ""
	
	For Each objItem in WMIQueryWithNamespace("winmgmts:root/SecurityCenter2", "Select * from AntiVirusProduct")
		display_name = objItem.displayName
		product_state = Hex(objItem.productState)
		
		security_provider = CInt(Mid(product_state, 1, 2))
		active =  CInt(Mid(product_state, 3, 2))
		updated = CInt(Mid(product_state, 5, 2))

		firewall = "has no firewall"
		If(security_provider AND WSC_SECURITY_PROVIDER_FIREWALL) > 0 Then
			firewall = "has a firewall"
		End If
		
		antivirus = "has no antivirus"
		If(security_provider AND WSC_SECURITY_PROVIDER_ANTIVIRUS) > 0 Then
			antivirus = "has antivirus"
		End If
		
		antispyware = "has no antispyware"
		If(security_provider AND WSC_SECURITY_PROVIDER_ANTISPYWARE) > 0 Then
			antispyware = "has antispyware"
		End If
		
		uac = "UAC is disabled"
		If(security_provider AND WSC_SECURITY_PROVIDER_USER_ACCOUNT_CONTROL) > 0 Then
			uac = "UAC is enabled"
		End If
	
		If(active = 10) Then
			active = "active"
		Else
			active = "not active"
		End If
			
		If(updated = 0) Then
			updated = "up-to-date"
		Else
			updated = "out of date"
		End If	
	
		security_provider = firewall & ", " & antivirus & ", " & antispyware & ", " & uac
	
		security_software = "[" & product_state & "] " & display_name & " (" & active & ", " & updated & "): " & security_provider 
		
		all_security_software = all_security_software & "<br>" & security_software
		
	Next
	

	GetSecuritySoftware = all_security_software
	
End Function

Function GetArchitecture()
	GetArchitecture = GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Office Section

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Function GetOfficeSection()
	section = "Microsoft Office"
	Dim items()
	ReDim items(0)

	office = GetOffice()
	
	products = Split(office, "!!!")
	For Each product In products
		
		'word = Instr(product, "Word")
		'excel = Instr(product, "Excel")
		'powerpoint = Instr(product, "PowerPoint")
		'outlook = Instr(product, "Outlook")
		
		architecture = Instr(product, "64")
		If architecture = 0 Then
			architecture = " (32-bit)"
		Else
			architecture = ""
		End If
		
		product = product & architecture
		
		If Contains(product, "Excel") Then
			If Applicable("office_excel") Then
				product_status = GetStatus("office_excel", product)
				AddSimpleItem items, "", product, product_status	
			End If
		End If
		
		If Contains(product, "Word") Then
			If Applicable("office_word") Then
				product_status = GetStatus("office_word", product)
				AddSimpleItem items, "", product, product_status	
			End If
		End If
				
	Next
	
	GetOfficeSection = BuildSection(section, items)
End Function

Function GetOffice()
	office = ""
	For Each objItem in WMIQuery("SELECT Version FROM Win32_Product WHERE Name Like 'Microsoft Office%'")
		office = office & objItem.Name & "!!!"
	Next
	GetOffice = office
	
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' SQL Section

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Function GetSqlSection()
	section = "Microsoft SQL Instances"
	Dim items()
	ReDim items(0)

	If Applicable("sql") Then
		sql = GetSql()
		
		products = Split(sql, "!!!")
		For Each product In products
			product_status = GetStatus("sql", product)
			AddSimpleItem items, "", product, product_status		
		Next
	End If
	
	GetSqlSection = BuildSection(section, items)
End Function

Function GetSql()
	sql = ""
	
	Const HKLM = &H80000002

	keyPath = "SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL"
	Set reg = getObject( "Winmgmts:root\default:StdRegProv" )

	If reg.enumValues(HKLM, keyPath, valueNames, types ) = 0 Then
		if isArray(valueNames) then
			For i = 0 to UBound( valueNames )
				reg.getStringValue HKLM, keyPath, valueNames(i), value
				
				instance = valueNames(i)
				instanceName = value
				
				baseKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server\" & instanceName & "\Setup\"
				version = ReadKey(baseKey & "Version")
				patchLevel = ReadKey(baseKey & "PatchLevel")
				edition = ReadKey(baseKey & "Edition")
				
				arrVersion = Split(version, ".")
				'version = arrVersion(0) & "." & arrVersion(1)
				version = arrVersion(0)
				Select Case version
					Case "13"
						version = "2016"
					Case "12"
						version = "2014"
					Case "11"
						version = "2012"
					Case "10.50"
						If(arrVersion(1)) = "50" Then
							version = "2008 R2"
						Else
							version = "2008"
						End If
					Case "9"
						version = "2005"
					Case "8"
						version = "2000"
					Case Else
						version = "Unknown"
				End Select
						
				'friendlyName = "Microsoft SQL Server " & version & " (" & edition & "), Build " & patchLevel & " (" & instance & ")"
				friendlyName = "SQL Server " & version & " (" & edition & ")"
				
				sql = sql & friendlyName & "!!!"
			Next
		End if
	End if
	
	If(sql = "") Then
		sql = "Not installed"
	Else
		sql = Left(sql, Len(sql) - 3)
	End If
	
	GetSql = sql
	
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' .NET Framework Section

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Function GetNetFrameworkSection()
	section = ".NET Framework"
	Dim items()
	ReDim items(0)

	If Applicable("net_framework1") Then
		net_framework1 = GetNetFramework1()
		status = GetStatus("net_framework1", net_framework1)
		AddSimpleItem items, ".NET Framework 1", net_framework1, status
	End If

	If Applicable("net_framework1_1") Then
		net_framework1_1 = GetNetFramework1_1()
		status = GetStatus("net_framework1_1", net_framework1_1)
		AddSimpleItem items, ".NET Framework 1.1", net_framework1_1, status
	End If
	
	If Applicable("net_framework2") Then
		net_framework2 = GetNetFramework2()
		status = GetStatus("net_framework2", net_framework2)
		AddSimpleItem items, ".NET Framework 2", net_framework2, status
	End If	
	
	If Applicable("net_framework3") Then
		net_framework3 = GetNetFramework3()
		status = GetStatus("net_framework3", net_framework3)
		AddSimpleItem items, ".NET Framework 3", net_framework3, status
	End If	
	
	If Applicable("net_framework3_5") Then
		net_framework3_5 = GetNetFramework3_5()
		status = GetStatus("net_framework3_5", net_framework3_5)
		AddSimpleItem items, ".NET Framework 3.5", net_framework3_5, status
	End If	
	
	If Applicable("net_framework4") Then
		net_framework4 = GetNetFramework4()
		status = GetStatus("net_framework4", net_framework4)
		AddSimpleItem items, ".NET Framework 4", net_framework4, status
	End If	
	
	If Applicable("net_framework4_5") Then
		net_framework4_5 = GetNetFramework4_5()
		status = GetStatus("net_framework4_5", net_framework4_5)
		AddSimpleItem items, ".NET Framework 4.5", net_framework4_5, status
	End If	
	
	If Applicable("net_framework4_6") Then
		net_framework4_6 = GetNetFramework4_6()
		status = GetStatus("net_framework4_6", net_framework4_6)
		AddSimpleItem items, ".NET Framework 4.6", net_framework4_6, status
	End If	
	
	GetNetFrameworkSection = BuildSection(section, items)

End Function

Function GetNetFramework1_4(version, key)
	full_version = "Not installed"
	
	registry_version = ReadKey(key)
	
	If Not registry_version = "" Then
		full_version = "Installed (v" & registry_version & ")"
	End If
	
	GetNetFramework1_4 = full_version
	
End Function

Function GetNetFramework1()
	version = "1"
	key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\1.0.3705\Version"
	GetNetFramework1 = GetNetFramework1_4(version, key)
End Function

Function GetNetFramework1_1()
	full_version = "Not installed"
	
	key1 = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\1.1.4322\Install"
	registry_version1 = ReadKey(key1)
	
	key2 = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\NET Framework Setup\NDP\v1.1.4322\Install"
	registry_version2 = ReadKey(key2)
	
	If Not registry_version1 = "" Then
		full_version = "Installed (v1.1.4322)"
	End If	
	
	If Not registry_version2 = "" Then
		full_version = "Installed (v1.1.4322)"
	End If
	
	GetNetFramework1_1 = full_version
End Function

Function GetNetFramework2()
	version = "2"
	key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727\Version"
	GetNetFramework2 = GetNetFramework1_4(version, key)
End Function

Function GetNetFramework3()
	version = "3"
	key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.0\Version"
	GetNetFramework3 = GetNetFramework1_4(version, key)
End Function

Function GetNetFramework3_5()
	version = "3.5"
	key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5\Version"
	GetNetFramework3_5 = GetNetFramework1_4(version, key)
End Function

Function GetNetFramework4()
	version = "4"
	key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\Version"
	release_key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\Release"
	
	release = ReadKey(release_key)
	
	If release = "" Then
		GetNetFramework4 = GetNetFramework1_4(version, key)
		Exit Function
	End If
	
	full_version = "Not installed"
	GetNetFramework4 = full_version
	
End Function

Function GetNetFramework4_5()
	version = "4.5"
	key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\Version"
	release_key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\Release"
	
	release = ReadKey(release_key)
	
	If Not release = "" Then
		GetNetFramework4_5 = GetNetFramework1_4(version, key)	
		Exit Function
	End If
	
	full_version = "Not installed"
	GetNetFramework4_5 = full_version
	
End Function

Function GetNetFramework4_6()
	version = "4.6"
	key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\Version"
	release_key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\Release"
	
	release = ReadKey(release_key)
	
	If Not release = "" Then
		GetNetFramework4_6 = GetNetFramework1_4(version, key)	
		Exit Function
	End If
	
	full_version = "Not installed"
	GetNetFramework4_6 = full_version
	
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Utility Functions
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function Add(ByRef array, item)
	i = UBound(array) + 1
	ReDim Preserve array(i)
	array(i) = item
End Function

Sub WriteFile(Data, FileName)
   Dim fso, File
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set File = fso.CreateTextFile(FileName, True)
   File.WriteLine(Data)
   File.Close
End Sub

Function WMIQuery(Query)
	Set WMIQuery = WMIQueryWithNamespace("winmgmts:root/cimv2", Query)
End Function

Function WMIQueryWithNamespace(WMINamespace, Query)
	Dim WMI, Collection
	Set WMI = GetObject(WMINamespace)
	Set Collection = WMI.ExecQuery(Query)
	Set WMIQueryWithNamespace = Collection
End Function

Sub OpenBrowser(URL)
	wsh.Run """" + URL + """"
End Sub

Function ReadKey(key)
	value = ""
	
	On Error Resume Next
	value = wsh.RegRead(key)
	On Error Goto 0
	
	ReadKey = value
End Function

Function Contains(ByVal needle, ByVal haystack)
	needle = LCase(needle)
	haystack = LCase(haystack)
	If(Instr(needle, haystack)) > 0 Then
		Contains = True
	Else
		Contains = False
	End If
End Function


