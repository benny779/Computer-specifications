#SingleInstance force
#NoTrayIcon
ComObjError(0)

Gui, -Border +AlwaysOnTop +Disabled
Gui, Font, s20 bold, Arial
Gui, Add, Text, w250 h35 Center, ...אנא המתן
Gui, show, AutoSize


OutputFile := "C:\temp\ziud_details.txt"
IfExist, % OutputFile
	FileDelete, % OutputFile


if !A_Is64bitOS
{
	MsgBox, 1572880, % A_ScriptName, תכונה זו פועלת רק במחשבי 64 ביט
	ExitApp
}

a := a_args[1]
if a in h,help,?,-?,/?
{
	a =
	(
Obtaining computer information.
By default the data is exported to the "ziud_details.txt" file.

Usage:		ziud_details ComputerName [gui]

Options:
	gui	Do not export the data to a file.
		Will be displayed in a message.
	)
	MsgBox % a
	ExitApp
}

if !computer := a_args[1]
	computer := "."

if !ComputerReturnPing(computer="."?"localhost":computer)
{
	FileAppend, ERROR, % OutputFile, UTF-8
	ExitApp
}

ComputerName	:= GetComputerName(computer)
Manufacturer	:= GetManufacturer(computer)
Model			:= GetModel(computer,Manufacturer)
Serial			:= GetBiosSerialNumber(computer)
Processor		:= GetProcessorName(computer)
Memory			:= GetPhysicalMemory(computer)
Disk			:= GetDiskSize(computer)
Os				:= GetOperatingSystem(computer)
;~ Warranty		:= GetWarrantyExpirationDate(Serial,Manufacturer)
MacAddress		:= GetMacAddress(computer)

TotalComputerDetails := ""
	. ComputerName "|"
	. Manufacturer "|"
	. Serial "|"
	. Processor "|"
	. Memory "|"
	. Disk "|"
	. Os "|"
	. MacAddress "|"
	;~ . Warranty "|"
	. Model


Gui, Destroy
if (a_args[2] = "gui")
	MsgBox % StrReplace(TotalComputerDetails,"|","`n")
else
	FileAppend, % TotalComputerDetails, % OutputFile, UTF-8
;~ MsgBox % TotalComputerDetails


ExitApp



ComputerReturnPing(computer:="localhost")
{
	for Item in ComObjGet("winmgmts:").ExecQuery("Select StatusCode From Win32_PingStatus where Address = '" computer "'")
		if Item.StatusCode = 0
			return true
		return false
}

GetComputerName(computer:=".")
{
	for Item in ComObjGet( "winmgmts:\\" computer ).ExecQuery("Select CSName from Win32_OperatingSystem")
		return Item.CSName
}

GetManufacturer(computer:=".")
{
	for Item in ComObjGet( "winmgmts:\\" computer ).ExecQuery("Select Manufacturer from Win32_BIOS")
		a := Item.Manufacturer
	if a = American Megatrends Inc.
		return ;"slimpc"
	if a contains ASUS
		return "ASUS"
	if a contains Dell
		return "Dell"
	if a contains INSYDE
		return "DT"
	if a = Hewlett-Packard
		return "HP"
	return a
}

GetModel(computer:=".", Manufacturer:="")
{
	if Manufacturer = Lenovo
	{
		for Item in ComObjGet( "winmgmts:\\" computer ).ExecQuery("Select SystemFamily from Win32_ComputerSystem")
			a := Item.SystemFamily
		
		if a contains V530-
			a := "V530"
		else if (SubStr(a, 1, 11) = "ThinkCentre")
			a := SubStr(a, 13)
	}
	else if Manufacturer = HP
	{
		for Item in ComObjGet( "winmgmts:\\" computer ).ExecQuery("Select Model from Win32_ComputerSystem")
			a := Item.Model
		
		if a contains HP ProDesk
			a := SubStr(a, 12, 6)
		else if a contains HP ProOne
			a := "AIO-" SubStr(a, 11, 6)
		else if a contains Retail System Model
			a := SubStr(a, (InStr(a, "Model")+5) )
		else if (SubStr(a, 1, 2) = "HP")
			a := SubStr(a, 4)
	}
	else if Manufacturer = Dell
	{
		for Item in ComObjGet( "winmgmts:\\" computer ).ExecQuery("Select Model from Win32_ComputerSystem")
			a := Item.Model
		
		if (SubStr(a, 1, 8) = "OptiPlex")
			a := SubStr(a, 10)
			
	}
	;~ else
	;~ {
		;~ for Item in ComObjGet( "winmgmts:\\" computer ).ExecQuery("Select * from Win32_ComputerSystem")
			;~ a := Item.Model " - " Item.SystemFamily
	;~ }

;~ OptiPlex 5080

	if a contains To be filled by
		a := ""
	
	return StrReplace(a, A_Space)
}

GetBiosSerialNumber(computer:=".") {
	for Item in ComObjGet( "winmgmts:\\" computer ).ExecQuery("Select SerialNumber from Win32_BIOS")
		return ( StrReplace(Item.SerialNumber,A_Space) ? Item.SerialNumber : "" )
}

GetProcessorName(computer:=".") {
	For Item in ComObjGet( "winmgmts:\\" computer ).ExecQuery("SELECT Family FROM Win32_Processor")
		if Item.Family = 198
			return "Corei7"
		else if Item.Family = 205
			return "Corei5"
		else if Item.Family = 206
			return "i3"
		else if Item.Family = 207
			return "Corei9"
}

GetPhysicalMemory(computer:=".") {
	for Item in ComObjGet( "winmgmts:\\" computer ).ExecQuery("Select Capacity from Win32_PhysicalMemory")
		TotalMemory += Item.Capacity
	return Floor( ( TotalMemory // (1024 * 1024) ) / 1000 ) * 1000
}

GetDiskSize(computer:=".") {
	For Item in ComObjGet( "winmgmts:\\" computer ).ExecQuery("Select * FROM Win32_LogicalDiskToPartition")
	if ( RegExReplace( Item.Dependent , "^.*?""|"".*") = "c:" )
		a := RegExReplace( Item.Antecedent , "^.*?""|"".*")

	For Item in ComObjGet( "winmgmts:\\" computer ).ExecQuery("Select * FROM Win32_DiskDriveToDiskPartition")
		if ( RegExReplace( Item.Dependent , "^.*?""|"".*") = a )
			b := StrReplace( RegExReplace( Item.Antecedent , "^.*?""|"".*") , "\\" , "\" )

	for Item in ComObjGet( "winmgmts:\\" computer ).ExecQuery("Select * from Win32_DiskDrive")
		if Item.Name = b
		{
			DiskSise := Substr(Item.Size, 1, -9)
			Serial := Item.SerialNumber
		}
	
	for Item in ComObjGet("winmgmts:\\" computer "\root\Microsoft\Windows\Storage").ExecQuery("SELECT MediaType,BusType FROM MSFT_PhysicalDisk WHERE SerialNumber='" Serial "'")
	{
		DiskType := Item.MediaType
		DiskBusType := Item.BusType
	}

	return ( DiskSise = 1000 ? "1TB" : DiskSise ) . ( DiskBusType=17 ? "-NVMe" : DiskType=4 ? "-SSD" : "" )
}

GetOperatingSystem(computer:=".") {
	for Item in ComObjGet( "winmgmts:\\" computer ).ExecQuery("Select OSArchitecture,Version from Win32_OperatingSystem")
	{
		bit := SubStr(Item.OSArchitecture, 1, 2)
		a := StrSplit(Item.Version, ".")[1]
		if a = 10
			OsVer := "WIN10x" bit
		if a = 6
			OsVer := bit=32 ? "windows 7" : "WIN7x64"
		if a = 5
			OsVer := "WinXP Pro"
		return OsVer
	}
	RegRead, VAR, \\%computer%\HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion, ProductName
	if VAR = Microsoft Windows XP
		return "WinXP Pro"
}

GetMacAddress(computer:=".") {
	for Item in ComObjGet( "winmgmts:\\" computer ).ExecQuery("Select MACAddress from Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
	return StrReplace(Item.MACAddress,":")
}

GetWarrantyExpirationDate(serial,manufacturer)
{
	if manufacturer = Lenovo
	{
		XML := "xml=<wiInputForm source='ibase'><id>LSC3</id><pw>IBA4LSC3</pw><product></product><serial>" serial
		. "</serial><wiOptions><machine/><parts/><service/><upma/><entitle/></wiOptions></wiInputForm>"
		
		strMessage := XML
		objHTTP := ComObjCreate("Msxml2.XMLHTTP")
		objHTTP.open("post", "https://ibase.lenovo.com/POIRequest.aspx", False)
		objHTTP.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
		objHTTP.send(strMessage)
		xmlReturn := objHTTP.responseText
		
		objXml := ComObjCreate("MSXML2.DOMDocument.6.0")
		objXml.async := false
		objXml.loadXML(xmlReturn)
		if objXml.parseError.errorCode
		{
			MsgBox, 48, XML Load Error, % "Unable to load XML data.`n`nError: " objXml.parseError.errorCode "`nReason: " objXml.parseError.reason
			return
		}
		if ExpirationDate := objXml.selectSingleNode("wiOutputForm/warrantyInfo/serviceInfo/wed").text
			ExpirationDate := StrSplit(ExpirationDate, "-")
		else
			ExpirationDate := ""

		objRelease(objHTTP)
		objRelease(objXml)
		return ExpirationDate ? ( ExpirationDate[3] "/" ExpirationDate[2] "/" ExpirationDate[1] ) : ""
	}
}


esc::ExitApp


