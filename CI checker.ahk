#SingleInstance Force ; Force SingleInstance of this application
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


;URL Encoder is: https://www.urlencoder.org

;change also the value itself of versionMAIN variable because it's being used also explicitly
version = AssetManager CI Checker (v1.6)
versionMAIN = AssetManager CI Checker (ver 1.6)
timestamp = 22nd of May 2019, 06:52 p.m.

;check if variables X and Y (for position of the window) are empty, if yes, set the value normally
if (X = "") 
{
	X = 750
}

if (Y = "") 
{
	Y = 50
}

Gui, Add, Edit, x11 y22 w350 h20 vHostnameInput, 
Gui, Add, Text, x11 y2 w350 h20 , Hostname / SID / SGER to check:
Gui, Add, Button, x11 y50 w350 h30 gCheckTheCI, Check
Gui, Add, Text, x55 y90 w350 h50 Center, Created by`nAntonio Raffaele Iannaccone
Gui, Add, Button, x330 y87 w30 h30 gGetHelp, ?
Gui, Add, Button, x110 y657 w140 h35 gCopyResultsToClipboard, Copy results to clipboard
Gui, Add, Progress, x11 y125 w350 h20 -Smooth vProgressBar,0
Gui, Add, DropDownList, x11 y93 vDropDownSelection Center, General information|AutoDiscovery data|Interfaces
GuiControl,, ProgressBar, 0

;~ Gui, Font, s11, Calibri

Gui, Add, ListView, gListViewControls x11 y150 w350 h500 NoSortHdr, Name of the property|Value
GuiControl, Hide, ListView

;selects default value in DropDownList
GuiControl, ChooseString, DropDownSelection, General information

Gui, Show, x%X% y%Y% h150 w373, %versionMAIN%
GuiControl,, ProgressBar, 30

;Check if proxy is enabled or not
RegRead, ProxyStatus, HKCU,SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings, ProxyEnable
if (ProxyStatus != "0")
	{
	;Check if the user is running the latest version
	;The pastebin data are under the user of gorneisoratdzimejldotkom
		try 
		{
			URLDownloadToFile, https://pastebin.com/raw/u22gSDbX, %A_MyDocuments%\AM_CI_Checker_Temporary.txt
		}
		catch
		{
			return	
		}
	
	
	GuiControl,, ProgressBar, 60
		Loop, read, %A_MyDocuments%\AM_CI_Checker_Temporary.txt
		{
			Found = 
			ActualLine = %A_LoopReadLine%
			GuiControl,, ProgressBar, 90
			Found := RegExMatch(ActualLine, "Actual version is this: (.+)", testVar)
			if (Found)
			{
				GuiControl,, ProgressBar, 100
				GuiControl,, ProgressBar, 0
	
				TempVar = %testVar1%
			
				if (testVar1 != versionMAIN)
					MsgBox, 16, %version%, You are not using the latest version!`n`nPlease, update to the latest version by contacting me, as the creator of this application (Antonio Raffaele Iannaccone), at antonio-raffaele.iannaccone@company.com or at +421 55 785 4553`nI will then send you the latest application.
				;~ else MsgBox, All goood.
			}
		}

	FileDelete, %A_MyDocuments%\AM_CI_Checker_Temporary.txt
}

GuiControl,, ProgressBar, 0
return


;Encode the URL into URL Encoded text
uriEncode(str) {
   f = %A_FormatInteger%
   SetFormat, Integer, Hex
   If RegExMatch(str, "^\w+:/{0,2}", pr)
      StringTrimLeft, str, str, StrLen(pr)
   StringReplace, str, str, `%, `%25, All
   Loop
      If RegExMatch(str, "i)[^\w\.~%/"""":?=!--']", char)
         StringReplace, str, str, %char%, % "%" . SubStr(Asc(char),3), All
      Else Break
   SetFormat, Integer, %f%
   Return, pr . str
}


CheckTheCI:
WinGetPos X, Y, Width, Height, %versionMAIN%
Gui, Show, x%X% y%Y% h150 w373, %versionMAIN%
;~ GuiControl, Hide, ListView
LV_Delete()
GuiControl,, ProgressBar, 20
GuiControlGet, HostnameInput,, HostnameInput
GuiControlGet, DropDownSelection,, DropDownSelection

HostnameInput := RegExReplace(HostnameInput, "`t","")
HostnameInput := RegExReplace(HostnameInput, """","")
HostnameInput := RegExReplace(HostnameInput, "`n","")
HostnameInput := RegExReplace(HostnameInput, "        $","")
HostnameInput := RegExReplace(HostnameInput, "^        ","")
HostnameInput := RegExReplace(HostnameInput, "       $","")
HostnameInput := RegExReplace(HostnameInput, "^       ","")
HostnameInput := RegExReplace(HostnameInput, "      $","")
HostnameInput := RegExReplace(HostnameInput, "^      ","")
HostnameInput := RegExReplace(HostnameInput, "     $","")
HostnameInput := RegExReplace(HostnameInput, "^     ","")
HostnameInput := RegExReplace(HostnameInput, "    $","")
HostnameInput := RegExReplace(HostnameInput, "^    ","")
HostnameInput := RegExReplace(HostnameInput, "   $","")
HostnameInput := RegExReplace(HostnameInput, "^   ","")
HostnameInput := RegExReplace(HostnameInput, "  $","")
HostnameInput := RegExReplace(HostnameInput, "^  ","")
HostnameInput := RegExReplace(HostnameInput, " $","")
HostnameInput := RegExReplace(HostnameInput, "^ ","")

HostnameInputMAIN = %HostnameInput%

if (HostnameInput = "")
{
	WinGetPos X, Y, Width, Height, %versionMAIN%
	;~ GuiControl, Hide, ResultsTextBox
	GuiControl, Hide, ListView
	LV_Delete()
	GuiControl,, ProgressBar, 0
	Gui, Show, x%X% y%Y% h150 w373, %versionMAIN%
	MsgBox, 16, %version%, You did not write anything into the search column.
	return
}

checkedWithOneOption = 0

;Create WinHTTPRequest object and log into AssetManager API
WinHTTP := ComObjCreate("WinHTTP.WinHttpRequest.5.1")
WinHTTP.Open("GET", "http://UrlCanNotBeDisclosed.com", 0)
WinHTTP.SetRequestHeader("Content-Type", "application/json")
Body := "{''}"

try
	{
		WinHTTP.Send(Body)
	}
catch
	{
		try
			{
				WinHTTP.Open("GET", "http://UrlCanNotBeDisclosed.com", 0)
				WinHTTP.SetRequestHeader("Content-Type", "application/json")
				Body := "{''}"
				WinHTTP.Send(Body)
				
				checkedWithOneOption = 1
				
			}
			catch
			{
				GuiControl,, ProgressBar, 0
				MsgBox, 16, %version%, An error has occured probably because of one of these reasons:`n1.) You are not connected to internal company network`n2.) Your proxy settings do not allow connection to AssetManager`n3.) You are not connected to a functioning internet connection
				return
			}
	}

Result := WinHTTP.ResponseText
Status := WinHTTP.Status

GuiControl,, ProgressBar, 45

;check if AutoDiscovery is selected in the DropDownList and then continue accordingly
if (DropDownSelection = "General information")
{
	goto, CheckGeneralInfo
	return
}
else if (DropDownSelection = "AutoDiscovery data")
{
	goto, CheckAutoDiscovery
	return
}
else if (DropDownSelection = "Interfaces")
{
	goto, CheckInterfaces
	return
}

;Show a message box if the selection from DropDownList does not match anything
GuiControl,, ProgressBar, 0
MsgBox, 16, %version%, An error occured.`n`nSeems like your selection from the drop-down list did not match any pre-defined possibilities.`nPlease, try to select your option from the drop-down list again and try it again.
return


CheckGeneralInfo:
RegExMatch(HostnameInput, "^S[1-9]", matchedSGERbyRegEx)

if (matchedSGERbyRegEx != "")
	{
		GuiControl,, ProgressBar, 60
		Result =
		
		
		if (checkedWithOneOption = 1)
		{
			urlFirstPart = CanNotBeDisclosed.com			
		}
		else
		{
			urlFirstPart = CanNotBeDisclosed.com
		}
		
		;Initiate URL and Encode it to URL standards
		url = http://%urlFirstPart%/service/api/GetLogSysInfoEx?whereclause=AssetTag = '%HostnameInput%'
		url := UriEncode(url)		
		WinHTTP.Open("GET", url, 0)
		WinHTTP.SetRequestHeader("Content-Type", "application/json")
		

		try
		{
			WinHTTP.Send(Body)
		}
		catch
		{
			GuiControl,, ProgressBar, 0
			MsgBox, 16, %version%, An error has occured probably because of one of these reasons:`n1.) You are not connected to internal company network`n2.) Your proxy settings do not allow connection to AssetManager`n3.) You are not connected to a functioning internet connection
			return
		}
		
		
		Result := WinHTTP.ResponseText
		Status := WinHTTP.Status

	}
else
	{
		GuiControl,, ProgressBar, 60
		
		Result =
		
		if (checkedWithOneOption = 1)
		{
			urlFirstPart = CanNotBeDisclosed.com			
		}
		else
		{
			urlFirstPart = CanNotBeDisclosed.com
		}
		
		url = http://%urlFirstPart%/service/api/GetLogSysInfoEx?whereclause=AddSysName LIKE '%HostnameInput%'
		url := UriEncode(url)
		
		WinHTTP.Open("GET", url, 0)
		WinHTTP.SetRequestHeader("Content-Type", "application/json")
		
		try
		{
			WinHTTP.Send(Body)
		}
		catch
		{
			GuiControl,, ProgressBar, 0
			MsgBox, 16, %version%, An error has occured probably because of one of these reasons:`n1.) You are not connected to internal company network`n2.) Your proxy settings do not allow connection to AssetManager`n3.) You are not connected to a functioning internet connection
			return
		}
		
		Result := WinHTTP.ResponseText
		Status := WinHTTP.Status
		
	
		Found := RegExMatch(result,"U)<Error_Desc>", Error_Desc)
		
		if (Found)
			{	
					GuiControl,, ProgressBar, 60
					
					Result =
					
					if (checkedWithOneOption = 1)
					{
						urlFirstPart = CanNotBeDisclosed.com		
					}
					else
					{
						urlFirstPart = CanNotBeDisclosed.com
					}
					
					
					url = http://%urlFirstPart%/service/api/GetLogSysInfoEx?whereclause=Portfolio.name LIKE '%HostnameInput%'
					url := UriEncode(url)
					
					WinHTTP.Open("GET", url, 0)
					WinHTTP.SetRequestHeader("Content-Type", "application/json")
					
					try
					{
						WinHTTP.Send(Body)
					}
					catch
					{
						GuiControl,, ProgressBar, 0
						MsgBox, 16, %version%, An error has occured probably because of one of these reasons:`n1.) You are not connected to internal company network`n2.) Your proxy settings do not allow connection to AssetManager`n3.) You are not connected to a functioning internet connection
						return
					}
					
					Result := WinHTTP.ResponseText
					Status := WinHTTP.Status
					
					Found =
			}
		
		
	}
	

GuiControl,, ProgressBar, 70

;Check if the API was not able to find the CI, then show a Message Box and return
FoundNew := RegExMatch(result,"U)<Error_Desc>(.*?)</Error_Desc>", Error_Desc)
if (FoundNew)
	{	WinGetPos X, Y, Width, Height, %versionMAIN%
		;~ GuiControl, Hide, ResultsTextBox
		GuiControl, Hide, ListView
		GuiControl,, ProgressBar, 0
		Gui, Show, x%X% y%Y% h150 w373, %versionMAIN%
		
		NotConnectedToAMvariable = %Error_Desc1%
		
		if (NotConnectedToAMvariable = "Not Connected to AM!")
			{
				MsgBox, 16, %version%, I was not able to find anything in the AssetManager with the value "%HostnameInputMAIN%".`n`nThe error is specifically this: %Error_Desc1%`n`nIt seems like the AssetManager's API interface is down. `nPlease, wait a couple of minutes and try again.`n`nThis happens sometimes since the backbone of AssetManager itself can get overloaded and the API interface can crash.`nNo worries, just wait.
			}
		else
			{
				MsgBox, 16, %version%, I was not able to find anything in the AssetManager with the value "%HostnameInputMAIN%".`n`nThe error is specifically this: %Error_Desc1%`n`nAre you sure that you used the correct case and CI name/SGER in your query? (Upper-case letters, small-case letters, spaces, etc.)`n`nExamples of proper usage:`nS20154997`nDCPWIN20154997`nAF6HL234
			}

		return
	}


;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<OSBuildVersion>(.*?)</OSBuildVersion>", OSBuildVersion)
if (Found)
	{	 
		OSBuildVersion = %OSBuildVersion1%
		;MsgBox, % "OSBuildVersion: " OSBuildVersion1
	}
else
	{
		OSBuildVersion = N/A
		;MsgBox, % "OSBuildVersion: " OSBuildVersion
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<OSDirectory>(.*?)</OSDirectory>", OSDirectory)
if (Found)
	{	 
		OSDirectory = %OSDirectory1%
		;MsgBox, % "OSBuildVersion: " OSDirectory1
	}
else
	{
		OSDirectory = N/A
		;MsgBox, % "OSBuildVersion: " OSDirectory
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<AvailabilitySLAname>(.*?)</AvailabilitySLAname>", AvailabilitySLAname)
if (Found)
	{	 
		AvailabilitySLAname = %AvailabilitySLAname1%
		;MsgBox, % "AvailabilitySLAname1: " AvailabilitySLAname1
	}
else
	{
		AvailabilitySLAname = N/A
		;MsgBox, % "AvailabilitySLAname1: " AvailabilitySLAname
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<AvailabilitySLAcode>(.*?)</AvailabilitySLAcode>", AvailabilitySLAcode)
if (Found)
	{	 
		AvailabilitySLAcode = %AvailabilitySLAcode1%
		;MsgBox, % "AvailabilitySLAcode1: " AvailabilitySLAcode1
	}
else
	{
		AvailabilitySLAcode = N/A
		;MsgBox, % "AvailabilitySLAcode1: " AvailabilitySLAcode
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<StorageType>(.*?)</StorageType>", StorageType)
if (Found)
	{	 
		StorageType = %StorageType1%
		;MsgBox, % "StorageType1: " StorageType1
	}
else
	{
		StorageType = N/A
		;MsgBox, % "StorageType: " StorageType
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<SecuritySet>(.*?)</SecuritySet>", SecuritySet)
if (Found)
	{	 
		SecuritySet = %SecuritySet1%
		;MsgBox, % "SecuritySet: " SecuritySet1
	}
else
	{
		SecuritySet = N/A
		;MsgBox, % "SecuritySet: " SecuritySet
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<SLAname>(.*?)</SLAname>", SLAname)
if (Found)
	{	 
		SLAname = %SLAname1%
		;MsgBox, % "SLAname: " SLAname1
	}
else
	{
		SLAname = N/A
		;MsgBox, % "SLAname: " SLAname
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<lSlices>(.*?)</lSlices>", lSlices)
if (Found)
	{	 
		lSlices = %lSlices1%
		;MsgBox, % "lSlices: " lSlices1
	}
else
	{
		lSlices = N/A
		;MsgBox, % "lSlices: " lSlices
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<companyOperatingMode>(.*?)</companyOperatingMode>", companyOperatingMode)
if (Found)
	{	 
		companyOperatingMode = %companyOperatingMode1%
		;MsgBox, % "companyOperatingMode: " companyOperatingMode1
	}
else
	{
		companyOperatingMode = N/A
		;MsgBox, % "companyOperatingMode: " companyOperatingMode
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<iNumberLogProc>(.*?)</iNumberLogProc>", iNumberLogProc)
if (Found)
	{	 
		iNumberLogProc = %iNumberLogProc1%
		;MsgBox, % "iNumberLogProc: " iNumberLogProc1
	}
else
	{
		iNumberLogProc = N/A
		;MsgBox, % "iNumberLogProc: " iNumberLogProc
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<iNumberCoresPerCPU>(.*?)</iNumberCoresPerCPU>", iNumberCoresPerCPU)
if (Found)
	{	 
		iNumberCoresPerCPU = %iNumberCoresPerCPU1%
		;MsgBox, % "iNumberCoresPerCPU: " iNumberCoresPerCPU1
	}
else
	{
		iNumberCoresPerCPU = N/A
		;MsgBox, % "iNumberCoresPerCPU: " iNumberCoresPerCPU
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<OlaOperationCategory>(.*?)</OlaOperationCategory>", OlaOperationCategory)
if (Found)
	{	 
		OlaOperationCategory = %OlaOperationCategory1%
		;MsgBox, % "OlaOperationCategory: " OlaOperationCategory1
	}
else
	{
		OlaOperationCategory = N/A
		;MsgBox, % "OlaOperationCategory: " OlaOperationCategory
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<OlaServiceClass>(.*?)</OlaServiceClass>", OlaServiceClass)
if (Found)
	{	 
		OlaServiceClass = %OlaServiceClass1%
		;MsgBox, % "OlaServiceClass: " OlaServiceClass1
	}
else
	{
		OlaServiceClass = N/A
		;MsgBox, % "OlaServiceClass: " OlaServiceClass
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<Nature>(.*?)</Nature>", Nature)
if (Found)
	{	 
		Nature = %Nature1%
		;MsgBox, % "Nature: " Nature1
	}
else
	{
		Nature = N/A
		;MsgBox, % "Nature: " Nature
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<ExternalSystem>(.*?)</ExternalSystem>", ExternalSystem)
if (Found)
	{	 
		ExternalSystem = %ExternalSystem1%
		;MsgBox, % "ExternalSystem: " ExternalSystem1
	}
else
	{
		ExternalSystem = N/A
		;MsgBox, % "ExternalSystem: " ExternalSystem
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<ExternalID>(.*?)</ExternalID>", ExternalID)
if (Found)
	{	 
		ExternalID = %ExternalID1%
		;MsgBox, % "ExternalID: " ExternalID1
	}
else
	{
		ExternalID = N/A
		;MsgBox, % "ExternalID: " ExternalID
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<Assettag>(.*?)</Assettag>", Assettag)
if (Found)
	{	 
		Assettag = %Assettag1%
		;MsgBox, % "Assettag: " Assettag1
	}
else
	{
		Assettag = N/A
		;MsgBox, % "Assettag: " Assettag
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<Parent_Assettag>(.*?)</Parent_Assettag>", Parent_Assettag)
if (Found)
	{	 
		Parent_Assettag = %Parent_Assettag1%
		;MsgBox, % "Parent_Assettag: " Parent_Assettag1
	}
else
	{
		Parent_Assettag = N/A
		;MsgBox, % "Parent_Assettag: " Parent_Assettag
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<Name>(.*?)</Name>", Name)
if (Found)
	{	 
		Name = %Name1%
		;MsgBox, % "Name: " Name1
	}
else
	{
		Name = N/A
		;MsgBox, % "Name: " Name
	}


;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<AssignmentGroup>(.*?)</AssignmentGroup>", AssignmentGroup)
if (Found)
	{	 
		AssignmentGroup = %AssignmentGroup1%
		;MsgBox, % "AssignmentGroup: " AssignmentGroup1
	}
else
	{
		AssignmentGroup = N/A
		;MsgBox, % "AssignmentGroup: " AssignmentGroup
	}


;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<IncidentAG>(.*?)</IncidentAG>", IncidentAG)
if (Found)
	{	 
		IncidentAG = %IncidentAG1%
		;MsgBox, % "IncidentAG: " IncidentAG1
	}
else
	{
		IncidentAG = N/A
		;MsgBox, % "IncidentAG: " IncidentAG
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<CO_CC>(.*?)</CO_CC>", CO_CC)
if (Found)
	{	 
		CO_CC = %CO_CC1%
		;MsgBox, % "CO_CC: " CO_CC1
	}
else
	{
		CO_CC = N/A
		;MsgBox, % "CO_CC: " CO_CC
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<OlaClassSystem>(.*?)</OlaClassSystem>", OlaClassSystem)
if (Found)
	{	 
		OlaClassSystem = %OlaClassSystem1%
		;MsgBox, % "OlaClassSystem: " OlaClassSystem1
	}
else
	{
		OlaClassSystem = N/A
		;MsgBox, % "OlaClassSystem: " OlaClassSystem
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<Status>(.*?)</Status>", Status)
if (Found)
	{	 
		Status = %Status1%
		;MsgBox, % "Status: " Status1
	}
else
	{
		Status = N/A
		;MsgBox, % "Status: " Status
	}


;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<Priority>(.*?)</Priority>", Priority)
if (Found)
	{	 
		Priority = %Priority1%
		;MsgBox, % "Priority: " Priority1
	}
else
	{
		Priority = N/A
		;MsgBox, % "Priority: " Priority
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<SC_Location_ID>(.*?)</SC_Location_ID>", SC_Location_ID)
if (Found)
	{	 
		SC_Location_ID = %SC_Location_ID1%
		;MsgBox, % "SC_Location_ID: " SC_Location_ID1
	}
else
	{
		SC_Location_ID = N/A
		;MsgBox, % "SC_Location_ID: " SC_Location_ID
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<Location_Code>(.*?)</Location_Code>", Location_Code)
if (Found)
	{	 
		Location_Code = %Location_Code1%
		;MsgBox, % "Location_Code: " Location_Code1
	}
else
	{
		Location_Code = N/A
		;MsgBox, % "Location_Code: " Location_Code
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<Usage>(.*?)</Usage>", Usage)
if (Found)
	{	 
		Usage = %Usage1%
		;MsgBox, % "Location_Code: " Usage1
	}
else
	{
		Usage = N/A
		;MsgBox, % "Usage: " Usage
	}
	
GuiControl,, ProgressBar, 90

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<dtInvent>(.*?)</dtInvent>", dtInvent)
if (Found)
	{	 
		dtInvent = %dtInvent1%
		dtInvent := RegExReplace(dtInvent, ":", ".")
		;~ MsgBox, % "dtInvent: " dtInvent
	}
else
	{
		dtInvent = N/A
		;MsgBox, % "dtInvent: " dtInvent
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<ControlCenter>(.*?)</ControlCenter>", ControlCenter)
if (Found)
	{	 
		ControlCenter = %ControlCenter1%
		;MsgBox, % "ControlCenter: " ControlCenter1
	}
else
	{
		ControlCenter = N/A
		;MsgBox, % "ControlCenter: " ControlCenter
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<ControlCenter2>(.*?)</ControlCenter2>", ControlCenter2)
if (Found)
	{	 
		ControlCenter2 = %ControlCenter21%
		;MsgBox, % "ControlCenter2: " ControlCenter21
	}
else
	{
		ControlCenter2 = N/A
		;MsgBox, % "ControlCenter2: " ControlCenter2
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<CPUType>(.*?)</CPUType>", CPUType)
if (Found)
	{	 
		CPUType = %CPUType1%
		;MsgBox, % "CPUType: " CPUType1
	}
else
	{
		CPUType = N/A
		;MsgBox, % "CPUType: " CPUType
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<lCPUNumber>(.*?)</lCPUNumber>", lCPUNumber)
if (Found)
	{	 
		lCPUNumber = %lCPUNumber1%
		;MsgBox, % "lCPUNumber: " lCPUNumber1
	}
else
	{
		lCPUNumber = N/A
		;MsgBox, % "lCPUNumber: " lCPUNumber
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<lCPUSpeedMHz>(.*?)</lCPUSpeedMHz>", lCPUSpeedMHz)
if (Found)
	{	 
		lCPUSpeedMHz = %lCPUSpeedMHz1%
		;MsgBox, % "lCPUSpeedMHz: " lCPUSpeedMHz1
	}
else
	{
		lCPUSpeedMHz = N/A
		;MsgBox, % "lCPUSpeedMHz: " lCPUSpeedMHz
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<lMemorySizeMb>(.*?)</lMemorySizeMb>", lMemorySizeMb)
if (Found)
	{	 
		lMemorySizeMb = %lMemorySizeMb1%
		;MsgBox, % "lMemorySizeMb: " lMemorySizeMb1
	}
else
	{
		lMemorySizeMb = N/A
		;MsgBox, % "lMemorySizeMb: " lMemorySizeMb
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<OperatingSystem>(.*?)</OperatingSystem>", OperatingSystem)
if (Found)
	{	 
		OperatingSystem = %OperatingSystem1%
		;MsgBox, % "OperatingSystem: " OperatingSystem1
	}
else
	{
		OperatingSystem = N/A
		;MsgBox, % "OperatingSystem: " OperatingSystem
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<ComputerType>(.*?)</ComputerType>", ComputerType)
if (Found)
	{	 
		ComputerType = %ComputerType1%
		;MsgBox, % "ComputerType: " ComputerType1
	}
else
	{
		ComputerType = N/A
		;MsgBox, % "ComputerType: " ComputerType
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<OrgUnit>(.*?)</OrgUnit>", OrgUnit)
if (Found)
	{	 
		OrgUnit = %OrgUnit1%
		;MsgBox, % "OrgUnit: " OrgUnit1
	}
else
	{
		OrgUnit = N/A
		;MsgBox, % "OrgUnit: " OrgUnit
	}

;;;;;;;;;;;;;;;;;;;;;
;~ This is not working/old RegEx: Found := RegExMatch(result,"U)<Remarks>([^;]*)<\/Remarks>", Remarks)
;~ Found := RegExMatch(result,"U)<Remarks>([\s\S]*)<\/Remarks>", Remarks)
;~ AE6HG031 nefunguje dobre Org.Remarks - uz funguje - OPRAVENE
Found := RegExMatch(result,"U)<Remarks>([\s\S]*)<\/Remarks>", Remarks)
if (Found)
	{	 
		Remarks = `n%Remarks1%
		;MsgBox, % "Remarks: " Remarks1
	}
else
	{
		Remarks = N/A
		;MsgBox, % "Remarks: " Remarks
	}


;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<Security_Unit>(.*?)</Security_Unit>", Security_Unit)
if (Found)
	{	 
		Security_Unit = %Security_Unit1%
		;MsgBox, % "Security_Unit: " Security_Unit1
	}
else
	{
		Security_Unit = N/A
		;MsgBox, % "Security_Unit: " Security_Unit
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<SoxRelevant>(.*?)</SoxRelevant>", SoxRelevant)
if (Found)
	{	 
		SoxRelevant = %SoxRelevant1%
		;MsgBox, % "SoxRelevant: " SoxRelevant1
	}
else
	{
		SoxRelevant = N/A
		;MsgBox, % "SoxRelevant: " SoxRelevant
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<EuroSoxRelevant>(.*?)</EuroSoxRelevant>", EuroSoxRelevant)
if (Found)
	{	 
		EuroSoxRelevant = %EuroSoxRelevant1%
		;MsgBox, % "EuroSoxRelevant: " EuroSoxRelevant1
	}
else
	{
		EuroSoxRelevant = N/A
		;MsgBox, % "EuroSoxRelevant: " EuroSoxRelevant
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<Virtualization>(.*?)</Virtualization>", Virtualization)
if (Found)
	{	 
		Virtualization = %Virtualization1%
		;MsgBox, % "Virtualization: " Virtualization1
	}
else
	{
		Virtualization = N/A
		;MsgBox, % "Virtualization: " Virtualization
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<HyperThreading>(.*?)</HyperThreading>", HyperThreading)
if (Found)
	{	 
		HyperThreading = %HyperThreading1%
		;MsgBox, % "HyperThreading: " HyperThreading1
	}
else
	{
		HyperThreading = N/A
		;MsgBox, % "HyperThreading: " HyperThreading
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<AutoDiscStatus>(.*?)</AutoDiscStatus>", AutoDiscStatus)
if (Found)
	{	 
		AutoDiscStatus = %AutoDiscStatus1%
		;MsgBox, % "AutoDiscStatus: " AutoDiscStatus1
	}
else
	{
		AutoDiscStatus = N/A
		;MsgBox, % "AutoDiscStatus: " AutoDiscStatus
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<AddSystemName>(.*?)</AddSystemName>", AddSystemName)
if (Found)
	{	 
		AddSystemName = %AddSystemName1%
		;MsgBox, % "AddSystemName: " AddSystemName1
	}
else
	{
		AddSystemName = N/A
		;MsgBox, % "AddSystemName: " AddSystemName
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<ServicePack>(.*?)</ServicePack>", ServicePack)
if (Found)
	{	 
		ServicePack = %ServicePack1%
		;MsgBox, % "ServicePack: " ServicePack1
	}
else
	{
		ServicePack = N/A
		;MsgBox, % "ServicePack: " ServicePack
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<IBANo>(.*?)</IBANo>", IBANo)
if (Found)
	{	 
		IBANo = %IBANo1%
		;MsgBox, % "IBANo: " IBANo1
	}
else
	{
		IBANo = N/A
		;MsgBox, % "IBANo: " IBANo
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<Criticality>(.*?)</Criticality>", Criticality)
if (Found)
	{	 
		Criticality = %Criticality1%
		;MsgBox, % "Criticality: " Criticality1
	}
else
	{
		Criticality = N/A
		;MsgBox, % "Criticality: " Criticality
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<Customer>(.*?)</Customer>", Customer)
if (Found)
	{	 
		Customer = %Customer1%
		;MsgBox, % "Customer: " Customer1
	}
else
	{
		Customer = N/A
		;MsgBox, % "Customer: " Customer
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<CustomerLink>(.*?)</CustomerLink>", CustomerLink)
if (Found)
	{	 
		CustomerLink = %CustomerLink1%
		;MsgBox, % "CustomerLink: " CustomerLink1
	}
else
	{
		CustomerLink = N/A
		;MsgBox, % "CustomerLink: " CustomerLink
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<CODescription>(.*?)</CODescription>", CODescription)
if (Found)
	{	 
		CODescription = %CODescription1%
		;MsgBox, % "CODescription: " CODescription1
	}
else
	{
		CODescription = N/A
		;MsgBox, % "CODescription: " CODescription
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<AdminLanIP>(.*?)</AdminLanIP>", AdminLanIP)
if (Found)
	{	 
		AdminLanIP = %AdminLanIP1%
		;MsgBox, % "AdminLanIP: " AdminLanIP1
	}
else
	{
		AdminLanIP = N/A
		;MsgBox, % "AdminLanIP: " AdminLanIP
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<Location>(.*?)</Location>", Location)
if (Found)
	{	 
		Location = %Location1%
		;MsgBox, % "Location: " Location1
	}
else
	{
		Location = N/A
		;MsgBox, % "Location: " Location
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<SC_Location_Name>(.*?)</SC_Location_Name>", SC_Location_Name)
if (Found)
	{	 
		SC_Location_Name = %SC_Location_Name1%
		;SC_Location_Name := StrReplace(StrReplace(StrReplace(SC_Location_Name, "`n", ""), "`v", ""), "`r", "")
		;MsgBox, % "SC_Location_Name: " SC_Location_Name1
	}
else
	{
		SC_Location_Name = N/A
		;MsgBox, % "SC_Location_Name: " SC_Location_Name
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<OSBackup>(.*?)</OSBackup>", OSBackup)
if (Found)
	{	 
		OSBackup = %OSBackup1%
		;MsgBox, % "OSBackup: " OSBackup1
	}
else
	{
		OSBackup = N/A
		;MsgBox, % "OSBackup: " OSBackup
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<fCPUNumber>(.*?)</fCPUNumber>", fCPUNumber)
if (Found)
	{	 
		fCPUNumber = %fCPUNumber1%
		;MsgBox, % "fCPUNumber: " fCPUNumber1
	}
else
	{
		fCPUNumber = N/A
		;MsgBox, % "fCPUNumber: " fCPUNumber
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<iTotalNumberofCores>(.*?)</iTotalNumberofCores>", iTotalNumberofCores)
if (Found)
	{	 
		iTotalNumberofCores = %iTotalNumberofCores1%
		;MsgBox, % "iTotalNumberofCores: " iTotalNumberofCores1
	}
else
	{
		iTotalNumberofCores = N/A
		;MsgBox, % "iTotalNumberofCores: " iTotalNumberofCores
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<ParentSerialNo>(.*?)</ParentSerialNo>", ParentSerialNo)
if (Found)
	{	 
		ParentSerialNo = %ParentSerialNo1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		ParentSerialNo = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<OPM>(.*?)</OPM>", OPM)
if (Found)
	{	 
		OPM = %OPM1%
		;MsgBox, % "OPM: " OPM1
	}
else
	{
		OPM = N/A
		;MsgBox, % "OPM: " OPM
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<CustomerCostCenter>(.*?)</CustomerCostCenter>", CustomerCC)
if (Found)
	{	 
		CustomerCC = %CustomerCC1%
		;MsgBox, % "CustomerCC: " CustomerCC1
	}
else
	{
		CustomerCC = N/A
		;MsgBox, % "CustomerCC: " CustomerCC
	}

GuiControl,, ProgressBar, 100

;~ MsgBox,,,%Remarks%

OutputResult = Hostname: %Name%`nSID (Second Name): %AddSystemName%`nSGER: %Assettag%`nParent AGER: %Parent_Assettag%`nStatus: %Status%`nVirtualization: %Virtualization%`nAG: %AssignmentGroup%`nIncident AG: %IncidentAG%`nCriticality: %Criticality%`nPriority: %Priority%`nUsage: %Usage%`nGCC Std.: %ControlCenter%`nGCC Non-Std.: %ControlCenter2%`nOperating Mode: %companyOperatingMode%`nComputer Type: %ComputerType%`nOS: %OperatingSystem%`nService Pack: %ServicePack%`nOS Build Version: %OSBuildVersion%`nOS Directory: %OSDirectory%`nNumber of CPUs: %fCPUNumber%`nNumber of Logical CPUs: %iNumberLogProc%`nNumber of Cores per CPU: %iNumberCoresPerCPU%`nCPU Type: %CPUType%`nCPU Speed (MHz): %lCPUSpeedMHz%`nSlices: %lSlices%`nMemory Size (MB): %lMemorySizeMb%`nOS Backup: %OSBackup%`nParent Serial Number: %ParentSerialNo%`nSLA: %SLAname%`nAvailability SLA: %AvailabilitySLAname%`nAvailability SLA code: %AvailabilitySLAcode%`nOLA: %OlaOperationCategory%`nOLA Service Class: %OlaServiceClass%`nOLA Class: %OlaClassSystem%`nCustomer: %Customer%`nCustomer Link: %CustomerLink%`nCustomer Description: %CODescription%`nNature: %Nature%`nStorage Type: %StorageType%`nSecurity Set: %SecuritySet%`nExternal System: %ExternalSystem%`nExternal ID: %ExternalID%`nCust. CostCenter: %CustomerCC%`nCust. Order / CC: %CO_CC%`nLocation Code: %Location_Code%`nInstallation Date: %dtInvent%`nLocation: %Location%`nSC Location: %SC_Location_Name%`nSC Location ID: %SC_Location_ID%`nOrg.Unit: %OrgUnit%`nSecurity Unit: %Security_Unit%`nSOX Relevant: %SoxRelevant%`nEURO SOX Relevant: %EuroSoxRelevant%`nHyper Threading: %HyperThreading%`nAuto-Discovery Target Status: %AutoDiscStatus%`nIBA Number: %IBANo%`nAdminLAN IP: %AdminLanIP%`nOPM: %OPM%`nOrg.Remarks: %Remarks%

;No. of CPUs (2): %lCPUNumber%`n
;Total number of CPU cores: %iTotalNumberofCores%`n
	
GuiControl,, ProgressBar, 0
WinGetPos X, Y, Width, Height, %versionMAIN%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;Populate ListView
LoopCount = 0

loop, parse, OutputResult, `n
{
	arry := StrSplit(A_LoopField, ": ")
	LV_Add("", arry[1], arry[2])
}


Gui, Show, x%X% y%Y% h700 w373, %versionMAIN%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;Disconnect after query
url = http://%urlFirstPart%/service/api/Disconnect
url := UriEncode(url)		
WinHTTP.Open("GET", url, 0)
WinHTTP.SetRequestHeader("Content-Type", "application/json")
Body := "{''}"

try
	{
		WinHTTP.Send(Body)
	}
catch
	{
		GuiControl,, ProgressBar, 0
		MsgBox, 16, %version%, An error has occured probably because of one of these reasons:`n1.) You are not connected to internal company network`n2.) Your proxy settings do not allow connection to AssetManager`n3.) You are not connected to a functioning internet connection
		return
	}

Result := WinHTTP.ResponseText
Status := WinHTTP.Status

;~ MsgBox, % Result
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;


url =
urlFirstPart = 

return

ListViewControls:
LV_GetText(name,A_EventInfo,1)
LV_GetText(value,A_EventInfo,2)
selectedColumnToClipboard = %name%: %value%
Tooltip, %name%: %value%
Clipboard = %selectedColumnToClipboard%
SetTimer, ToolTipOff, -1300
Return

ToolTipOff:
Tooltip
Return


CopyResultsToClipboard:
Clipboard = %OutputResult%
return

GetHelp:
MsgBox,, %version%, `t`t    >>>>>>>>> HELP <<<<<<<<<`n`n1.) You can choose if you want to search for general information from AssetManager, or, if you want to search for the AutoDiscovery data from AssetManager, or, if you want to search for the interfaces that belong to a specific CI. For this you need to select what you want to search for in the drop-down list.`n`nNOTICE: For the AutoDiscovery data the CI needs to have a working connection towards Audit Server X (via WinAudit or SiUX).`n`n2.) You can double click on a column to copy only one value to clipboard and/or to show the entire value if it's too long to be displayed in GUI.`n`n3.) You can click the "Copy to clipboard" button to copy all of the shown data to clipboard.`n`nIn case you'll experience any issues with this application, please, let me know by writing me an e-mail at antonio-raffaele.iannaccone@company.com`n`n`n`tCreated by Antonio Raffaele Iannaccone`nService Chain Operation Manager for Heineken Non-SAP Hosting`n`nBuild timestamp: %timestamp%
return


#IfWinActive, AssetManager CI Checker (ver 1.6)
Enter::
GoTo, CheckTheCI
return

NumpadEnter::
GoTo, CheckTheCI
return
#IfWinActive
return

CheckAutoDiscovery:
WinGetPos X, Y, Width, Height, %versionMAIN%
LV_Delete()

if (checkedWithOneOption = 1)
	{
		urlFirstPart = CanNotBeDisclosed.com			
	}
else
	{
		urlFirstPart = CanNotBeDisclosed.com
	}


;Initiate URL and Encode it to URL standards
url = http://%urlFirstPart%/service/api/GetAutoDiscoveryData?systemid=%HostnameInput%
url := UriEncode(url)		
WinHTTP.Open("GET", url, 0)

GuiControl,, ProgressBar, 50

WinHTTP.SetRequestHeader("Content-Type", "application/json")
Body := "{''}"
try
{
	WinHTTP.Send(Body)
}
catch
{
	GuiControl,, ProgressBar, 0
	MsgBox, 16, %version%, An error has occured probably because of one of these reasons:`n1.) You are not connected to internal company network`n2.) Your proxy settings do not allow connection to AssetManager`n3.) You are not connected to a functioning internet connection
	return
}
Result := WinHTTP.ResponseText
Status := WinHTTP.Status

;Try to search for SGER if it's not entered by the user and try to search for AutoDiscovery data again
FoundNewer := RegExMatch(result,"U)<Error_Desc>(.*?)</Error_Desc>", Error_Desc)
if (FoundNewer)
	{	
		NoRecordFoundVariable = %Error_Desc1%
		
		if (NoRecordFoundVariable = "No record found!")
			{
				
				RegExMatch(HostnameInput, "^S[1-9]", matchedSGERbyRegEx)
				
				if (matchedSGERbyRegEx != "")
					{
						GuiControl,, ProgressBar, 60
						Result =
						
						
						if (checkedWithOneOption = 1)
						{
							urlFirstPart = CanNotBeDisclosed.com		
						}
						else
						{
							urlFirstPart = CanNotBeDisclosed.com
						}
						
						;Initiate URL and Encode it to URL standards
						url = http://%urlFirstPart%/service/api/GetLogSysInfoEx?whereclause=AssetTag = '%HostnameInput%'
						url := UriEncode(url)		
						WinHTTP.Open("GET", url, 0)
						WinHTTP.SetRequestHeader("Content-Type", "application/json")
						
				
						try
						{
							WinHTTP.Send(Body)
						}
						catch
						{
							GuiControl,, ProgressBar, 0
							MsgBox, 16, %version%, An error has occured probably because of one of these reasons:`n1.) You are not connected to internal company network`n2.) Your proxy settings do not allow connection to AssetManager`n3.) You are not connected to a functioning internet connection
							return
						}
						
						
						Result := WinHTTP.ResponseText
						Status := WinHTTP.Status
				
					}
				else
					{
						GuiControl,, ProgressBar, 60
						
						Result =
						
						if (checkedWithOneOption = 1)
						{
							urlFirstPart = CanNotBeDisclosed.com		
						}
						else
						{
							urlFirstPart = CanNotBeDisclosed.com
						}
						
						url = http://%urlFirstPart%/service/api/GetLogSysInfoEx?whereclause=AddSysName LIKE '%HostnameInput%'
						url := UriEncode(url)
						
						WinHTTP.Open("GET", url, 0)
						WinHTTP.SetRequestHeader("Content-Type", "application/json")
						
						try
						{
							WinHTTP.Send(Body)
						}
						catch
						{
							GuiControl,, ProgressBar, 0
							MsgBox, 16, %version%, An error has occured probably because of one of these reasons:`n1.) You are not connected to internal company network`n2.) Your proxy settings do not allow connection to AssetManager`n3.) You are not connected to a functioning internet connection
							return
						}
						
						Result := WinHTTP.ResponseText
						Status := WinHTTP.Status
						
					
						Found := RegExMatch(result,"U)<Error_Desc>", Error_Desc)
						
						if (Found)
							{	
									GuiControl,, ProgressBar, 60
									
									Result =
									
									if (checkedWithOneOption = 1)
									{
										urlFirstPart = CanNotBeDisclosed.com		
									}
									else
									{
										urlFirstPart = CanNotBeDisclosed.com
									}
									
									
									url = http://%urlFirstPart%/service/api/GetLogSysInfoEx?whereclause=Portfolio.name LIKE '%HostnameInput%'
									url := UriEncode(url)
									
									WinHTTP.Open("GET", url, 0)
									WinHTTP.SetRequestHeader("Content-Type", "application/json")
									
									try
									{
										WinHTTP.Send(Body)
									}
									catch
									{
										GuiControl,, ProgressBar, 0
										MsgBox, 16, %version%, An error has occured probably because of one of these reasons:`n1.) You are not connected to internal company network`n2.) Your proxy settings do not allow connection to AssetManager`n3.) You are not connected to a functioning internet connection
										return
									}
									
									Result := WinHTTP.ResponseText
									Status := WinHTTP.Status
									
									Found =
							}
						
						
					}
			}
		;;;;;;;;;;;;;;;;;;;;;
		Found := RegExMatch(result,"U)<Assettag>(.*?)</Assettag>", Assettag)
		if (Found)
			{	 
				Assettag = %Assettag1%
				;MsgBox, % "Assettag: " Assettag1
			}
		else
			{
				Assettag = N/A
				;MsgBox, % "Assettag: " Assettag
			}
			
			
		if (checkedWithOneOption = 1)
			{
				urlFirstPart = CanNotBeDisclosed.com		
			}
		else
			{
				urlFirstPart = CanNotBeDisclosed.com
			}
			
			GuiControl,, ProgressBar, 75
		
		
		;Initiate URL and Encode it to URL standards
		url = http://%urlFirstPart%/service/api/GetAutoDiscoveryData?systemid=%Assettag%
		url := UriEncode(url)		
		WinHTTP.Open("GET", url, 0)
		WinHTTP.SetRequestHeader("Content-Type", "application/json")
		Body := "{''}"
		try
		{
			WinHTTP.Send(Body)
		}
		catch
		{
			GuiControl,, ProgressBar, 0
			MsgBox, 16, %version%, An error has occured probably because of one of these reasons:`n1.) You are not connected to internal company network`n2.) Your proxy settings do not allow connection to AssetManager`n3.) You are not connected to a functioning internet connection
			return
		}
		Result := WinHTTP.ResponseText
		Status := WinHTTP.Status
			
			

	}
	
FoundNewest := RegExMatch(result,"U)<Error_Desc>(.*?)</Error_Desc>", Error_Desc)
if (FoundNewest)
	{
		WinGetPos X, Y, Width, Height, %versionMAIN%
		GuiControl,, ProgressBar, 0
		MsgBox, 16, %version%, I was not able to find anything in the AssetManager with the value "%HostnameInputMAIN%".`n`nThe error is specifically this: %Error_Desc1%`n`nAre you sure that you used the correct case and CI name/SGER in your query? (Upper-case letters, small-case letters, spaces, etc.)`n`nExamples of proper usage:`nS20154997`nDCPWIN20154997`nAF6HL234
		return				
	}


GuiControl,, ProgressBar, 85

;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<Assettag>(.*?)</Assettag>", Assettag)
if (Found)
	{	 
		Assettag = %Assettag1%
		;MsgBox, % "Assettag: " Assettag1
	}
else
	{
		Assettag = N/A
		;MsgBox, % "Assettag: " Assettag
	}



ResultAutoDiscovery = 

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", Device := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<Device>(.*?)<\/Device>", Match, Pos + StrLen(Match))) {
	Device[++count] := Match1
}
loop % Device.MaxIndex() 
    str .= Device[A_Index] "`n"
;~ MsgBox %str%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", IPAddress := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<IPAddress>(.*?)<\/IPAddress>", Match, Pos + StrLen(Match))) {
	IPAddress[++count] := Match1
}
loop % IPAddress.MaxIndex() 
    str .= IPAddress[A_Index] "`n"
;~ MsgBox %str%

;Vypise pocet matchnutych IP adries (tuto hodnotu potom by bolo dobre pouzit pre loopu na vypisovanie veci)
;~ MsgBox % IPAddress.Length()

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", ModifiedOn := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<ModifiedOn>(.*?)<\/ModifiedOn>", Match, Pos + StrLen(Match))) {
	ModifiedOn[++count] := Match1
}
loop % ModifiedOn.MaxIndex() 
    str .= ModifiedOn[A_Index] "`n"
;~ MsgBox %str%


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", PhysicalAddress := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<PhysicalAddress>(.*?)<\/PhysicalAddress>", Match, Pos + StrLen(Match))) {
	PhysicalAddress[++count] := Match1
}
loop % PhysicalAddress.MaxIndex() 
    str .= PhysicalAddress[A_Index] "`n"
;~ MsgBox %str%


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", ScanDate := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<ScanDate>(.*?)<\/ScanDate>", Match, Pos + StrLen(Match))) {
	ScanDate[++count] := Match1
}
loop % ScanDate.MaxIndex() 
    str .= ScanDate[A_Index] "`n"
;~ MsgBox %str%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", Source := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<Source>(.*?)<\/Source>", Match, Pos + StrLen(Match))) {
	Source[++count] := Match1
}
loop % Source.MaxIndex() 
    str .= Source[A_Index] "`n"
;~ MsgBox %str%


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", Speed := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<Speed>(.*?)<\/Speed>", Match, Pos + StrLen(Match))) {
	Speed[++count] := Match1
}
loop % Speed.MaxIndex() 
    str .= Speed[A_Index] "`n"
;~ MsgBox %str%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", AdditionalIP := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<AdditionalIP>(.*?)<\/AdditionalIP>", Match, Pos + StrLen(Match))) {
	AdditionalIP[++count] := Match1
}
loop % AdditionalIP.MaxIndex() 
    str .= AdditionalIP[A_Index] "`n"
;~ MsgBox %str%


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", Connection := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<Connection>(.*?)<\/Connection>", Match, Pos + StrLen(Match))) {
	Connection[++count] := Match1
}
loop % Connection.MaxIndex() 
    str .= Connection[A_Index] "`n"
;~ MsgBox %str%


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", DuplexMode := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<DuplexMode>(.*?)<\/DuplexMode>", Match, Pos + StrLen(Match))) {
	DuplexMode[++count] := Match1
}
loop % DuplexMode.MaxIndex() 
    str .= DuplexMode[A_Index] "`n"
;~ MsgBox %str%


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", VLAN := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<VLAN>(.*?)<\/VLAN>", Match, Pos + StrLen(Match))) {
	VLAN[++count] := Match1
}
loop % VLAN.MaxIndex() 
    str .= VLAN[A_Index] "`n"
;~ MsgBox %str%


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", RegisteredOwner := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<RegisteredOwner>(.*?)<\/RegisteredOwner>", Match, Pos + StrLen(Match))) {
	RegisteredOwner[++count] := Match1
}
loop % RegisteredOwner.MaxIndex() 
    str .= RegisteredOwner[A_Index] "`n"
;~ MsgBox %str%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;Toto nie je zahrnute v outpute - netreba
Pos := 1, Match := "", SystemIDs := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<SystemID>(.*?)<\/SystemID>", Match, Pos + StrLen(Match))) {
	SystemIDs[++count] := Match1
}
loop % SystemIDs.MaxIndex() 
    str .= SystemIDs[A_Index] "`n"
;~ MsgBox %str%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", AutoDiscoveryInterfaces := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<AutoDiscoveryInterfaces>", Match, Pos + StrLen(Match))) {
	AutoDiscoveryInterfaces[++count] := Match1
}
loop % AutoDiscoveryInterfaces.MaxIndex() 
    str .= AutoDiscoveryInterfaces[A_Index] "`n"
;~ MsgBox %str%


;Create a variable to store the count of matched values
countOfFoundValues := % AutoDiscoveryInterfaces.Length()

;Loop through the count of found values
Loop % countOfFoundValues
	{
		if (Device[A_Index] = "")
		{
			Device[A_Index] := "N/A"
		}
		
		if (IPAddress[A_Index] = "")
		{
			IPAddress[A_Index] := "N/A"
		}
		
		if (ModifiedOn[A_Index] = "")
		{
			ModifiedOn[A_Index] := "N/A"
		}
		
		if (PhysicalAddress[A_Index] = "")
		{
			PhysicalAddress[A_Index] := "N/A"
		}
		
		if (ScanDate[A_Index] = "")
		{
			ScanDate[A_Index] := "N/A"
		}
		
		if (Source[A_Index] = "")
		{
			Source[A_Index] := "N/A"
		}
		
		if (Speed[A_Index] = "")
		{
			Speed[A_Index] := "N/A"
		}
		
		if (AdditionalIP[A_Index] = "")
		{
			AdditionalIP[A_Index] := "N/A"
		}
		
		if (Connection[A_Index] = "")
		{
			Connection[A_Index] := "N/A"
		}
		
		if (DuplexMode[A_Index] = "")
		{
			DuplexMode[A_Index] := "N/A"
		}
		
		if (VLAN[A_Index] = "" )
		{
			VLAN[A_Index] := "N/A"
		}
		
		if (RegisteredOwner[A_Index] = "" )
		{
			RegisteredOwner[A_Index] := "N/A"
		}
		
		if (RegisteredOwner[A_Index] = "*" )
		{
			RegisteredOwner[A_Index] := "N/A"
		}
				
		if (SystemIDs[A_Index] = "" )
		{
			SystemIDs[A_Index] := "N/A"
		}

		
		ResultAutoDiscovery := ResultAutoDiscovery . "Interface number: "  A_Index "`nName of the NIC: " Device[A_Index] "`nIP address: " IPAddress[A_Index] "`nMAC address: " PhysicalAddress[A_Index] "`nScan date: " ScanDate[A_Index] "`nSource: " Source[A_Index] "`nSpeed: " Speed[A_Index] "`nAdditional IP address: " AdditionalIP[A_Index] "`nConnection: " Connection[A_Index] "`nDuplex Mode: " DuplexMode[A_Index] "`nVLAN: " VLAN[A_Index] "`nModified on: " ModifiedOn[A_Index] "`nRegistered owner: " RegisteredOwner[A_Index] "`n`n"
		
		
		
	}

;~ MsgBox % ResultAutoDiscovery

GuiControl,, ProgressBar, 90
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<CustomData5>(.*?)</CustomData5>", CustomData5)
if (Found)
	{	 
		CustomData5 = %CustomData51%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		CustomData5 = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<CustomData4>(.*?)</CustomData4>", CustomData4)
if (Found)
	{	 
		CustomData4 = %CustomData41%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		CustomData4 = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<CustomData3>(.*?)</CustomData3>", CustomData3)
if (Found)
	{	 
		CustomData3 = %CustomData31%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		CustomData3 = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<CustomData2>(.*?)</CustomData2>", CustomData2)
if (Found)
	{	 
		CustomData2 = %CustomData21%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		CustomData2 = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<CustomData1>(.*?)</CustomData1>", CustomData1)
if (Found)
	{	 
		CustomData1 = %CustomData11%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		CustomData1 = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<VirtualCompType>(.*?)</VirtualCompType>", VirtualCompType)
if (Found)
	{	 
		VirtualCompType = %VirtualCompType1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		VirtualCompType = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<OSKernelMode>(.*?)</OSKernelMode>", OSKernelMode)
if (Found)
	{	 
		OSKernelMode = %OSKernelMode1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		OSKernelMode = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<OSArchitecture>(.*?)</OSArchitecture>", OSArchitecture)
if (Found)
	{	 
		OSArchitecture = %OSArchitecture1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		OSArchitecture = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<OSInstDate>(.*?)</OSInstDate>", OSInstDate)
if (Found)
	{	 
		OSInstDate = %OSInstDate1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		OSInstDate = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<OSSubVersion>(.*?)</OSSubVersion>", OSSubVersion)
if (Found)
	{	 
		OSSubVersion = %OSSubVersion1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		OSSubVersion = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<LOSMajorVersion>(.*?)</LOSMajorVersion>", LOSMajorVersion)
if (Found)
	{	 
		LOSMajorVersion = %LOSMajorVersion1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		LOSMajorVersion = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<OSType>(.*?)</OSType>", OSType)
if (Found)
	{	 
		OSType = %OSType1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		OSType = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<DTComputerBootTime>(.*?)</DTComputerBootTime>", DTComputerBootTime)
if (Found)
	{	 
		DTComputerBootTime = %DTComputerBootTime1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		DTComputerBootTime = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<Type>(.*?)</Type>", Type)
if (Found)
	{	 
		Type = %Type1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		Type = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<DTScanDate>(.*?)</DTScanDate>", DTScanDate)
if (Found)
	{	 
		DTScanDate = %DTScanDate1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		DTScanDate = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<DifferentFromSystem>(.*?)</DifferentFromSystem>", DifferentFromSystem)
if (Found)
	{	 
		DifferentFromSystem = %DifferentFromSystem1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		DifferentFromSystem = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<Hostid>(.*?)</Hostid>", Hostid)
if (Found)
	{	 
		Hostid = %Hostid1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		Hostid = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<Oslanguage>(.*?)</Oslanguage>", Oslanguage)
if (Found)
	{	 
		Oslanguage = %Oslanguage1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		Oslanguage = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<Environment>(.*?)</Environment>", Environment)
if (Found)
	{	 
		Environment = %Environment1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		Environment = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<Dttimestamp>(.*?)</Dttimestamp>", Dttimestamp)
if (Found)
	{	 
		Dttimestamp = %Dttimestamp1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		Dttimestamp = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<Source>(.*?)</Source>", Source)
if (Found)
	{	 
		Source = %Source1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		Source = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<SerialNo>(.*?)</SerialNo>", SerialNo)
if (Found)
	{	 
		SerialNo = %SerialNo1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		SerialNo = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<OS>(.*?)</OS>", OS)
if (Found)
	{	 
		OS = %OS1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		OS = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<Name>(.*?)</Name>", Name)
if (Found)
	{	 
		Name = %Name1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		Name = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<Model>(.*?)</Model>", Model)
if (Found)
	{	 
		Model = %Model1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		Model = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<LMemoryMB>(.*?)</LMemoryMB>", LMemoryMB)
if (Found)
	{	 
		LMemoryMB = %LMemoryMB1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		LMemoryMB = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<LCPUSpeedMHz>(.*?)</LCPUSpeedMHz>", LCPUSpeedMHz)
if (Found)
	{	 
		LCPUSpeedMHz = %LCPUSpeedMHz1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		LCPUSpeedMHz = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<LCPUCount>(.*?)</LCPUCount>", LCPUCount)
if (Found)
	{	 
		LCPUCount = %LCPUCount1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		LCPUCount = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<Assettag>(.*?)</Assettag>", Assettag)
if (Found)
	{	 
		Assettag = %Assettag1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		Assettag = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
;;;;;;;;;;;;;;;;;;;;;
Found := RegExMatch(result,"U)<CPUType>(.*?)</CPUType>", CPUType)
if (Found)
	{	 
		CPUType = %CPUType1%
		;MsgBox, % "ParentSerialNo: " ParentSerialNo1
	}
else
	{
		CPUType = N/A
		;MsgBox, % "ParentSerialNo: " ParentSerialNo
	}
	
GuiControl,, ProgressBar, 100


OutputFromAutoDiscovery = SGER: %AssetTag%`nHostname from AutoDiscovery: %Name%`n`n%ResultAutoDiscovery%Additional info 1: %CustomData1%`nAdditional info 2: %CustomData2%`nAdditional info 3: %CustomData3%`nAdditional info 4: %CustomData4%`nAdditional info 5: %CustomData5%`nVirtual Computer Type: %VirtualCompType%`nOS: %OS%`nOS Type: %OSType%`nOS Kernel Mode: %OSKernelMode%`nOS Architecture: %OSArchitecture%`nOS Installation Date: %OSInstDate%`nOS Version: %LOSMajorVersion%`nOS Sub-version: %OSSubVersion%`nLast boot date: %DTComputerBootTime%`nType: %Type%`nLast scan date: %DTScanDate%`nHost ID: %Hostid%`nOS Language: %Oslanguage%`nEnvironment: %Environment%`nSource: %Source%`nSerial Number: %SerialNo%`nModel: %Model%`nRAM Memory (MB): %LMemoryMB%`nCPU Type: %CPUType%`nCPU Count: %LCPUCount%`nCPU Speed (MHz): %LCPUSpeedMHz%`nDifferent from system: %DifferentFromSystem%

;DT Time Stamp: %Dttimestamp%`n

;~ MsgBox, % OutputFromAutoDiscovery
GuiControl,, ProgressBar, 0
WinGetPos X, Y, Width, Height, %versionMAIN%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;Populate ListView
LoopCount = 0

loop, parse, OutputFromAutoDiscovery, `n
{
	arry := StrSplit(A_LoopField, ": ")
	LV_Add("", arry[1], arry[2])
}


Gui, Show, x%X% y%Y% h700 w373, %versionMAIN%


OutputResult = %OutputFromAutoDiscovery%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;Disconnect after query
url = http://%urlFirstPart%/service/api/Disconnect
url := UriEncode(url)		
WinHTTP.Open("GET", url, 0)
WinHTTP.SetRequestHeader("Content-Type", "application/json")
Body := "{''}"

try
	{
		WinHTTP.Send(Body)
	}
catch
	{
		GuiControl,, ProgressBar, 0
		MsgBox, 16, %version%, An error has occured probably because of one of these reasons:`n1.) You are not connected to internal company network`n2.) Your proxy settings do not allow connection to AssetManager`n3.) You are not connected to a functioning internet connection
		return
	}

Result := WinHTTP.ResponseText
Status := WinHTTP.Status

;~ MsgBox, % Result
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;


url =
urlFirstPart = 


return


CheckInterfaces:
WinGetPos X, Y, Width, Height, %versionMAIN%
LV_Delete()

if (checkedWithOneOption = 1)
	{
		urlFirstPart = CanNotBeDisclosed.com		
	}
else
	{
		urlFirstPart = CanNotBeDisclosed.com
	}


;Initiate URL and Encode it to URL standards
url = http://%urlFirstPart%/service/api/GetInterfaceInfoEx?whereclause=Computer.AssetTag LIKE '%HostnameInput%'
url := UriEncode(url)		
WinHTTP.Open("GET", url, 0)

GuiControl,, ProgressBar, 50

WinHTTP.SetRequestHeader("Content-Type", "application/json")
Body := "{''}"
try
{
	WinHTTP.Send(Body)
}
catch
{
	GuiControl,, ProgressBar, 0
	MsgBox, 16, %version%, An error has occured probably because of one of these reasons:`n1.) You are not connected to internal company network`n2.) Your proxy settings do not allow connection to AssetManager`n3.) You are not connected to a functioning internet connection
	return
}
Result := WinHTTP.ResponseText
Status := WinHTTP.Status

;Try to search for SGER if it's not entered by the user and try to search for AutoDiscovery data again
FoundNewer := RegExMatch(result,"U)<Error_Desc>(.*?)</Error_Desc>", Error_Desc)
if (FoundNewer)
	{	
		NoRecordFoundVariable = %Error_Desc1%
		
		if (NoRecordFoundVariable = "No record found!")
			{
				
				RegExMatch(HostnameInput, "^S[1-9]", matchedSGERbyRegEx)
				
				if (matchedSGERbyRegEx != "")
					{
						GuiControl,, ProgressBar, 60
						Result =
						
						
						if (checkedWithOneOption = 1)
						{
							urlFirstPart = CanNotBeDisclosed.com		
						}
						else
						{
							urlFirstPart = CanNotBeDisclosed.com
						}
						
						;Initiate URL and Encode it to URL standards
						url = http://%urlFirstPart%/service/api/GetLogSysInfoEx?whereclause=AssetTag = '%HostnameInput%'
						url := UriEncode(url)		
						WinHTTP.Open("GET", url, 0)
						WinHTTP.SetRequestHeader("Content-Type", "application/json")
						
				
						try
						{
							WinHTTP.Send(Body)
						}
						catch
						{
							GuiControl,, ProgressBar, 0
							MsgBox, 16, %version%, An error has occured probably because of one of these reasons:`n1.) You are not connected to internal company network`n2.) Your proxy settings do not allow connection to AssetManager`n3.) You are not connected to a functioning internet connection
							return
						}
						
						
						Result := WinHTTP.ResponseText
						Status := WinHTTP.Status
				
					}
				else
					{
						GuiControl,, ProgressBar, 60
						
						Result =
						
						if (checkedWithOneOption = 1)
						{
							urlFirstPart = CanNotBeDisclosed.com		
						}
						else
						{
							urlFirstPart = CanNotBeDisclosed.com
						}
						
						url = http://%urlFirstPart%/service/api/GetLogSysInfoEx?whereclause=AddSysName LIKE '%HostnameInput%'
						url := UriEncode(url)
						
						WinHTTP.Open("GET", url, 0)
						WinHTTP.SetRequestHeader("Content-Type", "application/json")
						
						try
						{
							WinHTTP.Send(Body)
						}
						catch
						{
							GuiControl,, ProgressBar, 0
							MsgBox, 16, %version%, An error has occured probably because of one of these reasons:`n1.) You are not connected to internal company network`n2.) Your proxy settings do not allow connection to AssetManager`n3.) You are not connected to a functioning internet connection
							return
						}
						
						Result := WinHTTP.ResponseText
						Status := WinHTTP.Status
						
					
						Found := RegExMatch(result,"U)<Error_Desc>", Error_Desc)
						
						if (Found)
							{	
									GuiControl,, ProgressBar, 60
									
									Result =
									
									if (checkedWithOneOption = 1)
									{
										urlFirstPart = CanNotBeDisclosed.com		
									}
									else
									{
										urlFirstPart = CanNotBeDisclosed.com
									}
									
									
									url = http://%urlFirstPart%/service/api/GetLogSysInfoEx?whereclause=Portfolio.name LIKE '%HostnameInput%'
									url := UriEncode(url)
									
									WinHTTP.Open("GET", url, 0)
									WinHTTP.SetRequestHeader("Content-Type", "application/json")
									
									try
									{
										WinHTTP.Send(Body)
									}
									catch
									{
										GuiControl,, ProgressBar, 0
										MsgBox, 16, %version%, An error has occured probably because of one of these reasons:`n1.) You are not connected to internal company network`n2.) Your proxy settings do not allow connection to AssetManager`n3.) You are not connected to a functioning internet connection
										return
									}
									
									Result := WinHTTP.ResponseText
									Status := WinHTTP.Status
									
									Found =
							}
						
						
					}
			}
		;;;;;;;;;;;;;;;;;;;;;
		Found := RegExMatch(result,"U)<Assettag>(.*?)</Assettag>", Assettag)
		if (Found)
			{	 
				Assettag = %Assettag1%
				;MsgBox, % "Assettag: " Assettag1
			}
		else
			{
				Assettag = N/A
				;MsgBox, % "Assettag: " Assettag
			}
			
			
		if (checkedWithOneOption = 1)
			{
				urlFirstPart = CanNotBeDisclosed.com			
			}
		else
			{
				urlFirstPart = CanNotBeDisclosed.com
			}
			
			GuiControl,, ProgressBar, 75
			
		HostnameInput = %Assettag1%
		
		
		;Initiate URL and Encode it to URL standards
		url = http://%urlFirstPart%/service/api/GetInterfaceInfoEx?whereclause=Computer.AssetTag LIKE '%HostnameInput%'
		url := UriEncode(url)		
		WinHTTP.Open("GET", url, 0)
		WinHTTP.SetRequestHeader("Content-Type", "application/json")
		Body := "{''}"
		try
		{
			WinHTTP.Send(Body)
		}
		catch
		{
			GuiControl,, ProgressBar, 0
			MsgBox, 16, %version%, An error has occured probably because of one of these reasons:`n1.) You are not connected to internal company network`n2.) Your proxy settings do not allow connection to AssetManager`n3.) You are not connected to a functioning internet connection
			return
		}
		Result := WinHTTP.ResponseText
		Status := WinHTTP.Status
			
			

	}
	
FoundNewest := RegExMatch(result,"U)<Error_Desc>(.*?)</Error_Desc>", Error_Desc)
if (FoundNewest)
	{
		WinGetPos X, Y, Width, Height, %versionMAIN%
		GuiControl,, ProgressBar, 0
		MsgBox, 16, %version%, I was not able to find anything in the AssetManager with the value "%HostnameInputMAIN%".`n`nThe error is specifically this: %Error_Desc1%`n`nAre you sure that you used the correct case and CI name/SGER in your query? (Upper-case letters, small-case letters, spaces, etc.)`n`nExamples of proper usage:`nS20154997`nDCPWIN20154997`nAF6HL234
		return				
	}


GuiControl,, ProgressBar, 85

ResultInterfaces = 

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", Port := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<Port>(.*?)<\/Port>", Match, Pos + StrLen(Match))) {
	Port[++count] := Match1
}
loop % Port.MaxIndex() 
    str .= Port[A_Index] "`n"
;~ MsgBox %str%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", Vlan := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<Vlan>(.*?)<\/Vlan>", Match, Pos + StrLen(Match))) {
	Vlan[++count] := Match1
}
loop % Vlan.MaxIndex() 
    str .= Vlan[A_Index] "`n"
;~ MsgBox %str%

;Vypise pocet matchnutych IP adries (tuto hodnotu potom by bolo dobre pouzit pre loopu na vypisovanie veci)
;~ MsgBox % IPAddress.Length()

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", IPv6address := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<IPv6address>(.*?)<\/IPv6address>", Match, Pos + StrLen(Match))) {
	IPv6address[++count] := Match1
}
loop % IPv6address.MaxIndex() 
    str .= IPv6address[A_Index] "`n"
;~ MsgBox %str%


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", NetworkArea := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<NetworkArea>(.*?)<\/NetworkArea>", Match, Pos + StrLen(Match))) {
	NetworkArea[++count] := Match1
}
loop % NetworkArea.MaxIndex() 
    str .= NetworkArea[A_Index] "`n"
;~ MsgBox %str%


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", bDeleted := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<bDeleted>(.*?)<\/bDeleted>", Match, Pos + StrLen(Match))) {
	bDeleted[++count] := Match1
}
loop % bDeleted.MaxIndex() 
    str .= bDeleted[A_Index] "`n"
;~ MsgBox %str%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", TcpIpAddress := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<TcpIpAddress>(.*?)<\/TcpIpAddress>", Match, Pos + StrLen(Match))) {
	TcpIpAddress[++count] := Match1
}
loop % TcpIpAddress.MaxIndex() 
    str .= TcpIpAddress[A_Index] "`n"
;~ MsgBox %str%


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", PhysicalAddress := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<PhysicalAddress>(.*?)<\/PhysicalAddress>", Match, Pos + StrLen(Match))) {
	PhysicalAddress[++count] := Match1
}
loop % PhysicalAddress.MaxIndex() 
    str .= PhysicalAddress[A_Index] "`n"
;~ MsgBox %str%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", Computer := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<Computer>(.*?)<\/Computer>", Match, Pos + StrLen(Match))) {
	Computer[++count] := Match1
}
loop % Computer.MaxIndex() 
    str .= Computer[A_Index] "`n"
;~ MsgBox %str%


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", Application := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<Application>(.*?)<\/Application>", Match, Pos + StrLen(Match))) {
	Application[++count] := Match1
}
loop % Application.MaxIndex() 
    str .= Application[A_Index] "`n"
;~ MsgBox %str%


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", LogicalNetwork := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<LogicalNetwork>(.*?)<\/LogicalNetwork>", Match, Pos + StrLen(Match))) {
	LogicalNetwork[++count] := Match1
}
loop % LogicalNetwork.MaxIndex() 
    str .= LogicalNetwork[A_Index] "`n"
;~ MsgBox %str%


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", Description := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<Description>(.*?)<\/Description>", Match, Pos + StrLen(Match))) {
	Description[++count] := Match1
}
loop % Description.MaxIndex() 
    str .= Description[A_Index] "`n"
;~ MsgBox %str%


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", Status := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<Status>(.*?)<\/Status>", Match, Pos + StrLen(Match))) {
	Status[++count] := Match1
}
loop % Status.MaxIndex() 
    str .= Status[A_Index] "`n"
;~ MsgBox %str%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;Toto nie je zahrnute v outpute - netreba
Pos := 1, Match := "", DNSAlias := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<DNSAlias>(.*?)<\/DNSAlias>", Match, Pos + StrLen(Match))) {
	DNSAlias[++count] := Match1
}
loop % DNSAlias.MaxIndex() 
    str .= DNSAlias[A_Index] "`n"
;~ MsgBox %str%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", DNSName := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<DNSName>(.*?)<\/DNSName>", Match, Pos + StrLen(Match))) {
	DNSName[++count] := Match1
}
loop % DNSName.MaxIndex() 
    str .= DNSName[A_Index] "`n"
;~ MsgBox %str%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", NatIpAddress := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<NatIpAddress>(.*?)<\/NatIpAddress>", Match, Pos + StrLen(Match))) {
	NatIpAddress[++count] := Match1
}
loop % NatIpAddress.MaxIndex() 
    str .= NatIpAddress[A_Index] "`n"
;~ MsgBox %str%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", SubnetMask := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<SubnetMask>(.*?)<\/SubnetMask>", Match, Pos + StrLen(Match))) {
	SubnetMask[++count] := Match1
}
loop % SubnetMask.MaxIndex() 
    str .= SubnetMask[A_Index] "`n"
;~ MsgBox %str%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", DNSServers := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<DNSServers>(.*?)<\/DNSServers>", Match, Pos + StrLen(Match))) {
	DNSServers[++count] := Match1
}
loop % DNSServers.MaxIndex() 
    str .= DNSServers[A_Index] "`n"
;~ MsgBox %str%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", DNSSuffixes := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<DNSSuffixes>(.*?)<\/DNSSuffixes>", Match, Pos + StrLen(Match))) {
	DNSSuffixes[++count] := Match1
}
loop % DNSSuffixes.MaxIndex() 
    str .= DNSSuffixes[A_Index] "`n"
;~ MsgBox %str%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", DefaultGateway := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<DefaultGateway>(.*?)<\/DefaultGateway>", Match, Pos + StrLen(Match))) {
	DefaultGateway[++count] := Match1
}
loop % DefaultGateway.MaxIndex() 
    str .= DefaultGateway[A_Index] "`n"
;~ MsgBox %str%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", Type := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<Type>(.*?)<\/Type>", Match, Pos + StrLen(Match))) {
	Type[++count] := Match1
}
loop % Type.MaxIndex() 
    str .= Type[A_Index] "`n"
;~ MsgBox %str%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", DHCPEnabled := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<DHCPEnabled>(.*?)<\/DHCPEnabled>", Match, Pos + StrLen(Match))) {
	DHCPEnabled[++count] := Match1
}
loop % DHCPEnabled.MaxIndex() 
    str .= DHCPEnabled[A_Index] "`n"
;~ MsgBox %str%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", Code := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<Code>(.*?)<\/Code>", Match, Pos + StrLen(Match))) {
	Code[++count] := Match1
}
loop % Code.MaxIndex() 
    str .= Code[A_Index] "`n"
;~ MsgBox %str%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", DCHPServers := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<DCHPServers>(.*?)<\/DCHPServers>", Match, Pos + StrLen(Match))) {
	DCHPServers[++count] := Match1
}
loop % DCHPServers.MaxIndex() 
    str .= DCHPServers[A_Index] "`n"
;~ MsgBox %str%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", Remarks := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<Remarks>(.*?)<\/Remarks>", Match, Pos + StrLen(Match))) {
	Remarks[++count] := Match1
}
loop % Remarks.MaxIndex() 
    str .= Remarks[A_Index] "`n"
;~ MsgBox %str%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", ExternalID := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<ExternalID>(.*?)<\/ExternalID>", Match, Pos + StrLen(Match))) {
	ExternalID[++count] := Match1
}
loop % ExternalID.MaxIndex() 
    str .= ExternalID[A_Index] "`n"
;~ MsgBox %str%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Pos := 1, Match := "", ExternalSystem := Object(), count := 0, str := ""
While (Pos := RegExMatch(Result, "U)<ExternalSystem>(.*?)<\/ExternalSystem>", Match, Pos + StrLen(Match))) {
	ExternalSystem[++count] := Match1
}
loop % ExternalSystem.MaxIndex() 
    str .= ExternalSystem[A_Index] "`n"
;~ MsgBox %str%


;Create a variable to store the count of matched values
countOfFoundValues := % Computer.Length()

;Loop through the count of found values
Loop % countOfFoundValues
	{
		if (Port[A_Index] = "")
		{
			Port[A_Index] := "N/A"
		}
		
		if (Vlan[A_Index] = "")
		{
			Vlan[A_Index] := "N/A"
		}
		
		if (IPv6address[A_Index] = "")
		{
			IPv6address[A_Index] := "N/A"
		}
		
		if (NetworkArea[A_Index] = "")
		{
			NetworkArea[A_Index] := "N/A"
		}
		
		if (bDeleted[A_Index] = "")
		{
			bDeleted[A_Index] := "N/A"
		}
		
		if (TcpIpAddress[A_Index] = "")
		{
			TcpIpAddress[A_Index] := "N/A"
		}
		
		if (PhysicalAddress[A_Index] = "")
		{
			PhysicalAddress[A_Index] := "N/A"
		}
		
		if (Computer[A_Index] = "")
		{
			Computer[A_Index] := "N/A"
		}
		
		if (Application[A_Index] = "")
		{
			Application[A_Index] := "N/A"
		}
		
		if (LogicalNetwork[A_Index] = "")
		{
			LogicalNetwork[A_Index] := "N/A"
		}
		
		if (Description[A_Index] = "" )
		{
			Description[A_Index] := "N/A"
		}
		
		if (Status[A_Index] = "" )
		{
			Status[A_Index] := "N/A"
		}
		
		if (DNSAlias[A_Index] = "" )
		{
			DNSAlias[A_Index] := "N/A"
		}
				
		if (DNSName[A_Index] = "" )
		{
			DNSName[A_Index] := "N/A"
		}

		if (NatIpAddress[A_Index] = "" )
		{
			NatIpAddress[A_Index] := "N/A"
		}

		if (SubnetMask[A_Index] = "" )
		{
			SubnetMask[A_Index] := "N/A"
		}

		if (DNSServers[A_Index] = "" )
		{
			DNSServers[A_Index] := "N/A"
		}

		if (DNSSuffixes[A_Index] = "" )
		{
			DNSSuffixes[A_Index] := "N/A"
		}

		if (DefaultGateway[A_Index] = "" )
		{
			DefaultGateway[A_Index] := "N/A"
		}

		if (Type[A_Index] = "" )
		{
			Type[A_Index] := "N/A"
		}

		if (DHCPEnabled[A_Index] = "" )
		{
			DHCPEnabled[A_Index] := "N/A"
		}

		if (Code[A_Index] = "" )
		{
			Code[A_Index] := "N/A"
		}
		
		if (DCHPServers[A_Index] = "" )
		{
			DCHPServers[A_Index] := "N/A"
		}

		if (Remarks[A_Index] = "" )
		{
			Remarks[A_Index] := "N/A"
		}

		if (ExternalID[A_Index] = "" )
		{
			ExternalID[A_Index] := "N/A"
		}

		if (ExternalSystem[A_Index] = "" )
		{
			ExternalSystem[A_Index] := "N/A"
		}

		

		GuiControl,, ProgressBar, 100

		
		ResultInterfaces := ResultInterfaces . "Interface number: "  A_Index "`nBelongs to this CI: " Computer[A_Index] "`nStatus: " Status[A_Index] "`nType: " Type[A_Index] "`nIPv6 address: " IPv6address[A_Index] "`nIPv4 address: " TcpIpAddress[A_Index] "`nSubnet Mask: " SubnetMask[A_Index] "`nPort: " Port[A_Index] "`nVLAN: " Vlan[A_Index] "`nDefault Gateway: " DefaultGateway[A_Index] "`nNAT IP address: " NatIpAddress[A_Index] "`nRemarks: " Remarks[A_Index] "`nDNS Alias: " DNSAlias[A_Index] "`nDNS Name: " DNSName[A_Index] "`nDNS Servers: " DNSServers[A_Index] "`nDNS Suffixes: " DNSSuffixes[A_Index] "`nDHCP Enabled?: " DHCPEnabled[A_Index] "`nDHCP Servers: " DCHPServers[A_Index] "`nNetwork Area: " NetworkArea[A_Index] "`nMAC address: " PhysicalAddress[A_Index] "`nApplication: " Application[A_Index] "`nLogical network: " LogicalNetwork[A_Index] "`nDescription: " Description[A_Index] "`nExternal ID: " ExternalID[A_Index] "`nExternal System: " ExternalSystem[A_Index] "`nCode: " Code[A_Index] "`nDeleted?: " bDeleted[A_Index] "`n`n`n"
		
}

ResultInterfaces := "Total number of interfaces: " countOfFoundValues "`n`n" . ResultInterfaces

GuiControl,, ProgressBar, 0
WinGetPos X, Y, Width, Height, %versionMAIN%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;Populate ListView
LoopCount = 0

loop, parse, ResultInterfaces, `n
{
	arry := StrSplit(A_LoopField, ": ")
	LV_Add("", arry[1], arry[2])
}


Gui, Show, x%X% y%Y% h700 w373, %versionMAIN%


OutputResult = %ResultInterfaces%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;Disconnect after query
url = http://%urlFirstPart%/service/api/Disconnect
url := UriEncode(url)		
WinHTTP.Open("GET", url, 0)
WinHTTP.SetRequestHeader("Content-Type", "application/json")
Body := "{''}"

try
	{
		WinHTTP.Send(Body)
	}
catch
	{
		GuiControl,, ProgressBar, 0
		MsgBox, 16, %version%, An error has occured probably because of one of these reasons:`n1.) You are not connected to internal company network`n2.) Your proxy settings do not allow connection to AssetManager`n3.) You are not connected to a functioning internet connection
		return
	}

Result := WinHTTP.ResponseText
Status := WinHTTP.Status

;~ MsgBox, % Result
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

url =
urlFirstPart = 

return

GuiClose:
ExitApp
