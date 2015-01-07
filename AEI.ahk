; AEI.ahk - by joedf
; Revision Date : 13:01 2015/01/07
; Tested On AutoHotkey Version: 1.1.19.01
#NoTrayIcon
#SingleInstance, Off

CHECK_FOR_UPDATES := true
ScriptName:="AutoHotkey Environment Information"

gosub,GetInfo

VarList=
(
AutoHotkey
SystemModel
SystemCPU
SystemGPU
SystemType
SystemRAM
SystemOS
SystemUptime
SystemLocale
SystemScreen
A_AhkPath
Ahk_CompilerPath
Ahk_WindowSpyPath
MPRESS_IsPresent
)

;;////////////////////////////// SETUP Display and Variable parse...
;{

	Gui +LastFound -Caption +Border +hwndhGUI
	Gui, Margin, 4, 4
	Gui, Color, 0x202020, 0x202020
	Gui, Font, cFFFFFF s8, Consolas
	Gui, Add, Progress, c202020 Background828790 w450 h16, 100
	Gui, Add, Text, +center wp hp-1 xp yp+1 +c3399FF +BackgroundTrans, % " [ " ScriptName " ] "
	Gui, Add, ListView, wp r14 -Hdr +ReadOnly +Background0A0A0A vLV +LV0x4000 gLV_eventHandler, Key|Value
	Gui, Add, Picture,w16 h16 Icon1, %A_ahkpath%
	Gui, Add, Picture,wp hp Icon1 x+4 yp, %Ahk_CompilerPath%
	Gui, Add, Picture,wp hp Icon1 x+4 yp, %Ahk_WindowSpyPath%
	Gui, Add, Button,gCopy hp x+4 yp w176,Copy All to Clipboard
	Gui, Add, Button,gGuiClose hp w100 x+4 yp,Exit
	Gui, Add, Text, +center w106 x+4 yp+0 h16 cGray +Border vUpdateInfo, Checking...

	Loop, Parse, VarList, `n, `r
		if (i:=A_LoopField)
			d:=%i%, Parse_append(Message,i,d), LV_Add("",i,d)
	StringTrimRight,Message,Message,1

	LV_ModifyCol(1,112)
	LV_ModifyCol(2,330)
	Gui, Show,, % ScriptName
	GroupAdd,MainGUI,ahk_id %hGUI%
	OnMessage(0x201, "WM_LBUTTONDOWN") ;Enable Draggable GUI
	gosub,CheckUpdate
	return

GuiClose:
GuiEscape:
ExitApp

;///////////////////////////// the Rest ////////////////////////////////////////;{
WM_LBUTTONDOWN() {
	#IfWinActive ahk_group MainGUI
		#If Control("Static")
		PostMessage 0xA1,2  ;-- Goyyah/SKAN trick
		;http://www.autohotkey.com/board/topic/80594-how-to-enable-drag-for-a-gui-without-a-titlebar/#entry60075
		#If
	#IfWinActive
}
Copy:
	Clipboard:=Message
	return
GetInfo:
	Ahk_Flavour			:=	(A_PtrSize*8) "-bit"
	Ahk_IsInstalled		:=	isInstalled()
	Ahk_CompilerPath	:=	getCompiler()
	Ahk_WindowSpyPath	:=	getWindowSpy()
	MPRESS_IsPresent	:=	(!!FileExist(Ahk_CompilerPath "\..\MPRESS.exe"))?"1 (true)":"0 (false)"
	SystemLocale		:=	getSysLocale() " (0x" A_Language ")"
	SystemScreen		:=	A_ScreenWidth "x" A_ScreenHeight " (" A_ScreenDPI " DPI)"
	SystemName			:=	A_UserName "@" A_ComputerName, e:=GetOSVersionInfo()
	SystemOS			:=	A_OSVersion " " (A_Is64bitOS?"64":"32") "-bit " e.ServicePackString " v" e.EasyVersion " (" A_OSType ")"
	AutoHotkey			:=	"v" A_AhkVersion " " (A_IsUnicode?"Unicode":"ANSI") " " Ahk_Flavour " (" (Ahk_IsInstalled?"Installed":"Portable") ")"
	SystemUptime		:=	getSysUptime()
	
	objWMIService:=ComObjGet("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	
	Win32_ComputerSystem(objWMIService,SystemModel,SystemType,SystemRAM)
	Win32_VideoController(objWMIService,SystemGPU)
	Win32_Processor(objWMIService,SystemCPU)
	return
CheckUpdate:
	if (CHECK_FOR_UPDATES) {
		try {
			whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
			whr.Open("GET", "http://ahkscript.org/download/1.1/version.txt")
			whr.Send()
			UpdateVersion := whr.ResponseText
		} catch {
			UpdateVersion := "ERROR"
		}
		if !RegExMatch(UpdateVersion,"\d+\.\d+\.\d+\.\d+") {
			GuiControl,,UpdateInfo, % "Check Error !"
			Gui, Font, cRed
			GuiControl, Font, UpdateInfo
		} else {
			if UpdateVersion > %A_AhkVersion%
			{
				GuiControl,,UpdateInfo, % UpdateVersion " " chr(10008)
				Gui, Font, cRed
				GuiControl, Font, UpdateInfo
			} else {
				GuiControl,,UpdateInfo, % A_AhkVersion " " chr(10004)
				Gui, Font, cLime
				GuiControl, Font, UpdateInfo
			}
		}
	} else {
		GuiControl,,UpdateInfo, % "Not Checked ..."
	}
	return
LV_eventHandler:
	; to do .....
	return
;#################################################
;#################################################
;////////////////////// functions ;{
Parse_append(ByRef out, key, val) {
	k:=StrLen(key), out .= key " : `t"
	Loop % (k<9)?3:((k<16)?2:((k<17)?1:0))
		out .= "`t"
	out .= val "`n"
}
getSelectedLVtext() {
	ret:="", RowNumber := 0  ; This causes the first loop iteration to start the search at the top of the list.
	Loop
	{
		RowNumber := LV_GetNext(RowNumber)  ; Resume the search at the row after that found by the previous iteration.
		if not RowNumber  ; The above returned zero, so there are no more selected rows.
			break
		LV_GetText(Key, RowNumber, 1)
		LV_GetText(Value, RowNumber, 2)
		ret .= 
	}
}
isInstalled() {
	RegRead,InstallDir,HKLM,SOFTWARE\AutoHotkey,InstallDir
	return InStr(A_AhkPath,InstallDir)
}
getCompiler() {
	SplitPath,A_AhkPath,,Dir
	if FileExist(p:=(Dir "\Compiler\Ahk2Exe.exe"))
		return p
	return "Not Found"
}
getWindowSpy() {
	SplitPath,A_AhkPath,,Dir
	if FileExist(p:=(Dir "\AU3_Spy.exe"))
		return p
	return "Not Found"
}
getSysLocale() { ; fork of http://stackoverflow.com/a/7759505/883015
	VarSetCapacity(buf_a,9,0), VarSetCapacity(buf_b,9,0)
	f:="GetLocaleInfo" (A_IsUnicode?"W":"A")
	DllCall(f,"Int",LOCALE_SYSTEM_DEFAULT:=0x800
			,"Int",LOCALE_SISO639LANGNAME:=89
			,"Str",buf_a,"Int",9)
	DllCall(f,"Int",LOCALE_SYSTEM_DEFAULT
			,"Int",LOCALE_SISO3166CTRYNAME:=90
			,"Str",buf_b,"Int",9)
	return buf_a "-" buf_b
}
Control(ClassNN) {
    MouseGetPos,,,,control
    return InStr(control,ClassNN)
}
getSysUptime() {
	t_UpTime := A_TickCount // 1000 ; Elapsed seconds since start
	SystemUptime := % t_UpTime // 86400 " days " mod(t_UpTime // 3600, 24) " hours " mod(t_UpTime // 60, 60) " mins " mod(t_UpTime, 60) " seconds"
	return SystemUptime
}
If ( OSVersion := GetOSVersionInfo() )
  MsgBox % "OS Version`t:`t" OSVersion.EasyVersion
       . "`nOS Service pack`t:`t" . OSVersion.ServicePackString
       . "`nIs Workstation`t:`t" (OSVersion.ProductType = 1 ? "Yes" : "No")
       
GetOSVersionInfo() { ; from Shajul  //  http://www.autohotkey.com/board/topic/54639-getosversion/#entry414249
	static Ver
	If !Ver
	{
		VarSetCapacity(OSVer, 284, 0)
		NumPut(284, OSVer, 0, "UInt")
		If !DllCall("GetVersionExW", "Ptr", &OSVer)
		return 0 ; GetSysErrorText(A_LastError)
		Ver := Object()
		Ver.MajorVersion      := NumGet(OSVer, 4, "UInt")
		Ver.MinorVersion      := NumGet(OSVer, 8, "UInt")
		Ver.BuildNumber       := NumGet(OSVer, 12, "UInt")
		Ver.PlatformId        := NumGet(OSVer, 16, "UInt")
		Ver.ServicePackString := StrGet(&OSVer+20, 128, "UTF-16")
		Ver.ServicePackMajor  := NumGet(OSVer, 276, "UShort")
		Ver.ServicePackMinor  := NumGet(OSVer, 278, "UShort")
		Ver.SuiteMask         := NumGet(OSVer, 280, "UShort")
		Ver.ProductType       := NumGet(OSVer, 282, "UChar")
		Ver.EasyVersion       := Ver.MajorVersion . "." . Ver.MinorVersion . "." . Ver.BuildNumber
	}
	return Ver
}
Win32_ComputerSystem(o,ByRef SystemModel,ByRef SystemType,ByRef SystemRAM) {
	sys_l := o.ExecQuery("Select * from Win32_ComputerSystem")._NewEnum
	while sys_l[sys]
	{
	SystemModel:=sys.Model
	SystemType:=sys.SystemType
	SystemRAM:=Round(sys.TotalPhysicalMemory/(1024*1024),0) . " MB"
	break
	}
}
Win32_VideoController(o,ByRef SystemGPU) {
	colItems := o.ExecQuery("SELECT * FROM Win32_VideoController")._NewEnum
	while colItems[objItem]
	{
	SystemGPU:=objItem.Name " v" objItem.DriverVersion " @ " Round((objItem.AdapterRAM / (1024 ** 2)), 2) " MB RAM"
	break
	}
}
Win32_Processor(o,ByRef SystemCPU) {
	colItems := o.ExecQuery("SELECT * FROM Win32_Processor")._NewEnum
	while colItems[objItem]
	{
	SystemCPU:=RegExReplace(objItem.Name,"(\s{2,}|\t)"," ")
	break
	}
}
;}
;}
;}