; AEI.ahk - by joedf
; Revision Date : 00:04 2017/04/07
; Tested On AutoHotkey Version: 1.1.19.02
;#NoTrayIcon
#SingleInstance, Off
SetWinDelay, 0
SetBatchLines, 0

AHK_UPDATE_URL		:= "https://autohotkey.com/download/1.1/version.txt"
AHK_DOWNLOAD_EXE	:= "https://autohotkey.com/download/ahk-install.exe"

CHECK_FOR_UPDATES		:=	true
UpdateCachedDownload	:=	true

ScriptName:="AutoHotkey Environment Information"
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
	Ahk_CompilerPath	:=	getCompiler()
	Ahk_WindowSpyPath	:=	getWindowSpy()
	
	BS_FLAT:="0x8000"
	
	ListViewNRows:=StrCount(VarList,"`n")+1

	Gui +LastFound +Border +OwnDialogs +hwndhGUI -Caption +e0x80000
	Gui, Margin, 4, 4
	Gui, Color, 0x202020, 0x202020
	Gui, Font, cFFFFFF s8, Consolas
	Gui, Add, Text, +center w450 h16 +c3399FF +BackgroundTrans +Border, % " [ " ScriptName " ] "
	Gui, Add, ListView, wp r%ListViewNRows% -Hdr +ReadOnly +Background0A0A0A vLV +LV0x4000, Key|Value
	Gui, Add, Picture,w16 h16 Icon1, %A_ahkpath%
	Gui, Add, Picture,wp hp Icon1 x+4 yp, %Ahk_CompilerPath%
	Gui, Add, Picture,wp hp Icon1 x+4 yp, %Ahk_WindowSpyPath%
	Gui, Add, Button,gCopy hp x+4 yp w176 vCopyButton +%BS_FLAT% -theme hwndhCopyButton,Copy All to Clipboard
	Gui, Add, Button,gGuiClose hp w100 x+4 yp +%BS_FLAT% -theme hwndhExitButton,Exit
	Gui, Add, Text, +center w106 x+4 yp+0 h16 cGray +Border vUpdateInfo gOpenUpdate, Not Checked...

	LV_Add("","Loading","System & Environment information ...")
	LV_ModifyCol(1,112)
	LV_ModifyCol(2,330)
	OnMessage(0x201, "WM_LBUTTONDOWN") ;Enable Draggable GUI
	GuiControl,+Disabled,CopyButton
	WinSet, Transparent, 0
	Gui, Show,, % ScriptName
	GroupAdd,MainGUI,ahk_id %hGUI%
	WinSet, Redraw ; uses th last found window
	WinFade("ahk_id " hGUI,255,20)
	
	gosub,GetInfo
	
	LV_Delete()
	Loop, Parse, VarList, `n, `r
		if (i:=A_LoopField)
			d:=%i%, Parse_append(Message,i,d), LV_Add("",i,d)
	StringTrimRight,Message,Message,1
	
	GuiControl,-Disabled,CopyButton
	
	if (CHECK_FOR_UPDATES)
		gosub,CheckUpdate
	return

GuiClose:
GuiEscape:
WinFade("ahk_id " hGUI,0,20)
ExitApp

;///////////////////////////// the Rest ////////////////////////////////////////;{
;############ LABELS ############### ;{
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
	Ahk_IsInstalled		:=	!isPortable()
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
	Gui, Font, cGray
	GuiControl, Font, UpdateInfo
	GuiControl,,UpdateInfo, Checking...
	CheckChar:=chr(10004), CrossChar:=chr(10008)
	if isWinXPOrOlder()
		CheckChar:=":)", CrossChar:="X"
	try {
		whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
		whr.Open("GET", AHK_UPDATE_URL)
		whr.Send()
		UpdateVersion := whr.ResponseText
	} catch {
		UpdateVersion := "ERROR"
	}
	if !RegExMatch(UpdateVersion,"\d+\.\d+\.\d+\.\d+") {
		GuiControl,,UpdateInfo, Check Error !
		Gui, Font, cRed
		GuiControl, Font, UpdateInfo
	} else {
		if UpdateVersion > %A_AhkVersion%
		{
			GuiControl,,UpdateInfo, % "Outdated " CrossChar
			Gui, Font, cRed
			GuiControl, Font, UpdateInfo
		} else {
			GuiControl,,UpdateInfo, % "Up to date " CheckChar
			Gui, Font, cLime
			GuiControl, Font, UpdateInfo
		}
	}
	return
OpenUpdate:
	Gui +OwnDialogs 
	GuiControlGet,UpdateInfoText,,UpdateInfo
	UpdateQuestion:="Do you want to check for updates now?"
	if UpdateInfoText contains Not,Err
	{
		if InStr(UpdateInfoText,"Not")
			MsgBox, 36, [AEI] AutoHotkey Update, No check for updates has been performed.`n%UpdateQuestion%
		if InStr(UpdateInfoText,"Error")
			MsgBox, 36, [AEI] AutoHotkey Update, An error occurred when checking for updates.`n%UpdateQuestion%
		IfMsgBox,Yes
			gosub,CheckUpdate
	}
	if InStr(UpdateInfoText,"Out") {
		MsgBox, 36, [AEI] AutoHotkey Update,
		(ltrim
		You are using an outdated version of AutoHotkey.
		Current Version :`t%A_AhkVersion%
		Latest Version :`t%UpdateVersion%
		Do you want to update now?
		)
		IfMsgBox,Yes
		{
			;Run, http://ahkscript.org/download/ahk-install.exe
			Gui +Disabled
			if (!UpdateCachedDownload || !FileExist(UpdateFile:=(A_Temp "\AutoHotkey_Install-v" UpdateVersion ".exe")) )
				DownloadFile(AHK_DOWNLOAD_EXE,UpdateFile)
			Run, %UpdateFile%
			
			;Smooth app exit...
			WinWaitActive,ahk_exe %UpdateFile%,,2
			WinWaitActive,AutoHotkey Setup,,2
			Sleep 250
			goto GuiClose
		}
	}
	if InStr(UpdateInfoText,"Up t") {
		MsgBox, 36, [AEI] AutoHotkey Update,
		(ltrim
		You are using the latest version.
		Current Version :`t%A_AhkVersion%
		%UpdateQuestion%
		)
		IfMsgBox,Yes
			gosub,CheckUpdate
	}
	return
;}
;############ FUNCTIONS ############### ;{
Parse_append(ByRef out, key, val) {
	k:=StrLen(key), out .= key " : `t"
	Loop % (k<9)?3:((k<16)?2:((k<17)?1:0))
		out .= "`t"
	out .= val "`n"
}
isWinXPOrOlder() {
	if A_OSVersion in WIN_NT4,WIN_95,WIN_98,WIN_ME,WIN_2003,WIN_XP,WIN_2000	
		return true
	return false
}
isPortable() {
	isInstalled(p)
	return !(p "\AutoHotkey.exe" = A_AhkPath)
}
isInstalled(ByRef path="") {
	RegRead,InstallDir,HKLM,SOFTWARE\AutoHotkey,InstallDir
	return InStr(A_AhkPath,path:=InstallDir)
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
StrCount(Haystack,Needle) {
	StringReplace, Haystack, Haystack, %Needle%, %Needle%, UseErrorLevel
	return ErrorLevel
}
getSysUptime() {
	t_UpTime := A_TickCount // 1000 ; Elapsed seconds since start
	SystemUptime := % t_UpTime // 86400 " days " mod(t_UpTime // 3600, 24) " hours " mod(t_UpTime // 60, 60) " mins " mod(t_UpTime, 60) " seconds"
	return SystemUptime
}
_GetOSVersionInfo() { ; from Shajul  //  http://www.autohotkey.com/board/topic/54639-getosversion/#entry414249
	static Ver
	If !Ver
	{
		VarSetCapacity(OSVer, 284, 0)
		NumPut(284, OSVer, 0, "UInt")
		If !DllCall("GetVersionExW", "Ptr", &OSVer)
		return 0 ; GetSysErrorText(A_LastError)
		Ver := Object()
		Ver.MajorVersion		:= NumGet(OSVer, 4, "UInt")
		Ver.MinorVersion		:= NumGet(OSVer, 8, "UInt")
		Ver.BuildNumber			:= NumGet(OSVer, 12, "UInt")
		Ver.PlatformId			:= NumGet(OSVer, 16, "UInt")
		Ver.ServicePackString	:= StrGet(&OSVer+20, 128, "UTF-16")
		Ver.ServicePackMajor	:= NumGet(OSVer, 276, "UShort")
		Ver.ServicePackMinor	:= NumGet(OSVer, 278, "UShort")
		Ver.SuiteMask			:= NumGet(OSVer, 280, "UShort")
		Ver.ProductType			:= NumGet(OSVer, 282, "UChar")
		Ver.EasyVersion			:= Ver.MajorVersion . "." . Ver.MinorVersion . "." . Ver.BuildNumber
	}
	return Ver
}
GetOSVersionInfo() { ; Thanks jNizM  //  http://ahkscript.org/boards/viewtopic.php?f=6&t=5825#p36105
	static Ver
	If !Ver
	{
		static RTL_OSVIEX, init := VarSetCapacity(RTL_OSVIEX,284,0) && NumPut(284,RTL_OSVIEX,"UInt")
		if (DllCall("ntdll.dll\RtlGetVersion","Ptr",&RTL_OSVIEX) != 0)
			return _GetOSVersionInfo() ; "Error in RtlGetVersion"
		Ver := Object()
		Ver.MajorVersion		:= NumGet(RTL_OSVIEX,4,"UInt")
		Ver.MinorVersion		:= NumGet(RTL_OSVIEX,8,"UInt")
		Ver.BuildNumber			:= NumGet(RTL_OSVIEX,12,"UInt")
		Ver.PlatformId			:= NumGet(RTL_OSVIEX,16,"UInt")
		Ver.ServicePackString	:= StrGet(&RTL_OSVIEX+20,128,"UTF-16")
		Ver.ServicePackMajor	:= NumGet(RTL_OSVIEX,276,"UShort")
		Ver.ServicePackMinor	:= NumGet(RTL_OSVIEX,278,"UShort")
		Ver.SuiteMask			:= NumGet(RTL_OSVIEX,280,"UShort")
		Ver.ProductType			:= NumGet(RTL_OSVIEX,282,"UChar")
		Ver.EasyVersion       	:= Ver.MajorVersion . "." . Ver.MinorVersion . "." . Ver.BuildNumber
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
		return SystemGPU:=objItem.Name " v" objItem.DriverVersion " @ " Round((objItem.AdapterRAM / (1024 ** 2)), 2) " MB RAM"
}
Win32_Processor(o,ByRef SystemCPU) {
	colItems := o.ExecQuery("SELECT * FROM Win32_Processor")._NewEnum
	while colItems[objItem]
		return SystemCPU:=RegExReplace(objItem.Name,"(\s{2,}|\t)"," ")
}
winfade(w:="",t:=128,i:=1,d:=10) {
	w:=(w="")?("ahk_id " WinActive("A")):w
	t:=(t>255)?255:(t<0)?0:t
	WinGet,s,Transparent,%w%
	s:=(s="")?255:s ;prevent trans unset bug
	WinSet,Transparent,%s%,%w% 
	i:=(s<t)?abs(i):-1*abs(i)
	while(k:=(i<0)?(s>t):(s<t)&&WinExist(w)) {
		WinGet,s,Transparent,%w%
		s+=i
		WinSet,Transparent,%s%,%w%
		sleep %d%
	}
}
DownloadFile(UrlToFile, SaveFileAs, Overwrite := True, UseProgressBar := True, ProgressBarTitle:="Downloading...") {
	; DownloadFile() by brutosozialprodukt
	; http://ahkscript.org/boards/viewtopic.php?f=6&t=1674

	; Revision: joedf
	; Changes:  - Changed progress bar style & colors
	;           - Changed Display Information
	;           - Commented-out Size calculation
	;           - Added ShortURL()
	;           - Added short delay 100 ms to show the progress bar if download was too fast
	;           - Added ProgressBarTitle
	;           - Try-Catch "backup download code"
	; ----------------------------------------------------------------------------------

    ;Check if the file already exists and if we must not overwrite it
      If (!Overwrite && FileExist(SaveFileAs))
          Return
    ;Check if the user wants a progressbar
      If (UseProgressBar) {
          _surl:=ShortURL(UrlToFile)
          ;Initialize the WinHttpRequest Object
            WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
          ;Download the headers
            WebRequest.Open("HEAD", UrlToFile)
            WebRequest.Send()
          ;Store the header which holds the file size in a variable:
          try
            FinalSize := WebRequest.GetResponseHeader("Content-Length")
          catch
          {
            ;throw Exception("could not get Content-Length for URL: " UrlToFile)
            Progress, CW202020 CTFFFFFF CB3399FF w330 h52 B1 FS8 WM700 WS700 FM8 ZH12 ZY3 C11, , %ProgressBarTitle%, %_surl%, Consolas
            UrlDownloadToFile, %UrlToFile%, %SaveFileAs%
            Sleep 100
            Progress, Off
            return
          }
          ;Create the progressbar and the timer
            Progress, CW202020 CTFFFFFF CB3399FF w330 h52 B1 FS8 WM700 WS700 FM8 ZH12 ZY3 C11, , %ProgressBarTitle%, %_surl%, Consolas
            SetTimer, __UpdateProgressBar, 100
      }
    ;Download the file
      UrlDownloadToFile, %UrlToFile%, %SaveFileAs%
    ;Remove the timer and the progressbar because the download has finished
      If (UseProgressBar) {
          Sleep 100
          Progress, Off
          SetTimer, __UpdateProgressBar, Off
      }
    Return
    
    ;The label that updates the progressbar
      __UpdateProgressBar:
         ;Get the current filesize and tick
            CurrentSize := FileOpen(SaveFileAs, "r").Length ;FileGetSize wouldn't return reliable results
            CurrentSizeTick := A_TickCount
          ;Calculate the downloadspeed
            ;Speed := Round((CurrentSize/1024-LastSize/1024)/((CurrentSizeTick-LastSizeTick)/1000)) . " Kb/s"
          ;Save the current filesize and tick for the next time
            ;LastSizeTick := CurrentSizeTick
            ;LastSize := FileOpen(SaveFileAs, "r").Length
          ;Calculate percent done
            PercentDone := Round(CurrentSize/FinalSize*100)
          ;Update the ProgressBar
          _csize:=Round(CurrentSize/1024,1)
          _fsize:=Round(FinalSize/1024)
            Progress, %PercentDone%, Downloading:  %_csize% KB / %_fsize% KB  [ %PercentDone%`% ], %_surl%
      Return
}
ShortURL(p,l=50) {
    VarSetCapacity(_p, (A_IsUnicode?2:1)*StrLen(p) )
    DllCall("shlwapi\PathCompactPathEx"
        ,"str", _p
        ,"str", p
        ,"uint", abs(l)
        ,"uint", 0)
    return _p
}
;}
;}
;}