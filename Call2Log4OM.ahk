; ***************
; * Call2Log4OM *
; * v1.0.0      *
; * de IZ3XNJ   *
; ***************

; +------+
; | Main |
; +------+

	#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
	;#Warn ; Enable warnings to assist with detecting common errors. To use only in debug
	SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
	SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
	#Persistent ; only one copy running allowed
		
	version := "1.0.0"

	; splash
	SplashTextOn, 200, 75, Call -> Log4OM de IZ3XNJ, Call2Log4OM`nAutomatic callsign to Log4OM`nv%version%`nPlease wait...
	
	; setup tray menu
	Gosub, SetupTrayMenu

	; setup
	gosub, C2L_Setup
		
	if !C2L_IsAppRunning(true)
		; if log4OM is not running, just exit
		ExitApp
	else
	{
		; delay, just to wait LOg4OM is fully loaded
		if (flgDelay)
		{
			SplashTextOff
			MsgBox, 8240, Call -> Log4OM de IZ3XNJ, Waiting for Log4OM`, please click OK when is fully loaded and ready., 60
			SplashTextOn, 200, 75, Call -> Log4OM de IZ3XNJ, Call2Log4OM`nAutomatic callsign to Log4OM`nv%version%`nPlease wait...
		}
		
		; config
		
		gosub, C2L_Config

		if (flgBandMaster)
			; launch & setup BandMaster
			gosub, StartBandMaster

		if (flgClipBoard)
			; setup clipboard event
			OnClipboardChange("C2L_ClipChanged")

		; setup OnExit event
		OnExit, C2L_End
		
		;  setup timer
		SetTimer, C2L_CtrlApps, 5000
				
		; close splash text
		SplashTextOff

	}
	
return


; +-----------+
; | C2L_Setup |
; +-----------+
C2L_Setup:

	; forced windows activation
	#WinActivateForce 
	; the window's title must start with the specified WinTitle to be a match
	SetTitleMatchMode, 1
	; invisible windows are "seen" by the script
	DetectHiddenWindows, On 

	; default command line params
	flgTrayTip := false
	flgBandMaster := false
	flgBMCtrlClick := false
	flgClipBoard := false
	flgNoParam := true
	flgDelay := false
	
	; read command line params
	for n, param in A_Args  
	{
		; convert every  param to upper case
		StringUpper, param, param
		
		if (param = "-D")
			flgDelay := true
				
		if (param = "-TT")
		{
			flgNoParam := false
			flgTrayTip := true
		}
		
		if (param = "-BM")
		{
			flgNoParam := false
			flgBandMaster := true
		}
		
		if (param = "-CC")
		{
			flgNoParam := false
			flgBMCtrlClick := true
		}
		
		if (param = "-CL")
		{
			flgNoParam := false
			flgClipBoard := true
		}
		
	} 
	
	; if there is no params (excluding -D delay param) ....
	; there is no any defined source of callsigns to send to Log4OM...
	; so there is no use in running :-)
	if (flgNoParam)
	{
		SplashTextOff
		gosub, showHelp
		ExitApp
	}
	
	; setup control's name
	lblLog4OM = Log4OM [User Profile:
	
return


; +------------+
; | C2L_Config |
; +------------+
C2L_Config:

	flgStop := false
	; try to connect until you stop it
	While (!flgStop)
	{
		; activates the window  and makes it foremost
		WinActivate, %lblLog4OM% 
		
		; click CLR button to clear previous call if necessary
		ControlClick, CLR, %lblLog4OM%		
		
		; send a stringl
		SendInput {Raw}XNJ
		
		; wait a moment
		Sleep 1000
		
		; read same string to detect control id
		clsNNCall := GetLog4OmCtrl("XNJ")
		if (clsNNCall = "")
		{
			SplashTextOff
			MsgBox, 8229,  Call -> Log4OM de IZ3XNJ, Log4OM did not answer`, do you wanto to retry?
			IfMsgBox, Cancel
			{
				flgStop := true	
				MsgBox, 16, Call -> Log4OM de IZ3XNJ, Call -> Log4OM de IZ3XNJ`nErrore clsNNCall
				ExitApp
			}
		}	
		else
		{
			; click CLR button to clear string from callsign field
			ControlClick, CLR, %lblLog4OM%		

			; clsNNTab
			clsNNTab := GetLog4OmCtrl("QSO Information (F7)")
			if (clsNNTab = "")
			{
				SplashTextOff
				MsgBox, 8229,  Call -> Log4OM de IZ3XNJ, Log4OM did not answer`, do you wanto to retry?
				IfMsgBox, Cancel
				{
					flgStop := true	
					MsgBox, 16,  Call -> Log4OM de IZ3XNJ, Error in clsNNTab
					ExitApp
				}
			}
			else
				flgStop := true
		}
	}

return


; +---------+
; | C2L_End |
; +---------+
C2L_End:

	; stop BandMaster, if started
	if(flgBandMaster)
		Gosub, StopBandMaster
	
	; delete timer
	SetTimer, C2L_CtrlApps, Delete 
	ExitApp

return


; +--------------+
; | C2L_CtrlApps |
; +--------------+
C2L_CtrlApps:

	; check if Log4OM is still running
	if !C2L_IsAppRunning(false)
		goto C2L_End

return


; +------------------+
; | C2L_IsAppRunning |
; +------------------+
C2L_IsAppRunning(bMsg)
{
	
	global	
	
	flgStop := false
	
	if (!bMsg)
	{
		; suspend timer
		SetTimer, C2L_CtrlApps, Off 
	}
	
		; check if Log4OM running
		IfWinNotExist, Log4OM Communicator
		{
			if (bMsg)
			{
				SplashTextOff
				MsgBox, 16, Call -> Log4OM de IZ3XNJ, Log4OM not running
				return false
			}
			else
				return false
		}

	if (!bMsg)
		; restart timer
		SetTimer, C2L_CtrlApps, On 

	return true	
	
}


; +--------------+
; | C2L_Callsign |
; +--------------+
C2L_Callsign(callsign, source)
{
	global 
	
	; convert to upper case
	StringUpper, callsign, callsign
	
	; check if the text in clipboard could be a callsign
	if (C2L_isCallsign(callsign) or source = "bm")
	{			
		; activates the window  and makes it foremost
		WinActivate, %lblLog4OM% 
		
		; read prevoius call
		ControlGetText, prevCall, %clsNNCall%, %lblLog4OM%  

		; only if different
		if (prevCall != callsign)
		{	
			if (flgTrayTip)
				TrayTip, Call -> Log4OM, %callsign%, 40, 17
			
			; click CLR button to clear previous call
			ControlClick, CLR, %lblLog4OM% 
			
			; copy clipboard to the Callsign field
			ControlSetText, %clsNNCall%, %callsign%, %lblLog4OM% 
		}
		
		; QSO Information tab {F7} -> Push QSO Information Tab
		ControlSend, %clsNNTab%, {F7}, %lblLog4OM%  
	}

}


; +-----------------;
; | C2L_ClipChanged |
; +-----------------;
C2L_ClipChanged(Type) 
{
	
	global 

	; suspend timer
	SetTimer, C2L_CtrlApps, Off 
	
	; this event raise up when clipoard changes
	; type = 1  means clipboard contains something that can be expressed as text 
	; (this includes files copied from an Explorer window)
	if Type = 1
		C2L_Callsign(clipboard, "clip")
	
	; restart timer
	SetTimer, C2L_CtrlApps, On 
	
	return
	
}


; +----------------+
; | C2L_isCallsign |
; +----------------+
C2L_isCallsign(call)
{

	; check if the text read from clipboard could be a callsign
	
	; if the clipboard contains tabs or spaces, is not a callsign
	if call contains  %A_Space%, %A_Tab%
		return false
	
	; if it is too long or too short, is not a callsign
	if (StrLen(call) > 13 or StrLen(call) < 3)
		return false
	
	return true
	
}


; +---------------+
; | GetLog4OmCtrl |
; +---------------+
GetLog4OmCtrl(txtLbl)
{
	
	local hwnd
	local controls 
	local txtRead
	
	hWnd := WinExist(lblLog4OM)

	; retrieve all controls in the main window
	WinGet, controls, ControlListHwnd, Log4OM [User Profile:

	; for each control
	Loop, Parse, controls, `n
	{
		; retrieve text from control
		ControlGetText, txtRead,, ahk_id %A_LoopField%
		if (txtRead = txtLbl)
		{
			ctrlName := Control_GetClassNN(hWnd, A_LoopField) 
			break
		}
	}
	
	return ctrlName
	
}


; +--------------------+
; | Control_GetClassNN |
; +--------------------+
Control_GetClassNN(hWnd, hCtrl) 
{
	
	; SKAN: www.autohotkey.com/forum/viewtopic.php?t=49471
	WinGet, CH, ControlListHwnd, ahk_id %hWnd%
	WinGet, CN, ControlList, ahk_id %hWnd%
	Clipboard := CN
	LF:= "`n",  CH:= LF CH LF, CN:= LF CN LF,  S:= SubStr( CH, 1, InStr( CH, LF hCtrl LF ) )
	StringReplace, S, S,`n,`n, UseErrorLevel
	StringGetPos, P, CN, `n, L%ErrorLevel%

	Return SubStr( CN, P+2, InStr( CN, LF, 0, P+2 ) -P-2 )

}


; +---------------+
; | SetupTrayMenu |
; +---------------+
SetupTrayMenu:

	; set tray Icon  & menues
	try
	{
		Menu, Tray, Icon, Call2Log4OM.ico
	}
	catch e
	{
		; close splash text
		SplashTextOff
		MsgBox, 16, Call -> Log4OM de IZ3XNJ, Error %e% in Tray Icon
	}

	Menu, Tray, NoStandard
	Menu, Tray, Add, Config, C2L_Config
	Menu, Tray, Add, Help..., showHelp
	Menu, Tray, Add, Exit, C2L_End
	
return


; +-------------------------------+
; | BandMasterEngine_SpotSelected |
; +-------------------------------+
BandMasterEngine_SpotSelected(bmSpot, bmSelect)
{
	Global
	
	; this event is raised when you click on a spot on band master
	
	; suspend timer
	SetTimer, C2L_CtrlApps, Off 

	bmCall := bmSpot.Call
	
	if (bmSelect <> 0)
		if (!flgBMCtrlClick or bmSelect = 2)
			C2L_Callsign(bmCall, "bm")
	
	; restart timer
	SetTimer, C2L_CtrlApps, On 

}


; +----------+
; | showHelp |
; +----------+
showHelp:
	
	SplashTextOff

	MsgBox, 64,  Call -> Log4OM de IZ3XNJ - v%version%, Params:`n-D : wait for Log4OM loading`n-TT : show calls in tray`n-BM : launch & read calls from BandMaster`n-CC : read calls from BandMaster on Control+Click, else also on Click`n-CL : read calls from Clipboard
	
return


; +----------------+
; | StopBandMaster |
; +----------------+
StopBandMaster:

	; stop BandMaster engine
	BandMasterEngine := ""
	
return


; +-----------------+
; | StartBandMaster |
; +-----------------+
StartBandMaster:

	try
	{
		; start BandMaster engine
		BandMasterEngine := ComObjCreate("BandMaster.BandMasterApp") 

		; Connects events to corresponding script functions with the prefix "BandMasterEngine_".
		ComObjConnect(BandMasterEngine, "BandMasterEngine_")
		
		; show BandMaster main windows as visible
		BandMasterEngine.visible := true
	}
	catch e
	{
		; close splash text
		SplashTextOff
		MsgBox, 16, Call -> Log4OM de IZ3XNJ, Error %e% in StartBandMaster
		flgBandMaster := false
	}

return	
