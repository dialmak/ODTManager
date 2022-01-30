; ODTManager v1.0
; Copyright dialmak 2022
; https://github.com/dialmak/ODTManager

;#NoTrayIcon
#SingleInstance, Ignore
#NoEnv
SetBatchLines, -1
SendMode Input
#Persistent
global req
global Param1
global Param2
global ExeDir
global Parameter
global Activate
global ODTPath
global XMLPath


FileEncoding, UTF-8-RAW

IconFile := A_ScriptDir "\ODTManager.ico"
if FileExist(IconFile)
	Menu, Tray, Icon, %IconFile%

Loop, %0% 
{ 
  Name_Count += 1 
  Param%Name_Count% := %A_Index%
}

if Param1 
	{
	SetWorkingDir, %Param1%
	ExeDir := Param1
	}
else
	{
	SetWorkingDir, %A_ScriptDir%
	ExeDir := A_ScriptDir	
	}	

;MsgBox, % Param1 "`r`n" Param2 "`r`n" A_ScriptDir

full_command_line := DllCall("GetCommandLine", "str")

if not (A_IsAdmin or RegExMatch(full_command_line, " /restart(?!\S)"))
{
    try
    {
        if A_IsCompiled
            Run *RunAs "%A_ScriptFullPath%" /restart
        else
            Run *RunAs "%A_AhkPath%" /restart "%A_ScriptFullPath%" "%Param1%" "%Param2%"
    }
    ExitApp
}	

;MsgBox A_IsAdmin: %A_IsAdmin%`nCommand line: %full_command_line%
	
Settings :=
( LTrim Join Comments
{
	; ODT path
	"ODTPath": "",
	
	; File path for XML file
	"XMLPatch": "",

	; Editor (colors are 0xRRGGBB)
	"FGColor": 0xEEEEDF,
	"BGColor": 0x373A40,
	"TabSize": 2,
	"WordWrap": False,
	"Font": {
		"Typeface": "Consolas",
		"Size": 11,
		"Bold": False
	},
	"Gutter": {
	; Width in pixels. Make this larger when using
	; larger fonts. Set to 0 to disable the gutter.
		"Width": 40,
		"FGColor": 0x9FAFAF,
		"BGColor": 0x262626
	},

	; Highlighter (colors are 0xRRGGBB)
	"UseHighlighter": True,
	"Highlighter": "HighlightXML",
	"HighlightDelay": 200, ; Delay until the user is finished typing
	"Colors": {
		"Comments":  	0x609F60,
		"Punctuation":  0x97C0EB,
		"Strings":      0xF6C6BD,
		"Attributes":   0xFAFCC2,
		"Tags":         0xEEEEEE		
	},

	; Auto-Indenter
	"Indent": "`t",

	; AutoComplete
	"UseAutoComplete": True,
	"ACListRebuildDelay": 500 ; Delay until the user is finished typing
}
)

; Overlay any external settings onto the above defaults
if FileExist("Settings.ini")
{
	ExtSettings := Ini_Load(FileOpen("Settings.ini", "r").Read())
	for k, v in ExtSettings
		if IsObject(v)
			v.base := Settings[k]
	ExtSettings.base := Settings
	Settings := ExtSettings
}

Tester := new ODTManager(Settings)
Tester.RegisterCloseCallback(Func("TesterClose"))

XMLPath := Settings.XMLPath	
if !FileExist(XMLPath)
    XMLPath := "template.xml"
if !FileExist(XMLPath)
	XMLPath := ""

ODTPath := Settings.ODTPath		
if !FileExist(ODTPath)
    ODTPath := "setup.exe"
if !FileExist(ODTPath) 
	DownloadODTRuntime()
if !FileExist(ODTPath)
	ODTPath := ""
	
;MsgBox, % Settings.XMLPatch "`r`n" ODTPath	

return


#If

TesterClose(Tester)
{
    
	ExitApp
}



class RichCode
{
	static Msftedit := DllCall("LoadLibrary", "Str", "Msftedit.dll")
	static IID_ITextDocument := "{8CC497C0-A1DF-11CE-8098-00AA0047BE5D}"
	static MenuItems := ["Cut", "Copy", "Paste", "Delete", "", "Select All", "", "UPPERCASE", "lowercase", "TitleCase", "", "Comment", "Uncomment"]
	
	_Frozen := False
	
	; --- Static Methods ---
	
	BGRFromRGB(RGB)
	{
		return RGB>>16&0xFF | RGB&0xFF00 | RGB<<16&0xFF0000
	}
	
	; --- Properties ---
	
	Value[]
	{
		get {
			GuiControlGet, Code,, % this.hWnd
			return Code
		}
		
		set {
			this.Highlight(Value)
			return Value
		}
	}
	
	; TODO: reserve and reuse memory
	Selection[i:=0]
	{
		get {
			VarSetCapacity(CHARRANGE, 8, 0)
			this.SendMsg(0x434, 0, &CHARRANGE) ; EM_EXGETSEL
			Out := [NumGet(CHARRANGE, 0, "Int"), NumGet(CHARRANGE, 4, "Int")]
			return i ? Out[i] : Out
		}
		
		set {
			if i
				Temp := this.Selection, Temp[i] := Value, Value := Temp
			VarSetCapacity(CHARRANGE, 8, 0)
			NumPut(Value[1], &CHARRANGE, 0, "Int") ; cpMin
			NumPut(Value[2], &CHARRANGE, 4, "Int") ; cpMax
			this.SendMsg(0x437, 0, &CHARRANGE) ; EM_EXSETSEL
			return Value
		}
	}
	
	SelectedText[]
	{
		get {
			Selection := this.Selection, Length := Selection[2] - Selection[1]
			VarSetCapacity(Buffer, (Length + 1) * 2) ; +1 for null terminator
			if (this.SendMsg(0x43E, 0, &Buffer) > Length) ; EM_GETSELTEXT
				throw Exception("Text larger than selection! Buffer overflow!")
			Text := StrGet(&Buffer, Selection[2]-Selection[1], "UTF-16")
			return StrReplace(Text, "`r", "`n")
		}
		
		set {
			this.SendMsg(0xC2, 1, &Value) ; EM_REPLACESEL
			this.Selection[1] -= StrLen(Value)
			return Value
		}
	}
	
	EventMask[]
	{
		get {
			return this._EventMask
		}
		
		set {
			this._EventMask := Value
			this.SendMsg(0x445, 0, Value) ; EM_SETEVENTMASK
			return Value
		}
	}
	
	UndoSuspended[]
	{
		get {
			return this._UndoSuspended
		}
		
		set {
			try ; ITextDocument is not implemented in WINE
			{
				if Value
					this.ITextDocument.Undo(-9999995) ; tomSuspend
				else
					this.ITextDocument.Undo(-9999994) ; tomResume
			}
			return this._UndoSuspended := !!Value
		}
	}
	
	Frozen[]
	{
		get {
			return this._Frozen
		}
		
		set {
			if (Value && !this._Frozen)
			{
				try ; ITextDocument is not implemented in WINE
					this.ITextDocument.Freeze()
				catch
					GuiControl, -Redraw, % this.hWnd
			}
			else if (!Value && this._Frozen)
			{
				try ; ITextDocument is not implemented in WINE
					this.ITextDocument.Unfreeze()
				catch
					GuiControl, +Redraw, % this.hWnd
			}
			return this._Frozen := !!Value
		}
	}
	
	Modified[]
	{
		get {
			return this.SendMsg(0xB8, 0, 0) ; EM_GETMODIFY
		}
		
		set {
			this.SendMsg(0xB9, Value, 0) ; EM_SETMODIFY
			return Value
		}
	}
	
	; --- Construction, Destruction, Meta-Functions ---
	
	__New(Settings, Options:="")
	{
		static Test
		this.Settings := Settings
		FGColor := this.BGRFromRGB(Settings.FGColor)
		BGColor := this.BGRFromRGB(Settings.BGColor)
		
		Gui, Add, Custom, ClassRichEdit50W hWndhWnd +0x5031b1c4 +E0x20000 %Options%
		this.hWnd := hWnd
		
		; Enable WordWrap in RichEdit control ("WordWrap" : true)
		if this.Settings.WordWrap
			SendMessage, 0x0448, 0, 0, , % "ahk_id " . This.HWND
		
		; Register for WM_COMMAND and WM_NOTIFY events
		; NOTE: this prevents garbage collection of
		; the class until the control is destroyed
		this.EventMask := 1 ; ENM_CHANGE
		CtrlEvent := this.CtrlEvent.Bind(this)
		GuiControl, +g, %hWnd%, %CtrlEvent%
		
		; Set background color
		this.SendMsg(0x443, 0, BGColor) ; EM_SETBKGNDCOLOR
		
		; Set character format
		VarSetCapacity(CHARFORMAT2, 116, 0)
		NumPut(116,                    CHARFORMAT2, 0,  "UInt")       ; cbSize      = sizeof(CHARFORMAT2)
		NumPut(0xE0000000,             CHARFORMAT2, 4,  "UInt")       ; dwMask      = CFM_COLOR|CFM_FACE|CFM_SIZE
		NumPut(FGColor,                CHARFORMAT2, 20, "UInt")       ; crTextColor = 0xBBGGRR
		NumPut(Settings.Font.Size*20,  CHARFORMAT2, 12, "UInt")       ; yHeight     = twips
		StrPut(Settings.Font.Typeface, &CHARFORMAT2+26, 32, "UTF-16") ; szFaceName  = TCHAR
		this.SendMsg(0x444, 0, &CHARFORMAT2) ; EM_SETCHARFORMAT
		
		; Set tab size to 4 for non-highlighted code
		VarSetCapacity(TabStops, 4, 0), NumPut(Settings.TabSize*4, TabStops, "UInt")
		this.SendMsg(0x0CB, 1, &TabStops) ; EM_SETTABSTOPS
		
		; Change text limit from 32,767 to max
		this.SendMsg(0x435, 0, -1) ; EM_EXLIMITTEXT
		
		; Bind for keyboard events
		; Use a pointer to prevent reference loop
		this.OnMessageBound := this.OnMessage.Bind(&this)
		OnMessage(0x100, this.OnMessageBound) ; WM_KEYDOWN
		OnMessage(0x205, this.OnMessageBound) ; WM_RBUTTONUP
		
		; Bind the highlighter
		this.HighlightBound := this.Highlight.Bind(&this)
		
		; Create the right click menu
		this.MenuName := this.__Class . &this
		RCMBound := this.RightClickMenu.Bind(&this)
		for Index, Entry in this.MenuItems
			Menu, % this.MenuName, Add, %Entry%, %RCMBound%
		
		; Get the ITextDocument object
		VarSetCapacity(pIRichEditOle, A_PtrSize, 0)
		this.SendMsg(0x43C, 0, &pIRichEditOle) ; EM_GETOLEINTERFACE
		this.pIRichEditOle := NumGet(pIRichEditOle, 0, "UPtr")
		this.IRichEditOle := ComObject(9, this.pIRichEditOle, 1), ObjAddRef(this.pIRichEditOle)
		this.pITextDocument := ComObjQuery(this.IRichEditOle, this.IID_ITextDocument)
		this.ITextDocument := ComObject(9, this.pITextDocument, 1), ObjAddRef(this.pITextDocument)
	}
	
	RightClickMenu(ItemName, ItemPos, MenuName)
	{
		if !IsObject(this)
			this := Object(this)
		
		if (ItemName == "Cut")
			Clipboard := this.SelectedText, this.SelectedText := ""
		else if (ItemName == "Copy")
			Clipboard := this.SelectedText
		else if (ItemName == "Paste")
			this.SelectedText := Clipboard
		else if (ItemName == "Delete")
			this.SelectedText := ""
		else if (ItemName == "Select All")
			this.Selection := [0, -1]
		else if (ItemName == "UPPERCASE")
			this.SelectedText := Format("{:U}", this.SelectedText)
		else if (ItemName == "lowercase")
			this.SelectedText := Format("{:L}", this.SelectedText)
		else if (ItemName == "TitleCase")
			this.SelectedText := Format("{:T}", this.SelectedText)
		else if (ItemName == "Comment")
			this.SelectedText := "<!-- " this.SelectedText " -->"
		else if (ItemName == "Uncomment")	
			this.SelectedText := RegExReplace(this.SelectedText, "s)<!-- ?(.+?) ?-->", "$1")
	}
	
	__Delete()
	{
		; Release the ITextDocument object
		this.ITextDocument := "", ObjRelease(this.pITextDocument)
		this.IRichEditOle := "", ObjRelease(this.pIRichEditOle)
		
		; Release the OnMessage handlers
		OnMessage(0x100, this.OnMessageBound, 0) ; WM_KEYDOWN
		OnMessage(0x205, this.OnMessageBound, 0) ; WM_RBUTTONUP
		
		; Destroy the right click menu
		Menu, % this.MenuName, Delete
		
		HighlightBound := this.HighlightBound
		SetTimer, %HighlightBound%, Delete
	}
	
	; --- Event Handlers ---
	
	OnMessage(wParam, lParam, Msg, hWnd)
	{
		if !IsObject(this)
			this := Object(this)
		if (hWnd != this.hWnd)
			return
		
		if (Msg == 0x100) ; WM_KEYDOWN
		{
			if (wParam == GetKeyVK("Tab"))
			{
				; Indentation
				Selection := this.Selection
				if GetKeyState("Shift")
					this.IndentSelection(True) ; Reverse
				else if (Selection[2] - Selection[1]) ; Something is selected
					this.IndentSelection()
				else
				{
					; TODO: Trim to size needed to reach next TabSize
					this.SelectedText := this.Settings.Indent
					this.Selection[1] := this.Selection[2] ; Place cursor after
				}
				return False
			}
			else if (wParam == GetKeyVK("Escape")) ; Normally closes the window
				return False
			else if (wParam == GetKeyVK("v") && GetKeyState("Ctrl"))
			{
				this.SelectedText := Clipboard ; Strips formatting
				this.Selection[1] := this.Selection[2] ; Place cursor after
				return False
			}
		}
		else if (Msg == 0x205) ; WM_RBUTTONUP
		{
			Menu, % this.MenuName, Show
			return False
		}
	}
	
	CtrlEvent(CtrlHwnd, GuiEvent, EventInfo, _ErrorLevel:="")
	{
		if (GuiEvent == "Normal" && EventInfo == 0x300) ; EN_CHANGE
		{
			; Delay until the user is finished changing the document
			HighlightBound := this.HighlightBound
			SetTimer, %HighlightBound%, % -Abs(this.Settings.HighlightDelay)
		}
	}
	
	; --- Methods ---
	
	; First parameter is taken as a replacement value
	; Variadic form is used to detect when a parameter is given,
	; regardless of content
	Highlight(NewVal*)
	{
		if !IsObject(this)
			this := Object(this)
		if !(this.Settings.UseHighlighter && this.Settings.Highlighter)
		{
			if NewVal.Length()
				GuiControl,, % this.hWnd, % NewVal[1]
			return
		}
		
		; Freeze the control while it is being modified, stop change event
		; generation, suspend the undo buffer, buffer any input events
		PrevFrozen := this.Frozen, this.Frozen := True
		PrevEventMask := this.EventMask, this.EventMask := 0 ; ENM_NONE
		PrevUndoSuspended := this.UndoSuspended, this.UndoSuspended := True
		PrevCritical := A_IsCritical
		Critical, 1000
		
		; Run the highlighter
		Highlighter := this.Settings.Highlighter
		RTF := %Highlighter%(this.Settings, NewVal.Length() ? NewVal[1] : this.Value)
		
		; "TRichEdit suspend/resume undo function"
		; https://stackoverflow.com/a/21206620
		
		; Save the rich text to a UTF-8 buffer
		VarSetCapacity(Buf, StrPut(RTF, "UTF-8"), 0)
		StrPut(RTF, &Buf, "UTF-8")
		
		; Set up the necessary structs
		VarSetCapacity(ZOOM,      8, 0) ; Zoom Level
		VarSetCapacity(POINT,     8, 0) ; Scroll Pos
		VarSetCapacity(CHARRANGE, 8, 0) ; Selection
		VarSetCapacity(SETTEXTEX, 8, 0) ; SetText Settings
		NumPut(1, SETTEXTEX, 0, "UInt") ; flags = ST_KEEPUNDO
		
		; Save the scroll and cursor positions, update the text,
		; then restore the scroll and cursor positions
		MODIFY := this.SendMsg(0xB8, 0, 0)    ; EM_GETMODIFY
		this.SendMsg(0x4E0, &ZOOM, &ZOOM+4)   ; EM_GETZOOM
		this.SendMsg(0x4DD, 0, &POINT)        ; EM_GETSCROLLPOS
		this.SendMsg(0x434, 0, &CHARRANGE)    ; EM_EXGETSEL
		this.SendMsg(0x461, &SETTEXTEX, &Buf) ; EM_SETTEXTEX
		this.SendMsg(0x437, 0, &CHARRANGE)    ; EM_EXSETSEL
		this.SendMsg(0x4DE, 0, &POINT)        ; EM_SETSCROLLPOS
		this.SendMsg(0x4E1, NumGet(ZOOM, "UInt")
		, NumGet(ZOOM, 4, "UInt"))        ; EM_SETZOOM
		this.SendMsg(0xB9, MODIFY, 0)         ; EM_SETMODIFY
		
		; Restore previous settings
		Critical, %PrevCritical%
		this.UndoSuspended := PrevUndoSuspended
		this.EventMask := PrevEventMask
		this.Frozen := PrevFrozen
	}
	
	IndentSelection(Reverse:=False, Indent:="")
	{
		; Freeze the control while it is being modified, stop change event
		; generation, buffer any input events
		PrevFrozen := this.Frozen, this.Frozen := True
		PrevEventMask := this.EventMask, this.EventMask := 0 ; ENM_NONE
		PrevCritical := A_IsCritical
		Critical, 1000
		
		if (Indent == "")
			Indent := this.Settings.Indent
		IndentLen := StrLen(Indent)
		
		; Select back to the start of the first line
		Min := this.Selection[1]
		Top := this.SendMsg(0x436, 0, Min) ; EM_EXLINEFROMCHAR
		TopLineIndex := this.SendMsg(0xBB, Top, 0) ; EM_LINEINDEX
		this.Selection[1] := TopLineIndex
		
		; TODO: Insert newlines using SetSel/ReplaceSel to avoid having to call
		; the highlighter again
		Text := this.SelectedText
		if Reverse
		{
			; Remove indentation appropriately
			Loop, Parse, Text, `n, `r
			{
				if (InStr(A_LoopField, Indent) == 1)
				{
					Out .= "`n" SubStr(A_LoopField, 1+IndentLen)
					if (A_Index == 1)
						Min -= IndentLen
				}
				else
					Out .= "`n" A_LoopField
			}
			this.SelectedText := SubStr(Out, 2)
			
			; Move the selection start back, but never onto the previous line
			this.Selection[1] := Min < TopLineIndex ? TopLineIndex : Min
		}
		else
		{
			; Add indentation appropriately
			Trailing := (SubStr(Text, 0) == "`n")
			Temp := Trailing ? SubStr(Text, 1, -1) : Text
			Loop, Parse, Temp, `n, `r
				Out .= "`n" Indent . A_LoopField
			this.SelectedText := SubStr(Out, 2) . (Trailing ? "`n" : "")
			
			; Move the selection start forward
			this.Selection[1] := Min + IndentLen
		}
		
		this.Highlight()
		
		; Restore previous settings
		Critical, %PrevCritical%
		this.EventMask := PrevEventMask
		
		; When content changes cause the horizontal scrollbar to disappear,
		; unfreezing causes the scrollbar to jump. To solve this, jump back
		; after unfreezing. This will cause a flicker when that edge case
		; occurs, but it's better than the alternative.
		VarSetCapacity(POINT, 8, 0)
		this.SendMsg(0x4DD, 0, &POINT) ; EM_GETSCROLLPOS
		this.Frozen := PrevFrozen
		this.SendMsg(0x4DE, 0, &POINT) ; EM_SETSCROLLPOS
	}
	
	; --- Helper/Convenience Methods ---
	
	SendMsg(Msg, wParam, lParam)
	{
		SendMessage, Msg, wParam, lParam,, % "ahk_id" this.hWnd
		return ErrorLevel
	}
}
GenHighlighterCache(Settings)
{
	if Settings.HasKey("Cache")
		return
	Cache := Settings.Cache := {}
	
	
	; --- Process Colors ---
	Cache.Colors := Settings.Colors.Clone()
	
	; Inherit from the Settings array's base
	BaseSettings := Settings
	while (BaseSettings := BaseSettings.Base)
		for Name, Color in BaseSettings.Colors
			if !Cache.Colors.HasKey(Name)
				Cache.Colors[Name] := Color
	
	; Include the color of plain text
	if !Cache.Colors.HasKey("Plain")
		Cache.Colors.Plain := Settings.FGColor
	
	; Create a Name->Index map of the colors
	Cache.ColorMap := {}
	for Name, Color in Cache.Colors
		Cache.ColorMap[Name] := A_Index
	
	
	; --- Generate the RTF headers ---
	RTF := "{\urtf"
	
	; Color Table
	RTF .= "{\colortbl;"
	for Name, Color in Cache.Colors
	{
		RTF .= "\red"   Color>>16 & 0xFF
		RTF .= "\green" Color>>8  & 0xFF
		RTF .= "\blue"  Color     & 0xFF ";"
	}
	RTF .= "}"
	
	; Font Table
	if Settings.Font
	{
		FontTable .= "{\fonttbl{\f0\fmodern\fcharset0 "
		FontTable .= Settings.Font.Typeface
		FontTable .= ";}}"
		RTF .= "\fs" Settings.Font.Size * 2 ; Font size (half-points)
		if Settings.Font.Bold
			RTF .= "\b"
	}
	
	; Tab size (twips)
	RTF .= "\deftab" GetCharWidthTwips(Settings.Font) * Settings.TabSize
	
	Cache.RTFHeader := RTF
}

GetCharWidthTwips(Font)
{
	static Cache := {}
	
	if Cache.HasKey(Font.Typeface "_" Font.Size "_" Font.Bold)
		return Cache[Font.Typeface "_" font.Size "_" Font.Bold]
	
	; Calculate parameters of CreateFont
	Height := -Round(Font.Size*A_ScreenDPI/72)
	Weight := 400+300*(!!Font.Bold)
	Face := Font.Typeface
	
	; Get the width of "x"
	hDC := DllCall("GetDC", "UPtr", 0)
	hFont := DllCall("CreateFont"
	, "Int", Height ; _In_ int     nHeight,
	, "Int", 0      ; _In_ int     nWidth,
	, "Int", 0      ; _In_ int     nEscapement,
	, "Int", 0      ; _In_ int     nOrientation,
	, "Int", Weight ; _In_ int     fnWeight,
	, "UInt", 0     ; _In_ DWORD   fdwItalic,
	, "UInt", 0     ; _In_ DWORD   fdwUnderline,
	, "UInt", 0     ; _In_ DWORD   fdwStrikeOut,
	, "UInt", 0     ; _In_ DWORD   fdwCharSet, (ANSI_CHARSET)
	, "UInt", 0     ; _In_ DWORD   fdwOutputPrecision, (OUT_DEFAULT_PRECIS)
	, "UInt", 0     ; _In_ DWORD   fdwClipPrecision, (CLIP_DEFAULT_PRECIS)
	, "UInt", 0     ; _In_ DWORD   fdwQuality, (DEFAULT_QUALITY)
	, "UInt", 0     ; _In_ DWORD   fdwPitchAndFamily, (FF_DONTCARE|DEFAULT_PITCH)
	, "Str", Face   ; _In_ LPCTSTR lpszFace
	, "UPtr")
	hObj := DllCall("SelectObject", "UPtr", hDC, "UPtr", hFont, "UPtr")
	VarSetCapacity(SIZE, 8, 0)
	DllCall("GetTextExtentPoint32", "UPtr", hDC, "Str", "x", "Int", 1, "UPtr", &SIZE)
	DllCall("SelectObject", "UPtr", hDC, "UPtr", hObj, "UPtr")
	DllCall("DeleteObject", "UPtr", hFont)
	DllCall("ReleaseDC", "UPtr", 0, "UPtr", hDC)
	
	; Convert to twpis
	Twips := Round(NumGet(SIZE, 0, "UInt")*1440/A_ScreenDPI)
	Cache[Font.Typeface "_" Font.Size "_" Font.Bold] := Twips
	return Twips
}

EscapeRTF(Code)
{
	for each, Char in ["\", "{", "}", "`n"]
		Code := StrReplace(Code, Char, "\" Char)
	return StrReplace(StrReplace(Code, "`t", "\tab "), "`r")
}
;"ID|Description|SourcePath|Version|OfficeClientEdition|Channel|DownloadPath|AllowCdnFallback|MigrateArch|OfficeMgmtCOM|PIDKEY|Fallback|TargetProduct|Level|AcceptEULA|Path|Name|Value|All|Enabled|UpdatePath|TargetVersion|Deadline|IgnoreProduct|Key|Type|App"

HighlightXML(Settings, ByRef Code, Bare:=False)
{
	static Needle := "
	( LTrim Join Comments
		ODims)
		(\<\!--.*?--\>)       ; Comments
		|(<(?:\/\s*)?)(\w+)   ; Tags
		|([<>\/])             ; Punctuation
		|((?<=[>;])[^<>&]+)   ; Text
		|(""[^""]*""|'[^']*') ; Strings
		|(\w+\s*)(=)          ; Attributes
	)"
	
	GenHighlighterCache(Settings)
	Map := Settings.Cache.ColorMap
	
	Pos := 1
	while (FoundPos := RegExMatch(Code, Needle, Match, Pos))
	{
		RTF .= "\cf" Map.Plain " "
		RTF .= EscapeRTF(SubStr(Code, Pos, FoundPos-Pos))
		
		; Flat block of if statements for performance
		if (Match.Value(1) != "")
			RTF .= "\cf" Map.Comments " " EscapeRTF(Match.Value(1))
		else if (Match.Value(2) != "")
		{
			RTF .= "\cf" Map.Punctuation " " EscapeRTF(Match.Value(2))
			RTF .= "\cf" Map.Tags " " EscapeRTF(Match.Value(3))
		}
		else if (Match.Value(4) != "")
			RTF .= "\cf" Map.Punctuation " " Match.Value(4)
		else if (Match.Value(5) != "")
			RTF .= "\cf" Map.Plain " " EscapeRTF(Match.Value(5))
		else if (Match.Value(6) != "")
			RTF .= "\cf" Map.Strings " " EscapeRTF(Match.Value(6))
		else if (Match.Value(7) != "")
		{
			RTF .= "\cf" Map.Attributes " " EscapeRTF(Match.Value(7))
			RTF .= "\cf" Map.Punctuation " " Match.Value(8)
		}
		
		Pos := FoundPos + Match.Len()
	}
	
	if Bare
		return RTF . "\cf" Map.Plain " " EscapeRTF(SubStr(Code, Pos))
	
	return Settings.Cache.RTFHeader . RTF
	. "\cf" Map.Plain " " EscapeRTF(SubStr(Code, Pos)) "\`n}"
}

class ODTManager
{
	static Msftedit := DllCall("LoadLibrary", "Str", "Msftedit.dll")
	Title := "ODTManager"
	
	__New(Settings)
	{
		this.Settings := Settings		
		this.Shell := ComObjCreate("WScript.Shell")		
		this.Bound := []
		this.Bound.RunButton := this.RunButton.Bind(this)
		this.Bound.GuiSize := this.GuiSize.Bind(this)
		this.Bound.OnMessage := this.OnMessage.Bind(this)
		this.Bound.UpdateStatusBar := this.UpdateStatusBar.Bind(this)
		this.Bound.UpdateAutoComplete := this.UpdateAutoComplete.Bind(this)
		this.Bound.CheckIfRunning := this.CheckIfRunning.Bind(this)
		this.Bound.Highlight := this.Highlight.Bind(this)
		this.Bound.SyncGutter := this.SyncGutter.Bind(this)
		
		Buttons := new this.MenuButtons(this)
		this.Bound.Indent := Buttons.Indent.Bind(Buttons)
		this.Bound.Unindent := Buttons.Unindent.Bind(Buttons)
		Menus :=
		( LTrim Join Comments
		[
			["&File", [
				["&Run`tF5", Buttons.Execute.Bind(Buttons)],
				[],
				["&New`tCtrl+N", Buttons.New.Bind(Buttons)],
				["&New Blank`tCtrl+Shift+N", Buttons.NewBlank.Bind(Buttons)],
				["&Open`tCtrl+O", Buttons.Open.Bind(Buttons)],
				["&Open URL `tCtrl+Shift+O", Buttons.Fetch.Bind(Buttons)],
				["&Save`tCtrl+S", Buttons.Save.Bind(Buttons, False)],
				["&Save as`tCtrl+Shift+S", Buttons.Save.Bind(Buttons, True)],
				["Open Working Dir", Buttons.OpenFolder.Bind(Buttons)],
				[],			
				["E&xit`tCtrl+W", this.GuiClose.Bind(this)]
			]], ["&Edit", [
				["Find`tCtrl+F", Buttons.Find.Bind(Buttons)],
				[],
				["Comment Lines`tCtrl+K", Buttons.Comment.Bind(Buttons)],
				["Uncomment Lines`tCtrl+Shift+K", Buttons.Uncomment.Bind(Buttons)],
				[],
				["Indent Lines", this.Bound.Indent],
				["Unindent Lines", this.Bound.Unindent],
				[],
				["&Options", Buttons.ScriptOpts.Bind(Buttons)]
			]], ["&Tools", [
				["Re&indent`tCtrl+I", Buttons.AutoIndent.Bind(Buttons)],
				["&AlwaysOnTop`tAlt+A", Buttons.ToggleOnTop.Bind(Buttons)],
				[],
				["&Highlighter", Buttons.Highlighter.Bind(Buttons)],
				["AutoComplete", Buttons.AutoComplete.Bind(Buttons)]
			]], ["&Help", [
				["Online &Help`tCtrl+H", Buttons.Help.Bind(Buttons)],
				["&About", Buttons.About.Bind(Buttons)]
			]]
		]
		)
		
		Gui, New, +Resize +hWndhMainWindow -AlwaysOnTop
		this.AlwaysOnTop := False
		this.hMainWindow := hMainWindow
		this.Menus := CreateMenus(Menus)
		Gui, Menu, % this.Menus[1]
		
		; If set as default, check the highlighter option
		if this.Settings.UseHighlighter
			Menu, % this.Menus[4], Check, &Highlighter
		
		; If set as default, check the AutoComplete option
		if this.Settings.UseAutoComplete
			Menu, % this.Menus[4], Check, AutoComplete
		
		; Register for events
		WinEvents.Register(this.hMainWindow, this)
		for each, Msg in [0x111, 0x100, 0x101, 0x201, 0x202, 0x204] ; WM_COMMAND, WM_KEYDOWN, WM_KEYUP, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_RBUTTONDOWN
			OnMessage(Msg, this.Bound.OnMessage)
		
		; Add code editor and gutter for line numbers
		this.RichCode := new RichCode(this.Settings, "-E0x20000")
		RichEdit_AddMargins(this.RichCode.hWnd, 3, 3)
		if Settings.Gutter.Width
			this.AddGutter()	
			
		FilePath := Settings.XMLPatch
		if !FileExist(FilePath)
			FilePath := ExeDir "\template.xml"
		
		if (FilePath ~= "^https?://")
			this.RichCode.Value := UrlDownloadToVar(FilePath)
		else if (FilePath = "Clipboard")
			this.RichCode.Value := Clipboard
		else if InStr(FileExist(FilePath), "A")
		{
			this.RichCode.Value := FileOpen(FilePath, "r").Read()
			this.RichCode.Modified := False
			
			if (FilePath == Settings.XMLPatch)
			{
				; Place cursor after the default template text
				this.RichCode.Selection := [0, 0]
			}
			else
			{
				; Keep track of the file currently being edited
				this.FilePath := GetFullPathName(FilePath)
				
				; Follow the directory of the most recently opened file
				SetWorkingDir, %FilePath%\..
			}
		}
		else
			this.RichCode.Value := "<!-- Powered by ODTManager -->`r`n<Configuration>`r`n`r`n</Configuration>"
		
		; Add run button
		Gui, Add, Button, hWndhRunButton, &Run
		this.hRunButton := hRunButton
		BoundFunc := Buttons.Execute.Bind(Buttons)
		GuiControl, +g, %hRunButton%, %BoundFunc%
		
		; Add status bar
		Gui, Add, StatusBar, hWndhStatusBar
		this.UpdateStatusBar()
		ControlGetPos,,,, StatusBarHeight,, ahk_id %hStatusBar%
		this.StatusBarHeight := StatusBarHeight
		
		; Initialize the AutoComplete
		this.AC := new this.AutoComplete(this, this.settings.UseAutoComplete)
		
		this.UpdateTitle()
		Gui, Show, w840 h600
	}
	
	AddGutter()
	{
		s := this.Settings, f := s.Font, g := s.Gutter
		
		; Add the RichEdit control for the gutter
		Gui, Add, Custom, ClassRichEdit50W hWndhGutter +0x5031b1c6 -HScroll -VScroll
		this.hGutter := hGutter
		
		; Set the background and font settings
		FGColor := RichCode.BGRFromRGB(g.FGColor)
		BGColor := RichCode.BGRFromRGB(g.BGColor)
		VarSetCapacity(CF2, 116, 0)
		NumPut(116,        &CF2+ 0, "UInt") ; cbSize      = sizeof(CF2)
		NumPut(0xE<<28,    &CF2+ 4, "UInt") ; dwMask      = CFM_COLOR|CFM_FACE|CFM_SIZE
		NumPut(f.Size*20,  &CF2+12, "UInt") ; yHeight     = twips
		NumPut(FGColor,    &CF2+20, "UInt") ; crTextColor = 0xBBGGRR
		StrPut(f.Typeface, &CF2+26, 32, "UTF-16") ; szFaceName = TCHAR
		SendMessage(0x444, 0, &CF2,    hGutter) ; EM_SETCHARFORMAT
		SendMessage(0x443, 0, BGColor, hGutter) ; EM_SETBKGNDCOLOR
		
		RichEdit_AddMargins(hGutter, 3, 3, -3, 0)
	}
	
	RunButton()
	{
		if !FileExist(ODTPath)
			ODTPath := "setup.exe"
		if !FileExist(ODTPath)
		{
			DownloadODTRuntime()
			return
		}
		if (this.Exec.Status == 0) ; Running
		{	
			;this.Exec.Terminate() ; CheckIfRunning updates the GUI
			Gui, +OwnDialogs
			MsgBox, 304, Warning, Office Deployment Tool runtime is already running.
		}
		else 
		{
			this.Exec := ExecScript(ODTPath, "/" Parameter, this.RichCode.Value)						
			GuiControl,, % this.hRunButton, Please wait ...			
			SetTimer(this.Bound.CheckIfRunning, 240)
		}
	}
	
	CheckIfRunning()
	{
		this.UpdateStatusBar()
		if (this.Exec.Status == 1)
		{
			SetTimer(this.Bound.CheckIfRunning, "Delete")
			GuiControl,, % this.hRunButton, &Run
			if (this.Exec.ExitCode == 0)
			{
				Gui, +OwnDialogs
				MsgBox, 0x40040, Success, Execution successful.
			}
			else 
			{
				Gui, +OwnDialogs
				MsgBox, 0x40010, Error, Execution failed.
			}
					
		}
	}
	
	LoadCode(Code, FilePath:="")
	{
		; Do nothing if nothing is changing
		if (this.FilePath == FilePath && this.RichCode.Value == Code)
			return
		
		; Confirm the user really wants to load new code
		Gui, +OwnDialogs
		MsgBox, 308,  % this.Title " - Confirm Overwrite", Are you sure you want to overwrite your ODT configuration?
		IfMsgBox, No
			return
		
		; If we're changing the open file mark as modified
		; If we're loading a new file mark as unmodified
		this.RichCode.Modified := this.FilePath == FilePath
		this.FilePath := FilePath
		
		; Update the GUI
		this.RichCode.Value := Code
		this.UpdateStatusBar()
	}
	
	OnMessage(wParam, lParam, Msg, hWnd)
	{
		if (hWnd == this.hMainWindow && Msg == 0x111 ; WM_COMMAND
			&& lParam == this.RichCode.hWnd)         ; for RichEdit
		{
			Command := wParam >> 16
			
			if (Command == 0x400) ; An event that fires on scroll
			{
				this.SyncGutter()
				
				; If the user is scrolling too fast it can cause some messages
				; to be dropped. Set a timer to make sure that when the user stops
				; scrolling that the line numbers will be in sync.
				SetTimer(this.Bound.SyncGutter, -50)
			}
			else if (Command == 0x200) ; EN_KILLFOCUS
				if this.Settings.UseAutoComplete
					this.AC.Fragment := ""
		}
		else if (hWnd == this.RichCode.hWnd)
		{
			; Call UpdateStatusBar after the edit handles the keystroke
			SetTimer(this.Bound.UpdateStatusBar, -0)
			
			if this.Settings.UseAutoComplete
			{
				SetTimer(this.Bound.UpdateAutoComplete
				, -Abs(this.Settings.ACListRebuildDelay))
				
				if (Msg == 0x100) ; WM_KEYDOWN
					return this.AC.WM_KEYDOWN(wParam, lParam)
				else if (Msg == 0x201) ; WM_LBUTTONDOWN
					this.AC.Fragment := ""
			}
		}
		else if (hWnd == this.hGutter
			&& {0x100:1,0x101:1,0x201:1,0x202:1,0x204:1}[Msg]) ; WM_KEYDOWN, WM_KEYUP, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_RBUTTONDOWN
		{
			; Disallow interaction with the gutter
			return True
		}
	}
	
	SyncGutter()
	{
		static BUFF, _ := VarSetCapacity(BUFF, 16, 0)
		
		if !this.Settings.Gutter.Width
			return
		
		SendMessage(0x4E0, &BUFF, &BUFF+4, this.RichCode.hwnd) ; EM_GETZOOM
		SendMessage(0x4DD, 0, &BUFF+8, this.RichCode.hwnd)     ; EM_GETSCROLLPOS
		
		; Don't update the gutter unnecessarily
		State := NumGet(BUFF, 0, "UInt") . NumGet(BUFF, 4, "UInt")
		. NumGet(BUFF, 8, "UInt") . NumGet(BUFF, 12, "UInt")
		if (State == this.GutterState)
			return
		
		NumPut(-1, BUFF, 8, "UInt") ; Don't sync horizontal position
		Zoom := [NumGet(BUFF, "UInt"), NumGet(BUFF, 4, "UInt")]
		PostMessage(0x4E1, Zoom[1], Zoom[2], this.hGutter)     ; EM_SETZOOM
		PostMessage(0x4DE, 0, &BUFF+8, this.hGutter)           ; EM_SETSCROLLPOS
		this.ZoomLevel := Zoom[1] / Zoom[2]
		if (this.ZoomLevel != this.LastZoomLevel)
			SetTimer(this.Bound.GuiSize, -0), this.LastZoomLevel := this.ZoomLevel
		
		this.GutterState := State
	}
	
	GetKeywordFromCaret()
	{
		; https://autohotkey.com/boards/viewtopic.php?p=180369#p180369
		static Buffer
		IsUnicode := !!A_IsUnicode
		
		rc := this.RichCode
		sel := rc.Selection
		
		; Get the currently selected line
		LineNum := rc.SendMsg(0x436, 0, sel[1]) ; EM_EXLINEFROMCHAR
		
		; Size a buffer according to the line's length
		Length := rc.SendMsg(0xC1, sel[1], 0) ; EM_LINELENGTH
		VarSetCapacity(Buffer, Length << !!A_IsUnicode, 0)
		NumPut(Length, Buffer, "UShort")
		
		; Get the text from the line
		rc.SendMsg(0xC4, LineNum, &Buffer) ; EM_GETLINE
		lineText := StrGet(&Buffer, Length)
		
		; Parse the line to find the word
		LineIndex := rc.SendMsg(0xBB, LineNum, 0) ; EM_LINEINDEX
		RegExMatch(SubStr(lineText, 1, sel[1]-LineIndex), "[#\w]+$", Start)
		RegExMatch(SubStr(lineText, sel[1]-LineIndex+1), "^[#\w]+", End)
		
		return Start . End
	}
	
	UpdateStatusBar()
	{
		; Delete the timer if it was called by one
		SetTimer(this.Bound.UpdateStatusBar, "Delete")
		
		; Get the document length and cursor position
		VarSetCapacity(GTL, 8, 0), NumPut(1200, GTL, 4, "UInt")
		Len := this.RichCode.SendMsg(0x45F, &GTL, 0) ; EM_GETTEXTLENGTHEX (Handles newlines better than GuiControlGet on RE)
		ControlGet, Row, CurrentLine,,, % "ahk_id" this.RichCode.hWnd
		ControlGet, Col, CurrentCol,,, % "ahk_id" this.RichCode.hWnd
		
		; Get Selected Text Length
		; If the user has selected 1 char further than the end of the document,
		; which is allowed in a RichEdit control, subtract 1 from the length
		Sel := this.RichCode.Selection
		Sel := Sel[2] - Sel[1] - (Sel[2] > Len)

		MessageStatus := "Ready"
		; Get the syntax status
		if (this.Exec.Status == 0)
		{
			MessageStatus := "Office Deployment Tool runtime running"
		} 
		else 
		{
			if this.Exec.ExitCode>=0
				MessageStatus := "Ready"
		}
		
		this.MessageStatus := MessageStatus		
		; Update the Status Bar text
		Gui, % this.hMainWindow ":Default"
		SB_SetText("Len " Len ", Line " Row ", Col " Col
		. (Sel > 0 ? ", Sel " Sel : "") "     " this.MessageStatus)
		
		; Update the title Bar
		this.UpdateTitle()
		
		; Update the gutter to match the document
		if this.Settings.Gutter.Width
		{
			ControlGet, Lines, LineCount,,, % "ahk_id" this.RichCode.hWnd
			if (Lines != this.LineCount)
			{
				Loop, %Lines%
					Text .= A_Index "`n"
				GuiControl,, % this.hGutter, %Text%
				this.SyncGutter()
				this.LineCount := Lines
			}
		}
	}
	
	UpdateTitle()
	{
		Title := this.Title
		
		; Show the current file name
		if this.FilePath
		{
			SplitPath, % this.FilePath, FileName
			Title .= " - " FileName
		}
		
		; Show the curernt modification status
		if this.RichCode.Modified
			Title .= "*"
		
		; Return if the title doesn't need to be updated
		if (Title == this.VisibleTitle)
			return
		this.VisibleTitle := Title
		
		HiddenWindows := A_DetectHiddenWindows
		DetectHiddenWindows, On
		WinSetTitle, % "ahk_id" this.hMainWindow,, %Title%
		DetectHiddenWindows, %HiddenWindows%
	}
	
	UpdateAutoComplete()
	{
		; Delete the timer if it was called by one
		SetTimer(this.Bound.UpdateAutoComplete, "Delete")
		
		this.AC.BuildWordList()
	}
	
	RegisterCloseCallback(CloseCallback)
	{
		this.CloseCallback := CloseCallback
	}
	
	GuiSize()
	{
		static RECT, _ := VarSetCapacity(RECT, 16, 0)
		if A_Gui
			gw := A_GuiWidth, gh := A_GuiHeight
		else
		{
			DllCall("GetClientRect", "UPtr", this.hMainWindow, "Ptr", &RECT, "UInt")
			gw := NumGet(RECT, 8, "Int"), gh := NumGet(RECT, 12, "Int")
		}
		gtw := 3 + Round(this.Settings.Gutter.Width) * (this.ZoomLevel ? this.ZoomLevel : 1), sbh := this.StatusBarHeight
		GuiControl, Move, % this.RichCode.hWnd, % "x" 0+gtw "y" 0         "w" gw-gtw "h" gh-28-sbh
		if this.Settings.Gutter.Width
			GuiControl, Move, % this.hGutter  , % "x" 0     "y" 0         "w" gtw    "h" gh-28-sbh
		GuiControl, Move, % this.hRunButton   , % "x" 0     "y" gh-28-sbh "w" gw     "h" 28
	}
	
	GuiDropFiles(hWnd, Files)
	{
		; TODO: support multiple file drop
		this.LoadCode(FileOpen(Files[1], "r").Read(), Files[1])
	}
		
	GuiClose()
	{
		
		if (this.Exec.Status == 0) ; Running
		{
			;SetTimer(this.Bound.CheckIfRunning, "Delete")
			Gui, +OwnDialogs
			MsgBox, 308, % this.Title " - Confirm Exit", Are you sure you want to exit?`r`nODT runtime is already running.`r`nPostprocessing fails after closing.
			IfMsgBox, No
				return true
			;this.Exec.Terminate()
		}
		
		if this.RichCode.Modified
		{
			Gui, +OwnDialogs
			MsgBox, 308, % this.Title " - Confirm Exit", There are unsaved changes.`r`nAre you sure you want to exit?
			IfMsgBox, No
				return true
		}
		
		; Free up the AC class
		this.AC := ""
		
		; Release wm_message hooks
		for each, Msg in [0x100, 0x201, 0x202, 0x204] ; WM_KEYDOWN, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_RBUTTONDOWN
			OnMessage(Msg, this.Bound.OnMessage, 0)
		
		; Delete timers
		SetTimer(this.Bound.SyncGutter, "Delete")
		SetTimer(this.Bound.GuiSize, "Delete")
		
		; Break all the BoundFunc circular references
		this.Delete("Bound")
		
		; Release WinEvents handler
		WinEvents.Unregister(this.hMainWindow)
		
		; Release GUI window and control glabels
		Gui, Destroy
		
		; Release menu bar (Has to be done after Gui, Destroy)
		for each, MenuName in this.Menus
			Menu, %MenuName%, DeleteAll
		
		this.CloseCallback()
	}
	
	class Execute
	{
		__New(Parent)
		{
			this.Parent := Parent
			
			ParentWnd := this.Parent.hMainWindow
			Gui, New, +Owner%ParentWnd% +ToolWindow +hWndhWnd
			this.hWnd := hWnd
			;Gui, Margin, 7, 7
			
			; 0x200 for vcenter
			Gui, Add, Text, w180 h22, Select operation:
			
			Gui, Add, Checkbox, hWndhWnd w180 h22 Checked, ODTManagerStart.cmd
			this.hODTManagerStart := hWnd			
			
			Gui, Add, DropDownList, hWndhWnd w180, Configure||Download|Customize
			this.hParameter := hWnd
			BoundHide := this.Hide.Bind(this)
			GuiControl, +g, % this.hParameter, %BoundHide%
						
			Gui, Add, Checkbox, hWndhWnd w180 h22 Checked, ODTManagerFinish.cmd
			this.hODTManagerFinish := hWnd
			
			Gui, Add, Button, hWndhWnd w120 h120 ys-1 Default, Execute
			this.hButton := hWnd
			BoundRun := this.Run.Bind(this)
			GuiControl, +g, % this.hButton, %BoundRun%
			
			Gui, Show,, % this.Parent.Title " - Select and Confirm"
						
			WinEvents.Register(this.hWnd, this)
		}
		
		GuiClose()
		{
			GuiControl, -g, % this.hButton
			WinEvents.Unregister(this.hWnd)
			Gui, Destroy
		}
		
		GuiEscape()
		{
			this.GuiClose()
		}
		
		Hide()
		{
			GuiControlGet, Parameter,, % this.hParameter
			if (Parameter == "Configure") 
			{
				GuiControl,  Show, % this.hODTManagerStart
				GuiControl,  Show, % this.hODTManagerFinish			
				GuiControl,  Enable, % this.hODTManagerStart
				GuiControl,  Enable, % this.hODTManagerFinish	
			}	
			else
			{
				GuiControl,  Hide, % this.hODTManagerStart
				GuiControl,  Hide, % this.hODTManagerFinish	
				GuiControl,  Disable, % this.hODTManagerStart
				GuiControl,  Disable, % this.hODTManagerFinish					
			}				
		}
		
		Run()
		{
			GuiControlGet, Parameter,, % this.hParameter
			GuiControlGet, ODTManagerStart,, % this.hODTManagerStart
			GuiControlGet, ODTManagerFinish,, % this.hODTManagerFinish
			;MsgBox, % Parameter "`r`n" ODTManagerStart "`r`n" ODTManagerFinish
			this.GuiClose()
			;this.Parent.Bound.RunButton()
		}
	}
	

	class Find
	{
		__New(Parent)
		{
			this.Parent := Parent
			
			ParentWnd := this.Parent.hMainWindow
			Gui, New, +Owner%ParentWnd% +ToolWindow +hWndhWnd
			this.hWnd := hWnd
			Gui, Margin, 5, 5
			
			
			; Search
			Gui, Add, Edit, hWndhWnd w200
			SendMessage, 0x1501, True, &cue := "Search Text",, ahk_id %hWnd% ; EM_SETCUEBANNER
			this.hNeedle := hWnd
			
			Gui, Add, Button, yp-1 x+m w75 Default hWndhWnd, Find Next
			Bound := this.BtnFind.Bind(this)
			GuiControl, +g, %hWnd%, %Bound%
			
			Gui, Add, Button, yp x+m w75 hWndhWnd, Coun&t All
			Bound := this.BtnCount.Bind(this)
			GuiControl, +g, %hWnd%, %Bound%
			
			
			; Replace
			Gui, Add, Edit, hWndhWnd w200 xm Section
			SendMessage, 0x1501, True, &cue := "Replacement",, ahk_id %hWnd% ; EM_SETCUEBANNER
			this.hReplace := hWnd
			
			Gui, Add, Button, yp-1 x+m w75 hWndhWnd, &Replace
			Bound := this.Replace.Bind(this)
			GuiControl, +g, %hWnd%, %Bound%
			
			Gui, Add, Button, yp x+m w75 hWndhWnd, Replace &All
			Bound := this.ReplaceAll.Bind(this)
			GuiControl, +g, %hWnd%, %Bound%
			
			
			; Options
			Gui, Add, Checkbox, hWndhWnd xm, &Case Sensitive
			this.hOptCase := hWnd
			Gui, Add, Checkbox, hWndhWnd, Re&gular Expressions
			this.hOptRegEx := hWnd
			;Gui, Add, Checkbox, hWndhWnd, Transform`, &Deref
			;this.hOptDeref := hWnd
			
			
			Gui, Show,, % this.Parent.Title " - Find"
			
			WinEvents.Register(this.hWnd, this)
		}
		
		GuiClose()
		{
			GuiControl, -g, % this.hButton
			WinEvents.Unregister(this.hWnd)
			Gui, Destroy
		}
		
		GuiEscape()
		{
			this.GuiClose()
		}
		
		GetNeedle()
		{
			Opts := this.Case ? "`n" : "i`n"
			Opts .= this.Needle ~= "^[^\(]\)" ? "" : ")"
			if this.RegEx
				return Opts . this.Needle
			else
				return Opts "\Q" StrReplace(this.Needle, "\E", "\E\\E\Q") "\E"
		}
		
		Find(StartingPos:=1, WrapAround:=True)
		{
			Needle := this.GetNeedle()
			
			; Search from StartingPos
			NextPos := RegExMatch(this.Haystack, Needle, Match, StartingPos)
			
			; Search from the top
			if (!NextPos && WrapAround)
				NextPos := RegExMatch(this.Haystack, Needle, Match)
			
			return NextPos ? [NextPos, NextPos+StrLen(Match)] : False
		}
		
		Submit()
		{
			; Options
			;GuiControlGet, Deref,, % this.hOptDeref
			GuiControlGet, Case,, % this.hOptCase
			this.Case := Case
			GuiControlGet, RegEx,, % this.hOptRegEx
			this.RegEx := RegEx
			
			; Search Text/Needle
			GuiControlGet, Needle,, % this.hNeedle
			if Deref
				Transform, Needle, Deref, %Needle%
			this.Needle := Needle
			
			; Replacement
			GuiControlGet, Replace,, % this.hReplace
			if Deref
				Transform, Replace, Deref, %Replace%
			this.Replace := Replace
			
			; Haystack
			this.Haystack := StrReplace(this.Parent.RichCode.Value, "`r")
		}
		
		BtnFind()
		{
			Gui, +OwnDialogs
			this.Submit()
			
			; Find and select the item or error out
			if (Pos := this.Find(this.Parent.RichCode.Selection[1]+2))
				this.Parent.RichCode.Selection := [Pos[1] - 1, Pos[2] - 1]
			else
				MsgBox, 0x30, % this.Parent.Title " - Find", Search text not found
		}
		
		BtnCount()
		{
			Gui, +OwnDialogs
			this.Submit()
			
			; Find and count all instances
			Count := 0, Start := 1
			while (Pos := this.Find(Start, False))
				Start := Pos[1]+1, Count += 1
			
			MsgBox, 0x40, % this.Parent.Title " - Find", %Count% instances found
		}
		
		Replace()
		{
			this.Submit()
			
			; Get the current selection
			Sel := this.Parent.RichCode.Selection
			
			; Find the next occurrence including the current selection
			Pos := this.Find(Sel[1]+1)
			
			; If the found item is already selected
			if (Sel[1]+1 == Pos[1] && Sel[2]+1 == Pos[2])
			{
				; Replace it
				this.Parent.RichCode.SelectedText := this.Replace
				
				; Update the haystack to include the replacement
				this.Haystack := StrReplace(this.Parent.RichCode.Value, "`r")
				
				; Find the next item *not* including the current selection
				Pos := this.Find(Sel[1]+StrLen(this.Replace)+1)
			}
			
			; Select the next found item or error out
			if Pos
				this.Parent.RichCode.Selection := [Pos[1] - 1, Pos[2] - 1]
			else
				MsgBox, 0x30, % this.Parent.Title " - Find", No more instances found
		}
		
		ReplaceAll()
		{
			rc := this.Parent.RichCode
			this.Submit()
			
			Needle := this.GetNeedle()
			
			; Replace the text in a way that pushes to the undo buffer
			rc.Frozen := True
			Sel := rc.Selection
			rc.Selection := [0, -1]
			rc.SelectedText := RegExReplace(this.Haystack, Needle, this.Replace, Count)
			rc.Selection := Sel
			rc.Frozen := False
			
			MsgBox, 0x40, % this.Parent.Title " - Find", %Count% instances replaced
		}
	}
	class ScriptOpts
	{
		__New(Parent)
		{
			this.Parent := Parent
			
			; Create a GUI
			ParentWnd := this.Parent.hMainWindow
			Gui, New, +Owner%ParentWnd% +ToolWindow +hWndhWnd
			this.hWnd := hWnd
			WinEvents.Register(this.hWnd, this)
			
			; Add ODT path button
			Gui, Add, Button, xm ym w95 hWndhButton, ODT Path
			BoundSelectFile := this.SelectFile.Bind(this)
			GuiControl, +g, %hButton%, %BoundSelectFile%
			
			; Add ODT path visualization field
			Gui, Add, Edit, ym w300 ReadOnly hWndhODTPath, % this.Parent.Settings.ODTPath
			this.hODTPath := hODTPath
			GuiControl,, % this.hODTPath, %ODTPath%
			
			; Add XML path button
			Gui, Add, Button, xm w95 hWndhXPButton, XML Path
			BoundSelectXMLFile := this.SelectXMLFile.Bind(this)
			GuiControl, +g, %hXPButton%, %BoundSelectXMLFile%
			
			; Add XML path visualization field
			Gui, Add, Edit, x+m w300 ReadOnly hWndhXMLPath, % this.Parent.Settings.XMLPath
			this.hXMLPath := hXMLPath
			GuiControl,, % this.hXMLPath, %XMLPath%

			; Show the GUI
			Gui, Show,, % this.Parent.Title " - Options"
		}
		
		SelectFile()
		{
			GuiControlGet, ODTPath,, % this.hODTPath
			FileSelectFile, ODTPath, 1, %ODTPath%, Select an ODT file, Executables (*.exe)
			if !ODTPath
				return
			this.Parent.Settings.ODTPath := ODTPath
			GuiControl,, % this.hODTPath, %ODTPath%
		}

		SelectXMLFile()
		{
			GuiControlGet, XMLPath,, % this.hXMLPath
			FileSelectFile, XMLPath, 1, %XMLPath%, Select an XML file, XML (*.xml)
			if !XMLPath
				return
			this.Parent.Settings.XMLPath := XMLPath
			GuiControl,, % this.hXMLPath, %XMLPath%
		}
		
		GuiClose()
		{
			WinEvents.Unregister(this.hWnd)
			Gui, Destroy
		}
		
		GuiEscape()
		{
			this.GuiClose()
		}
	}
	class MenuButtons
	{
		__New(Parent)
		{
			this.Parent := Parent
		}
		
		Save(SaveAs)
		{
			if (SaveAs || !this.Parent.FilePath)
			{
				Gui, +OwnDialogs
				FileSelectFile, FilePath, S18, configuration.xml, % this.Parent.Title " - Save XML file", XML file (*.xml)
				if ErrorLevel
					return
				IfNotInString, FilePath, .xml
					FilePath = %FilePath%.xml
				this.Parent.FilePath := FilePath
			}
			
			FileOpen(this.Parent.FilePath, "w").Write(this.Parent.RichCode.Value)
			
			this.Parent.RichCode.Modified := False
			this.Parent.UpdateStatusBar()
		}
		
		
		Open()
		{
			Gui, +OwnDialogs
			FileSelectFile, FilePath, 3,, % this.Parent.Title " - Open XML file", XML file (*.xml)
			if ErrorLevel
				return
			this.Parent.LoadCode(FileOpen(FilePath, "r").Read(), FilePath)
			
			; Follow the directory of the most recently opened file
			SetWorkingDir, %FilePath%\..
			this.Parent.ScriptOpts.UpdateFields()
		}
		
		OpenFolder()
		{
			Run, explorer.exe "%A_WorkingDir%"
		}
		
		New()
		{
			;Run *RunAs "%A_AhkPath%" /restart "%A_ScriptFullPath%"
			if Param1	
				ExeDir := Param1
			else
				ExeDir := A_ScriptDir

			FilePath := Settings.XMLPatch
			if !FileExist(FilePath)
				FilePath := ExeDir "\template.xml"
			if !FileExist(FilePath) {
				this.Parent.LoadCode("<!-- Powered by ODTManager -->`r`n<Configuration>`r`n`r`n</Configuration>`r`n")
				this.Parent.RichCode.Selection := [0, 0] 
			} else {
				this.Parent.LoadCode(FileOpen(ExeDir "\template.xml", "r").Read(), ExeDir "\template.xml")
			}				
		}
		
		NewBlank()
		{
			this.Parent.LoadCode("<!-- Powered by ODTManager -->`r`n<Configuration>`r`n`r`n</Configuration>`r`n")
			this.Parent.RichCode.Selection := [0, 0]
		}
		
		Execute()
		{ ; TODO: Recycle PubInstance
			if WinExist("ahk_id" this.PubInstance.hWnd)
				WinActivate, % "ahk_id" this.PubInstance.hWnd
			else
				this.PubInstance := new this.Parent.Execute(this.Parent)
		}
		
		Fetch()
		{
			Gui, +OwnDialogs
			InputBox, Url, % this.Parent.Title " - Open XML file from URL", Enter a URL to fetch ODT configuration from Internet.,,,130
			if (Url := Trim(Url))
				this.Parent.LoadCode(UrlDownloadToVar(Url))
		}
		
		Find()
		{ ; TODO: Recycle FindInstance
			if WinExist("ahk_id" this.FindInstance.hWnd)
				WinActivate, % "ahk_id" this.FindInstance.hWnd
			else
				this.FindInstance := new this.Parent.Find(this.Parent)
		}
		
		ScriptOpts()
		{
			if WinExist("ahk_id" this.Parent.ScriptOptsInstance.hWnd)
				WinActivate, % "ahk_id" this.Parent.ScriptOptsInstance.hWnd
			else
				this.Parent.ScriptOptsInstance := new this.Parent.ScriptOpts(this.Parent)
		}
		
		ToggleOnTop()
		{
			if (this.Parent.AlwaysOnTop := !this.Parent.AlwaysOnTop)
			{
				Menu, % this.Parent.Menus[4], Check, &AlwaysOnTop`tAlt+A
				Gui, +AlwaysOnTop
			}
			else
			{
				Menu, % this.Parent.Menus[4], Uncheck, &AlwaysOnTop`tAlt+A
				Gui, -AlwaysOnTop
			}
		}
		
		Highlighter()
		{
			if (this.Parent.Settings.UseHighlighter := !this.Parent.Settings.UseHighlighter)
				Menu, % this.Parent.Menus[4], Check, &Highlighter
			else
				Menu, % this.Parent.Menus[4], Uncheck, &Highlighter
			
			; Force refresh the code, adding/removing any highlighting
			this.Parent.RichCode.Value := this.Parent.RichCode.Value
		}
		
		
		AutoIndent()
		{
			this.Parent.LoadCode(AutoIndent(this.Parent.RichCode.Value
			, this.Parent.Settings.Indent), this.Parent.FilePath)
		}
		
		Help()
		{
;			HelpFile.Open(this.Parent.GetKeywordFromCaret())
		}
		
		About()
		{
			Gui, +OwnDialogs
			MsgBox,, % this.Parent.Title " - About", ODTManager written by dialmak
		}
		
		Comment()
		{
		this.Parent.RichCode.SelectedText := "<!-- " this.Parent.RichCode.SelectedText " -->"		
		}
		
		Uncomment()
		{
		this.Parent.RichCode.SelectedText := RegExReplace(this.Parent.RichCode.SelectedText, "s)<!-- ?(.+?) ?-->", "$1")		
		}
		
		Indent()
		{
			this.Parent.RichCode.IndentSelection()
		}
		
		Unindent()
		{
			this.Parent.RichCode.IndentSelection(True)
		}
		
		AutoComplete()
		{
			if (this.Parent.Settings.UseAutoComplete := !this.Parent.Settings.UseAutoComplete)
				Menu, % this.Parent.Menus[4], Check, AutoComplete
			else
				Menu, % this.Parent.Menus[4], Uncheck, AutoComplete
			
			this.Parent.AC.Enabled := this.Parent.Settings.UseAutoComplete
		}
	}
	/*
		Implements functionality necessary for AutoCompletion of keywords in the
		RichCode control. Currently works off of values stored in the provided
		Parent object, but could be modified to work off a provided RichCode
		instance directly.
		
		The class is mostly self contained and could be easily extended to other
		projects, and even other types of controls. The main method of interacting
		with the class is by passing it WM_KEYDOWN messages. Another way to interact
		is by modifying the Fragment property, especially to clear it when you want
		to cancel autocompletion.
		
		Depends on CQT.ahk, and optionally on HelpFile.ahk
	*/
	
	class AutoComplete
	{
		; Maximum number of suggestions to be displayed in the dialog
		static MaxSuggestions := 20
		
		; Minimum length for a word to be entered into the word list
		static MinWordLen := 3
		
		; Minimum length of fragment before suggestions should be displayed
		static MinSuggestLen := 2
		
		; Stores the initial caret position for newly typed fragments
		static CaretX := 0, CaretY := 0
		
		
		; --- Properties ---
		
		Fragment[]
		{
			get
			{
				return this._Fragment
			}
			
			set
			{
				this._Fragment := Value
				
				; Give suggestions when a fragment of sufficient
				; length has been provided
				if (StrLen(this._Fragment) >= 2)
					this._Suggest()
				else
					this._Hide()
				
				return this._Fragment
			}
		}
		
		Enabled[]
		{
			get
			{
				return this._Enabled
			}
			
			set
			{
				this._Enabled := Value
				if (Value)
					this.BuildWordList()
				else
					this.Fragment := ""
				return Value
			}
		}
		
		
		; --- Constructor, Destructor ---
		
		__New(Parent, Enabled:=True)
		{
			this.Parent := Parent
			this.Enabled := Enabled
			this.WineVer := DllCall("ntdll.dll\wine_get_version", "AStr")
			
			; Create the tool GUI for the floating list
			hParentWnd := this.Parent.hMainWindow
			Gui, +hWndhDefaultWnd
			Relation := this.WineVer ? "Parent" Parent.RichCode.hWnd : "Owner" Parent.hMainWindow
			Gui, New, +%Relation% -Caption +ToolWindow +hWndhWnd
			this.hWnd := hWnd
			Gui, Margin, 0, 0
			
			; Create the ListBox control withe appropriate font and styling
			Font := this.Parent.Settings.Font
			Gui, Font, % "s" Font.Size, % Font.Typeface
			Gui, Add, ListBox, x0 y0 r1 0x100 AltSubmit hWndhListBox, Item
			this.hListBox := hListBox
			
			; Finish GUI creation and restore the default GUI
			Gui, Show, Hide, % this.Parent.Title " - AutoComplete"
			Gui, %hDefaultWnd%:Default
			
			; Get relevant dimensions of the ListBox for later resizing
			SendMessage, 0x1A1, 0, 0,, % "ahk_id" this.hListBox ; LB_GETITEMHEIGHT
			this.ListBoxItemHeight := ErrorLevel
			VarSetCapacity(ListBoxRect, 16, 0)
			DllCall("User32.dll\GetClientRect", "Ptr", this.hListBox, "Ptr", &ListBoxRect)
			this.ListBoxMargins := NumGet(ListBoxRect, 12, "Int") - this.ListBoxItemHeight
			
			; Set up the GDI Device Context for later text measurement in _GetWidth
			this.hDC := DllCall("GetDC", "UPtr", this.hListBox, "UPtr")
			SendMessage, 0x31, 0, 0,, % "ahk_id" this.hListBox ; WM_GETFONT
			this.hFont := DllCall("SelectObject", "UPtr", this.hDC, "UPtr", ErrorLevel, "UPtr")
			
			; Record the total screen width for later user. If the monitors get
			; rearranged while the script is still running this value will be
			; inaccurate. However, this will not likely be a significant issue,
			; and the issues caused by it would be minimal.
			SysGet, ScreenWidth, 78
			this.ScreenWidth := ScreenWidth
			
			; List of default words.
			ODTChannel := "BetaChannel|CurrentPreview|Current|MonthlyEnterprise|SemiAnnualPreview|SemiAnnual|InsiderFast|Insiders|Monthly|Targeted|Broad|PerpetualVL2019"
			OfficeProductID := "O365BusinessRetail|O365SmallBusPremRetail|O365ProPlusRetail|O365HomePremRetail|O365EduCloudRetail|ProPlus2019Volume|ProPlus2019Retail|Standard2019Volume|Standard2019Retail|Professional2019Retail|HomeBusiness2019Retail|HomeStudent2019Retail|Personal2019Retail|MondoVolume|MondoRetail|ProPlusVolume|ProPlusRetail|StandardVolume|StandardRetail|ProfessionalPipcRetail|ProfessionalRetail|HomeBusinessPipcRetail|HomeBusinessRetail|HomeStudentRetail|HomeStudentVNextRetail|PersonalPipcRetail|PersonalRetail"
			OtherOfficeProductID := "LanguagePack|AccessRetail|AccessRuntimeRetail|Access2019Retail|Access2019Volume|ExcelRetail|Excel2019Retail|Excel2019Volume|OneNoteRetail|OutlookRetail|Outlook2019Retail|Outlook2019Volume|PowerPointRetail|PowerPoint2019Retail|PowerPoint2019Volume|ProjectProXVolume|ProjectPro2019Retail|ProjectPro2019Volume|ProjectProRetail|ProjectStdRetail|ProjectStdXVolume|ProjectStd2019Retail|ProjectStd2019Volume|PublisherRetail|Publisher2019Retail|Publisher2019Volume|SkypeforBusinessRetail|SkypeforBusinessEntryRetail|SkypeforBusiness2019Retail|SkypeforBusiness2019Volume|SkypeforBusinessEntry2019Retail|VisioProRetail|VisioProXVolume|VisioPro2019Retail|VisioPro2019Volume|VisioStdRetail|VisioStdXVolume|VisioStd2019Retail|VisioStd2019Volume|WordRetail|Word2019Retail|Word2019Volume"
			ExcludeAppID := "Access|Excel|Groove|Lync|OneDrive|OneNote|Outlook|PowerPoint|Publisher|Teams|Word"
			ODTKeywords := "ID|Configuration|Info|Description|Add|Remove|SourcePath|Version|OfficeClientEdition|Channel|DownloadPath|AllowCdnFallback|MigrateArch|OfficeMgmtCOM|Product|PIDKEY|Language|ProofingTools|ExcludeApp|Property|Name|AUTOACTIVATE|FORCEAPPSHUTDOWN|PACKAGEGUID|SharedComputerLicensing|SCLCacheOverride|SCLCacheOverrideDirectory|DeviceBasedLicensing|PinIconsToTaskBar|RemoveMSI|IgnoreProduct|Updates|Enabled|UpdatePath|TargetVersion|Deadline|Display|Level|AcceptEULA|Logging|Level|Path|True|False|Value|None|All|Full|Off|Standard|AppSettings|Type|App"
			LanguageID := "MatchOS|MatchPreviousMSI|MatchInstalled|Fallback|TargetProduct"
			this.DefaultWordList := ""
			this.DefaultWordList .= "|" ODTKeywords "|" OfficeProductID "|" OtherOfficeProductID "|" ODTChannel "|" ExcludeAppID "|" LanguageID
			; MsgBox % this.DefaultWordList
			this.BuildWordList()
		}
		
		__Delete()
		{
			Gui, % this.hWnd ":Destroy"
			this.Visible := False
			DllCall("SelectObject", "UPtr", this.hDC, "UPtr", this.hFont, "UPtr")
			DllCall("ReleaseDC", "UPtr", this.hListBox, "UPtr", this.hDC)
		}
		
		
		; --- Private Methods ---
		
		; Gets the pixel-based width of a provided text snippet using the GDI font
		; selected into the ListBox control
		_GetWidth(Text)
		{
			MaxWidth := 0
			Loop, Parse, Text, |
			{
				DllCall("GetTextExtentPoint32", "UPtr", this.hDC, "Str", A_LoopField
				, "Int", StrLen(A_LoopField), "Int64*", Size), Size &= 0xFFFFFFFF
				
				if (Size > MaxWidth)
					MaxWidth := Size
			}
			
			return MaxWidth
		}
		
		; Shows the suggestion dialog with contents of the provided DisplayList
		_Show(DisplayList)
		{
			; Insert the new list
			GuiControl,, % this.hListBox, %DisplayList%
			GuiControl, Choose, % this.hListBox, 1
			
			; Resize to fit contents
			StrReplace(DisplayList, "|",, Rows)
			Height := Rows * this.ListBoxItemHeight + this.ListBoxMargins
			Width := this._GetWidth(DisplayList) + 10
			GuiControl, Move, % this.hListBox, w%Width% h%Height%
			
			; Keep the dialog from running off the screen
			X := this.CaretX, Y := this.CaretY + 20
			if ((X + Width) > this.ScreenWidth)
				X := this.ScreenWidth - Width
			
			; Make the dialog visible
			Gui, % this.hWnd ":Show", x%X% y%Y% AutoSize NoActivate
			this.Visible := True
		}
		
		; Hides the dialog if it is visible
		_Hide()
		{
			if !this.Visible
				return
			
			Gui, % this.hWnd ":Hide"
			this.Visible := False
		}
		
		; Filters the word list for entries starting with the fragment, then
		; shows the dialog with the filtered list as suggestions
		_Suggest()
		{
			; Filter the list for words beginning with the fragment
			Suggestions := LTrim(RegExReplace(this.WordList
			, "i)\|(?!" this.Fragment ")[^\|]+"), "|")
			
			; Fail out if there were no matches
			if !Suggestions
				return true, this._Hide()
			
			; Pull the first MaxSuggestions suggestions
			if (Pos := InStr(Suggestions, "|",,, this.MaxSuggestions))
				Suggestions := SubStr(Suggestions, 1, Pos-1)
			this.Suggestions := Suggestions
			
			this._Show("|" Suggestions)
		}
		
		; Finishes the fragment with the selected suggestion
		_Complete()
		{
			; Get the text of the selected item
			GuiControlGet, Selected,, % this.hListBox
			Suggestion := StrSplit(this.Suggestions, "|")[Selected]
			
			; Replace fragment preceding cursor with selected suggestion
			RC := this.Parent.RichCode
			RC.Selection[1] -= StrLen(this.Fragment)
			RC.SelectedText := Suggestion
			RC.Selection[1] := RC.Selection[2]
			
			; Clear out the fragment in preparation for further typing
			this.Fragment := ""
		}
		
		
		; --- Public Methods ---
		
		; Interpret WM_KEYDOWN messages, the primary means of interfacing with the
		; class. These messages can be provided by registering an appropriate
		; handler with OnMessage, or by forwarding the events from another handler
		; for the control.
		WM_KEYDOWN(wParam, lParam)
		{
			if (!this._Enabled)
				return
			
			; Get the name of the key using the virtual key code. The key's scan
			; code is not used here, but is available in bits 16-23 of lParam and
			; could be used in future versions for greater reliability.
			Key := GetKeyName(Format("vk{:02x}", wParam))
			
			; Treat Numpad variants the same as the equivalent standard keys
			Key := StrReplace(Key, "Numpad")
			
			; Handle presses meant to interact with the dialog, such as
			; navigational, confirmational, or dismissive commands.
			if (this.Visible)
			{
				if (Key == "Tab" || Key == "Enter")
					return False, this._Complete()
				else if (Key == "Up")
					return False, this.SelectUp()
				else if (Key == "Down")
					return False, this.SelectDown()
			}
			
			; Ignore standalone modifier presses, and some modified regular presses
			if Key in Shift,Control,Alt
				return
			
			; Reset on presses with the control modifier
			if GetKeyState("Control")
				return "", this.Fragment := ""
			
			; Subtract from the end of fragment on backspace
			if (Key == "Backspace")
				return "", this.Fragment := SubStr(this.Fragment, 1, -1)
			
			; Apply Shift and CapsLock
			if GetKeyState("Shift")
				Key := StrReplace(Key, "-", "_")
			if (GetKeyState("Shift") ^ GetKeyState("CapsLock", "T"))
				Key := Format("{:U}", Key)
			
			; Reset on unwanted presses -- Allow numbers but not at beginning
			if !(Key ~= "^[A-Za-z_]$" || (this.Fragment != "" && Key ~= "^[0-9]$"))
				return "", this.Fragment := ""
			
			; Record the starting position of new fragments
			if (this.Fragment == "")
			{
				CoordMode, Caret, % this.WineVer ? "Client" : "Screen"
				
				; Round "" to 0, which can prevent errors in the unlikely case that
				; input is received while the control is not focused.
				this.CaretX := Round(A_CaretX), this.CaretY := Round(A_CaretY)
			}
			
			; Update fragment with the press
			this.Fragment .= Key
		}
		
		; Triggers a rebuild of the word list from the RichCode control's contents
		BuildWordList()
		{
			if (!this._Enabled)
				return
			
			; Replace non-word chunks with delimiters
			List := RegExReplace(this.Parent.RichCode.Value, "\W+", "|")
			
			; Ignore numbers at the beginning of words
			List := RegExReplace(List, "\b[0-9]+")
			
			; Ignore words that are too small
			List := RegExReplace(List, "\b\w{1," this.MinWordLen-1 "}\b")
			
			; Append default entries, remove duplicates, and save the list
			List .= this.DefaultWordList
			Sort, List, U D| Z
			this.WordList := "|" Trim(List, "|")
		}
		
		; Moves the selected item in the dialog up one position
		SelectUp()
		{
			GuiControlGet, Selected,, % this.hListBox
			if (--Selected < 1)
				Selected := this.MaxSuggestions
			GuiControl, Choose, % this.hListBox, %Selected%
		}
		
		; Moves the selected item in the dialog down one position
		SelectDown()
		{
			GuiControlGet, Selected,, % this.hListBox
			if (++Selected > this.MaxSuggestions)
				Selected := 1
			GuiControl, Choose, % this.hListBox, %Selected%
		}
	}
}

class WinEvents ; static class
{
	static _ := WinEvents.AutoInit()
	
	AutoInit()
	{
		this.Table := []
		OnMessage(2, this.Destroy.bind(this))
	}
	
	Register(ID, HandlerClass, Prefix="Gui")
	{
		Gui, %ID%: +hWndhWnd +LabelWinEvents_
		this.Table[hWnd] := {Class: HandlerClass, Prefix: Prefix}
	}
	
	Unregister(ID)
	{
		Gui, %ID%: +hWndhWnd
		this.Table.Delete(hWnd)
	}
	
	Dispatch(Type, Params*)
	{
		Info := this.Table[Params[1]]
		return (Info.Class)[Info.Prefix . Type](Params*)
	}
	
	Destroy(wParam, lParam, Msg, hWnd)
	{
		this.Table.Delete(hWnd)
	}
}

WinEvents_Close(Params*) {
	return WinEvents.Dispatch("Close", Params*)
} WinEvents_Escape(Params*) {
	return WinEvents.Dispatch("Escape", Params*)
} WinEvents_Size(Params*) {
	return WinEvents.Dispatch("Size", Params*)
} WinEvents_ContextMenu(Params*) {
	return WinEvents.Dispatch("ContextMenu", Params*)
} WinEvents_DropFiles(Params*) {
	return WinEvents.Dispatch("DropFiles", Params*)
}

AutoIndent(Code, Indent = "`t", Newline = "`r`n")
{
	IndentRegEx =
	( LTrim Join
	Configuration|Info|Add|Remove|Product|Language|ExcludeApp|Property|RemoveMSI|Updates|Display|Logging|AppSettings|Setup|User
	)
	
	; Lock and Block are modified ByRef by Current
	Lock := [], Block := []
	ParentIndent := Braces := 0
	ParentIndentObj := []
	
	for each, Line in StrSplit(Code, "`n", "`r")
	{
		Text := Trim(RegExReplace(Line, "\s;.*")) ; Comment removal
		First := SubStr(Text, 1, 1), Last := SubStr(Text, 0, 1)
		FirstTwo := SubStr(Text, 1, 2)
		
		IsExpCont := (Text ~= "i)^\s*(&&|OR|AND|\.|\,|\|\||:|\?)")
		IndentCheck := (Text ~= "iA)}?\s*\b(" IndentRegEx ")\b")
		
		if (First == "(" && Last != ")")
			Skip := True
		if (Skip)
		{
			if (First == ")")
				Skip := False
			Out .= Newline . RTrim(Line)
			continue
		}
		
		if (FirstTwo == "*/")
			Block := [], ParentIndent := 0
		
		if Block.MinIndex()
			Current := Block, Cur := 1
		else
			Current := Lock, Cur := 0
		
		; Round converts "" to 0
		Braces := Round(Current[Current.MaxIndex()].Braces)
		ParentIndent := Round(ParentIndentObj[Cur])
		
		if (First == "}")
		{
			while ((Found := SubStr(Text, A_Index, 1)) ~= "}|\s")
			{
				if (Found ~= "\s")
					continue
				if (Cur && Current.MaxIndex() <= 1)
					break
				Special := Current.Pop().Ind, Braces--
			}
		}
		
		if (First == "{" && ParentIndent)
			ParentIndent--
		
		Out .= Newline
		Loop, % Special ? Special-1 : Round(Current[Current.MaxIndex()].Ind) + Round(ParentIndent)
			Out .= Indent
		Out .= Trim(Line)
		
		if (FirstTwo == "/*")
		{
			if (!Block.MinIndex())
			{
				Block.Push({ParentIndent: ParentIndent
				, Ind: Round(Lock[Lock.MaxIndex()].Ind) + 1
				, Braces: Round(Lock[Lock.MaxIndex()].Braces) + 1})
			}
			Current := Block, ParentIndent := 0
		}
		
		if (Last == "{")
		{
			Braces++, ParentIndent := (IsExpCont && Last == "{") ? ParentIndent-1 : ParentIndent
			Current.Push({Braces: Braces
			, Ind: ParentIndent + Round(Current[Current.MaxIndex()].ParentIndent) + Braces
			, ParentIndent: ParentIndent + Round(Current[Current.MaxIndex()].ParentIndent)})
			ParentIndent := 0
		}
		
		if ((ParentIndent || IsExpCont || IndentCheck) && (IndentCheck && Last != "{"))
			ParentIndent++
		if (ParentIndent > 0 && !(IsExpCont || IndentCheck))
			ParentIndent := 0
		
		ParentIndentObj[Cur] := ParentIndent
		Special := 0
	}
	
	if Braces
		throw Exception("Segment Open!")
	
	return SubStr(Out, StrLen(Newline)+1)
}

ExecScript(ODTFile, Params, XMLFile)
{
	static Shell := ComObjCreate("WScript.Shell")
	EnvGet, EnvTemp, Temp
	FileOpen(TempXMLFile := EnvTemp "\Config_" A_Now ".xml", "w", "UTF-8-RAW").Write(XMLFile)
	Exec := Shell.Exec(ODTFile " " Params " " TempXMLFile)
	;Exec := Shell.Exec(ODTFile " " TempXMLFile)
	;Exec := Shell.Exec("cmd.exe /c ping google.com")
	return Exec
}

UrlDownloadToVar(Url) {
   req := ComObjCreate("Msxml2.XMLHTTP")
   ; Open a request with async enabled.
   req.open("GET", Url, true)
   ; Set our callback function [requires v1.1.17+].
   req.onreadystatechange := Func("Ready")
   ; Send the request.  Ready() will be called when it's complete.
   req.send()
   while req.readyState != 4
   sleep 100
   if (req.status == 200)
     return req.responseText
}
 
Ready() {
   if (req.readyState == 4) {
      if (req.status == 200) {
	    return
	  }
      else
	  {
		Gui, +OwnDialogs
        MsgBox 16,, % "Error: " req.status  
	  }		
   }
}

; Helper function, to make passing in expressions resulting in function objects easier
SetTimer(Label, Period)
{
	SetTimer, %Label%, %Period%
}

SendMessage(Msg, wParam, lParam, hWnd)
{
	; DllCall("SendMessage", "UPtr", hWnd, "UInt", Msg, "UPtr", wParam, "Ptr", lParam, "UPtr")
	SendMessage, Msg, wParam, lParam,, ahk_id %hWnd%
	return ErrorLevel
}

PostMessage(Msg, wParam, lParam, hWnd)
{
	PostMessage, Msg, wParam, lParam,, ahk_id %hWnd%
	return ErrorLevel
}

CreateMenus(Menu)
{
	static MenuName := 0
	Menus := ["Menu_" MenuName++]
	for each, Item in Menu
	{
		Ref := Item[2]
		if IsObject(Ref) && Ref._NewEnum()
		{
			SubMenus := CreateMenus(Ref)
			Menus.Push(SubMenus*), Ref := ":" SubMenus[1]
		}
		Menu, % Menus[1], Add, % Item[1], %Ref%
	}
	return Menus
}

Ini_Load(Contents)
{
	Section := Out := []
	loop, Parse, Contents, `n, `r
	{
		if ((Line := Trim(StrReplace(A_LoopField, """"))) ~= "^;|^$")
			continue
		else if RegExMatch(Line, "^\[(.+)\]$", Match)
			Out[Match1] := (Section := [])
		else if RegExMatch(Line, "^(.+?)=(.*)$", Match)
			Section[Trim(Match1)] := Trim(Match2)
	}
	return Out
}

GetFullPathName(FilePath)
{
	VarSetCapacity(Path, A_IsUnicode ? 520 : 260, 0)
	DllCall("GetFullPathName", "Str", FilePath
	, "UInt", 260, "Str", Path, "Ptr", 0, "UInt")
	return Path
}

RichEdit_AddMargins(hRichEdit, x:=0, y:=0, w:=0, h:=0)
{
	static WineVer := DllCall("ntdll.dll\wine_get_version", "AStr")
	VarSetCapacity(RECT, 16, 0)
	if (x | y | w | h)
	{
		if WineVer
		{
			; Workaround for bug in Wine 3.0.2.
			; This code will need to be updated this code
			; after future Wine releases that fix it.
			NumPut(x, RECT,  0, "Int"), NumPut(y, RECT,  4, "Int")
			NumPut(w, RECT,  8, "Int"), NumPut(h, RECT, 12, "Int")
		}
		else
		{
			if !DllCall("GetClientRect", "UPtr", hRichEdit, "UPtr", &RECT, "UInt")
				throw Exception("Couldn't get RichEdit Client RECT")
			NumPut(x + NumGet(RECT,  0, "Int"), RECT,  0, "Int")
			NumPut(y + NumGet(RECT,  4, "Int"), RECT,  4, "Int")
			NumPut(w + NumGet(RECT,  8, "Int"), RECT,  8, "Int")
			NumPut(h + NumGet(RECT, 12, "Int"), RECT, 12, "Int")
		}
	}
	SendMessage(0xB3, 0, &RECT, hRichEdit)
}

DownloadODTRuntime()
{
	Gui, +OwnDialogs +Owner
	MsgBox, 0x40034, Warning, Office Deployment Tool runtime not found.`r`nDownload ODT from microsoft site?
	IfMsgBox, Yes
	{
		EnvGet, EnvTemp, Temp
		TempODT := EnvTemp "\officedeploymenttool.exe"
		UrlDownloadToFile, https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_14729-20228.exe , %TempODT%
		if ErrorLevel {
			Gui, +OwnDialogs
			MsgBox, 0x40010, Error , The ODT downloaded was failed.
		}
		else
		{
			RunWait, "%TempODT%" /extract:"%EnvTemp%\ODTManager" /quiet
			if ErrorLevel {
				Gui, +OwnDialogs
				MsgBox, 0x40010, Error , The ODT download was failed.
			}
			else
			{
				FileCopy, %EnvTemp%\ODTManager\setup.exe, %ExeDir%\setup.exe, 1
				if ErrorLevel {
					Gui, +OwnDialogs
					MsgBox, 0x40010, Error , The ODT downloaded was failed.
				}
				else
				{
					Gui, +OwnDialogs
					MsgBox, 0x40040, Success, The ODT downloaded was successfully.
				}
			}
		}
	}
}
