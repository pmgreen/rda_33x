#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_outfile=rda_3xx_serials.exe
#AutoIt3Wrapper_Run_Tidy=y
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****
; **************************************************************
; requested by: jwb
; from: 20130913
; for: SeERS
; insert i as 18th byte in ldr
; insert �e rda in 040 if not present
; insert 040 if not present
; insert 336, 337, 338
;~ 336 $3 serial $a text $2 rdacontent
;~ 336 $3 CD-ROM $a text $2 rdacontent
;~ 337 $3 serial $a unmediated $2 rdamedia
;~ 337 $3 CD-ROM $a computer $2 rdamedia
;~ 338 $3 serial $a volume $2 rdacarrier
;~ 338 $3 CD-ROM $a computer disc $2 rdacarrier
; **************************************************************

Opt("WinTitleMatchMode", 2) ; allow partial window matches

main()

; this is a separate function to allow cancelling an action without exiting the macro.
Func main()

	HotKeySet("+{F2}", "change_ldr")

	While 1
		Sleep(100)
	WEnd

EndFunc   ;==>main

Func change_ldr()
	WinActivate("Voyager Cataloging") ; change this to Bib to require maximized window
	WinWaitActive("Voyager Cataloging")
	Sleep(100)
	Send("{ALTDOWN}m{ALTUP}")
	Sleep(100)

	Send("!l")
	WinWaitActive("Leader", "", 20)
	ControlClick("[CLASS:VSFlexGridL; INSTANCE:1]", "", "") ; click within window
	Sleep(100)
	Send("{DOWN 6}")
	Send("i") ; isbd
	Sleep(100)
	Send("!o") ; OK

	insert_040()

EndFunc   ;==>change_ldr


Func insert_040()
	; tab different number of times based on existence of 006 and 007. Otherwise, chaos.
	If ControlGetText("", "", "[CLASS:ThunderRT6ComboBox; INSTANCE:2]") <> "" Then ; 007
		Send("{TAB 7}")
	ElseIf ControlGetText("", "", "[CLASS:ThunderRT6ComboBox; INSTANCE:1]") <> "" Then ; 006
		Send("{TAB 6}")
	Else
		Send("{TAB 5}") ; if no 006 or 007, tab 5 times
	EndIf

	Sleep(10)
	ClipPut("") ; clear out clipboard first
	Sleep(10)
	Send("{CTRLDOWN}{HOME 2}{CTRLUP}")
	Sleep(10)

	; check the first tag first
	Send("{F8}")
	Send("^c")
	Sleep(50)
	$a = ClipGet()
	Sleep(50)
	If ($a <> "" And $a < 040) Then
		Do
			ClipPut("")
			Send("{DOWN}")
			Send("{F8}")
			Send("^c")
			Sleep(50)
			$a = ClipGet()
			Sleep(50) ; this pause in necessary on some W7 machines
		Until ($a = 040) Or ($a >= 041) Or ($a == "")
	EndIf
	If $a = 040 Then
		ClipPut("")
		Send("{TAB 3}")
		Send("^+{END}")
		Send("^c")
		$field040 = ClipGet()

		If Not StringInStr($field040, "�e rda") Then
			ClipPut(StringRegExpReplace($field040, "�c", "�e rda �c"))
			Send("^v")
			Sleep(50)
			ClipPut("")
		EndIf
		Sleep(10)
		If Not StringInStr($field040, "�b eng") Then
			ClipPut(StringRegExpReplace($field040, "�e", "�b eng �e"))
			Send("^v")
			Sleep(50)
			ClipPut("")
		EndIf
		Sleep(10)
	ElseIf $a >= 041 Then
		Send("{F3}")
		Send("040")
		Send("{TAB}")
		Send("0")
		Send("{TAB}")
		Send("0")
		Send("{TAB}")
		Send("NjP �b eng �e rda �c NjP")
	ElseIf $a == "" Or $a < 040 Then
		ClipPut("") ; clear out clipboard first
		Sleep(10)
		Send("{CTRLDOWN}{HOME 2}{CTRLUP}")
		Sleep(10)
		Send("{F8}")
		Send("^c")
		Sleep(50)
		$b = ClipGet()
		Sleep(50)
		If $b < 040 Then
			Do
				ClipPut("")
				Send("{DOWN}")
				Send("{F8}")
				Send("^c")
				Sleep(50)
				$b = ClipGet()
				Sleep(50) ; this pause in necessary on some W7 machines
			Until ($b > 040)
		EndIf
		If $b > 040 Then
			Send("{F3}")
			Send("040")
			Send("{TAB}")
			Send("0")
			Send("{TAB}")
			Send("0")
			Send("{TAB}")
			Send("NjP �b eng �e rda �c NjP")
		EndIf
	EndIf
	ClipPut("")

	add_3xx()

EndFunc   ;==>insert_040

Func add_3xx()
	$336 = "�3 serial �a text �2 rdacontent"
	$3362 = "�3 CD-ROM �a text �2 rdacontent"
	$337 = "�3 serial �a unmediated �2 rdamedia"
	$3372 = "�3 CD-ROM �a computer �2 rdamedia"
	$338 = "�3 serial �a volume �2 rdacarrier"
	$3382 = "�3 CD-ROM �a computer disc �2 rdacarrier"
	ClipPut("") ; overkill, just to be sure
	Send("{CTRLDOWN}{HOME 2}{CTRLUP}")

	Do
		ClipPut("")
		Send("{DOWN}")
		Send("{F8}")
		Send("^c")
		Sleep(50)
		$a = ClipGet()
		Sleep(25) ; this pause in necessary on some W7 machines
	Until ($a >= 339) Or ($a == "") Or $a = 336

	If $a < 336 Then
		; If $a hit the bottom of the field list and was less than 336, enter 336, 337, then 338
		Send("{F4}")
		Send("336")
		Send("{TAB 3}")
		Send("{BACKSPACE 3}")
		ClipPut($336)
		Send("^v")
		Sleep(25)
		ClipPut("")

		Send("{F4}")
		Send("336")
		Send("{TAB 3}")
		Send("{BACKSPACE 3}")
		ClipPut($3362)
		Send("^v")
		Sleep(25)
		ClipPut("")

		Send("{F4}")
		Send("337")
		Send("{TAB 3}")
		Send("{BACKSPACE 3}")
		ClipPut($337)
		Send("^v")
		Sleep(25)
		ClipPut("")

		Send("{F4}")
		Send("337")
		Send("{TAB 3}")
		Send("{BACKSPACE 3}")
		ClipPut($3372)
		Send("^v")
		Sleep(25)
		ClipPut("")

		Send("{F4}")
		Send("338")
		Send("{TAB 3}")
		Send("{BACKSPACE 3}")
		ClipPut($338)
		Send("^v")
		Sleep(25)
		ClipPut("")

		Send("{F4}")
		Send("338")
		Send("{TAB 3}")
		Send("{BACKSPACE 3}")
		ClipPut($3382)
		Send("^v")
		Sleep(25)
		ClipPut("")

	ElseIf $a >= 339 Then
		; If/when $a is more than 336 enter in reverse order: 338, 337 then 336
		Send("{F3}")
		Send("338")
		Send("{TAB 3}")
		Send("{BACKSPACE 3}")
		ClipPut($3382)
		Send("^v")
		Sleep(25)
		ClipPut("")

		Send("{F3}")
		Send("338")
		Send("{TAB 3}")
		Send("{BACKSPACE 3}")
		ClipPut($338)
		Send("^v")
		Sleep(25)
		ClipPut("")

		Send("{F3}")
		Send("337")
		Send("{TAB 3}")
		Send("{BACKSPACE 3}")
		ClipPut($3372)
		Send("^v")
		Sleep(25)
		ClipPut("")

		Send("{F3}")
		Send("337")
		Send("{TAB 3}")
		Send("{BACKSPACE 3}")
		ClipPut($337)
		Send("^v")
		Sleep(25)
		ClipPut("")

		Send("{F3}")
		Send("336")
		Send("{TAB 3}")
		Send("{BACKSPACE 3}")
		ClipPut($3362)
		Send("^v")
		Sleep(25)
		ClipPut("")

		Send("{F3}")
		Send("336")
		Send("{TAB 3}")
		Send("{BACKSPACE 3}")
		ClipPut($336)
		Send("^v")
		Sleep(25)
		ClipPut("")

	ElseIf $a >= 336 Then
		$msg = MsgBox(4, "", "There's already a 336 field. Continue to add 336, 337 and 338?")
		If $msg = 6 Then
			Send("{F3}")
			Send("338")
			Send("{TAB 3}")
			Send("{BACKSPACE 3}")
			ClipPut($3382)
			Send("^v")
			Sleep(25)
			ClipPut("")

			Send("{F3}")
			Send("338")
			Send("{TAB 3}")
			Send("{BACKSPACE 3}")
			ClipPut($338)
			Send("^v")
			Sleep(25)
			ClipPut("")

			Send("{F3}")
			Send("337")
			Send("{TAB 3}")
			Send("{BACKSPACE 3}")
			ClipPut($3372)
			Send("^v")
			Sleep(25)
			ClipPut("")

			Send("{F3}")
			Send("337")
			Send("{TAB 3}")
			Send("{BACKSPACE 3}")
			ClipPut($337)
			Send("^v")
			Sleep(25)
			ClipPut("")

			Send("{F3}")
			Send("336")
			Send("{TAB 3}")
			Send("{BACKSPACE 3}")
			ClipPut($3362)
			Send("^v")
			Sleep(25)
			ClipPut("")

			Send("{F3}")
			Send("336")
			Send("{TAB 3}")
			Send("{BACKSPACE 3}")
			ClipPut($336)
			Send("^v")
			Sleep(25)
			ClipPut("")
		Else
			Send("{CTRLDOWN}{HOME 2}{CTRLUP}")
			main()
		EndIf
	EndIf
	Sleep(10)
	;finished()
EndFunc   ;==>add_3xx

Func finished()
	; this is just for fun -- random messages on completion
	Local $num = Random(1, 3, 1)
	Switch $num
		Case 1
			$msg = "Finis."
		Case 2
			$msg = "All done."
		Case 3
			$msg = "Congrats."
		Case Else
			$msg = "Take a break!"
	EndSwitch
	MsgBox(0, "done", $msg)
EndFunc   ;==>finished