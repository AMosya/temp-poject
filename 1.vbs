ib = INPUTBOX ("Hello!", "PARAMETR", "VVEDITE TEXT",20,20)
BTN = MsgBox (IB,VBOKONLY +VBINFORMATION + VBSYSTEMMODAL,"Ю бнр х бюь рейяр")

SET P = CreateObject("WScript.Shell")
p.Popup BTN,3,"comment text", VBINFORMATION + VBSYSTEMMODAL

