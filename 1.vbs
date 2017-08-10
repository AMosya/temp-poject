ib = INPUTBOX ("Привет, пупик!", "ЗАПРОС СТРОКИ", "ВВЕДИТЕ СЮДА ЧТО-НИТЬ",20,20)
BTN = MsgBox (IB,VBOKONLY +VBINFORMATION + VBSYSTEMMODAL,"а ВОТ И ВАШ ТЕКСТ")

SET P = CreateObject("WScript.Shell")
p.Popup BTN,3,"бла-бла", VBINFORMATION + VBSYSTEMMODAL

