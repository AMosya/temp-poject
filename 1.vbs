ib = INPUTBOX ("������, �����!", "������ ������", "������� ���� ���-����",20,20)
BTN = MsgBox (IB,VBOKONLY +VBINFORMATION + VBSYSTEMMODAL,"� ��� � ��� �����")

SET P = CreateObject("WScript.Shell")
p.Popup BTN,3,"���-���", VBINFORMATION + VBSYSTEMMODAL

