CWID =  CreateObject("WScript.Network").UserName
code =  Wscript.Arguments(0)
date1 = Wscript.Arguments(1)
date2 = Wscript.Arguments(2)

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/n/BAY0/AK_ERGEBNIS"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/btn%_S_WERK_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "YA21"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "YA23"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtS_MKWERK-LOW").text = "YA21"
session.findById("wnd[0]/usr/ctxtS_HSDAT-LOW").text = date1
session.findById("wnd[0]/usr/ctxtS_HSDAT-HIGH").text = date2
session.findById("wnd[0]/usr/btn%_S_MERKM_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[23]").press
session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "clip.txt"
session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\Users\"&CWID&"\OneDrive - Bayer\Desktop\QualityData\Trend2\Files\"&CWID&"\"
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]").sendVKey 8
session.findById("wnd[0]/usr/radP_PC").select
session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = code
session.findById("wnd[0]/usr/txtPC_FILE").text = "C:\Users\"&CWID&"\OneDrive - Bayer\Desktop\QualityData\Trend2\rawdata\results\code"&code&".txt"
session.findById("wnd[0]/usr/txtP_ANZAHL").text = "50000000"
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]").sendVKey 3
