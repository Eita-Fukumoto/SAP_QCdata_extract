CWID =  CreateObject("WScript.Network").UserName

Function Getqa33(ByVal MaterialCode)
	session.findById("wnd[0]/usr/radP_UD").select
	session.findById("wnd[0]/usr/ctxtQL_ENSTD-LOW").text = ""
	session.findById("wnd[0]/usr/ctxtQL_ENSTD-HIGH").text = ""
	session.findById("wnd[0]/usr/ctxtQL_MATNR-LOW").text = MaterialCode
	session.findById("wnd[0]/usr/ctxtQL_WERKS-LOW").text = ""
        session.findById("wnd[0]/usr/txtQL_MAX_R").text = ""
	session.findById("wnd[0]/usr/ctxtVARIANT").text = "/YA-1STREL"
	session.findById("wnd[0]/usr/ctxtVARIANT").setFocus
	session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition = 10
	session.findById("wnd[0]/tbar[1]/btn[8]").press
	session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select
	session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
	session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\"&CWID&"\OneDrive - Bayer\Desktop\QualityData\Trend2\rawdata\qa33\"
	session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "q_"&MaterialCode&".txt"
	session.findById("wnd[1]/tbar[0]/btn[11]").press
	session.findById("wnd[0]/tbar[0]/btn[15]").press
End Function

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
session.findById("wnd[0]/tbar[0]/okcd").text = "qa33"
session.findById("wnd[0]").sendVKey 0


' ファイルパス
strFile = "Files/"&CWID&"/table6.csv"

'ファイルシステムオブジェクト作成
Set objFS = CreateObject("Scripting.FileSystemObject")

' ファイルオープン
Set objText = objFS.OpenTextFile(strFile)

' 全行読み込む
strText = objText.ReadAll

' ファイルクローズ
objText.Close

' 改行で分割
arrText = Split(strText, vbCrLf)

' 配列の内容を1行ずつ表示
Dim cols
For i = 1 To UBound(arrText)
   cols = arrText(i)
   col = Split(cols, ",")
   WScript.StdOut.Write i & "|" & Now() & "|" & col(0) & "|" & col(1) & vbCrLf 
   Getqa33 col(0)

Next

