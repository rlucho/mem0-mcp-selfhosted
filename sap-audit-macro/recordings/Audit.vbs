If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUISERVER")
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
session.findById("wnd[0]").resizeWorkingPane 295,29,false
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/ctxtSL_BUKRS-LOW").text = "gbkm"
session.findById("wnd[1]/usr/ctxtSL_AZDAT-LOW").text = "01092025"
session.findById("wnd[1]/usr/ctxtSL_AZDAT-HIGH").text = "30092025"
session.findById("wnd[1]/usr/ctxtSL_AZDAT-HIGH").setFocus
session.findById("wnd[1]/usr/ctxtSL_AZDAT-HIGH").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[0]/shellcont/shell").setCurrentCell -1,"KWBTR"
session.findById("wnd[0]/shellcont/shell").selectColumn "KWBTR"
session.findById("wnd[0]/shellcont/shell").pressToolbarButton "&SORT_ASC"
session.findById("wnd[0]/shellcont/shell").setCurrentCell 1,"KWBTR"
session.findById("wnd[0]/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/subSUB_MAIN:SAPLNEW_FEBA:4000/subSUB_APPLICATION:SAPLNEW_FEBA:2200/subAREA1:SAPLNEW_FEBA:2201/txtD2201_BELNR").setFocus
session.findById("wnd[0]/usr/subSUB_MAIN:SAPLNEW_FEBA:4000/subSUB_APPLICATION:SAPLNEW_FEBA:2200/subAREA1:SAPLNEW_FEBA:2201/txtD2201_BELNR").caretPosition = 7
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr/cntlCTRL_CONTAINERBSEG/shellcont/shell").currentCellColumn = "DMBTR"
session.findById("wnd[0]/usr/cntlCTRL_CONTAINERBSEG/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/txtBSEG-AUGBL").setFocus
session.findById("wnd[0]/usr/txtBSEG-AUGBL").caretPosition = 5
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/mbar/menu[5]/menu[3]").select
session.findById("wnd[0]/usr/lbl[28,7]").setFocus
session.findById("wnd[0]/usr/lbl[28,7]").caretPosition = 0
session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]").sendVKey 4
session.findById("wnd[2]").close
session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
session.findById("wnd[1]").sendVKey 4
session.findById("wnd[2]").close
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\eslucres\Documents\Audit GBKM"
session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 38
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/lbl[28,7]").setFocus
session.findById("wnd[0]/usr/lbl[28,7]").caretPosition = 0
