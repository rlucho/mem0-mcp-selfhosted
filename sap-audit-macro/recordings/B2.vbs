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
session.findById("wnd[0]/usr/radX_AISEL").select
session.findById("wnd[0]/usr/ctxtKD_BUKRS-LOW").text = "GBKM"
session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = "01092025"
session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = "30092025"
session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").setFocus
session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").caretPosition = 8
session.findById("wnd[0]/tbar[1]/btn[16]").press
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/txt%%DYN011-LOW").setFocus
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/txt%%DYN011-LOW").caretPosition = 0
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN011_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/txt%%DYN011-LOW").caretPosition = 10
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/lbl[164,8]").setFocus
session.findById("wnd[0]/usr/lbl[164,8]").caretPosition = 5
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr/lbl[164,8]").setFocus
session.findById("wnd[0]/usr/lbl[164,8]").caretPosition = 5
session.findById("wnd[0]/tbar[1]/btn[41]").press
session.findById("wnd[0]/usr/lbl[164,8]").setFocus
session.findById("wnd[0]/usr/lbl[164,8]").caretPosition = 3
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr/lbl[164,8]").setFocus
session.findById("wnd[0]/usr/lbl[164,8]").caretPosition = 3
session.findById("wnd[0]/tbar[1]/btn[40]").press
session.findById("wnd[0]/usr/lbl[164,10]").setFocus
session.findById("wnd[0]/usr/lbl[164,10]").caretPosition = 4
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/mbar/menu[4]/menu[3]").select
session.findById("wnd[0]/usr/chk[1,10]").setFocus
session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\eslucres\Documents\Audit 2"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Higher Vendor.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/lbl[164,8]").setFocus
session.findById("wnd[0]/usr/lbl[164,8]").caretPosition = 3
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr/lbl[164,8]").setFocus
session.findById("wnd[0]/usr/lbl[164,8]").caretPosition = 3
session.findById("wnd[0]/tbar[1]/btn[40]").press
session.findById("wnd[0]/usr/lbl[164,8]").setFocus
session.findById("wnd[0]/usr/lbl[164,8]").caretPosition = 3
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr/lbl[164,8]").setFocus
session.findById("wnd[0]/usr/lbl[164,8]").caretPosition = 3
session.findById("wnd[0]/tbar[1]/btn[41]").press
session.findById("wnd[0]/usr/lbl[164,10]").setFocus
session.findById("wnd[0]/usr/lbl[164,10]").caretPosition = 4
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/titl/shellcont[1]/shell").pressContextButton "%GOS_TOOLBOX"
session.findById("wnd[0]/titl/shellcont[1]/shell").selectContextMenuItem "%GOS_VIEW_ATTA"
session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").setCurrentCell 1,"BITM_DESCR"
session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").contextMenu
session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").selectContextMenuItem "%ATTA_EXPORT"
session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\Users\eslucres\Documents\Audit 2"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "invoice 2.PDF"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 9
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/lbl[164,10]").setFocus
session.findById("wnd[0]/usr/lbl[164,10]").caretPosition = 4
