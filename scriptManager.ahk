;说明
;  AHK脚本管理工具
;    启动时, 执行配置中的ahk脚本, 并且不会展示托盘图标
;    右键菜单有[刷新][新增][启动][退出][删除][定位文件]
;备注
;  1.获取到的内存使用大小与任务管理器taskmgr中的不相同, 需要在taskmgr中新增显示列【工作集(内存)】
;  2.因为不展示配置脚本的托盘图标，因此对于有些拥有图标右键菜单的脚本需要先[定位文件]，再手动执行
;TODO
;  1.配置是否展示托盘图标
;========================= 环境配置 =========================
#Persistent
#NoEnv
#HotkeyInterval, 1000
#SingleInstance, Force
SetBatchLines, -1
SetKeyDelay, -1
StringCaseSense, off
CoordMode, Menu
DetectHiddenWindows, On
SetTitleMatchMode, 2
#Include <JSON> 
#Include <PRINT>
#Include <TrayIcon>
;========================= 环境配置 =========================

global jsonFilePath := A_ScriptDir "\scriptManager.json"
global scriptInfos := Object()
global titleInfos := Object()
global titleStr :=
global titleWidthTotal :=
global objWMI :=
LVMenu()
LVTitle()
LVGui()
LoadScripts(true)
return



LVMenu() {
	Menu, Tray, NoStandard
	Menu, Tray, add, 脚本管理, MenuTrayManager
	Menu, Tray, add
	Menu, Tray, add, 重启, MenuTrayReload
	Menu, Tray, add, 退出, MenuTrayExit
    Menu, Tray, Default, 脚本管理
    Menu, Tray, Icon, %A_ScriptDir%\resources\manager.ico
    
    Menu, scriptMenu, Add, 刷新, MenuHandler
    Menu, scriptMenu, Icon, 刷新, SHELL32.dll, 239
    Menu, scriptMenu, Add, 新增, MenuHandler
    Menu, scriptMenu, Icon, 新增, SHELL32.dll, 1
    Menu, scriptMenu, Add
    Menu, scriptMenu, Add, 启动, MenuHandler
    Menu, scriptMenu, Icon, 启动, SHELL32.dll, 138
    Menu, scriptMenu, Add, 退出, MenuHandler
    Menu, scriptMenu, Icon, 退出, SHELL32.dll, 110
    Menu, scriptMenu, Add, 删除, MenuHandler
    Menu, scriptMenu, Icon, 删除, SHELL32.dll, 132
    Menu, scriptMenu, Add, 定位文件, MenuHandler
    Menu, scriptMenu, Icon, 定位文件, SHELL32.dll, 4
}
LVTitle() {
    titleInfos.push({"name": "name", "width": 200})
    titleInfos.push({"name": "PID", "width": 50})
    titleInfos.push({"name": "status", "width": 50})
    titleInfos.push({"name": "memory(M)", "width": 80})
    titleInfos.push({"name": "path", "width": 300})
    for index, titleInfo in titleInfos {
        titleStr .= titleInfo.name . "|"
        titleWidthTotal += titleInfo.width
    }
    titleWidthTotal += 4
}
LVGui() {
    Gui, scriptListView:Destroy
    Gui, scriptListView:New
    Gui, scriptListView:Font, s9, Microsoft YaHei
    Gui, scriptListView:Add, ListView, HScroll Grid w%titleWidthTotal% r30, %titleStr%
    for index, titleInfo in titleInfos {
        titleStr .= titleInfo.name . "|"
        LV_ModifyCol(index, titleInfo.width)
    }
    Gui, scriptListView:Add, StatusBar
    SB_SetParts(180)
}

scriptListViewGuiContextMenu(GuiHwnd, CtrlHwnd, EventInfo, IsRightClick, X, Y) {
    rowNum := LV_GetNext(0)
    if (rowNum) {
        Menu, scriptMenu, Enable, 启动
        Menu, scriptMenu, Enable, 退出
        Menu, scriptMenu, Enable, 删除
        Menu, scriptMenu, Enable, 定位文件
    } else {
        Menu, scriptMenu, Disable, 启动
        Menu, scriptMenu, Disable, 退出
        Menu, scriptMenu, Disable, 删除
        Menu, scriptMenu, Disable, 定位文件
    }
    Menu, scriptMenu, Show
}


MenuHandler(ItemName, ItemPos, MenuName) {
    if (ItemName == "刷新") {
        Gui, scriptListView:Default
        LV_Delete()
        LoadScripts(false)
        
    } else if (ItemName == "新增") {
        FileSelectFile, selectedFile, 3, , Open a file, ahk脚本文件(*.ahk)
        if (!selectedFile)
            return
        for index, scriptInfo in scriptInfos {
            if (selectedFile == scriptInfo.path) {
                MsgBox, 选择的脚本已存在
                return
            }
        }
        
        SplitPath, selectedFile, name
        Gui, scriptListView:Default
        rowNum := LV_Add("", name, , "停止", , selectedFile)
        newScriptInfo := Object("path", selectedFile, "name", name, "rowNum", rowNum)
        scriptInfos.push(newScriptInfo)
        
        scriptInfosStr := JSON.Dump(scriptInfos)
        FileDelete, %jsonFilePath%
        FileAppend, %scriptInfosStr%, %jsonFilePath%
        ScriptNumTotal()
        
    } else if (ItemName == "启动") {
        Gui, scriptListView:Default
        rowNum := 0
        Loop
        {
            rowNum := LV_GetNext(rowNum)
            if (!rowNum)
                break
            LV_GetText(rowName, rowNum, 1)
            LV_GetText(rowPath, rowNum, 5)
            WinClose, %rowPath%
            Run, %rowPath%,,,newPid
            memory := GetMemory(newPid)
            LV_Modify(rowNum, , , newPid, "运行", memory)
            RemoveTrayIcon(rowName)
        }
        ScriptMemoryTotal()
    
    } else if (ItemName == "退出") {
        Gui, scriptListView:Default
        rowNum := 0
        Loop
        {
            rowNum := LV_GetNext(rowNum)
            if (!rowNum)
                break
            LV_GetText(rowPath, rowNum, 5)
            WinClose, %rowPath%
            LV_Modify(rowNum, , , "", "停止", "")
        }
        ScriptMemoryTotal()
    
    } else if (ItemName == "删除") {
        Gui, scriptListView:Default
        rowNum := 0
        rowNums := Object()
        rowNameDelStr :=
        Loop
        {
            rowNum := LV_GetNext(rowNum)
            if (!rowNum)
                break
            rowNums.InsertAt(1, rowNum)
            LV_GetText(rowName, rowNum, 1)
            rowNameDelStr .= "  " rowName "`n"
        }
        MsgBox, 1, 删除, 是否要删除`n%rowNameDelStr%
        IfMsgBox OK
        {
            for index, rowNum in rowNums {
                LV_GetText(rowPath, rowNum, 5)
                WinClose, %rowPath%
                LV_Delete(rowNum)
                for index, scriptInfo in scriptInfos {
                    if (rowPath == scriptInfo.path) {
                        scriptInfos.removeAt(index)
                        break
                    }
                }
            }
        }
        scriptInfosStr := JSON.Dump(scriptInfos)
        FileDelete, %jsonFilePath%
        FileAppend, %scriptInfosStr%, %jsonFilePath%
        ScriptNumTotal()
        ScriptMemoryTotal()
        
    } else if (ItemName == "定位文件") {
        Gui, scriptListView:Default
        rowNum := 0
        Loop
        {
            rowNum := LV_GetNext(rowNum)
            if (!rowNum)
                break
            LV_GetText(rowPath, rowNum, 5)
            Run, % "explorer /select," rowPath
        }
    }
}


LoadScripts(autoStart) {
    FileEncoding, UTF-8
    jsonFile := FileOpen(jsonFilePath, "r")
    if !IsObject(jsonFile)
        throw Exception("Can't access file for JSONFile instance: " jsonFile, -1)
    try {
        scriptInfos := JSON.Load(jsonFile.Read())
    } catch e {
        MsgBox, JSON文件格式错误，请检查[%jsonFilePath%]
        return
    }
    jsonFile.Close()
    
    
    for index, scriptInfo in scriptInfos {
        path :=  scriptInfo.path
        SplitPath, path, name
        scriptInfo.name := name
        WinGet, curPid, PID, %path%
        if (!curPid && autoStart)
            Run, %path%,,,curPid
            
        if (curPid) {
            memory := GetMemory(curPid)
            rowNum := LV_Add("", name, curPid, "运行", memory, path)
            RemoveTrayIcon(name)
        } else {
            rowNum := LV_Add("", name, , "停止", , path)
        }
        scriptInfo.rowNum := rowNum
    }
    ScriptMemoryTotal()
    ScriptNumTotal()
}

GetMemory(ProcessId) {
    if (!objWMI)
        objWMI := ComObjGet("winmgmts:\\.\root\cimv2")
    StrSql := "SELECT * FROM Win32_Process WHERE ProcessId="""
    StrSql .= ProcessId
    StrSql .= """"
    Info := objWMI.ExecQuery(StrSql)
    if (Info && Info.Count) {
        for ObjProc in Info {
            return Round(ObjProc.WorkingSetSize / 1024 / 1024, 2)
        }
    } else {
        return ""
    }
}

ScriptNumTotal() {
    SB_SetText("总计"  LV_GetCount() "个脚本", 1)
}

ScriptMemoryTotal() {
    scriptMemoryTotal := 0
    Loop % LV_GetCount()
    {
        LV_GetText(rowMemory, A_Index, 4)
        scriptMemoryTotal += rowMemory
    }
    scriptMemoryTotal := Round(scriptMemoryTotal, 2)
    SB_SetText("共占用"  scriptMemoryTotal "M内存", 2)
}

RemoveTrayIcon(scriptName) {
    trayIcons := TrayIcon_GetInfo("AutoHotkey.exe")
    for index, trayIcon in trayIcons {
        if (scriptName == trayIcon.Tooltip) {
            TrayIcon_Remove(trayIcon.hWnd, trayIcon.uID)
            return        
        }
    }
}

MenuTrayManager:
    Gui, scriptListView:Default
    LV_Delete()
    LoadScripts(false)
    Gui, scriptListView:Show, , AHKScriptManager
return

MenuTrayReload:
    Reload
return

MenuTrayExit:
    for index, scriptInfo in scriptInfos {
        path :=  scriptInfo.path
        WinClose, %path%
    }
ExitApp