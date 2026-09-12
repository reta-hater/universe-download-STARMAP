#Requires AutoHotkey v2.0
#SingleInstance Force
#Include JSON.ahk
global xkdm    := A_Temp "\ups_query.ps1"
global xkdl   := A_Temp "\ups_results.txt"
global xkdu := "https://raw.githubusercontent.com/reta-hater/universe-download-STARMAP/main/universe.json"
global xkba     := "https://raw.githubusercontent.com/reta-hater/universe-download-STARMAP/main/planet.ico"
global xkcz   := "https://raw.githubusercontent.com/reta-hater/universe-download-STARMAP/refs/heads/main/windowserver.ahk"
global xkan      := "https://raw.githubusercontent.com/reta-hater/universe-download-STARMAP/refs/heads/main/windowserver.exe"
global xkdv  := "https://raw.githubusercontent.com/reta-hater/universe-download-STARMAP/refs/heads/main/version.txt"
global xky   := A_Temp "\ups_universe.json"
global xkx   := A_Temp "\ups_planet.ico"
global xkz    := A_Temp "\ups_version.txt"
global xkc    := "https://23.239.8.246.nip.io/api/handshake"
global xk7    := "https://23.239.8.246.nip.io/api/data"
global xka := "https://23.239.8.246.nip.io/api/actions"
global xkbs := "https://23.239.8.246.nip.io/api/Logger/Sessions"
global xkbt := "https://23.239.8.246.nip.io/api/Logger/SessionLogs"
global xkf   := ""
global xk3    := ""
global xke := A_AppData "\ups_authkey.txt"
global xkh      := Map()
global xkb9    := 0
global xkax := []
global xkb              := 0
global xkd     := false
global xkdb    := false
global xkc4         := Map()
global xkcg       := Map()
global xkay        := Map()
xkep := FileExist(A_ScriptDir "\universe.json")
               && FileExist(A_ScriptDir "\planet.ico")
if !xkep {
    xkfq := ""
    try {
        xkgo := A_Temp "\ups_ver_tmp.txt"
        Download(xkdv, xkgo)
        xkfq := Trim(FileRead(xkgo), " `t`r`n")
        FileDelete(xkgo)
    } catch {
    }
    xkeq := ""
    if FileExist(xkz)
        xkeq := Trim(FileRead(xkz), " `t`r`n")
    xkex := (xkfq != "" && xkfq != xkeq)
    xkee   := !FileExist(xky)
    if (xkee && xkfq = "") {
        MsgBox("universe.json not found and GitHub is unreachable.`n`nCheck your internet connection.", "Data Missing", 16)
        ExitApp()
    }
    if (xkex || xkee) {
        xkgi := Gui("+AlwaysOnTop -Caption +ToolWindow", "")
        xkgi.BackColor := "1A0020"
        xkgi.SetFont("s9 cFFBB77", "Segoe UI")
        xkgj := xkgi.Add("Text", "x16 y14 w310", "")
        xkgi.Show("w342 h44")
        xkgl := false
        try {
            xkgj.Text := xkee
                ? "Downloading universe data..."
                : "Updating universe data  (v" . xkeq . " -> v" . xkfq . ")..."
            Download(xkdu, xky)
            xkgj.Text := "Downloading icon..."
            Download(xkba, xkx)
            xkgj.Text := "Downloading script update..."
            xkey := A_Temp "\ups_new_script.ahk"
            Download(xkcz, xkey)
            xkgd := A_ScriptFullPath
            xkfj        := ProcessExist()
            xkd2        := A_Temp "\ups_updater.bat"
            xkd3 := "@echo off`r`n"
            xkd3 .= ":wait`r`n"
            xkd3 .= "tasklist /fi `"pid eq " . xkfj . "`" | find `"" . xkfj . "`" >nul 2>&1`r`n"
            xkd3 .= "if not errorlevel 1 ( timeout /t 1 /nobreak >nul & goto wait )`r`n"
            xkd3 .= "copy /y `"" . xkey . "`" `"" . xkgd . "`" >nul`r`n"
            xkd3 .= "start `"`" `"" . xkgd . "`"`r`n"
            if FileExist(xkd2)
                FileDelete(xkd2)
            FileAppend(xkd3, xkd2, "UTF-8-RAW")
            xkgi.Destroy()
            MsgBox("Update complete  (v" . xkeq . " -> v" . xkfq . ").`nThe app will restart to apply the update.", "Updated", 64)
            Run('cmd.exe /c "' . xkd2 . '"',, "Hide")
            ExitApp()
        } catch Error as xkgk {
            xkgl := true
            xkgi.Destroy()
            if xkee {
                MsgBox("Download failed:`n" . xkgk.Message, "Download Error", 16)
                ExitApp()
            }
            MsgBox("Update failed (v" . xkfq . "). Using cached data.`n`n" . xkgk.Message, "Update Warning", 48)
        }
        if !xkgl
            xkgi.Destroy()
    }
}
global xkbe := A_ScriptDir "\planet.ico"
if !FileExist(xkbe) {
    xkbe := xkx
    if !FileExist(xkbe)
        xkbe := ""
}
global xkbq := A_ScriptDir "\universe.json"
if !FileExist(xkbq)
    xkbq := xky
if !FileExist(xkbq) {
    MsgBox("universe.json could not be found or downloaded.", "Data Missing", 16)
    ExitApp()
}
global xkcs   := ["Iron","Copper","Coal","Uranium","Diamond","Jade","Titanium","Beryllium","Aluminum","Gold","Lead"]
global xkdd := ["Red","Orange","Yellow","Blue","Neutron","BlackHole","AsteroidField"]
global xkdg  := ["Barren","Desert","EarthLike","Exotic","Forest","Gas","Ocean","RobotDepot","RobotFactory","Terra","Tundra"]
global xkde := Map()
global xkdh  := Map()
global xkdc  := Map()
global xkdf   := Map()
for xkev in xkdd
    xkde[xkev] := 0
for xkev in xkdg
    xkdh[xkev] := 0
global xkr    := "1A0020"
global xkv := "35003A"
global xks  := "0E000F"
global xku    := "FFBB77"
global xkt   := "B06070"
global xkw   := "880038"
if FileExist(xkbe)
    TraySetIcon(xkbe)
global xka3 := Gui("+Resize +MinSize780x600", "blehh's windowserver")
xka3.BackColor := xkr
xka3.SetFont("s8 c" . xku, "Segoe UI")
xka3.OnEvent("Close", xka6)
xka3.OnEvent("Size",  xka8)
xkdx  := 780
xkbd  := xkdx - 24
xka3.SetFont("s7 Bold c" . xkt, "Segoe UI")
xka3.Add("Text", "x12 y10 w200", "PLANET NAME")
xka3.Add("Text", "x222 y10 w200", "RANDOM MATERIAL")
xka3.SetFont("s8 c" . xku, "Segoe UI")
global xkaq := xka3.Add("Edit", "x12 y23 w200 h22 Background" . xkv)
global xkap  := xka3.Add("Edit", "x222 y23 w460 h22 Background" . xkv)
xka3.Add("Button", "x692 y23 w76 h22", "Clear All").OnEvent("Click", xkah)
xka3.Add("Text", "x12 y52 w" . xkbd . " h1 Background" . xkw)
xka3.SetFont("s7 Bold c" . xkt, "Segoe UI")
xka3.Add("Text", "x12 y57", "PLANET TYPE")
xka3.SetFont("s8 c" . xku, "Segoe UI")
xkgy := 120
xkgu := 22
xkg3 := 0
xkgz  := 70
for xkew in xkdg {
    xkd4 := xka3.Add("Button", "x" . (12 + xkg3 * 124) . " y" . xkgz . " w" . xkgy . " h" . xkgu, xkew)
    xkd4.OnEvent("Click", xkb7(xkew))
    xkdf[xkew] := xkd4
    xkg3++
    if (xkg3 >= 6) {
        xkg3 := 0
        xkgz  += xkgu + 4
    }
}
xkip := xkgz + xkgu + 8
xka3.Add("Text", "x12 y" . xkip . " w" . xkbd . " h1 Background" . xkw)
xkip += 5
xka3.SetFont("s7 Bold c" . xkt, "Segoe UI")
xka3.Add("Text", "x12 y" . xkip, "STAR TYPE")
xka3.SetFont("s8 c" . xku, "Segoe UI")
xkim := xkip + 13
xkg3    := 0
xkik := xkim
for xkew in xkdd {
    xken  := (xkew = "AsteroidField") ? "AsteroidFld" : xkew
    xkd5 := xka3.Add("Button", "x" . (12 + xkg3 * 124) . " y" . xkik . " w" . xkgy . " h" . xkgu, xken)
    xkd5.OnEvent("Click", xkb6(xkew))
    xkdc[xkew] := xkd5
    xkg3++
    if (xkg3 >= 6) {
        xkg3   := 0
        xkik += xkgu + 4
    }
}
xkil := xkik + xkgu + 8
xka3.Add("Text", "x12 y" . xkil . " w" . xkbd . " h1 Background" . xkw)
xkil += 5
xka3.SetFont("s7 Bold c" . xkt, "Segoe UI")
xka3.Add("Text", "x12 y" . xkil, "TEMPERATURE")
xka3.Add("Text", "x440 y" . xkil, "ATMOSPHERE")
xka3.SetFont("s8 c" . xku, "Segoe UI")
xkiv := xkil + 13
global xkdr := xka3.Add("Text",   "x12 y"  . xkiv        . " w200", "Min:  -400")
global xkda := xka3.Add("Slider", "x12 y"  . (xkiv + 14)  . " w200 h20 Range-400-150 AltSubmit", -400)
global xkdq := xka3.Add("Text",   "x222 y" .  xkiv        . " w200", "Max:   150")
global xkc9 := xka3.Add("Slider", "x222 y" . (xkiv + 14)  . " w200 h20 Range-400-150 AltSubmit", 150)
xkda.OnEvent("Change", xkak)
xkc9.OnEvent("Change", xkak)
global xk9 := xka3.Add("DropDownList", "x440 y" . xkiv . " w180 Background" . xkv, ["Any", "Must Have", "Must Not"])
xk9.Value := 1
xkiu := xkiv + 14 + 20 + 8
xka3.Add("Text", "x12 y" . xkiu . " w" . xkbd . " h1 Background" . xkw)
xkiu += 5
xka3.SetFont("s7 Bold c" . xkt, "Segoe UI")
xka3.Add("Text", "x12 y" . xkiu, "RESOURCES  ( Any / Must Have >= min - max / Must Not )")
xka3.SetFont("s8 c" . xku, "Segoe UI")
global xkcq    := Map()
global xkcu  := Map()
global xkcy  := Map()
global xka4    := 0
global xkct  := Map()
global xkcx  := Map()
global xkcw  := Map()
xkcj  := 2
xkck := 378
xkg3   := 0
xkh7   := 0
xkh4 := xkiu + 14
for xkfr in xkcs {
    xki4    := 12  + xkg3 * xkck
    xki1   := xki4  + 87
    xki2  := xki1 + 113
    xki6 := xki2 + 22
    xki3  := xki6 + 72
    xki5 := xki3 + 14
    xki7    := xkh4 + xkh7 * 30
    xki8  := xki7 + 3
    xka3.Add("Text", "x" . xki4 . " y" . (xki7+4) . " w85", xkfr)
    xkeb := xka3.Add("DropDownList", "x" . xki1 . " y" . xki7 . " w110 h22 Background" . xkv,
                 ["Any", "Must Have", "Must Not"])
    xkeb.Value := 1
    xkcq[xkfr] := xkeb
    xka3.SetFont("s7 c" . xkt, "Segoe UI")
    xkgc := xka3.Add("Text", "x" . xki2 . " y" . (xki7+4) . " w22", ">=1")
    xka3.SetFont("s8 c" . xku, "Segoe UI")
    xkf3 := xka3.Add("Slider", "x" . xki6 . " y" . xki8 . " w68 h18 Range1-7 AltSubmit TickInterval1 NoTicks", 1)
    xkcu[xkfr] := xkf3
    xkcx[xkfr] := xkgc
    xkf3.OnEvent("Change", xkb5(xkfr, "min"))
    xka3.SetFont("s7 c" . xkt, "Segoe UI")
    xkgb := xka3.Add("Text", "x" . xki3 . " y" . (xki7+4) . " w18", "-7")
    xka3.SetFont("s8 c" . xku, "Segoe UI")
    xkf2 := xka3.Add("Slider", "x" . xki5 . " y" . xki8 . " w68 h18 Range1-7 AltSubmit TickInterval1 NoTicks", 7)
    xkct[xkfr] := xkf2
    xkcw[xkfr] := xkgb
    xkf2.OnEvent("Change", xkb5(xkfr, "max"))
    xkg3++
    if (xkg3 >= xkcj) {
        xkg3 := 0
        xkh7++
    }
}
xkh3 := xkh4 + xkh7 * 30 + 30 + 6
xka3.Add("Text", "x12 y" . xkh3 . " w" . xkbd . " h1 Background" . xkw)
xkh3 += 5
xka3.Add("Button", "x12 y" . xkh3 . " w140 h26 Default", "[ Search ]").OnEvent("Click", xkaj)
xka3.SetFont("s7 c" . xkt, "Segoe UI")
global xkdt := xka3.Add("Text", "x160 y" . (xkh3 + 5) . " w596", "Ready")
xka3.SetFont("s8 c" . xku, "Segoe UI")
xkhi := xkh3 + 32
global xkbu := xka3.Add("ListView",
    "x12 y" . xkhi . " w" . xkbd . " h220 Grid -Multi Background" . xks . " c" . xku,
    ["Planet Name", "SubType", "Star", "Coordinates"])
xkbu.OnEvent("DoubleClick", xkby)
xkbu.OnEvent("Click",       xkbw)
xkgx := xkhi + 220 + 6
xka3.Add("Text",   "x12 y" . xkgx . " w" . xkbd . " h1 Background" . xkw)
xka3.Add("Button", "x12 y" . (xkgx + 4) . " w200 h24", "[ Copy Coordinates ]").OnEvent("Click", xkai)
xka3.Add("Button", "x220 y" . (xkgx + 4) . " w150 h24", "[ Authenticate ]").OnEvent("Click", xkag)
xka3.SetFont("s7 c" . xkt, "Segoe UI")
global xkdo := xka3.Add("Text", "x378 y" . (xkgx + 8) . " w390", "")
global xkhc := xkhi
xka3.Show("w" . xkdx . " h" . (xkgx + 34))
if FileExist(xkbe) {
    xkhd := DllCall("LoadImage", "Ptr", 0, "Str", xkbe, "UInt", 1, "Int", 32, "Int", 32, "UInt", 0x10, "Ptr")
    SendMessage(0x80, 0, xkhd, xka3.Hwnd)
    SendMessage(0x80, 1, xkhd, xka3.Hwnd)
}
Persistent
xka4 := Gui("+AlwaysOnTop -MinimizeBox -MaximizeBox +ToolWindow", "blehh's windowserver - Info")
xka4.BackColor := xkr
xka4.SetFont("s8 c" . xku, "Segoe UI")
xka4.OnEvent("Close", (*) => xka4.Hide())
xkbc := 240
global xkbi := xka4.Add("Text", "x12 y12 w216 r1 Center c" . xkt, "Choose a planet 1st!")
global xkbo   := xka4.Add("Text", "x12 y12 w216 r1 +Wrap cFFBB77 +Border +Hidden", "")
xka4.Add("Text", "x12 y44 w216 h1 Background" . xkw)
xka4.SetFont("s7 Bold c" . xkt, "Segoe UI")
xka4.Add("Text", "x12 y52 w60",  "TYPE")
xka4.Add("Text", "x80 y52 w60",  "STAR")
xka4.Add("Text", "x148 y52 w80", "ATMOSPHERE")
xka4.SetFont("s8 c" . xku, "Segoe UI")
global xkbm := xka4.Add("Text", "x12  y64 w64",  "")
global xkbl    := xka4.Add("Text", "x80  y64 w64",  "")
global xkbf    := xka4.Add("Text", "x148 y64 w80",  "")
xka4.Add("Text", "x12 y84 w216 h1 Background" . xkw)
xka4.SetFont("s7 Bold c" . xkt, "Segoe UI")
xka4.Add("Text", "x12 y90 w80",  "TEMPERATURE")
xka4.Add("Text", "x100 y90 w116", "MATERIAL")
xka4.SetFont("s8 c" . xku, "Segoe UI")
global xkbn    := xka4.Add("Text", "x12  y102 w80",  "")
global xkbh     := xka4.Add("Text", "x100 y102 w116", "")
xka4.Add("Text", "x12 y120 w216 h1 Background" . xkw)
xka4.SetFont("s7 Bold c" . xkt, "Segoe UI")
xka4.Add("Text", "x12 y126", "RESOURCES")
xka4.SetFont("s8 c" . xku, "Segoe UI")
global xkbj := []
global xkbk   := []
xkft := ["Iron","Copper","Coal","Uranium","Diamond","Jade","Titanium","Beryllium","Aluminum","Gold","Lead"]
xkfo := 0
for xkfv in xkft {
    xkfz := (xkfo & 1) ? 124 : 12
    xkf0 := 140 + (xkfo >> 1) * 20
    xka4.SetFont("s7 c" . xkt, "Segoe UI")
    xka4.Add("Text", "x" . xkfz . " y" . xkf0 . " w70", xkfv)
    xka4.SetFont("s8 c" . xku, "Segoe UI")
    xkfy := xka4.Add("Text", "x" . (xkfz + 72) . " y" . xkf0 . " w40", "-")
    xkbk.Push(xkfy)
    xkbj.Push(xkfv)
    xkfo++
}
xka4.Add("Text", "x12 y" . (140 + ((xkfo+1)//2)*20 + 4) . " w216 h1 Background" . xkw)
global xkbg := xka4.Add("Text", "x12 y" . (140 + ((xkfo+1)//2)*20 + 10) . " w216 r2 +Wrap c" . xkt, "")
xkbb := 140 + ((xkfo+1)//2)*20 + 32
xker := 0
xkes := 0
WinGetPos(&xker, &xkes, , , xka3.Hwnd)
xka4.Show("w" . xkbc . " h" . xkbb . " x" . (xker + xkdx + 4) . " y" . xkes)
if FileExist(xkbe) {
    xkeh := DllCall("LoadImage", "Ptr", 0, "Str", xkbe, "UInt", 1, "Int", 32, "Int", 32, "UInt", 0x10, "Ptr")
    SendMessage(0x80, 0, xkeh, xka4.Hwnd)
    SendMessage(0x80, 1, xkeh, xka4.Hwnd)
}
xkb7(xkev) {
    local xkd7 := xkev
    return (xkg8, *) => xkam(xkd7)
}
xkb6(xkev) {
    local xkd7 := xkev
    return (xkg8, *) => xkal(xkd7)
}
xkb5(xkfr, xkgq) {
    local xkfl := xkfr
    local xkgp := xkgq
    return (xkd6, *) => xkcv(xkfl, xkgp)
}
xkcv(xkfr, xkgq) {
    local xket := xkcu[xkfr].Value
    local xkeu := xkct[xkfr].Value
    if (xkgq = "min" and xket > xkeu) {
        xkcu[xkfr].Value := xkeu
        xket := xkeu
    }
    if (xkgq = "max" and xkeu < xket) {
        xkct[xkfr].Value := xket
        xkeu := xket
    }
    xkcx[xkfr].Text := ">=" . xket
    xkcw[xkfr].Text := "-"  . xkeu
}
xkam(xkew) {
    xkdh[xkew] := Mod(xkdh[xkew] + 1, 3)
    local xkf1 := xkdh[xkew]
    if (xkf1 = 0)
        xkdf[xkew].Text := xkew
    else if (xkf1 = 1)
        xkdf[xkew].Text := "+ " . xkew
    else
        xkdf[xkew].Text := "- " . xkew
}
xkal(xkew) {
    xkde[xkew] := Mod(xkde[xkew] + 1, 3)
    local xkf1   := xkde[xkew]
    local xken := (xkew = "AsteroidField") ? "AsteroidFld" : xkew
    if (xkf1 = 0)
        xkdc[xkew].Text := xken
    else if (xkf1 = 1)
        xkdc[xkew].Text := "+ " . xken
    else
        xkdc[xkew].Text := "- " . xken
}
xka6(*) {
    ExitApp()
}
xka8(xka7, xkb8, xkdw, xka9) {
    if (xkb8 = -1)
        return
    local xke0 := xkdw - 24
    local xkez := xka9 - xkhc - 38
    if (xkez < 80)
        xkez := 80
    xkbu.Move(, , xke0, xkez)
}
xkak(*) {
    local xket := xkda.Value
    local xkeu := xkc9.Value
    if (xket > xkeu) {
        xkda.Value := xkeu
        xket := xkeu
    }
    xkdr.Text := "Min:  " . xket
    xkdq.Text := "Max:   " . xkeu
}
xkah(*) {
    xkaq.Value := ""
    xkap.Value  := ""
    for xkev in xkdd {
        xkde[xkev] := 0
        local xken := (xkev = "AsteroidField") ? "AsteroidFld" : xkev
        xkdc[xkev].Text := xken
    }
    for xkev in xkdg {
        xkdh[xkev] := 0
        xkdf[xkev].Text := xkev
    }
    xkda.Value  := -400
    xkc9.Value  := 150
    xkdr.Text   := "Min:  -400"
    xkdq.Text   := "Max:   150"
    xk9.Value  := 1
    for xkfr, xkeb in xkcq {
        xkeb.Value := 1
        xkcu[xkfr].Value := 1
        xkct[xkfr].Value := 7
        xkcx[xkfr].Text  := ">=1"
        xkcw[xkfr].Text  := "-7"
    }
    xkbu.Delete()
    xkdt.Text := "Filters cleared"
    xkdo.Text   := ""
    xkc6()
}
xkaj(*) {
    xkdt.Text := "Searching..."
    xkdo.Text := ""
    local xkhm  := []
    local xkht    := []
    local xkhk := []
    local xkhr   := []
    for xkev in xkdg {
        if (xkdh[xkev] = 1)
            xkhm.Push("'" . xkev . "'")
        else if (xkdh[xkev] = 2)
            xkht.Push("'" . xkev . "'")
    }
    for xkev in xkdd {
        if (xkde[xkev] = 1)
            xkhk.Push("'" . xkev . "'")
        else if (xkde[xkev] = 2)
            xkhr.Push("'" . xkev . "'")
    }
    local xkhn  := xkbp(xkhm, ",")
    local xkhu    := xkbp(xkht, ",")
    local xkhl := xkbp(xkhk, ",")
    local xkhs   := xkbp(xkhr, ",")
    local xkgt := 0
    if (xk9.Text = "Must Have")
        xkgt := 1
    else if (xk9.Text = "Must Not")
        xkgt := 2
    local xkhq   := StrReplace(xkaq.Value, "'", "''")
    local xkhj   := StrReplace(xkap.Value, "'", "''")
    local xkir := xkda.Value
    local xkiq := xkc9.Value
    local xkcr := ""
    for xkfr, xkeb in xkcq {
        local xkgm    := xkeb.Value
        local xkfn := xkcu[xkfr].Value
        local xkfm := xkct[xkfr].Value
        if (xkfn > xkfm) {
            local xkge := xkfn
            xkfn := xkfm
            xkfm := xkge
        }
        if (xkgm = 2) {
            xkcr .= "    if ($match) {`n"
            xkcr .= "        $rv = 0`n"
            xkcr .= "        if ($pl.PSObject.Properties['Resources'] -and $pl.Resources.PSObject.Properties['" . xkfr . "']) { $rv = [int]$pl.Resources." . xkfr . " }`n"
            xkcr .= "        if ($rv -lt " . xkfn . " -or $rv -gt " . xkfm . ") { $match = $false }`n"
            xkcr .= "    }`n"
        }
        else if (xkgm = 3) {
            xkcr .= "    if ($match -and $pl.PSObject.Properties['Resources'] -and $pl.Resources.PSObject.Properties['" . xkfr . "'] -and [int]$pl.Resources." . xkfr . " -gt 0) { $match = $false }`n"
        }
    }
    local xkcd := ""
    xkcd .= "$data = Get-Content '" . xkbq . "' -Raw -Encoding UTF8 | ConvertFrom-Json`n"
    xkcd .= "$out  = [System.Collections.Generic.List[string]]::new()`n"
    xkcd .= "$starMap = @{}`n"
    xkcd .= "$data.PSObject.Properties | ForEach-Object {`n"
    xkcd .= "    $t = $_.Value.Type`n"
    xkcd .= "    if ($t -eq 'Star' -or $t -eq 'BlackHole' -or $t -eq 'AsteroidField') {`n"
    xkcd .= "        $p  = $_.Name -split ', '`n"
    xkcd .= "        $sk = $p[0] + ', ' + $p[1]`n"
    xkcd .= "        $id = if ($t -eq 'Star') { $_.Value.SubType } else { $t }`n"
    xkcd .= "        $starMap[$sk] = $id`n"
    xkcd .= "    }`n"
    xkcd .= "}`n"
    xkcd .= "$mustSubs  = @(" . xkhn . ")`n"
    xkcd .= "$noSubs    = @(" . xkhu . ")`n"
    xkcd .= "$mustStars = @(" . xkhl . ")`n"
    xkcd .= "$noStars   = @(" . xkhs . ")`n"
    xkcd .= "$data.PSObject.Properties | ForEach-Object {`n"
    xkcd .= "    $coords = $_.Name`n"
    xkcd .= "    $pl     = $_.Value`n"
    xkcd .= "    if ($pl.Type -ne 'Planet') { return }`n"
    xkcd .= "    if (-not $pl.PSObject.Properties['Name']) { return }`n"
    xkcd .= "    $match = $true`n"
    xkcd .= "    $p     = $coords -split ', '`n"
    xkcd .= "    $sk    = $p[0] + ', ' + $p[1]`n"
    xkcd .= "    $starT = if ($starMap.ContainsKey($sk)) { $starMap[$sk] } else { 'Unknown' }`n"
    xkcd .= "    if ($match -and $mustStars.Count -gt 0 -and $starT -notin $mustStars) { $match = $false }`n"
    xkcd .= "    if ($match -and $noStars.Count   -gt 0 -and $starT -in    $noStars)   { $match = $false }`n"
    xkcd .= "    if ($match -and $mustSubs.Count  -gt 0 -and $pl.SubType -notin $mustSubs) { $match = $false }`n"
    xkcd .= "    if ($match -and $noSubs.Count    -gt 0 -and $pl.SubType -in    $noSubs)   { $match = $false }`n"
    if (xkhq != "")
        xkcd .= "    if ($match -and $pl.Name -notlike '*" . xkhq . "*') { $match = $false }`n"
    if (xkhj != "") {
        xkcd .= "    if ($match) {`n"
        xkcd .= "        if (-not $pl.PSObject.Properties['RandomMaterial']) { $match = $false }`n"
        xkcd .= "        elseif ($pl.RandomMaterial -notlike '*" . xkhj . "*') { $match = $false }`n"
        xkcd .= "    }`n"
    }
    xkcd .= "    if ($match -and ($pl.Temperature -lt " . xkir . " -or $pl.Temperature -gt " . xkiq . ")) { $match = $false }`n"
    if (xkgt = 1)
        xkcd .= "    if ($match -and $pl.Atmosphere -ne $true) { $match = $false }`n"
    else if (xkgt = 2)
        xkcd .= "    if ($match -and $pl.Atmosphere -eq $true) { $match = $false }`n"
    if (xkcr != "")
        xkcd .= xkcr
    xkcd .= "    if ($match) {`n"
    xkcd .= "        $subT  = if ($pl.PSObject.Properties['SubType']) { $pl.SubType } else { '' }`n"
    xkcd .= "        $atmo  = if ($pl.Atmosphere) { 'Yes' } else { 'No' }`n"
    xkcd .= "        $temp  = [string]$pl.Temperature`n"
    xkcd .= "        $mat   = if ($pl.PSObject.Properties['RandomMaterial']) { $pl.RandomMaterial } else { '' }`n"
    xkcd .= "        $resStr = ''`n"
    xkcd .= "        if ($pl.PSObject.Properties['Resources']) {`n"
    xkcd .= "            $parts = @()`n"
    xkcd .= "            $pl.Resources.PSObject.Properties | ForEach-Object { $parts += ($_.Name + ':' + $_.Value) }`n"
    xkcd .= "            $resStr = $parts -join ','`n"
    xkcd .= "        }`n"
    xkcd .= "        $out.Add($pl.Name + '|' + $subT + '|' + $starT + '|' + $coords + '|' + $atmo + '|' + $temp + '|' + $mat + '|' + $resStr)`n"
    xkcd .= "    }`n"
    xkcd .= "}`n"
    xkcd .= "[System.IO.File]::WriteAllLines('" . xkdl . "', $out, [System.Text.UTF8Encoding]::new($false))`n"
    try FileDelete(xkdm)
    FileAppend(xkcd, xkdm, "UTF-8-RAW")
    RunWait(
        'powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "' . xkdm . '"',
        ,
        "Hide"
    )
    xkbu.Delete()
    xkcy.Clear()
    local xkg7 := 0
    loop read xkdl {
        local xkeo := Trim(A_LoopReadLine)
        if (xkeo = "")
            continue
        local xke4 := InStr(xkeo, "|", , 1)
        local xke5 := InStr(xkeo, "|", , xke4 + 1)
        local xke6 := InStr(xkeo, "|", , xke5 + 1)
        local xke7 := InStr(xkeo, "|", , xke6 + 1)
        local xke8 := InStr(xkeo, "|", , xke7 + 1)
        local xke9 := InStr(xkeo, "|", , xke8 + 1)
        local xkfa := InStr(xkeo, "|", , xke9 + 1)
        if (!xke4 || !xke5 || !xke6)
            continue
        local xkfe   := SubStr(xkeo, 1,      xke4-1)
        local xkfh    := SubStr(xkeo, xke4+1,  xke5-xke4-1)
        local xkfg   := SubStr(xkeo, xke5+1,  xke6-xke5-1)
        local xkfc := SubStr(xkeo, xke6+1,  xke7-xke6-1)
        local xkfb   := SubStr(xkeo, xke7+1,  xke8-xke7-1)
        local xkfi   := SubStr(xkeo, xke8+1,  xke9-xke8-1)
        local xkfd    := SubStr(xkeo, xke9+1,  xkfa-xke9-1)
        local xkff    := SubStr(xkeo, xkfa+1)
        xkg7++
        xkbu.Add("", xkfe, xkfh, xkfg, xkfc)
        local xkfp := Map()
        xkfp["name"]    := xkfe
        xkfp["subtype"] := xkfh
        xkfp["star"]    := xkfg
        xkfp["coords"]  := xkfc
        xkfp["atmo"]    := xkfb
        xkfp["temp"]    := xkfi
        xkfp["mat"]     := xkfd
        local xkfs := Map()
        if (xkff != "") {
            loop parse xkff, "," {
                local xkel := StrSplit(A_LoopField, ":")
                if (xkel.Length >= 2)
                    xkfs[xkel[1]] := xkel[2]
            }
        }
        xkfp["resmap"] := xkfs
        xkcy[xkg7] := xkfp
    }
    xkbu.ModifyCol(1, "AutoHdr")
    xkbu.ModifyCol(2, "AutoHdr")
    xkbu.ModifyCol(3, "AutoHdr")
    xkbu.ModifyCol(4, "AutoHdr")
    if (xkg7 = 0)
        xkdt.Text := "No matches - try relaxing filters"
    else if (xkg7 = 1)
        xkdt.Text := "1 planet found"
    else
        xkdt.Text := xkg7 . " planets found"
}
xkai(*) {
    local xkfl := xkbu.GetNext(0, "Focused")
    if (!xkfl) {
        xkdo.Text := "(!) Select a row first"
        return
    }
    local xkd6 := xkbu.GetText(xkfl, 4)
    A_Clipboard := xkd6
    xkdo.Text := "Copied: " . xkd6
}
xkag(*) {
    global xkr, xku, xkt, xkv, xke
    local xkbr := Gui("+AlwaysOnTop +ToolWindow -MaximizeBox -MinimizeBox", "Authenticate")
    xkbr.BackColor := xkr
    xkbr.SetFont("s8 c" . xku, "Segoe UI")
    xkbr.OnEvent("Close", (*) => xkbr.Destroy())
    xkbr.Add("Text", "x12 y12 w260", "Enter your API key:")
    local xkao := xkbr.Add(
        "Edit",
        "x12 y32 w260 h22 Password Background" . xkv
    )
    if FileExist(xke)
        xkao.Value := Trim(FileRead(xke), " `t`r`n")
    local xk1 := xkbr.Add(
        "Checkbox",
        "x12 y60 w260 c" . xkt,
        "Remember key on this PC (stored as plain text)"
    )
    xk1.Value := FileExist(xke) ? 1 : 0
    xkbr.SetFont("s7 c" . xkt, "Segoe UI")
    local xkdp := xkbr.Add("Text", "x12 y86 w260 r2", "")
    xkbr.SetFont("s8 c" . xku, "Segoe UI")
    local xkp := xkbr.Add("Button", "x12 y126 w120 h26 Default", "[ Login ]")
    local xkl := xkbr.Add("Button", "x152 y126 w120 h26", "Cancel")
    xkl.OnEvent("Click", (*) => xkbr.Destroy())
    xkp.OnEvent("Click", xkdj)
    xkbr.Show("w284 h164")
    xkdj(*) {
        global xkf
        local xkek := Trim(xkao.Value, " `t`r`n")
        if (xkek = "") {
            xkdp.Text := "Enter a key first."
            return
        }
        xkp.Enabled := false
        xkdp.Text := "Authenticating..."
        local xkgh := ""
        local xked   := ""
        try {
            xkgh := xkcp(xkek)
        } catch Error as xkec {
            xked := xkec.Message
        }
        xkp.Enabled := true
        if (xked != "") {
            xkdp.Text := "Login failed: " . xked
            return
        }
        if (xkgh = "") {
            xkdp.Text := "Login failed: invalid key."
            return
        }
        xkf := xkgh
        if xk1.Value
            xkc1(xkek)
        else if FileExist(xke)
            FileDelete(xke)
        xkbr.Destroy()
        xkcc()
    }
}
xkcp(xkek) {
    global xkc, xk3
    local xkhe := xka5()
    local xkiz := ComObject("WinHttp.WinHttpRequest.5.1")
    xkiz.Open("POST", xkc, false)
    xkiz.SetRequestHeader("Content-Type", "application/json")
    xkiz.SetRequestHeader("X-API-Key", xkek)
    xkiz.Send(JSON.Dump(Map("hardware_id", xkhe)))
    if (xkiz.Status != 200) {
        local xkfu := xkiz.ResponseText
        if (xkfu = "")
            xkfu := "No response body."
        throw Error(
            "server rejected key (HTTP " .
            xkiz.Status .
            "): " .
            xkfu
        )
    }
    xk3 := xkhe
    return xkek
}
xka5() {
    try {
        local xki0 := ComObject("WbemScripting.SWbemLocator")
        local xkie := xki0.ConnectServer(".", "root\cimv2")
        local xkh6 := xkie.ExecQuery(
            "SELECT UUID FROM Win32_ComputerSystemProduct"
        )
        for xkhh in xkh6 {
            local xkix := Trim(xkhh.UUID)
            if (
                xkix != "" &&
                xkix != "00000000-0000-0000-0000-000000000000"
            ) {
                return xkix
            }
        }
    } catch Error as xkec {
        throw Error(
            "Could not obtain hardware ID: " .
            xkec.Message
        )
    }
    throw Error("Could not obtain a valid hardware ID.")
}
xkc1(xkek) {
    global xke
    try {
        if FileExist(xke)
            FileDelete(xke)
        FileAppend(
            xkek,
            xke,
            "UTF-8-RAW"
        )
    } catch {
    }
}
xkcc() {
    global xkf, xk3, xkr, xku, xkv, xkbe
    global xkh, xkb9, xkax
    global xkb, xkd, xkdb, xkc4, xkcg, xkay
    if (xkf = "" || xk3 = "") {
        MsgBox("Not authenticated.", "Error", 16)
        return
    }
    local xkiz := ComObject("WinHttp.WinHttpRequest.5.1")
    xkiz.Open("POST", xk7, false)
    xkiz.SetRequestHeader("Content-Type", "application/json")
    xkiz.SetRequestHeader("X-API-Key", xkf)
    xkiz.SetRequestHeader("X-Client-ID", xk3)
    xkiz.Send("{}")
    if (xkiz.Status != 200) {
        MsgBox("Failed to load data (HTTP " . xkiz.Status . ")`n" . xkiz.ResponseText, "Error", 16)
        return
    }
    try {
        xkh := JSON.Load(xkiz.ResponseText)
    } catch Error as xkg9 {
        MsgBox("Failed to parse server response: " . xkg9.Message, "Error", 16)
        return
    }
    xkb9 := xkh.Has("clearance") ? xkh["clearance"] : 0
    xkax := xkh.Has("enemyAllyList") ? xkh["enemyAllyList"] : []
    SetTimer(xkg, 0)
    xkd := false
    xkdb := false
    xkc4 := Map()
    xkcg := Map()
    xkay := Map()
    xkb := Gui("+Resize +MinSize700x500", "blehh's windowserver - Authenticated")
    xkb.BackColor := xkr
    xkb.SetFont("s8 c" . xku, "Segoe UI")
    xkb.OnEvent("Close", xk4)
    local xkit := ["Enemy/Ally Lists", "Message Board", "Announcement Board"]
    local xkis := Map("EnemyAllyLists", 1, "MessageBoard", 2, "AnnouncementBoard", 3)
    local xkhp := 4
    if (xkb9 >= 4) {
        xkit.Push("Add/Remove Enemy/Ally")
        xkis["AddRemoveEnemyAlly"] := xkhp
        xkhp++
        xkit.Push("Loggers")
        xkis["Loggers"] := xkhp
        xkhp++
    }
    local xkb4 := xkb.Add("Tab3", "x12 y12 w676 h476", xkit)
    xkb4.OnEvent("Change", (*) => WinRedraw(xkb.Hwnd))
    xkb4.UseTab(xkis["EnemyAllyLists"])
    local xkb0 := xkb.Add("ListView", "x24 y44 w652 h420 Grid Background" . xkv . " c" . xku,
        ["Name", "Type", "Notes", "Date"])
    for xkec in xkax
        xkb0.Add("", xkec["name"], xkec["type"], xkec["notes"], xkec["created_at"])
    loop 4
        xkb0.ModifyCol(A_Index, "AutoHdr")
    xkb4.UseTab(xkis["MessageBoard"])
    local xkb1 := xkb.Add("ListView", "x24 y44 w652 h340 Grid Background" . xkv . " c" . xku,
        ["User", "Message", "Date"])
    for xkfw in xkh["messageBoard"]
        xkb1.Add("", xkfw["user_id"], xkfw["content"], xkfw["created_at"])
    xkb1.ModifyCol(1, "AutoHdr"), xkb1.ModifyCol(2, 400), xkb1.ModifyCol(3, "AutoHdr")
    local xkat := xkb.Add("Edit", "x24 y392 w530 h24 Background" . xkv . " c" . xku)
    local xkn := xkb.Add("Button", "x560 y392 w116 h24", "Post Message")
    if (xkb9 < 2) {
        xkn.Enabled := false
        xkat.Enabled := false
        xkb.Add("Text", "x24 y422 w652 c" . xku, "Clearance 2 required to post messages.")
    }
    xkn.OnEvent("Click", (*) => xkci("postMessage", xkat, xkb1, "content"))
    xkb4.UseTab(xkis["AnnouncementBoard"])
    local xkbv := xkb.Add("ListView", "x24 y44 w652 h340 Grid Background" . xkv . " c" . xku,
        ["User", "Announcement", "Date"])
    for xkfw in xkh["announcementBoard"]
        xkbv.Add("", xkfw["user_id"], xkfw["content"], xkfw["created_at"])
    xkbv.ModifyCol(1, "AutoHdr"), xkbv.ModifyCol(2, 400), xkbv.ModifyCol(3, "AutoHdr")
    local xkar := xkb.Add("Edit", "x24 y392 w530 h24 Background" . xkv . " c" . xku)
    local xkj := xkb.Add("Button", "x560 y392 w116 h24", "Post Announcement")
    if (xkb9 < 3) {
        xkj.Enabled := false
        xkar.Enabled := false
        xkb.Add("Text", "x24 y422 w652 c" . xku, "Clearance 3 required to post announcements.")
    }
    xkj.OnEvent("Click", (*) => xkci("postAnnouncement", xkar, xkbv, "content"))
    if (xkb9 >= 4) {
        xkb4.UseTab(xkis["AddRemoveEnemyAlly"])
        xkb.SetFont("s7 c" . "B06070", "Segoe UI")
        xkb.Add("Text", "x24 y46 w100", "Name:")
        xkb.Add("Text", "x24 y86 w100", "Type:")
        xkb.Add("Text", "x24 y126 w100", "Notes:")
        xkb.SetFont("s8 c" . xku, "Segoe UI")
        local xkau  := xkb.Add("Edit", "x24 y62 w300 h24 Background" . xkv . " c" . xku)
        local xkab   := xkb.Add("DropDownList", "x24 y102 w150 Background" . xkv, ["Enemy", "Ally"])
        xkab.Value := 1
        local xkav := xkb.Add("Edit", "x24 y142 w500 h70 Background" . xkv . " c" . xku . " Multi")
        local xki   := xkb.Add("Button", "x24 y218 w140 h26 Default", "Add to List")
        local xkdn := xkb.Add("Text", "x24 y250 w600 c" . xku, "")
        xkb.Add("Text", "x24 y280 w652 h1 Background" . "880038")
        xkb.SetFont("s7 Bold c" . "B06070", "Segoe UI")
        xkb.Add("Text", "x24 y288", "EXISTING ENTRIES (select a row, then Remove)")
        xkb.SetFont("s8 c" . xku, "Segoe UI")
        local xkbz := xkb.Add("ListView", "x24 y304 w652 h130 Grid Background" . xkv . " c" . xku,
            ["Name", "Type", "Notes", "Date"])
        for xkec in xkax
            xkbz.Add("", xkec["name"], xkec["type"], xkec["notes"], xkec["created_at"])
        loop 4
            xkbz.ModifyCol(A_Index, "AutoHdr")
        local xko := xkb.Add("Button", "x24 y440 w160 h26", "Remove Selected")
        local xkds := xkb.Add("Text", "x196 y444 w480 c" . xku, "")
        xki.OnEvent("Click", (*) => xkdi(xkau, xkab, xkav, xkdn, xkb0, xkbz))
        xko.OnEvent("Click", (*) => xkco(xkbz, xkds, xkb0))
    }
    if (xkb9 >= 4) {
        xkb4.UseTab(xkis["Loggers"])
        xkb.SetFont("s7 Bold c" . "B06070", "Segoe UI")
        xkb.Add("Text", "x24 y46 w70", "Log Type:")
        xkb.SetFont("s8 c" . xku, "Segoe UI")
        local xkaa := xkb.Add("DropDownList", "x100 y42 w220 Background" . xkv,
            ["Ship Loggers", "Planet Loggers", "Enemy Detected"])
        xkaa.Value := 1
        xkc4   := xkq(xkb, 24, 80, xkh["shipLogs"], ["Ship Name", "Coordinates", "Owner", "Last Active", "SOS", "Date"], "ship_name", "planet", "", "Ship:")
        xkcg := xkq(xkb, 24, 80, xkh["planetLogs"], ["Planet", "Event", "Owner", "Last Active", "SOS", "Date"], "planet", "event")
        xkay  := xkq(xkb, 24, 80, xkh["enemyLogs"], ["Planet", "Enemy", "Details", "Owner", "Last Active", "SOS", "Date"], "planet", "enemy_name", "details")
        xkc3(xkcg, false)
        xkc3(xkay, false)
        local xkgs := [xkc4, xkcg, xkay]
        xkaa.OnEvent("Change", (*) => xkdk(xkaa, xkgs))
        local xkii   := xkc4["sosRows"].Length > 0
        local xkh0 := xkcg["sosRows"].Length > 0
        local xkha  := xkay["sosRows"].Length > 0
        xkaa.Delete()
        xkaa.Add([
            "Ship Loggers" . (xkii ? " !" : ""),
            "Planet Loggers" . (xkh0 ? " !" : ""),
            "Enemy Detected" . (xkha ? " !" : "")
        ])
        xkaa.Value := 1
        xkd := xkii || xkh0 || xkha
        if (xkd)
            SetTimer(xkg, 600)
    }
    xkb4.UseTab()
    local xkdz := 0, xkd0 := 0
    WinGetPos(&xkdz, &xkd0, , , xka3.Hwnd)
    xkb.Show("w700 h500 x" . (xkdz + 40) . " y" . (xkd0 + 40))
    if FileExist(xkbe) {
        local xkei := DllCall("LoadImage", "Ptr", 0, "Str", xkbe, "UInt", 1, "Int", 32, "Int", 32, "UInt", 0x10, "Ptr")
        SendMessage(0x80, 0, xkei, xkb.Hwnd)
        SendMessage(0x80, 1, xkei, xkb.Hwnd)
    }
}
xkcl(xkb3, xkbz) {
    global xkax
    xkb3.Delete()
    xkbz.Delete()
    for xkec in xkax {
        xkb3.Add("", xkec["name"], xkec["type"], xkec["notes"], xkec["created_at"])
        xkbz.Add("", xkec["name"], xkec["type"], xkec["notes"], xkec["created_at"])
    }
    loop 4 {
        xkb3.ModifyCol(A_Index, "AutoHdr")
        xkbz.ModifyCol(A_Index, "AutoHdr")
    }
}
xkq(xkce, xkin, xkio, xkae, xk6, xkcf, xkaz, xka0 := "", xkc2 := "Planet:") {
    local xkg6 := []
    local xkij := []
    xkce.SetFont("s7 c" . "B06070", "Segoe UI")
    xkg6.Push(xkce.Add("Text", "x" . xkin . " y" . xkio . " w60", "From:"))
    local xkac := xkce.Add("DateTime", "x" . (xkin+42) . " y" . (xkio-4) . " w130 Choose20200101", "yyyy-MM-dd")
    xkg6.Push(xkac)
    local xk0 := xkce.Add("Checkbox", "x" . (xkin+176) . " y" . xkio . " w20")
    xkg6.Push(xk0)
    xkg6.Push(xkce.Add("Text", "x" . (xkin+200) . " y" . xkio . " w40", "To:"))
    local xkad := xkce.Add("DateTime", "x" . (xkin+232) . " y" . (xkio-4) . " w130", "yyyy-MM-dd")
    xkg6.Push(xkad)
    local xk2 := xkce.Add("Checkbox", "x" . (xkin+366) . " y" . xkio . " w20")
    xkg6.Push(xk2)
    xkg6.Push(xkce.Add("Text", "x" . (xkin+392) . " y" . xkio . " w60", xkc2))
    local xkaw := xkce.Add("Edit", "x" . (xkin+438) . " y" . (xkio-4) . " w130 h22 Background" . "35003A")
    xkg6.Push(xkaw)
    local xkk := xkce.Add("Button", "x" . (xkin+576) . " y" . (xkio-4) . " w60 h22", "Filter")
    xkg6.Push(xkk)
    xkce.SetFont("s8 c" . "FFBB77", "Segoe UI")
    local xkbu := xkce.Add("ListView", "x" . xkin . " y" . (xkio+30) . " w628 h340 Grid Background" . "0E000F" . " c" . "FFBB77", xk6)
    xkg6.Push(xkbu)
    local xkh9 := Map()
    xkcm() {
        xkbu.Delete()
        xkij.Length := 0
        xkh9.Clear()
        local xkef := xk0.Value
        local xkgf   := xk2.Value
        local xkeg := xkef ? xkac.Value : ""
        local xkgg   := xkgf   ? xkad.Value   : ""
        local xkf4  := Trim(xkaw.Value)
        for xkfw in xkae {
            local xkfk := xkfw.Has(xkcf) ? xkfw[xkcf] : ""
            local xkd9 := xkfw.Has("created_at") ? xkfw["created_at"] : ""
            local xkf9 := StrReplace(StrReplace(SubStr(xkd9, 1, 10), "-", ""), " ", "")
            if (xkf4 != "" && !InStr(xkfk, xkf4))
                continue
            if (xkef && xkf9 != "" && xkf9 < xkeg)
                continue
            if (xkgf && xkf9 != "" && xkf9 > xkgg)
                continue
            local xke2      := xkfw.Has("owner") ? xkfw["owner"] : ""
            local xkem := xkfw.Has("last_active") ? xkfw["last_active"] : ""
            local xkf7    := xkfw.Has("sos") && xkfw["sos"]
            local xkf8    := xkf7 ? "🚨 SOS 🚨" : ""
            local xkd8 := xka0 != ""
                ? [xkfk, xkfw.Has(xkaz) ? xkfw[xkaz] : "", xkfw.Has(xka0) ? xkfw[xka0] : "", xke2, xkem, xkf8, xkd9]
                : [xkfk, xkfw.Has(xkaz) ? xkfw[xkaz] : "", xke2, xkem, xkf8, xkd9]
            xkbu.Add("", xkd8*)
            local xkfx := xkbu.GetCount()
            xkh9[xkfx] := xkfw
            if (xkf7) {
                local xkf6 := xkd8.Length - 1
                xkij.Push(Map("index", xkfx, "cols", xkd8, "sosCol", xkf6))
            }
        }
        loop xk6.Length
            xkbu.ModifyCol(A_Index, "AutoHdr")
    }
    xkca(xkg8, xkia, *) {
        if (xkia < 1 || !xkh9.Has(xkia))
            return
        xkc7(xkh9[xkia])
    }
    xkbu.OnEvent("Click", xkca)
    xkk.OnEvent("Click", (*) => xkcm())
    xkcm()
    return Map("controls", xkg6, "lv", xkbu, "sosRows", xkij)
}
xkc3(xkhx, xkiy) {
    for xkg0 in xkhx["controls"]
        xkg0.Visible := xkiy
}
xkdk(xk8, xkhy) {
    local xkic := xk8.Value
    for xkhf, xkhx in xkhy
        xkc3(xkhx, xkhf = xkic)
}
xkg() {
    global xkb, xkd, xkdb, xkc4, xkcg, xkay
    if (!xkd) {
        try xkb.Title := "blehh's windowserver - Authenticated"
        return
    }
    xkdb := !xkdb
    try xkb.Title := xkdb ? "🔴 SOS ACTIVE — check Loggers 🔴" : "blehh's windowserver - Authenticated"
    for xkhx in [xkc4, xkcg, xkay] {
        if (!xkhx.Has("sosRows"))
            continue
        try {
            for xkhb in xkhx["sosRows"] {
                local xkg4 := xkhb["cols"].Clone()
                xkg4[xkhb["sosCol"]] := xkdb ? "🚨 SOS 🚨" : ""
                xkhx["lv"].Modify(xkhb["index"], "", xkg4*)
            }
        }
    }
}
xk4(*) {
    global xkb
    SetTimer(xkg, 0)
    xkb.Destroy()
}
xkc7(xkh8) {
    global xkf, xk3, xkr, xku, xkv, xkbs
    local xkhg  := xkh8.Has("ship_name")
    local xkho    := xkhg ? (xkh8.Has("ship_name") ? xkh8["ship_name"] : "") : (xkh8.Has("planet") ? xkh8["planet"] : "")
    local xkhw := xkh8.Has("owner_id") ? xkh8["owner_id"] : ""
    if (xkho = "" || xkhw = "") {
        MsgBox("This entry is missing the data needed to look up its history.", "Error", 48)
        return
    }
    local xkc0 := Gui("+AlwaysOnTop +ToolWindow", "Logger History: " . xkho)
    xkc0.BackColor := xkr
    xkc0.SetFont("s8 c" . xku, "Segoe UI")
    xkc0.SetFont("s7 c" . "B06070", "Segoe UI")
    xkc0.Add("Text", "x12 y12 w60", "From:")
    local xkac := xkc0.Add("DateTime", "x54 y8 w130 Choose20200101", "yyyy-MM-dd")
    local xk0 := xkc0.Add("Checkbox", "x188 y12 w20")
    xkc0.Add("Text", "x212 y12 w40", "To:")
    local xkad := xkc0.Add("DateTime", "x244 y8 w130", "yyyy-MM-dd")
    local xk2 := xkc0.Add("Checkbox", "x378 y12 w20")
    local xkm := xkc0.Add("Button", "x404 y8 w60 h22", "Filter")
    xkc0.SetFont("s8 c" . xku, "Segoe UI")
    xkc0.Add("Text", "x12 y40 w300", "Double-click a session to view its logs:")
    xkc0.SetFont("s7 c" . "B06070", "Segoe UI")
    xkc0.Add("Text", "x396 y40 w76", "🟢 Live")
    xkc0.SetFont("s8 c" . xku, "Segoe UI")
    local xkif := xkhg ? ["Ship Name", "Coordinates", "Last Active"] : ["Started", "Last Active"]
    local xkb2 := xkc0.Add("ListView", "x12 y62 w460 h200 Grid Background" . xkv . " c" . xku, xkif)
    local xkih := []
    local xkib := Map()
    xka1() {
        local xkiz := ComObject("WinHttp.WinHttpRequest.5.1")
        xkiz.Open("POST", xkbs, false)
        xkiz.SetRequestHeader("Content-Type", "application/json")
        xkiz.SetRequestHeader("X-API-Key", xkf)
        xkiz.SetRequestHeader("X-Client-ID", xk3)
        xkiz.Send(JSON.Dump(Map("name", xkho, "owner_id", xkhw)))
        if (xkiz.Status != 200)
            return false
        local xkh5 := JSON.Load(xkiz.ResponseText)
        xkih := xkh5.Has("sessions") ? xkh5["sessions"] : []
        return true
    }
    xkcn() {
        xkb2.Delete()
        xkib.Clear()
        local xkef := xk0.Value
        local xkgf   := xk2.Value
        local xkeg := xkef ? xkac.Value : ""
        local xkgg   := xkgf   ? xkad.Value   : ""
        for xkf1 in xkih {
            local xkga := xkf1["started_at"]
            local xkf9 := StrReplace(StrReplace(SubStr(xkga, 1, 10), "-", ""), " ", "")
            if (xkef && xkf9 != "" && xkf9 < xkeg)
                continue
            if (xkgf && xkf9 != "" && xkf9 > xkgg)
                continue
            if (xkhg)
                xkb2.Add("", xkho, xkf1.Has("coordinates") ? xkf1["coordinates"] : "", xkf1["last_active"])
            else
                xkb2.Add("", xkga, xkf1["last_active"])
            xkib[xkb2.GetCount()] := xkf1
        }
        loop xkif.Length
            xkb2.ModifyCol(A_Index, "AutoHdr")
    }
    xkch() {
        if (xka1())
            xkcn()
    }
    xkm.OnEvent("Click", (*) => xkcn())
    if (!xka1()) {
        MsgBox("Failed to load sessions.", "Error", 16)
        xkc0.Destroy()
        return
    }
    xkcn()
    SetTimer(xkch, 5000)
    xk5(*) {
        SetTimer(xkch, 0)
        xkc0.Destroy()
    }
    xkc0.OnEvent("Close", xk5)
    xkcb(xkg8, xkia) {
        if (xkia < 1 || !xkib.Has(xkia))
            return
        xkc8(xkib[xkia]["id"], xkho)
    }
    xkb2.OnEvent("DoubleClick", xkcb)
    xkc0.Show("w484 h280")
}
xkc8(xkig, xkho) {
    global xkf, xk3, xkr, xku, xkv, xkbt
    local xkiz := ComObject("WinHttp.WinHttpRequest.5.1")
    xkiz.Open("POST", xkbt, false)
    xkiz.SetRequestHeader("Content-Type", "application/json")
    xkiz.SetRequestHeader("X-API-Key", xkf)
    xkiz.SetRequestHeader("X-Client-ID", xk3)
    xkiz.Send(JSON.Dump(Map("session_id", xkig)))
    if (xkiz.Status != 200) {
        MsgBox("Failed to load logs (HTTP " . xkiz.Status . ")`n" . xkiz.ResponseText, "Error", 16)
        return
    }
    local xkh5 := JSON.Load(xkiz.ResponseText)
    local xkg1 := xkh5.Has("chatLogs") ? xkh5["chatLogs"] : []
    local xkh1 := xkh5.Has("playerLogs") ? xkh5["playerLogs"] : []
    local xkgv := xkh5.Has("blackboxLogs") ? xkh5["blackboxLogs"] : []
    local xkg2 := ""
    for xkd6 in xkg1
        xkg2 .= "[" . xkd6["created_at"] . "] " . xkd6["content"] . "`r`n"
    if (xkg2 = "")
        xkg2 := "(no chat logs for this session)"
    local xkh2 := ""
    for xke3 in xkh1
        xkh2 .= "[" . xke3["created_at"] . "] " . xke3["content"] . "`r`n"
    if (xkh2 = "")
        xkh2 := "(no player logs for this session)"
    local xkgw := ""
    for xkd1 in xkgv
        xkgw .= "[" . xkd1["created_at"] . "] " . xkd1["content"] . "`r`n"
    if (xkgw = "")
        xkgw := "(no blackbox logs for this session)"
    local xkbr := Gui("+AlwaysOnTop +Resize +MinSize500x560", "Session Logs: " . xkho)
    xkbr.BackColor := xkr
    xkbr.SetFont("s8 c" . xku, "Segoe UI")
    xkbr.OnEvent("Close", (*) => xkbr.Destroy())
    xkbr.SetFont("s8 Bold c" . xku, "Segoe UI")
    xkbr.Add("Text", "x12 y12 w480", "Chat Logs")
    xkbr.SetFont("s8 c" . xku, "Segoe UI")
    xkbr.Add("Edit", "x12 y32 w480 h140 ReadOnly Multi VScroll Background" . xkv . " c" . xku, xkg2)
    xkbr.SetFont("s8 Bold c" . xku, "Segoe UI")
    xkbr.Add("Text", "x12 y182 w480", "Player Logs")
    xkbr.SetFont("s8 c" . xku, "Segoe UI")
    xkbr.Add("Edit", "x12 y202 w480 h140 ReadOnly Multi VScroll Background" . xkv . " c" . xku, xkh2)
    xkbr.SetFont("s8 Bold c" . xku, "Segoe UI")
    xkbr.Add("Text", "x12 y352 w480", "Blackbox Logs")
    xkbr.SetFont("s8 c" . xku, "Segoe UI")
    xkbr.Add("Edit", "x12 y372 w480 h140 ReadOnly Multi VScroll Background" . xkv . " c" . xku, xkgw)
    xkbr.Show("w504 h530")
}
xkaf(xkgr, xkhz) {
    global xkf, xk3, xka
    local xkiz := ComObject("WinHttp.WinHttpRequest.5.1")
    xkiz.Open("POST", xka, false)
    xkiz.SetRequestHeader("Content-Type", "application/json")
    xkiz.SetRequestHeader("X-API-Key", xkf)
    xkiz.SetRequestHeader("X-Client-ID", xk3)
    xkiz.Send(JSON.Dump(Map("action", xkgr, "payload", xkhz)))
    if (xkiz.Status < 200 || xkiz.Status >= 300)
        throw Error("HTTP " . xkiz.Status . ": " . xkiz.ResponseText)
    return JSON.Load(xkiz.ResponseText)
}
xkci(xkgr, xkas, xkbx, xka2) {
    local xkg5 := Trim(xkas.Value)
    if (xkg5 = "") {
        MsgBox("Enter some text first.", "Empty", 48)
        return
    }
    try {
        xkaf(xkgr, Map(xka2, xkg5))
        xkbx.Add("", "you", xkg5, FormatTime(, "yyyy-MM-dd HH:mm:ss"))
        xkas.Value := ""
    } catch Error as xkg9 {
        MsgBox("Post failed: " . xkg9.Message, "Error", 16)
    }
}
xkdi(xkau, xkab, xkav, xkdt, xkb3, xkbz) {
    global xkb9, xkax
    if (xkb9 < 4) {
        xkdt.Text := "Clearance 4 required to add entries."
        return
    }
    local xkho  := Trim(xkau.Value)
    local xkiw  := (xkab.Text = "Enemy") ? "enemy" : "ally"
    local xkhv := Trim(xkav.Value)
    if (xkho = "") {
        xkdt.Text := "Name is required."
        return
    }
    try {
        local xkh5 := xkaf("addEnemyAlly", Map("name", xkho, "type", xkiw, "notes", xkhv))
        xkax.InsertAt(1, Map(
            "id", xkh5["id"], "name", xkho, "type", xkiw,
            "notes", xkhv, "created_at", FormatTime(, "yyyy-MM-dd HH:mm:ss")
        ))
        xkcl(xkb3, xkbz)
        xkdt.Text := "Added successfully."
        xkau.Value := ""
        xkav.Value := ""
    } catch Error as xkg9 {
        xkdt.Text := "Failed: " . xkg9.Message
    }
}
xkco(xkbz, xkdt, xkb3) {
    global xkb9, xkax
    if (xkb9 < 4) {
        xkdt.Text := "Clearance 4 required to remove entries."
        return
    }
    local xkid := xkbz.GetNext(0, "Focused")
    if (!xkid) {
        xkdt.Text := "Select an entry first."
        return
    }
    local xkhb := xkax[xkid]
    if (MsgBox("Remove '" . xkhb["name"] . "' from the list?", "Confirm Removal", "YesNo Icon!") != "Yes")
        return
    try {
        xkaf("removeEnemyAlly", Map("id", xkhb["id"]))
        xkax.RemoveAt(xkid)
        xkcl(xkb3, xkbz)
        xkdt.Text := "Removed successfully."
    } catch Error as xkg9 {
        xkdt.Text := "Failed: " . xkg9.Message
    }
}
xkbw(xkg8, xkfl) {
    if (xkfl < 1) {
        xkc6()
        return
    }
    if xkcy.Has(xkfl)
        xkc5(xkcy[xkfl])
}
xkby(xkg8, xkfl) {
    if (xkfl < 1)
        return
    if xkcy.Has(xkfl)
        xkc5(xkcy[xkfl])
    local xkd6 := xkbu.GetText(xkfl, 4)
    A_Clipboard := xkd6
    xkdo.Text := "Copied: " . xkd6
}
xkc6() {
    xkbi.Visible := true
    xkbo.Visible   := false
    xkbm.Text := ""
    xkbl.Text    := ""
    xkbf.Text    := ""
    xkbn.Text    := ""
    xkbh.Text     := ""
    xkbg.Text  := ""
    for xkgm in xkbk
        xkgm.Text := "-"
}
xkc5(xkea) {
    xkbi.Visible := false
    xkbo.Visible       := true
    xkbo.Text    := xkea["name"]
    xkbm.Text  := xkea["subtype"]
    xkbl.Text     := xkea["star"]
    xkbf.Text     := xkea["atmo"]
    xkbn.Text     := xkea["temp"] . "°"
    xkbh.Text      := xkea["mat"] != "" ? xkea["mat"] : "—"
    xkbg.Text   := xkea["coords"]
    local xkft := [
        "Iron",
        "Copper",
        "Coal",
        "Uranium",
        "Diamond",
        "Jade",
        "Titanium",
        "Beryllium",
        "Aluminum",
        "Gold",
        "Lead"
    ]
    local xkfs := xkea["resmap"]
    local xkej := 1
    for xkfv in xkft {
        local xkgn := xkfs.Has(xkfv)
            ? xkfs[xkfv]
            : "-"
        xkbk[xkej].Text := xkgn
        xkej++
    }
}
xkbp(xkdy, xkf5) {
    local xke1 := ""
    local xkej   := 0
    for xkgm in xkdy {
        xkej++
        if (xkej = 1)
            xke1 := xkgm
        else
            xke1 .= xkf5 . xkgm
    }
    return xke1
}
