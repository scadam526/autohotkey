#Requires AutoHotkey v2.0


; Remap Calculator key to open SpeedCrunch
Launch_App2::
{
    ; Check if SpeedCrunch is already running
    if WinExist("ahk_exe speedcrunch.exe")
    {
        ; If it exists, activate the window
        WinActivate
    }
    else
    {
        ; If not running, launch it
        Run "C:\Program Files\SpeedCrunch\speedcrunch.exe"
    }
    return
}


; Ctrl+Shift+E to open or activate Everything
^+e::
{
    ; Check if Everything is already running and active window exists
    if WinExist("ahk_exe Everything.exe")
    {
        WinActivate
    }
    else
    {
        Run '"C:\Program Files\Everything\Everything.exe"'
    }
    return
}



; Ctrl+Alt+U to extract selected zip file and move to recycle bin
^!u::
{
    ; Get the selected file path from File Explorer or OneCommander
    zipPath := GetSelectedFilePath()

    if (zipPath = "")
    {
        MsgBox "No file selected. Please select a .zip file in File Explorer or OneCommander."
        return
    }

    ; Check if the selected file is a .zip file
    SplitPath zipPath, , , &fileExt
    if (StrLower(fileExt) != "zip")
    {
        MsgBox "Selected file is not a .zip file.`nDetected extension: " fileExt
        return
    }

    ; Check if the zip file exists
    if (!FileExist(zipPath))
    {
        MsgBox "The selected zip file does not exist."
        return
    }

    ; Get the directory and filename without extension
    SplitPath zipPath, &fileName, &fileDir, &fileExt, &fileNameNoExt

    ; Create the destination folder path
    destFolder := fileDir "\" fileNameNoExt

    ; Check if destination folder already exists
    if (DirExist(destFolder))
    {
        result := MsgBox("Folder '" fileNameNoExt "' already exists. Continue anyway?", "Folder Exists", "YesNo")
        if (result = "No")
            return
    }
    else
    {
        ; Create the destination folder
        DirCreate destFolder
    }

    ; Build the list of zip files to extract: the original plus any
    ; browser-style duplicates named "fileNameNoExt (n).zip" in the same folder
    zipFiles := [zipPath]
    Loop Files, fileDir "\" fileNameNoExt " (*).zip"
    {
        ; Only treat it as a duplicate if what's between the parens is purely digits
        suffix := SubStr(A_LoopFileName, StrLen(fileNameNoExt) + 2)
        if RegExMatch(suffix, "i)^\(\d+\)\.zip$")
            zipFiles.Push(A_LoopFileFullPath)
    }

    ; Extract each zip (original + duplicates) using Windows Shell
    try
    {
        ; Create Shell object
        shell := ComObject("Shell.Application")
        destFolderObj := shell.NameSpace(destFolder)

        for zf in zipFiles
        {
            ; Get the zip file object
            zipFolder := shell.NameSpace(zf)

            ; Extract all items (16 = respond "Yes to All" automatically)
            destFolderObj.CopyHere(zipFolder.Items(), 16)

            ; Wait a bit to ensure extraction is complete
            Sleep 1000
        }

        ; Send all extracted zip files (original + duplicates) to the recycle bin
        for zf in zipFiles
            FileRecycle zf

        ; MsgBox "Extraction complete! Files extracted to:`n" destFolder "`n`nThe zip file(s) have been moved to the recycle bin."
    }
    catch as err
    {
        MsgBox "Error during extraction: " err.Message
    }

    return
}


; Ctrl+Alt+Shift+U to extract selected zip file directly into its own folder and move to recycle bin
^!+u::
{
    ; Get the selected file path from File Explorer or OneCommander
    zipPath := GetSelectedFilePath()

    if (zipPath = "")
    {
        MsgBox "No file selected. Please select a .zip file in File Explorer or OneCommander."
        return
    }

    ; Check if the selected file is a .zip file
    SplitPath zipPath, , , &fileExt
    if (StrLower(fileExt) != "zip")
    {
        MsgBox "Selected file is not a .zip file.`nDetected extension: " fileExt
        return
    }

    ; Check if the zip file exists
    if (!FileExist(zipPath))
    {
        MsgBox "The selected zip file does not exist."
        return
    }

    ; Get the directory and filename without extension
    SplitPath zipPath, &fileName, &fileDir, &fileExt, &fileNameNoExt

    ; Extract straight into the directory containing the zip
    destFolder := fileDir

    ; Build the list of zip files to extract: the original plus any
    ; browser-style duplicates named "fileNameNoExt (n).zip" in the same folder
    zipFiles := [zipPath]
    Loop Files, fileDir "\" fileNameNoExt " (*).zip"
    {
        ; Only treat it as a duplicate if what's between the parens is purely digits
        suffix := SubStr(A_LoopFileName, StrLen(fileNameNoExt) + 2)
        if RegExMatch(suffix, "i)^\(\d+\)\.zip$")
            zipFiles.Push(A_LoopFileFullPath)
    }

    ; Extract each zip (original + duplicates) using Windows Shell
    try
    {
        ; Create Shell object
        shell := ComObject("Shell.Application")
        destFolderObj := shell.NameSpace(destFolder)

        for zf in zipFiles
        {
            ; Get the zip file object
            zipFolder := shell.NameSpace(zf)

            ; Extract all items (16 = respond "Yes to All" automatically)
            destFolderObj.CopyHere(zipFolder.Items(), 16)

            ; Wait a bit to ensure extraction is complete
            Sleep 1000
        }

        ; Send all extracted zip files (original + duplicates) to the recycle bin
        for zf in zipFiles
            FileRecycle zf
    }
    catch as err
    {
        MsgBox "Error during extraction: " err.Message
    }

    return
}


; Function to get the selected file path from File Explorer or OneCommander
GetSelectedFilePath()
{
    activeHwnd := WinExist("A")
    if !activeHwnd
        return ""

    ; If OneCommander is active, use the clipboard method to get the selected file path
    if WinActive("ahk_exe OneCommander.exe")
    {
        clipSaved := ClipboardAll()
        A_Clipboard := ""
        Send "^+c" ; OneCommander default shortcut to Copy Path
        if ClipWait(0.5)
        {
            filePath := A_Clipboard
            A_Clipboard := clipSaved
            lines := StrSplit(filePath, "`n", "`r")
            if (lines.Length > 0)
                return Trim(lines[1], '"' . A_Space . A_Tab)
        }

        ; Fallback to standard Ctrl+C if Ctrl+Shift+C doesn't yield anything
        A_Clipboard := ""
        Send "^c"
        if ClipWait(0.5)
        {
            filePath := A_Clipboard
            A_Clipboard := clipSaved
            lines := StrSplit(filePath, "`n", "`r")
            if (lines.Length > 0)
                return Trim(lines[1], '"' . A_Space . A_Tab)
        }

        A_Clipboard := clipSaved
        return ""
    }

    try
    {
        ; Get the active File Explorer window
        for window in ComObject("Shell.Application").Windows
        {
            if (window.HWND = activeHwnd)
            {
                ; Get selected items
                selectedItems := window.Document.SelectedItems()
                if (selectedItems.Count > 0)
                {
                    ; Return the path of the first selected item
                    return selectedItems.Item(0).Path
                }
            }
        }
    }
    return ""
}


; Ctrl + Shift + o
^+o::
{
    Send "{U+03A9}"  ; Sends the Ω symbol
}

; Ctrl + Shift + U
^+u::
{
    Send "{U+03BC}"  ; Sends the micro sign (µ)—actually a greek lowercase mu for SMATH
}

; Ctrl + Shift + m -> Em Dash (—)
^+m::
{
    Send "{U+2014}"
}


; --- Context-Specific Hotkey ---

; This block applies only if the active window is NOT (!) Explorer and NOT (!) OneCommander.
; "ahk_exe explorer.exe" covers both File Explorer windows and the Windows Desktop.
#HotIf !WinActive("ahk_exe explorer.exe") && !WinActive("ahk_exe OneCommander.exe")

; Ctrl + Shift + n -> En Dash (–)
^+n::
{
    Send "{U+2013}"
}

#HotIf ; Closes the conditional block


; Simulate a click of the Altium Selection Filter button
; Ctrl + Shift + S
; DISABLED: replaced by Altium's own shortcut for the Selection Filter popup.
; Remove the /* and */ lines to re-enable.
/*
#Requires AutoHotkey v2.0
^+s::
{
    ; Capture the active window ID
    active_hwnd := WinExist("A")
    if (active_hwnd)
        active_title := WinGetTitle("ahk_id " active_hwnd)
    else
        active_title := "(none)"

    CoordMode "Mouse", "Screen"
    MouseGetPos &origX, &origY

    try
    {
        ; Check if Altium is the active window
        isAltium := WinActive("ahk_exe X2.EXE") || (InStr(active_title, "Altium") > 0)

        if isAltium
        {
            image_dir := "C:\Users\Shawn\Nocturnal Dropbox\Shawn Adams\Scripts\autohotkey\images\Altium Selection Filter Button"

            ; 1. Get the position/size of the Active Altium Window
            ; This works regardless of which monitor the window is on.
            WinGetPos &winX, &winY, &winW, &winH, "ahk_id " active_hwnd

            ; 2. Define the search area limits based on the window
            X1 := winX
            Y1 := winY
            X2 := winX + winW
            Y2 := winY + winH

            CoordMode "Pixel", "Screen"

            Loop Files (image_dir "\*.png")
            {
                imgPath := A_LoopFileFullPath

                ; 3. Use the Window variables (X1, Y1, X2, Y2) for the search area
                if ImageSearch(&foundX, &foundY, X1, Y1, X2, Y2, "*50 " imgPath)
                {
                    Click(foundX + 5, foundY + 5)
                    return
                }
            }
        }
    }
    finally
    {
        if (IsSet(origX) && IsSet(origY))
            MouseMove origX, origY
    }
}
*/


; --- Altium: mouse side buttons -> layer up/down (same as Ctrl + Mouse Wheel) ---
; Only active when the mouse pointer is over an Altium window (X2.EXE).
; Checks the window UNDER THE POINTER rather than the active window, so it
; works correctly with multiple Altium windows / monitors. Matches on the
; process name only (not the title), so browser tabs etc. with "Altium" in
; the title keep normal Back/Forward behavior.
#HotIf MouseOverAltium()

XButton2::Send "^{WheelUp}"     ; Forward (front) side button -> layer up
XButton1::Send "^{WheelDown}"   ; Back (rear) side button -> layer down

#HotIf ; Closes the conditional block

MouseOverAltium()
{
    MouseGetPos , , &hwnd
    try return (WinGetProcessName("ahk_id " hwnd) = "X2.EXE")
    return false
}