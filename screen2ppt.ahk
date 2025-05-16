#Requires AutoHotkey v2.0

^1:: ; Ctrl-1 key combo, change to suit
{
A_Clipboard := "" ; empty clipboard

; Sleep 200

Send "{PrintScreen}"

Sleep 300

ClipWait(3,1)

oPPT := "" ; Required in v2 
; Get a reference to the PPT application, assumed only one window/file open
Try 
{ 
    oPPT := ComObjActive("Powerpoint.Application") 
} 
Catch 
{ 
    MsgBox "No active PPT application found." 
    ExitApp 
    ;Alternatively use ComObjCreate("Excel.Application") here instead of closing 
}

oPPT.Run("'Clipboard-image-collector.pptm'!Module3.InsertImageFromClipboard") ;

oPPT := "" ; Empty the variable

A_Clipboard := ""

}
