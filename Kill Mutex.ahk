#NoEnv  
SendMode Input 
SetWorkingDir %A_ScriptDir% 


HandleName := "0000112" 
handlePath := "C:\Users\mis_reynan\Desktop\Vector\Handle\handle64.exe" 

pid := 0
handleID := getHandle(pid, HandleName)

if (handleID != "") {
    
    DestroyMutex(handlePath, pid, handleID)
} else {
    MsgBox, Handle not found!
}

DestroyMutex(HandlePath, PID, HexID) {
    StringReplace, PID, PID, PID:, , All  
    PID := Trim(PID)  

    StringReplace, HexID, HexID, Mutant, , All 
    HexID := Trim(HexID) 


    command := HandlePath . " -p " . PID . " -c " . HexID . " -y"
    Run, %comspec% /c %command%,, Hide

    if (ErrorLevel = 0) {
        
    } else {
        MsgBox, Failed to close handle. Error level: %ErrorLevel%.
    }
}


getHandle(ByRef pid, name) {
    command := "C:\Users\mis_reynan\Desktop\Vector\Handle\handle64.exe -a " . name
    stdout := runStdout(command)

    needle := "No matching"
    if InStr(stdout, needle) {
        return ""
    }
    
    RegExMatch(stdout, "pid:\s*(\d+)", pid) 
    RegExMatch(stdout, "Mutant\s+([0-9A-Fa-f]+)", handleID) 

    return handleID
}

runStdout(command) {
    shell := ComObjCreate("WScript.Shell")
    exec := shell.Exec(comspec " /c " command)
    stdout := exec.StdOut.ReadAll()
    
    return stdout
}
