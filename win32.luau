
-- // Dependencies
local process = require("@lune/process")
local stdio = require("@lune/stdio")
local fs = require("@lune/fs")


-- check if the os is windows, if not error
if process.os ~= "windows" then
    stdio.write(stdio.color("red")); stdio.write(stdio.style("bold")); stdio.write("fatal"); stdio.write(stdio.color("reset")); stdio.write(stdio.style("reset")); stdio.write(": cannot use win32.luau on "..process.os)
    process.exit(1)
end

-- // Module Declaration
local win32 = {}

-- // Enums Declaration
local enums = {}

-- // Enums
local msgBoxIcons = {
    Critical = 16,
    Question = 32,
    Exclamation = 48,
    Information = 64
}

local msgBoxButtons = {
    OkOnly = 0,
    OkCancel = 1,
    AbortRetryIgnore = 2,
    YesNoCancel = 3,
    YesNo = 4,
    RetryCancel = 5
}

local msgBoxResult = {
    Ok = 1,
    Cancel = 2,
    Abort = 3,
    Retry = 4,
    Ignore = 5,
    Yes = 6,
    No = 7
}

enums.msgBoxResult = msgBoxResult
enums.msgBoxButtons = msgBoxButtons
enums.msgBoxIcons = msgBoxIcons

-- // Functions

local function shell(command: string): process.SpawnResult
    local result = process.spawn(command, nil, { shell = true })

    if result.ok then
        print(result.stdout)
    else
        print(result.stderr)
    end

    return result
end

local function msgbox(title: string, prompt: string, icon: number, buttons: number)
    local scriptContents = ('Set objFSO=CreateObject("Scripting.FileSystemObject") \n result = MsgBox("'..prompt..'",' ..buttons.. '+ ' ..icon.. ', "' ..title.. '") \n outFile="./msgboxresult.txt" \n Set objFile = objFSO.CreateTextFile(outFile,True) \n objFile.Write result \n objFile.Close')

    fs.writeFile("./script.vbs", scriptContents)
    shell("./script.vbs")
    
    local result = fs.readFile("./msgboxresult.txt")
    fs.removeFile("./msgboxresult.txt"); fs.removeFile("./script.vbs")

    return tonumber(result)
end

-- // Returning
win32.enums = enums
win32.shell = shell
win32.arch = process.arch
win32.env = process.env
win32.cwd = process.cwd
win32.args = process.args
win32.msgbox = msgbox

return win32