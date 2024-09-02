-- // Type Declarations
export type html = string
export type css = string


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

local function window(title: string, icon: string, height: number, width: number, content: html, style: css, scroll: "yes" | "no")
    local scriptContents = [[
    Set objFSO=CreateObject("Scripting.FileSystemObject")

        Class Window
'@description: Create a custom window with MSHTA.
'@author: Jeremy England (SimplyCoded).
  Private title, style, body, options, width, height, xpos, ypos  
  Private Sub Class_Initialize()
    title = "&nbsp;" : width = 350 : height = 250
    xpos = "(screen.width - " & width & ")/2"
    ypos = "(screen.height -" & height & ")/2" 
    style = "html{display:table;}body{display:table-cell;font-family:Arial;background-color:#f6f6f6;}html,body{width:100%;height:100%;margin:0;}"
  End Sub  
  Public Property Let SetTitle(str)     : title = str         : End Property   
  Public Property Let SetWidth(num)     : width = num         : End Property 
  Public Property Let SetHeight(num)    : height = num        : End Property
  Public Property Let SetXPosition(num) : xpos = num          : End Property 
  Public Property Let SetYPosition(num) : ypos = num          : End Property 
  Public Property Let AddStyles(css)    : style = style & css : End Property 
  Public Property Let AddContent(html)  : body = body & html  : End Property    
  Public Property Let ApplicationName(str)    : options = options & "applicationName='" & str & "' "    : End Property
  Public Property Let Border(thick_thin_none) : options = options & "border='" & thick_thin_none & "' " : End Property
  Public Property Let Caption(yes_no)         : options = options & "caption='" & yes_no & "' "         : End Property
  Public Property Let ContextMenu(yes_no)     : options = options & "contextMenu='" & yes_no & "' "     : End Property
  Public Property Let Icon(path)              : options = options & "icon='" & path & "' "              : End Property
  Public Property Let MaximizeButton(yes_no)  : options = options & "maximizeButton='" & yes_no & "' "  : End Property
  Public Property Let MinimizeButton(yes_no)  : options = options & "minimizeButton='" & yes_no & "' "  : End Property
  Public Property Let Scroll(yes_no)          : options = options & "scroll='" & yes_no & "' "          : End Property
  Public Property Let Selection(yes_no)       : options = options & "selection='" & yes_no & "' "       : End Property
  Public Property Let ShowInTaskBar(yes_no)   : options = options & "showInTaskBar='" & yes_no & "' "   : End Property
  Public Property Let SingleInstance(yes_no)  : options = options & "singleInstance='" & yes_no & "' "  : End Property
  Public Property Let SysMenu(yes_no)         : options = options & "sysMenu='" & yes_no & "' "         : End Property
  Public Property Let WindowState(normal_minimize_maximize) : options = options & "windowState='" & normal_minimize_maximize & "' " : End Property  
  Public Function Create()
    Create = CreateObject("WScript.Shell").Exec( _
      "mshta ""about:<!DOCTYPE html><html><head><meta http-equiv='X-UA-Compatible' content='IE=9'><title>" & title & _
      "</title><hta:application " & options & "/><style>" & style & "</style>" & _
      "<script>var c=true;function send(s){c=false;new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(s);close();}" & _
      "window.onbeforeunload=function(){if(c)send(0);};resizeTo(" & width & "," & height & ");moveTo(" & xpos & "," & ypos & ");</script>" & _
      "</head><body>" & body & "</body></html>""" _
    ).StdOut.ReadLine()
  End Function 
End Class

'Create a basic continue/cancel message.
Set win = New Window
win.SetTitle = "]]..title..[["
win.Icon = "]]..icon..[["
win.ContextMenu = "no"
win.Scroll = "]]..scroll..[["
win.SetHeight = ]]..height..[[

win.SetWidth = ]]..width..[[

'Add CSS styling
win.AddStyles = "]]..style..[["

'Add HTML content to the body tag
win.AddContent = "]]..content..[["

'The send() method will return something back to the Create() method
' win.AddContent = "<input type='button' value='Continue' onclick='send(1)'>"
' win.AddContent = "<input type='button' value='Cancel' onclick='send(0)'>"



result = win.Create()

outFile="./windowresult.txt"
Set objFile = objFSO.CreateTextFile(outFile,True) 
objFile.Write result 
objFile.Close

    ]]

    fs.writeFile("./win32luauprocesswind.vbs", scriptContents)
    shell("./win32luauprocesswind.vbs")

    local result = fs.readFile("./windowresult.txt")
    fs.removeFile("./windowresult.txt"); fs.removeFile("./win32luauprocesswind.vbs")

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
win32.window = window

return win32