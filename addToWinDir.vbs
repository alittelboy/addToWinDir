' Author : https://github.com/alittelboy
' this project website : https://github.com/alittelboy/addToWinDir

fullName = wscript.ScriptFullName
'msgbox( "当前文件路径是 " & fullName)

pos = InStrRev(fullName,"\")

pathName = Left(fullName,pos)

'msgbox( "当前 文件夹 路径是 " & pathName)
set objWsh = CreateObject("WScript.Shell")
strApp = objWsh.ExpandEnvironmentStrings("%AppData%") 
strSendTo = strApp & "\Microsoft\Windows\SendTo\"
' msgbox("Install Location:" & strSendTo)

strWinDir = strApp & "\WinDir\"

Dim fso
Set fso=CreateObject("Scripting.FileSystemObject")        
If fso.folderExists(strWinDir) Then         
        
Else 
    fso.CreateFolder(strWinDir)
End If 

Set WshShell = Wscript.CreateObject("Wscript.Shell")
strPath = WshShell.Environment("user").Item("path")
'msgbox strPath
if InStr(strPath, strWinDir)<=0 then
   
    WshShell.Environment("user").Item("path")=strWinDir &";"& WshShell.Environment("user").Item("path")
    Set WshShell = Nothing
end if 


'msgbox strSendTo=pathName

' 读取参数文件地址
dim argFullPath
Set oArgs = WScript.Arguments
    For Each s In oArgs
        argFullPath = s
    Next
Set oArgs = Nothing


if(not(strSendTo=pathName))then
    ' copy to
	msgbox("Install Location:" & strSendTo)
	' msgbox("Commands Location:" & strWinDir)
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile fullName, strSendTo
    ' msgbox "已经复制到sendto。你已经设置成功，右键任何文件，发送到，本软件，即可新建指令打开那个文件"
    msgbox "Successfuly copied to path SendTo. Installation finished. Right click the file you want, send to, this app, to create cmd"
	Set WshShell=WScript.CreateObject("WScript.shell")
    Set Shortcut=WshShell.CreateShortCut(strWinDir & "windir.lnk") 
    Shortcut.Hotkey = "" 
    Shortcut.IconLocation = strWinDir
    Shortcut.TargetPath = strWinDir
    Shortcut.Save 
	Set WshShell=WScript.CreateObject("WScript.shell")
    Set Shortcut=WshShell.CreateShortCut(strWinDir & "send2.lnk") 
    Shortcut.Hotkey = "" 
    Shortcut.IconLocation = strSendTo
    Shortcut.TargetPath = strSendTo
    Shortcut.Save 

else
    ' msgbox "is here"
    if(argFullPath="")then 
        ' msgbox("你已经设置成功，右键任何文件，发送到，本软件，即可新建指令打开那个文件")
		msgbox("You have set up! Right click the file you want, send to, this app, to create cmd")
        wscript.quit 
    end if
    rightName = Right(argFullPath,Len(argFullPath) - InStrRev(argFullPath,"\"))
    if InStr(rightName,".")<>0 then
        rightName = Left(rightName, InStrRev(rightName,".")-1)
    end if
    ' name = inputbox("这里设置的值，在运行(win+R)里输入即可打开你的文件。Author: ljtd","输入快捷指令",rightName)
    name = inputbox("set the command here, use win+R + command to open your file。Author: ljtd","Set Command",rightName)
    if(name="")then 
        wscript.quit 
    end if
    Set WshShell=WScript.CreateObject("WScript.shell")
    Set Shortcut=WshShell.CreateShortCut(strWinDir & name & ".lnk") 
    Shortcut.Hotkey = "" 
    Shortcut.IconLocation = argFullPath
    Shortcut.TargetPath = argFullPath
    Shortcut.Save 

end if
