' Author : https://github.com/alittelboy
' this project website : https://github.com/alittelboy/addToWinDir

fullName = wscript.ScriptFullName
'msgbox( "当前文件路径是 " & fullName)

pos = InStrRev(fullName,"\")

partName = Left(fullName,pos)

'msgbox( "当前 文件夹 路径是 " & partName)
set objWsh = CreateObject("WScript.Shell")
strApp = objWsh.ExpandEnvironmentStrings("%AppData%") 
strSendTo = strApp & "\Microsoft\Windows\SendTo\"
'msgbox(strSendTo)

'msgbox strSendTo=partName

dim argFullPath
Set oArgs = WScript.Arguments
    For Each s In oArgs
        argFullPath = s
    Next
Set oArgs = Nothing


if(not(strSendTo=partName))then
    ' copy to
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile fullName, strSendTo
    msgbox "已经复制到sendto。你已经设置成功，右键任何文件，发送到，本软件，即可新建指令打开那个文件"

else
    ' msgbox "is here"
    if(argFullPath="")then 
        msgbox("你已经设置成功，右键任何文件，发送到，本软件，即可新建指令打开那个文件")
        wscript.quit 
    end if
    rightName = Right(argFullPath,Len(argFullPath) - InStrRev(argFullPath,"\"))
    rightName = Left(rightName, InStrRev(rightName,".")-1)
    name = inputbox("这里设置的值，在运行(win+R)里输入即可打开你的文件。Author: ljtd","输入快捷指令",rightName)
    if(name="")then 
        wscript.quit 
    end if
    Set WshShell=WScript.CreateObject("WScript.shell")
    Set Shortcut=WshShell.CreateShortCut("C:\WINDOWS\" & name & ".lnk") 
    Shortcut.Hotkey = "" 
    Shortcut.IconLocation = argFullPath
    Shortcut.TargetPath = argFullPath
    Shortcut.Save 

end if