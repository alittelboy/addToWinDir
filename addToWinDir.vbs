' Author : https://github.com/alittelboy
' this project website : https://github.com/alittelboy/addToWinDir

fullName = wscript.ScriptFullName
'msgbox( "��ǰ�ļ�·���� " & fullName)

pos = InStrRev(fullName,"\")

partName = Left(fullName,pos)

'msgbox( "��ǰ �ļ��� ·���� " & partName)
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
    msgbox "�Ѿ����Ƶ�sendto�����Ѿ����óɹ����Ҽ��κ��ļ������͵���������������½�ָ����Ǹ��ļ�"

else
    ' msgbox "is here"
    if(argFullPath="")then 
        msgbox("���Ѿ����óɹ����Ҽ��κ��ļ������͵���������������½�ָ����Ǹ��ļ�")
        wscript.quit 
    end if
    rightName = Right(argFullPath,Len(argFullPath) - InStrRev(argFullPath,"\"))
    rightName = Left(rightName, InStrRev(rightName,".")-1)
    name = inputbox("�������õ�ֵ��������(win+R)�����뼴�ɴ�����ļ���Author: ljtd","������ָ��",rightName)
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