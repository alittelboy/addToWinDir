' Author : https://github.com/alittelboy
' this project website : https://github.com/alittelboy/addToWinDir

fullName = wscript.ScriptFullName
'msgbox( "��ǰ�ļ�·���� " & fullName)

pos = InStrRev(fullName,"\")

pathName = Left(fullName,pos)

'msgbox( "��ǰ �ļ��� ·���� " & pathName)
set objWsh = CreateObject("WScript.Shell")
strApp = objWsh.ExpandEnvironmentStrings("%AppData%") 
strSendTo = strApp & "\Microsoft\Windows\SendTo\"
'msgbox(strSendTo)

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

' ��ȡ�����ļ���ַ
dim argFullPath
Set oArgs = WScript.Arguments
    For Each s In oArgs
        argFullPath = s
    Next
Set oArgs = Nothing


if(not(strSendTo=pathName))then
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
    if InStr(rightName,".")<>0 then
        rightName = Left(rightName, InStrRev(rightName,".")-1)
    end if
    name = inputbox("�������õ�ֵ��������(win+R)�����뼴�ɴ�����ļ���Author: ljtd","������ָ��",rightName)
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