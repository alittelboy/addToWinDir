# addToWinDir
设置运行里的自定义指令，运行里输入**指令**使用。

例如，你可以通过此脚本设置字符串“QQ”对应 腾讯qq。然后win+R打开运行，输入QQ，回车，就打开腾讯QQ了！

vbs脚本，仅Windows可用。

~~有些win电脑无法使用，是因为C:\\Windows目录下无法新建超链接文件，暂无找到解决办法。~~

## 下载
点击右上角code，download zip，解压后得到addToWinDir.vbs文件。

## 第一次运行
双击addToWinDir.vbs即可配置。

## 之后使用
找到想快捷启动的文件或文件夹，右键，发送到，选择addToWinDir.vbs，设置自定义**指令**。

之后，在运行（win+R）里输入**指令**，即可打开文件。

## 致谢
本项目是受到此文的启发，
https://blog.csdn.net/hggjgff/article/details/84087589

## 工作原理
 - 右键文件，发送到功能。Windows下有个文件夹，地址是%appdata%\Microsoft\Windows\SendTo，右键发送到的选项都在这里设置。
 - 发送参数。右键发送到，只传递一个参数，就是文件地址。
 - 运行自定义指令。要实现自定义指令，只要把文件/文件夹快捷方式创建在C:\\Windows目录下即可，快捷方式的名字就是自定义指令。
 - 新的实验表明，只要是path环境变量下的文件夹，都可以存放指令超链接。


## 我的工作
 - 首先，原脚本是使用bat生成vbs实现的，不优雅，我直接改成了纯vbs实现。
 - 原脚本是针对Windows xp的，现在已经无法使用，故修改了SendTo文件地址。
 - 原脚本把文件超链接生成在了C:\\Windows目录下,由于win10的权限设置，这个办法在有的电脑上不能成功。于是我改成了在%appdata%\windir目录下生成超链接，并把这个文件夹放在了环境变量里，这就解决了这个问题。
