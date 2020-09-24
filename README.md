<div align="center">

## UniSock \- Winsock control replacement


</div>

### Description

A new class implementation of Winsock API. By style close to the native Winsock control and other class implementations such as CSocket and CSocketMaster, so you don't need to learn or rewrite much of existing code. The new and cool part about this class is that it is just one class file. Also, it performs better (by speed) and handles errors a bit more cleanly (you aren't forced to close the socket each time an error occurs). Other speciality is transparent Unicode support: when you switch to text mode, you start receiving TextArrival event instead of DataArrival and start getting individual lines. These lines are automatically Unicode if received line is UTF-8 or UTF-16! ANSI lines require you to use StrConv to get an usable string, thus you have the power on what to do with the raw data before any conversion has affected it. 

----

Support thread and documentation: http://www.vbforums.com/showthread.php?t=534580 

----

Special thanks to Paul Caton and LaVolpe for their work on SelfSub, SelfHook and SelfCallback. 

----

UPDATE 2008-08-12: Now Vista compatible!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2008-08-07 11:59:54
**By**             |[Vesa Piittinen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vesa-piittinen.md)
**Level**          |Intermediate
**User Rating**    |4.8 (67 globes from 14 users)
**Compatibility**  |VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[UniSock\_\-\_2123328122008\.zip](https://github.com/Planet-Source-Code/vesa-piittinen-unisock-winsock-control-replacement__1-70932/archive/master.zip)








