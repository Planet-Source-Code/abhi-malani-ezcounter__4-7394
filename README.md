<div align="center">

## ezcounter


</div>

### Description

EzCounter is a Session Counter. In this the Counter increments only once in a Session

get the updated article from http://www.duelcom.com/malani/articles/1/
 
### More Info
 
<!--#include file="counter.asp"-->

or

<%Server.Execute("counter.asp")%>

You need to have read/write permission on the hits.txt file which is in the same directory as this file is.

Returns the no. of counts

To return the no. of counts:

<%=Session("count")%>

or

<%

Response.Write(Session("count"))

%>


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Abhi Malani](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/abhi-malani.md)
**Level**          |Beginner
**User Rating**    |5.0 (99 globes from 20 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__4-1.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/abhi-malani-ezcounter__4-7394/archive/master.zip)

### API Declarations

```
Copyright Abhi Malani
abhimalani@yahoo.com
```


### Source Code

```
"COUNTER.ASP" / VBSCRIPT VERSION
<%
If IsEmpty(Session("count")) Then
  Call Countthisuser
End if
'by Abhi Malani - http://www.duelcom.com/malani/
Sub CountThisUser()
Dim objFSO, objFS, file, intCount
file = Server.MapPath("hits.txt")
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(file) Then
  Set objFS = objFSO.OpenTextFile(file, 1)
Else
  Set objFS = objFSO.CreateTextFile(file, True)
End If
If Not objFS.AtEndOfStream Then
  intCount = objFS.ReadAll
Else
  intCount = 0
End If
objFS.Close
intCount = intCount + 1
Session("count")= intCount
Set objFS = objFSO.OpenTextFile(file, 2)
objFS.Write intCount
objFS.Close
Set objFSO = Nothing
Set objFS = Nothing
End Sub
'Session.Abandon 'for testing the code
%>
JSCRIPT VERSION/ "COUNTER.ASP"
<%@Language=Jscript%>
<%
//by Abhi Malani - http://www.duelcom.com/malani/
var count = Session("count");
if(count=="" || count==null || count=="undefined"){
  countuser();
}
function countuser(){
var ForReading = 1, ForWriting = 2, TristateUseDefault = -2;
var file = "hits.txt";
var fso = new ActiveXObject("Scripting.FileSystemObject");
var f = fso.GetFile(Server.Mappath(file));
var ts = f.OpenAsTextStream(ForReading, TristateUseDefault);
var count = ts.ReadAll();
count = parseInt(count);
count = ((isNaN(count)) ? 1 : count++);
ts.Close( );
ts = f.OpenAsTextStream(ForWriting, TristateUseDefault);
ts.Write(count);
ts.Close();
Session("count") = count;
}
//Session.Abandon //for testing the code
%>
```

