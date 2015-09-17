<HTML>
<head>
<link rel="stylesheet" type="text/css" href="/ccount/css/table.css" /> 
</head>
<BODY>
<%
' setup stuff:
CONST MAXWIDTH = 400 ' width of widest colored bar
startingFolder = Server.MapPath("/ccount/code") ' but could be absolute path directly

Session("STARTFOLDER") = startingFolder

' and then the code starts:
'
' the two needed objects:
Set counts = Server.CreateObject("Scripting.dictionary")
Set fso = Server.Createobject("scripting.filesystemobject")

' our recursive subroutine:
Sub ProcessOneDir( dir )
    For Each file In dir.files
        dotAt = InStrRev( file.name, "." )
        If dotAt > 0 Then
            extension = UCase( Mid( file.name, dotAt ) )
        Else
            extension = "[none]"
        End If
        If counts.exists(extension) Then
            counts(extension) = counts(extension) + 1
        Else
            counts.Add extension, 1
        End If
    Next
    For Each subdir In dir.subfolders
        ProcessOneDir subdir
    Next
End Sub

' and start it off from the starting folder:
ProcessOneDir fso.GetFolder( startingFolder )


' find total and maxcount and convert the dictionary to parallel arrays:
total    = 0
countmax = 0
Dim extArr(), cntArr()
ReDim extArr(counts.count-1), cntArr(counts.count-1)

cc = 0
For Each extension In counts.keys
    ' DEBUG Response.Write cc & "::" & extension & ""
    count = counts.item(extension)
    If count > countmax Then countmax = count
    total = total + count
    extArr(cc) = extension
    cntArr(cc) = count
    cc = cc + 1
Next

' then sort the arrays, in descending count order:
For outer = UBound(cntArr)-1 To 0 Step -1
    For inner = 0 To outer
        If cntArr(inner) < cntArr(inner+1) Then
            swap = cntArr(inner)
            cntArr(inner) = cntArr(inner+1)
            cntArr(inner+1) = swap
            swap = extArr(inner)
            extArr(inner) = extArr(inner+1)
            extArr(inner+1) = swap
        End If
    Next
Next

' finally ready for the display:
%>
<CENTER>
<h3><% server.htmlencode(request.serverVariables("SERVER_NAME"))%> Code Breakdown </h3>
<CENTER>
<TABLE BORDER=1 cellpadding=3 STYLE="font-family: helvetica, arial, times;font-size:10pt;">
<TR>
    <TH>Type</TH>
    <TH># Files</TH>
    <TH>Percent</TH>
    <TH>Graphical Count</TH>
</TR>
<%
colors = Array("Red","Blue","Brown","Green","Orange","Lightblue","Lightgreen","Purple")
colorCount = UBound(colors)+1
For cc = 0 To UBound(cntArr)
    count = cntArr(cc)
    extension = extArr(cc)
    pct = count/total
    color = colors( cc MOD colorCount )
    width = (count/countmax) * MAXWIDTH
%>
<TR>
    <TD><A href="showExtension.asp?ext=<%=extension%>"><%=extension%> Files</a></TD>
    <TD align=right><%=count%></TD>
    <TD align=right><%=FormatPercent(count/total)%></TD>
    <TD><TABLE><TR><TD style="background-color: <%=color%>; height: 12px; width: <%=width%>px;"></TD></TR></TABLE></TD>
</TR>
<%
Next
%>
</TABLE>
</CENTER>
</BODY>
</HTML>