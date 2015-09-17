<%
startingFolder = Trim("" & Session("STARTFOLDER"))
ext = Trim(Request("ext"))
If startingFolder = "" OR ext = "" Then
    Session.Transfer "fileCounts.asp" ' must not have visited this page first!
    Response.End
End If
%>
<html>
<head>
<title>Show details for Extension <%=ext%></title>
<link rel="stylesheet" type="text/css" href="/main/css/table.css" />

<style>
td { font-size: small; }
th { font-size: medium; }
</style>
</head>
<body>
<font face="Arial">
<a href="#" onClick="history.go(-1)"><font color="#000000" size="4">Back</font></a></font>
<br></br>
<table border=1 cellpadding=3>
<tr>
    <th>#</th>
    <th>file name</th>
    <th>full path</th>
</tr>
<%
Set fso = Server.Createobject("scripting.filesystemobject")

' our recursive subroutine:
counter = 0
Sub ProcessOneDir( dir )
    For Each file In dir.files
        dotAt = InStrRev( file.name, "." )
        If dotAt > 0 Then
            extension = UCase( Mid( file.name, dotAt ) )
        Else
            extension = "[none]"
        End If
        If extension = ext Then
             counter = counter + 1
%>
<tr>
    <td><%=counter%></td>
    <td><%=file.name%></td>
    <td style="font-size: x-small;"><%=file.path%></td>
</tr>
<%
        End If
    Next
    For Each subdir In dir.subfolders
        ProcessOneDir subdir
    Next
End Sub

' and start it off from the starting folder:
ProcessOneDir fso.GetFolder( startingFolder )
%>
</table>
</body>
</head>