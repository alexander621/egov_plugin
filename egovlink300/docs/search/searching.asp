<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="../../includes/start_modules.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: doctree.asp
' AUTHOR: Steve Loar	
' CREATED: 3/26/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Documents.
'
' MODIFICATION HISTORY
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
%>

<html>
<head>
	<title>E-Gov Services - <%=sOrgName%></title>

	<!-- Required CSS -->
	<link rel="stylesheet" type="text/css" href="../../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../../global.css" />
	<link rel="stylesheet" type="text/css" href="../../css/style_<%=iorgid%>.css" />


</head>

	<!--#Include file="../../include_top.asp"-->

	

	<tr><td valign="top">
		

<%
Dim oConn, iRow
'Dim sSql
Dim oRs
Dim sLocationName
Dim sPath, FSO

iRow = 0

Set FSO = CreateObject("Scripting.FileSystemObject")

sLocationName =  GetVirtualDirectyName()
'sPath = FSO.GetFolder(Server.Mappath("/public_documents300/" & sLocationName & "/published_documents/"))
sPath = FSO.GetFolder(Server.Mappath("/public_documents300/gpw/published_documents/"))
'sPath = "http://www.egovlink.com/public_documents300/eclink/published_documents/Recreation"

'sSql = "SELECT DocTitle, vpath, size, characterization FROM TABLE egovlink300..SCOPE('DEEP TRAVERSAL OF """ & sPath & """') ORDER BY size DESC"
sSql = "SELECT DocAuthor, vpath, size,  DocTitle, characterization, filename, rank FROM egovlink300..SCOPE('DEEP TRAVERSAL OF ""file:\\" & sPath & """') WHERE CONTAINS(Contents, '""holiday""') > 0  ORDER BY rank DESC"
response.write sSql & "<br /><br />"
'response.End 

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open "provider=MSIDXS"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Maxrecords = 30

Set oRs = oConn.Execute(sSql)

If oRs.EOF Then
	response.write "No results"
Else
	Do While Not oRs.EOF
		iRow = iRow + 1
		response.write "<p>" & oRs("rank") & "% <a href=""" & oRs("vpath") & """>" & oRs("filename") & "</a><br>" & vbcrlf
		response.write "Abstract: " & oRs("characterization") & "<br>" & vbcrlf
		response.write "Title: " & oRs("DocTitle") & "<br>" & vbcrlf
		response.write "File size: " & oRs("size") & " bytes</p>" & vbcrlf
		oRs.MoveNext
		If iRow = 10 Then
			Exit Do 
		End If 
	Loop 
End If


oRs.Close
Set oRs = Nothing 
oConn.Close
Set oConn = Nothing
'Set FSO = Nothing 

%>

</td></tr></table>

<!--#Include file="../../include_bottom.asp"-->


