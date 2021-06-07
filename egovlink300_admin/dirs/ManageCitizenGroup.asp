<!--#include file='header_simply.asp'-->


<%
' EGOVLINK PERMISSION CHECK
'If not HasPermission("CanEditCommittee") and not HasPermission("CanEdit"&CommitteeName) then
'	response.redirect "InvalidRole.asp?error="&server.urlencode(langInvalidRoleEditCommittee)
'End If
%>


<!-- <body onload1="javasript:opener.location.reload(true);" bgcolor="#c9def0"> -->
<body bgcolor="#c9def0">


<%
' CHECK FOR VALID CITIZEN ID PASSED VIA QUERYSTRING
If trim(request.querystring("userid"))="" Then
	response.write "<br />Error: Missing citizen ID!"
	response.end
End If

' SET WINDOW CAPTION INFORMATION
thisname=request.servervariables("script_name")
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
set rs = Server.CreateObject("ADODB.Recordset")
set rs.ActiveConnection = conn
rs.CursorLocation = 3 
rs.CursorType = 3 
strSQL="select userfname,userlname,userid from egov_users u where u.userid="&CLng(request.querystring("userid"))
rs.Open strSQL
Title="<strong>"&rs("userfname")&" "&rs("userlname")&"("&rs("userid")&")</strong>"
rs.close
set rs=nothing
%>


<!--  BEGIN: DISPLAY THE USER'S GROUP MEMBERSHIP --> 
<table border="0" cellspacing="0" cellpadding="5" align="center">
<tr><td align="center"><font size="+1">Name: <b> <%=Title %> </b></font><br></td></tr>
<tr><td>
<%
thisname=request.servervariables("script_name")
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
set rs1 = Server.CreateObject("ADODB.Recordset")
set rs1.ActiveConnection = conn
rs1.CursorLocation = 3 
rs1.CursorType = 3 
set rs2 = Server.CreateObject("ADODB.Recordset")
set rs2.ActiveConnection = conn
rs2.CursorLocation = 3 
rs2.CursorType = 3 

strSQL1 = "select ugp.groupid,ugp.groupname from vwCitizenGroups ugp where ugp.citizenid="&CLng(request.querystring("userid"))
strSQL2 = "select g.groupid,g.groupname from CitizenGroups g where (g.orgid=" & session("orgid") & ") AND g.groupid not in ( select ugp.groupid from egov_users u inner join vwCitizenGroups ugp on u.userid=ugp.citizenid where (g.orgid=" & session("orgid") & ") AND (u.userid="&CLng(request.querystring("userid")) &"))"

rs1.Open strSQL1
rs2.Open strSQL2

response.write "<table border=""0"" cellpadding=""5"" cellspacing=""0"" width=""350"" class=""tablelist""> <tr>"
response.write "<td>"
response.write "Already in Groups"
ExistingRoleList
response.write "</td><td>"
response.write "&nbsp;<a href='javascript:document.c1.submit();'><img src='../images/ieforward.gif' align=""absmiddle"" border=""0""></a>"
response.write "<br /><br />"
response.write "<a href='javascript:document.r1.submit();'><img src='../images/ieback.gif' align=""absmiddle"" border=""0""></a>"
response.write "</td><td>"
response.write "Available Groups"
TheRemainingMemberList
response.write "</td>"
response.write "</tr></table>"
rs1.close
set rs1=nothing
rs2.close
set rs2=nothing

conn.close
set conn=nothing
%>
<!--  END: DISPLAY THE USER'S GROUP MEMBERSHIP --> 

</td>
</tr>
</table>
</body>


<%
'--------------------------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'--------------------------------------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' SUB EXISTINGROLELIST
'--------------------------------------------------------------------------------------------------
Sub ExistingRoleList()
   
   response.write vbcrlf & "<form name=""c1"" method=""post"" action=""DeleteCitizenGroup.asp?userid="&trim(request.querystring("userid"))&""">"
   
   ' LIST USER'S GROUP MEMBERSHIP
   response.write vbcrlf & "<select size='10'  WIDTH=""150"" STYLE=""width:140px"" name=""ExistingList"" multiple=""multiple"">"
	   For i=0 to rs1.recordcount-1
			memberstr = rs1("GroupName") & "("&rs1("GroupID") & ")"
    		response.write vbcrlf & "<option value=" & rs1("GroupID") & ">" & memberstr & "</option>"
			rs1.movenext
		Next 
    response.write vbcrlf & "</select></p>"
   
	response.write vbcrlf & "</form>"

End Sub


'--------------------------------------------------------------------------------------------------
' SUB THEREMAININGMEMBERLIST
'--------------------------------------------------------------------------------------------------
Sub TheRemainingMemberList()
	
	response.write vbcrlf & "<form name=""r1"" method=""post"" action=""AddCitizenGroup.asp?userid="&trim(request.querystring("userid"))&""">"
   
	' LIST REMAINING GROUPS 
	response.write vbcrlf & "<select size=""10""  WIDTH=""150"" STYLE=""width:140px"" name=""RemainingList"" multiple=""multiple"">"
	   For i=0 to rs2.recordcount-1
			memberstr = rs2("GroupName") & "(" & rs2("GroupID") & ")"
    		response.write vbcrlf & "<option value=" & rs2("GroupID") & ">" & memberstr & "</option>"
			rs2.movenext
	   Next
    response.write vbcrlf & "</select></p>"	 
    
	response.write vbcrlf & "</form>"

End Sub
%>