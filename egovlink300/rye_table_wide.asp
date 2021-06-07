<% Option Explicit %>  
<%  
Dim Conn,strSQL,objRec  
Set Conn = Server.Createobject("ADODB.Connection")  
Conn.Open Application("DSN")

Dim PageLen,PageNo,TotalRecord,TotalPage,No,intID  ,MaxRows,MaxPage
PageLen = 50
PageNo = Request.QueryString("Page")  
if PageNo = "" Then PageNo = 1  

strSQL = "WITH YourCTE AS (SELECT u.userfname + ' ' + u.userlname as Name, ad.answer as Address, MIN(a.parcelidnumber) as ParcelID, " _
		& " sd.answer as StartDate, DATEADD(d,29,CONVERT(datetime,sd.answer)) as EndDate, rm.answer as RemovalType, " _
		& " ROW_NUMBER() OVER (ORDER BY sd.answer DESC) rn  " _
	& " FROM egov_actionline_requests ar " _
	& " INNER JOIN egov_users u ON u.userid = ar.userid " _
	& " INNER JOIN action_submitted_questions_and_answers ad ON ad.action_autoid = ar.action_autoid and ad.question = 'Address' " _
	& " INNER JOIN egov_residentaddresses a ON a.orgid=ar.orgid AND ad.answer = a.residentstreetnumber + ' ' + a.residentstreetname " _
	& " INNER JOIN action_submitted_questions_and_answers sd ON sd.action_autoid = ar.action_autoid and sd.question = 'Start Date' " _
	& " INNER JOIN action_submitted_questions_and_answers rm ON rm.action_autoid = ar.action_autoid and rm.question = 'Type of Removal' " _
	& " WHERE ar.category_id = '17890' AND ar.status = 'RESOLVED' " _
	& " GROUP BY a.residentstreetname,a.residentstreetnumber, u.userfname, u.userlname, ad.answer, sd.answer, rm.answer " _
	& " ) " _
	& " SELECT *, (SELECT MAX(rn) FROM YourCTE) AS 'TotalRows' FROM YourCTE WHERE rn BETWEEN " & ((PageNo - 1) * PageLen + 1) & " AND " & (PageNo * PageLen) & " "
'response.write strSQL & "<br />"

Set objRec = Server.CreateObject("ADODB.Recordset")  
objRec.Open strSQL, Conn, 1,3  
  
If objRec.EOF Then  
Response.write (" Not found record.")  
Else  
  
%>  
<head>
<style type="text/css" media="screen">

div, td, th, h2, h3, h4 {
  font-family: verdana,sans-serif;
	font-size:    12px;
	voice-family: "\"}\"";
	voice-family: inherit;
	font-size: 12px;
	color: #333;
}  


table { border: 1px solid; border-collapse: collapse; border-color: #82966d; }
th, td { border: 1px solid; border-color: #82966d; vertical-align: top; }
th { background-color: #82966d; color: #d4dbcc; font-weight: normal;}
h a:link,  th a:visited,  th a:active,  th a:hover {
	color: #d4dbcc;
}

.d1 { background-color: #d4dbcc }
.d2 { background-color: #f2faeb }




</style>
</head>
<body>
<table cellpadding="3" cellspacing="0" border="1">
<tr>  
		<th>Address</th>
		<th>Parcel ID (S/B/L)</th>
		<th>Rock Removal Type</th>
		<th>Start Date</th>
		<th>End Date</th>
</tr>  
<%  
Do While Not objRec.EOF 
	MaxRows = objRec("TotalRows")
      	%><tr><td><%=objRec("Address")%></td><td><%=objRec("ParcelID")%></td><td><%=objRec("RemovalType")%></td><td><%=objRec("StartDate")%></td><td><%=objRec("EndDate")%></td></tr><%
objRec.MoveNext  
Loop  

MaxPage = ceil(Cint(MaxRows) / PageLen)

%>  
<tr><td colspan=5>
<div style="position:relative;">

<div style="width:50%;margin: 0 auto;text-align:center;">Page: <%=PageNo%>/<%=MaxPage%></div>
<% if Cint(PageNo) > 1 then %>
<div style="position:absolute; top:0; left:0;"><a href="rye_table_test.asp?page=<%=PageNo-1%>">&lt;- Prev</a></div>
<% end if%>
<% if Cint(PageNo) < Cint(MaxPage) then %>
<div style="position:absolute; top:0; right:0;"><a href="rye_table_test.asp?page=<%=PageNo+1%>">Next -&gt;</a></div>
<% end if%>

</div>
</td></tr>

</table>  
 <% 
 end if
objRec.Close()  
Conn.Close()  
Set objRec = Nothing  
Set Conn = Nothing  
%>
</body>
<%
function ceil(x)
        dim temp
 
        temp = Round(x)
 
        if temp < x then
            temp = temp + 1
        end if
 
        ceil = temp
    end function
%>  
