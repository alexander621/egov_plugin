<%
dim pagesize, totalpages,totalrecords,groupmode
dim thisname,Currentpage,rs,strSQL,conn,numstartid,numendid,strDirectory,conn2,rs2,strSQL2,AllUsers,totalUsers,strUser,totalContacts,strContact,AdditonURL,deleteurl
dim RA,i,GroupNumber,editurl,temp,Str_Bgcolor,HowManyCommitteeDisplayed
HowManyCommitteeDisplayed=0
'groupmode=1, display individual group
'groupmode=2, display all member
pagesize=Session("PageSize")
totalpages=1

thisname=request.servervariables("script_name")
if not isempty(request.querystring("currentpage")) then
CurrentPage=clng(request.querystring("currentpage"))
else
currentpage=1
end if
%>
<%  
' the above value will be provided by the display_committee.asp
'page size, RA, pagerecord, currentpage values must be declared to global variables.
DisplayCommitteeRecords
dim bCanEdit
'dim rs,conn,strSQL,totalrecords,totalpages,pagesize,Currentpage,numstartid,numendid,deleteurl,groupnumber,
%>

<table border=0 cellpadding=5 cellspacing=0 width='100%' class=""tablelist"">
<form name=DeleteCommittee method=post action='"&deleteurl&"'>
<tr style="height:25px;"><th align=left></th></tr>

<tr><td></td></tr>

</form>
</table>
