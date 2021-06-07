<!--#include file="dbaction.asp"-->
<%
'---------- Insert New reocrd ------------------- 
	const ActNewPost=4  'Insert New record 
'---------- Save New reocrd ------------------- 
	const ActNewPostSave=5  'Insert New record 
'---- display the list of records ---------------
	const ActDisplayRecords=6
'----------- Update and save --------
	const ActUpdatePost=7
	const ActUpdatePostSave=8
'----------- Delete --------
	const ActDelete=9
%>
<%
dim aSorted()
dim amessages
dim currentpage
dim totalpages
sub ShowForum
	iOfAction = request("iOfaction")
	if isNumeric(iOfAction) then iOfAction = CLng(iOfAction) else iOfAction = 0
'	print vbCrlf&vbCrlf & "<!-- Begin UC New Student Automatic Pick Up System -->" & vbCrLf & vbCrLf
	select case iOfAction
'-------------- Add new record to the database, 
		case ActNewPost
			%>
			<!--#include file="NewPost.asp"-->
			<% 
'-------------- Save new record to the database, 
		case ActNewPostSave
        set of = new DBActionClass
			of.AddNew
			destroy of	
		response.write langResponseAfterPostSave&"<br>"
'-------------- display all or one record ----
		case ActDisplayRecords
			DisplayRecords
'----------  display --  Modify------------
		case ActUpdatePost
			id=clng(request.querystring("id"))
			UpdatePost id
		case ActUpdatePostSave
			id=clng(request.form(conTABLEFIELDS(IDCol)))
			 set of = new DBActionClass
			 of.saveupdate id
			 destroy of		
      response.write "<br><a href='javascript:history.go(-2)'>"&langGoBack&"</a>"
'-------------Delete -----------------------------------
		case ActDelete
		 set of = new DBActionClass
		 for each delete in request("delete")
		id=clng(delete)
		of.DeleteRecord id
		response.write "<br>"&langdelete&" ID=<FONT COLOR='red'>"&id&"</FONT>"
	 	next
		destroy of
      response.write "<br><a href='javascript:history.go(-2)'>"&langGoBack&"</a>"
'----------- default -----------------------------
		case else

		end select
	
'	print vbCrlf & vbCrlf & "<!-- Begin OpenForum -->" & vbCrLf & vbCrLf
'	print vbCrlf & vbCrlf & "<!-- End UC New Student Automatic Pick Up System -->" & vbCrLf & vbCrLf
end sub
%>

<%
sub DisplayRecords
if isempty(trim(request("currentpage"))) then currentpage=1 else  currentpage=clng(request("currentpage")) 
 ' for only display all the page, the minimum value is 1
'strSQL = "select * from " & conTABLENAME &" order by "&conTABLEFIELDS(0)&" DESC"
strSQL = "select * from " & conTABLENAME
set of = new DBActionClass
intcolumns=of.intcolumns
' -- input--
of.strSQL=strSQL
of.pagesize=cpagesize
of.currentpage=currentpage
'-- in order to get output, the following calling must be made--
RA=of.getrecordarray  ' RA= record array
'-- output ---
numstartid=of.numstartid
numendid=of.numendid
totalpages=of.totalpages
totalrecords=of.totalrecords
destroy of
if totalrecords>cpagesize  then DisPlayNavbar
response.write "<table border=0 cellpadding=5 class='tablelist' cellspacing=1 align=center width=400 bgcolor='" & cForumBgColor & "'>"
response.write "<tr>"
response.write " <td>"&langDelete&"</td>"

'RS = open Recordset object

for i=0 to intcolumns
if l_display(i)=1 then response.write " <td>"&conTABLEFIELDS(i)&"</td>"
next
response.write "<td>More Roles</td>"
response.write "</tr>"
editurl=thisname&"?iofaction="&ActUpdatePost
deleteurl=thisname&"?iofaction="&ActDelete
response.write "<form method=post action='"&deleteurl&"'>"
for i = numstartid to numendid
response.write "<tr>"& _
  " <td><input name=delete type=checkbox value='"&RA(IDCol,i)&"'></td>"& _
   " <td><a href='"&editurl&"&id="&RA(IDCol,i)&"'>"&RA(0,i)&"</a></td>"
  for j=1 to intcolumns   
  if l_display(j)=1 then 
  '----------------------------------
      extraStr="&nbsp;"
			if j=1 then
			set cDB2 = new DBClass
			cDB2.Open sConnForum
			strSQL = "select firstname+lastname as extra from users where userid="&RA(j,i)
			set rs2=cDB2.GetRS(strSQL)
			extraStr=rs2("extra")
			end if
			if j=2 then
			set cDB2 = new DBClass
			cDB2.Open sConnForum
			strSQL = "select rolename as extra from roles where roleid="&RA(j,i)
			set rs2=cDB2.GetRS(strSQL)
			extraStr=rs2("extra")
			end if
'			if j=2 then strSQL = "select groupname as extra from groups where groupid="&RA(j,i)
	
		destroy cDB2
  '---------------------------------------------
  response.write " <td>"&RA(j,i)&" <I>("&extraStr&")</I>"&"</td>" 
  end if
  Next
'------ from RA, get out the rolelist from group---------
set cDB = new DBClass
	cDB.Open sConnForum
strSQL = "select distinct rolename from "& _
		"( ( (users u inner join usersgroups ug on u.userid=ug.userid)"& _
		"			inner join groupsroles rg on ug.groupid=rg.groupid ) "& _
		"			inner join roles r on rg.roleid=r.roleid)"& _
		" where u.userid="&RA(IDCol,i)&"order by rolename"
set rs=cDB.GetRS(strSQL)
while not rs.eof
strGroupRoles=rs("rolename")&" "&strGroupRoles
rs.movenext
wend
destroy cDB
'---------------------------------------------
  response.write "<td><A HREF=''>&nbsp;"&strGroupRoles&"</A></td>"
  response.write " </tr>"
next     
response.write "</table>"
response.write "<CENTER><INPUT type=submit value="&langDelete&"></CENTER> </form>"
destroy RA
end sub 

sub UpdatePost(id)
set cDB = new DBClass
	cDB.Open sConnForum
strSQL = "select * from " & conTABLENAME & " where "&conTABLEFIELDS(IDCol)&"="& id
set rs=cDB.GetRS(strSQL)
%>
<!--#include file="UpdatePost.asp"-->
<%
set rs=nothing
destroy cDB
set cDB=nothing
end sub

sub DisplayNavBar
	sFirst = IIf(currentpage <> 1, "<a href=""" & GetPagingURL(1) & """>" & langFirst & "</a>", langFirst)
	sPrev = IIf(currentpage > 1, "<a href=""" & GetPagingURL(currentpage - 1) & """>" &langPrev & "</a>", langPrev)
	sNext = IIf(currentpage < totalpages, "<a href=""" & GetPagingURL(currentpage + 1) & """>" & langNext & "</a>", langNext)
	sLast = IIf(currentpage <> totalpages, "<a href=""" & GetPagingURL(totalpages) & """>" &langLast & "</a>", langLast)
	sNavBarHeader = "<table border=""0"" width=""100%"" bgcolor=""" & cNavBarBgColor & """><tr>"
	sNavBarFooter = "</tr></table>"
	sOfNavBar = "<td align=""center"">" & sFirst & " | " & sPrev & " | " & sNext & " | " & sLast & "    ("&currentpage&"/"&totalpages&")</td>"
	response.write sNavBarHeader & sOfNavBar & sNavBarFooter
	response.write "<img width=""1"" height=""" & cOfSeperatorHeight & """>"
end sub

sub list(j)
set conn = Server.CreateObject("ADODB.Connection")
conn.Open sConnForum
set rs = Server.CreateObject("ADODB.Recordset")
set rs.ActiveConnection = conn
rs.CursorLocation = 3 'adUseClient
rs.CursorType = 3 'adOpenStatic
select case j
case 1
openstr="select userid as id,(firstname+' '+lastname) as description from users"
case 2
openstr="select roleid as id,rolename as description from roles"
case else
exit sub
end select
'response.write openstr
rs.Open openstr,,, 2 'adCmdTable
response.write "<option>select one</option>"&vbcrlf
while not rs.eof
response.write "<option value="&rs("id")&">"&left(rs("description"),20)&"</option>"&vbcrlf
rs.movenext
wend
rs.close
conn.close
end sub

function GetPagingURL(i)
GetPagingUrl=thisname&"?iofaction="&request.querystring("iofaction")&"&currentpage="&i
end function
%>

