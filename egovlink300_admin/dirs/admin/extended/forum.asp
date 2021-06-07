<!--#include file="dbaction.asp"-->
<%
'--------   creat/drop table --------------
'	const ActNewTable = 1	 'Create a new table.
'	const ActDropTable = 2	 'Drop a existing table.
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
' ----   create and drop new table -------------
		case ActNewTable
             set of = new DBActionClass
	'		 of.createtable
			 destroy of	
		case ActDropTable
             set of = new DBActionClass
'             response.write "hello"& of.tablename
	'		 of.droptable
			 destroy of	
'-------------- Add new record to the database, 
		case ActNewPost
			%>
			<!--#include file="NewPost.asp"-->
			<% 
'-------------- Save new record to the database, 
		case ActNewPostSave
		response.write "<B>"&langAdminPropertyAdded&"</B><br>"
        set of = new DBActionClass
			of.AddNew
			destroy of	
'-------------- display all or one record ----
		case ActDisplayRecords
			DisplayRecords
'----------  display --  Modify------------
		case ActUpdatePost
			id=clng(request.querystring(conTABLEFIELDS(0)))
			UpdatePost id
		case ActUpdatePostSave
			id=clng(request.form(conTABLEFIELDS(0)))
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
strSQL = "select * from " & conTABLENAME &" where "&conTABLEFIELDS(1)&"="&clng(request.querystring("userid"))
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
'DisPlayNavbar
response.write "<table border=0 cellpadding=5 cellspacing=0 class='tablelist' cellspacing=1 align=center width=400 bgcolor='" & cForumBgColor & "'>"
response.write "<tr>"
response.write " <th align=left>"&langDelete&"</th>"

for i=2 to intcolumns
if l_display(i)=1 then response.write " <th align=left>"&conTABLEFIELDS(i)&"</th>"
next
response.write "</tr>"
editurl=thisname&"?iofaction="&ActUpdatePost&"&userid="&request.querystring("userid")
deleteurl=thisname&"?iofaction="&ActDelete&"&userid="&request.querystring("userid")
response.write "<form method=post action='"&deleteurl&"'>"
for i = numstartid to numendid
response.write "<tr>"& _
   " <td><input name=delete type=checkbox value='"&RA(0,i)&"'></td>"& _
   " <td><a href='"&editurl&"&"&conTABLEFIELDS(0)&"="&RA(0,i)&"'>"&RA(2,i)&"</a></td>"
  for j=3 to intcolumns   
  if l_display(j)=1 then response.write " <td>&nbsp;"&RA(j,i)&"</td>" 
  Next
  response.write " </tr>"
next     
response.write "</table>"
response.write "<CENTER><INPUT type=submit value="&langDelete&"></CENTER> </form>"
destroy RA
end sub 

sub UpdatePost(id)
set cDB = new DBClass
	cDB.Open sConnForum
response.write "<h4> "&Updateproperty&"</h4>"
strSQL = "select * from " & conTABLENAME & " where "&conTABLEFIELDS(0)&"="& id
set rs=cDB.GetRS(strSQL)
%>
<!--#include file="UpdatePost.asp"-->
<%
set rs=nothing
destroy cDB
set cDB=nothing
end sub

function GetPagingURL(i)
GetPagingUrl=thisname&"?iofaction="&request.querystring("iofaction")&"&currentpage="&i
end function
%>

