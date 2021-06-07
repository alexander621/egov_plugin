<!--#include virtual="/header_s.asp"-->
<!--#include virtual="/checkrole.asp"-->
<!--#include file="forum.asp"-->
<!--#include file="../NewStudentMenu.asp"-->
<%
Dim CheckUser,login
login=false
authorized=false
if not isempty(Session("userid")) then
authorid=session("userid")
login =true
else 
response.write "You didn't login in, or session expires, please login in first<br>"
'Response.redirect "member/login.asp?error=Sorry%20Your%20Password%20Was%20Wrong"
end if
if login then
	temp=request.servervariables("script_name")	
	j=0
	while  instr(temp,"/")>0 and j<3
	j=j+1
	totallen=len(temp)
	pos=instr(temp,"/")
	howmany=totallen-pos
	myright=right(temp,howmany)
	temp=myright
'	response.write "<br>special="&special&" temp="&temp
	wend
	temp=trim(replace(temp,".asp",""))
'	response.write "<br>*special="&special&" temp="&temp
    if checkrole(temp)  then 
	authorized=true
	end if
end if
if authorized then
'response.write "<a href="&thisname&"?iOfaction=" & ActNewTable&">Create User talbe</a><br>"
'response.write "<a href="&thisname&"?iOfaction=" & ActDropTable&">Drop User table</a><br>"
'response.write "<a href="&thisname&"?iOfaction=" & ActNewPost &">Add A new record</a><br>"
'response.write "<a href="&thisname&"?iOfaction=" & ActDisplayRecords  & ">Display list of Records</a><br>"
'response.write "<a href="&thisname&"?iOfaction=" & ActMemberlogin&">Login In</a><br>"
SHowForum
end if
%>
<!--#include virtual="/footer.asp"-->