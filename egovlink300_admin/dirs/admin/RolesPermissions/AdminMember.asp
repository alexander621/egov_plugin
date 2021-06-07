<!--#include virtual="/header_s.asp"-->
<!--#include file="forum.asp"-->
<!--#include virtual="/checkrole.asp"-->
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
	else
    response.write "<B><H4>This is protected area<br> only for the&nbsp; Authorized People </H4></B>"
	end if
end if

if authorized then
%>
<CENTER><U><FONT SIZE='4' COLOR='blue'>Administer All Member Information</FONT></U>&nbsp;&nbsp;<I><a href=/member/register.asp>New record</a></I>&nbsp;&nbsp;<I><a href='/admin//member/AdminMember.asp?iofaction=6&currentpage=1'>Page 1</a></I></CENTER>
<%
SHowForum
end if
%>
<!--#include virtual="/footer.asp"-->