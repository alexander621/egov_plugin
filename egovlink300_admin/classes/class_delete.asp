
<%
'--------------------------------------------------------------------------------------------------
' 
' AUTHOR: Steve Loar
' CREATED: 05/03/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'--------------------------------------------------------------------------------------------------
	Class_Delete( CLng(request("classid")) )
'	response.write "Success"
'	response.end
	
	' REDIRECT TO instructor management page
	response.redirect "class_list.asp"


%>

<!-- #include file="class_global_functions.asp" //-->

