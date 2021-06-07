<!-- #include file="class/classFamily.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: delete_family_member.asp
' AUTHOR: Steve Loar
' CREATED: 1/2/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This Deletes a family member. It is called from family_list.asp
'
' MODIFICATION HISTORY
' 1.0   1/2/2007	Steve Loar - Initial code 
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iUserId, oFamily, iDeletedById

iUserId = CLng(request("iUserId"))

iDeletedById = CLng(request("deletedbyid"))

Set oFamily = New classFamily

oFamily.DeleteFamilyMember iUserId, iDeletedById 

Set oFamily = Nothing 

response.redirect "family_list.asp"

%>