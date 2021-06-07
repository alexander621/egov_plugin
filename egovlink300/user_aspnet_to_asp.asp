<!DOCTYPE html>
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<% 
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: user_aspnet_to_asp.asp
' AUTHOR: ???
' CREATED: 2012
' COPYRIGHT: Copyright 2012 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Sets up the ASP session variables from ASPNET
'
' MODIFICATION HISTORY
' 1.0 11/07/2012 Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
%>
<html>
<head>
  <title>E-Gov Services</title>
</head>

<!--#include file="include_top.asp"-->

<%
  dim sUserSessionID = ""

  session("RedirectPage")    = ""
  session("RedirectLang")    = ""
  session("ManageURL")       = ""
  session("LoginDisplayMsg") = ""
  session("DisplayMsg")      = ""

  if request("usid") <> "" then
     if not containsApostrophe(request("usid")) then
        sUserSessionID = request("usid")

        sSQL = "SELECT usersessionid, "
        sSQL = sSQL & " sessionid, "
        sSQL = sSQL & " orgid, "
        sSQL = sSQL & " userid, "
        sSQL = sSQL & " RedirectPage, "
        sSQL = sSQL & " RedirectLang, "
        sSQL = sSQL & " ManageURL, "
        sSQL = sSQL & " LoginDisplayMsg, "
        sSQL = sSQL & " DisplayMsg "
        sSQL = sSQL & " FROM egov_aspnet_to_asp_usersessions "
        sSQL = sSQL & " WHERE usersessionid = '" & sUserSessionID & "' "

       	set oGetUserSessions = Server.CreateObject("ADODB.Recordset")
       	oGetUserSessions.Open sSQL, Application("DSN"), 3, 1

        if not oGetUserSessions.eof then
           session("RedirectPage")    = oGetUserSessions("RedirectPage")
           session("RedirectLang")    = oGetUserSessions("RedirectLang")
           session("ManageURL")       = oGetUserSessions("ManageURL")
           session("LoginDisplayMsg") = oGetUserSessions("LoginDisplayMsg")
           session("DisplayMsg")      = oGetUserSessions("DisplayMsg")
        end if

        oGetUserSessions.close
        set oGetUserSessions = nothing

     end if
  end if

  response.redirect "dtb_user_login.asp?testing=blahblahblah"


%>

<!--#include file="include_bottom.asp"-->
