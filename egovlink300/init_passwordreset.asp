<!DOCTYPE html>
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: forgot_password.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module sends a registered citizen their password.
'
' MODIFICATION HISTORY
' 1.0   ??
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 dim sError
%>
<html>
<head>
<title>E-Gov Services <%=sOrgName%></title>

<link type="text/css" rel="stylesheet" href="css/styles.css" />
<link type="text/css" rel="stylesheet" href="global.css" />
<link type="text/css" rel="stylesheet" href="css/style_<%=iorgid%>.css" />

<style type="text/css">
.fieldset,
.fieldset_doesnotexist
{
   margin: 10px;
   padding: 10px;
   border-radius: 6px;
}

.fieldset legend
{
   padding: 4px 8px;
   border: 1pt solid #808080;
   border-radius: 6px;
   color: #800000;
}

.fieldset_doesnotexist
{
   color: #800000;
   font-size: 1.25em;
}

#email
{
   width: 300px;
}

#passwordText
{
   margin: 5px 0px 10px 0px;
}

#buttonLookup,
#buttonLogin
{
   cursor: pointer;
}
</style>

</head>

<!--#Include file="include_top.asp"-->
<!--#Include file="inc_password_reset.asp"-->
<!--#Include file="include_bottom.asp"-->    
