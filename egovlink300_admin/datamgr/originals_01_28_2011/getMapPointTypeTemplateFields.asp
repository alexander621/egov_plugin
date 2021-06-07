<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: displayMapPointTypeTemplateFields.asp
' AUTHOR: David Boyer
' CREATED: 01/18/2011
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Displays the MapPoint Type "default" Template fields
'
' MODIFICATION HISTORY
' 1.0  01/18/2011 	David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 lcl_isRootAdmin   = False
 lcl_isLimited     = False
 lcl_isTemplate    = False
 lcl_isDisplayOnly = False

 if request("mappoint_typeid") <> "" then
    lcl_mappoint_typeid = request("mappoint_typeid")
 end if

 if request("isRootAdmin") = "Y" then
    lcl_isRootAdmin = True
 end if

 if request("isLimited") = "Y" then
    lcl_isLimited = True
 end if

 if request("isTemplate") = "Y" then
    lcl_isTemplate = True
 end if

 if request("isDisplayOnly") = "Y" then
    lcl_isDisplayOnly = True
 end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="mappoints_global_functions.asp" //-->
<html>
<head>
  <link rel="stylesheet" type="text/css" href="../global.css" />
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="<%=lcl_onload%>">
<%
'Retrieve any/all fields related to this Map-Point Type Template
 displayMPTypesFields "", lcl_mappoint_typeid, lcl_isRootAdmin, lcl_isLimited, lcl_isDisplayOnly

 response.write "</body>" & vbcrlf
 response.write "</html>" & vbcrlf
%>