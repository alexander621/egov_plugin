<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="events_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: outofrangedate.asp
' AUTHOR: Steve Loar
' CREATED: 09/09/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Displays a message when the selected date is out of range (+/- 5yrs from today).
'
' MODIFICATION HISTORY
' 1.0   09/09/2009	Steve Loar - INITIAL VERSION
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
%>

<html>
<head>
	<title>E-Gov Services - <%=sOrgName%></title>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />

	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />


	<script language="javascript">
	<!--

	//-->
	</script>

	<style type="text/css">

		p#outofrangemsg {
			font-size: 12pt;
			font-weight: bold;
			text-align: center;
			}
		
		p#outofrangemsg a {
			font-size: 12pt;
			font-weight: bold;
			}

	</style>

</head>

<!--#Include file="../include_top.asp"-->

<!--BODY CONTENT-->
	<p>&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</p>

	<p id="outofrangemsg">
		The date you have selected is out of the acceptable range of dates for our calendar.<br />
		<a href="calendar.asp">Click here</a> to return to the calendar showing the current month.
	</p>


<p>&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</p>

<!--#Include file="../include_bottom.asp"-->

<%
%>