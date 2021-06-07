<!-- #include file="../custom/includes/custom.asp" //-->

<%
' Database call to retrieve Database version
Dim sSQL, oRst, strDBVersion
sSql = "Select AboutDbVersion From About Where AboutID=1"
Set oRst = Server.CreateObject("ADODB.Recordset")
oRst.Open sSql, Application("DSN"), 3, 1
If Not oRst.EOF Then
	strDBVersion = oRst("AboutDbVersion")
Else
	strDBVersion = "<font color=""red"">unknown</font>"
End If 
%>


<html>
<head>
<meta name="GENERATOR" Content="Microsoft Visual Studio.NET 7.0">
<link href="../global.css" rel="stylesheet" type="text/css">
<title>About <%=Application("ProgramName")%></title>
</head>


<!--Begin Content Area-->
<body  TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0>
<form>
<table cellspacing="0" cellpadding="0" width="100%">
<tr bgcolor="#336699">
  <td colspan=3 ><img src="<% response.write "..\" & custGraphic & "home.jpg" %>"></td>
</tr>
<tr><td>&nbsp;</td></tr>
<tr><td class=aboutmelabel>&nbsp;&nbsp;&nbsp;Program Version: </td><td class=aboutmetext><%=Application("ProgramVersion")%></td>
<tr><td class=aboutmelabel>&nbsp;&nbsp;&nbsp;Database Version:</td><td class=aboutmetext><%=strDBVersion%></td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td colspan=2 class=aboutmetext >&nbsp;Copyright &#xA9; 2002 <a href="http://www.eclink.com" target="_blank"><i>electronic commerce link</i>, Inc.</a> </td>
<td valign=bottom align=center><input class=aboutme type=button value="Ok" onclick="window.close();">&nbsp;</td>
</tr>
<tr><td>&nbsp;</td></tr>
</table>
</form>
</body>
<!--End Content Area-->


<script language="Javascript">
<!--// Bring window to the front 
window.focus()
//-->
</script>

</html>
