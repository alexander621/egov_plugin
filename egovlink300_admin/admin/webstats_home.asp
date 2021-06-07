<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>E-Govlink Usage Statistics</title>
	<meta NAME="Generator" CONTENT="EditPlus" />
	<meta NAME="Author" CONTENT="" />
	<meta NAME="Keywords" CONTENT="" />
	<meta NAME="Description" CONTENT="" />

	<link rel="stylesheet" type="text/css" href="../global.css" />

</head>


<body>

	<!-- BEGIN: DISPLAY LIST OF ORGANIZATIONS-->
	<div style="padding:20 px;">
		<h3 class="webstattitle">E-Govlink Web Site Section Detail Statistics<hr align=left style="color:#000000;size:1 px; width:75%;" ></h3>
		<p>
		<%
			subListOrganizations
		%>
		</p>
	</div>
	<!-- END: DISPLAY LIST OF ORGANIZATIONS-->

	<!-- BEGIN: DISPLAY FOOTER-->
	<center>
		<hr  style="color:#000000;size:1 px; width:90%;" >

		<div class="copyright_text">
			Copyright &copy;2004-2007. <i>electronic commerce</i> link, inc. dba <a href="http://www.egovlink.com" target="_NEW">egovlink</a>.
		</div>
	</center>
	<!-- END: DISPLAY FOOTER-->

</body>
</html>


<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' SUB SUBLISTORGANIZATIONS()
'------------------------------------------------------------------------------------------------------------
Sub subListOrganizations()

	' GET HOME HITS
	sSQL = "SELECT * FROM Organizations ORDER BY OrgCity, OrgState"
	
	Set oOrg = Server.CreateObject("ADODB.Recordset")
	oOrg.Open sSQL, Application("DSN") , 3, 1

	If NOT oOrg.EOF Then
		response.write vbcrlf & "<div class=""webstat"">"
		response.write vbcrlf & "<ul>"
		Do While NOT oOrg.EOF
			response.write vbcrlf & "<li><a href=""webstats.asp?orgid=" & oOrg("OrgID") & """>" & oOrg("OrgCity") & ", " & oOrg("OrgState") & "</a></li>"
			oOrg.MoveNext
		Loop
		response.write vbcrlf & "</ul>"
		response.write vbcrlf & "</div>"
	Else
		response.write " No organizations found."
	End If 

End Sub



%>