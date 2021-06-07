<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: fontpreview.asp
' AUTHOR: Steve Loar
' CREATED: 11/19/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Displays a preview of selected font styling
'
' MODIFICATION HISTORY
' 1.0   11/19/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sFontStyle, sDisplayText

sFontStyle = request("fontstyle") & " " & request("fontweight") & " " & request("fontsize") & " '" & request("fontfamily") & "'"

If request("displaytext") <> "" Then
	sDisplayText = request("displaytext")
Else
	sDisplayText = "Font Preview Text."
End If 

%>

<html>
	<head>
		<link rel="stylesheet" type="text/css" href="../global.css" />
		<link rel="stylesheet" type="text/css" href="permits.css" />

		<script language="Javascript">
		<!--

		function doClose()
		{
			window.close();
			window.opener.focus();
		}

		//-->
		</script>

	</head>
	<body>
		<div id="content">
			<div id="centercontent">
				<br />
				<p style="margin-left: 30px;">
					<input type="button" class="button ui-button ui-widget ui-corner-all" value="Close" onclick="doClose();" />
				</p>
				<p>
					<span style="font: <%=sFontStyle%>;"> 
						<%=sDisplayText%>
					</span>
				</p>
				
			</div>
		</div>
	</body>
</html>

