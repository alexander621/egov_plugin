<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<%

Dim sImageUrl

sImageUrl = request("url")

%>
<html>
	<head>
		<title>E-Gov Administration Console</title>

		<link rel="stylesheet" type="text/css" href="../global.css" />
		<link rel="stylesheet" type="text/css" href="rentalsstyles.css" />

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

		<div id="imagedisplay">
			<img src="<%=sImageUrl%>" border="0" />
		
			<p id="imagedisplayclose">
				<input type="button" onclick="doClose();" value="Close" class="button" />
			</p>

		</div>

	</body>
</html>
