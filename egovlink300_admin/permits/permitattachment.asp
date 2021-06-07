<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitattachment.asp
' AUTHOR: Steve Loar
' CREATED: 07/08/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Select attachments to be uploaded and associated with a permit
'
' MODIFICATION HISTORY
' 1.0   07/08/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId

iPermitId = CLng(request("permitid"))

%>

<html>
	<head>
		<link rel="stylesheet" type="text/css" href="../global.css" />
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
		<link rel="stylesheet" type="text/css" href="permits.css" />

		<script language="JavaScript" src="../scripts/textareamaxlength.js"></script>
		<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

		<script language="Javascript">
		<!--

		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

		function doAddCheck()
		{
			if ($("permitattachment").value == '')
			{
				alert('Please select a file to upload.');
				$("permitattachment").focus();
				return;
			}
			if ($("permitattachment").value.indexOf('+') > 0)
			{
				alert('Sorry, you cannot upload a file with a plus sign in the name.');
				$("permitattachment").focus();
				return;
			}
			document.frmAddAttachment.submit();
		}

		function doLoad()
		{
			setMaxLength();
			<% 
			if request("success") <> "" then 
				if cLng(request("success")) > cLng(0) then 
					response.write "alert('Attachment Successfully Uploaded.');"
					response.write "parent.RefreshPageAfterVoid();"
				else
					response.write "alert('Your file was too large.\nPlease limit the file to 20MB.');"
					response.write "$(""permitattachment"").focus();"
				end if 
			end if 
			%>
		}

		//-->
		</script>

	</head>
	<body onload="doLoad();">
		<div id="content">
			<div id="centercontent">
				<form  name="frmAddAttachment" action="permitattachmentupload.asp" method="POST" enctype="multipart/form-data">
					<input type="hidden" name="permitid" value="<%=iPermitId%>" />
					<p>
						Attachment:<br /><input style="width:650px;" id="permitattachment" name="permitattachment" type="file" /><br />
						(20MB and smaller files)<br /><br />
						Description:<br /><textarea style="width:575px;height:50px;" name="attachmentdesc" maxlength="1024"></textarea>
					</p>
					<p>
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Add Attachment" onclick="doAddCheck();" /> &nbsp; &nbsp; 
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Close" onclick="doClose();" /> 
					</p>
				</form>
			</div>
		</div>
	</body>
</html>
