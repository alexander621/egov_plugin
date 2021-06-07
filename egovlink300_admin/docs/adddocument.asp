<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: adddocument.asp
' AUTHOR: Steve Loar
' CREATED: 08/31/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Documents Prototype page
'
' MODIFICATION HISTORY
' 1.0   08/31/2010	Steve Loar - INITIAL VERSION
' 1.1	01/10/2011	Steve Loar - Modified to check the link and direct text file name and content.
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSuccessFlag, sParentFolder, sFileName, sLoadMsg, lcl_show_continueupload_msg

sLevel = "../" ' Override of value from common.asp
lcl_show_continueupload_msg = False 

' check if page is online and user has permissions in one call not two
PageDisplayCheck "manage documents", sLevel	' In common.asp

sParentFolder = request("path")

sSuccessFlag = request("sf")
If sSuccessFlag = "df" Then
	sLoadMsg = "displayScreenMsg('The file you are attempting to upload already exists in this folder.');"
End If 
If sSuccessFlag = "tb" Then
	sLoadMsg = "displayScreenMsg('The file you are attempting to upload is too big. Files must be under 50 MB.');"
End If 
If sSuccessFlag = "su" Then
	sLoadMsg = "displayScreenMsg('The document has been successfully added to the folder.');"
	lcl_show_continueupload_msg = True 
	sFileName = request("filename")
End If 
If sSuccessFlag = "fa" Then
	sLoadMsg = "displayScreenMsg('The document has been successfully created and added to the folder.');"
End If 


%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="docstyles.css" />

	<script type="text/javascript" src="../scripts/jquery-1.4.2.min.js"></script>
	<script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

	<script language="Javascript">
	<!--

		function loader()
		{
			<%=sLoadMsg%>
		}

		function displayScreenMsg(iMsg) 
		{
			if(iMsg!="") 
			{
				$("#screenMsg").html("*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;");
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			$("#screenMsg").html("&nbsp;");
		}

		function showMethod( layerName ) 
		{
			if (layerName == "direct") 
			{
				$("div#link").hide("slow");
				$("div#upload").hide("slow");
				$("div#direct").show("slow");
				document.frmAddArticle.encoding = "application/x-www-form-urlencoded";
				document.frmAddArticle.action   = "adddocumentdo.asp?method=direct";
			}
			else if (layerName == "upload") 
			{
				$("div#link").hide("slow");
				$("div#direct").hide("slow");
				$("div#upload").show("slow");
				document.frmAddArticle.encoding = "multipart/form-data";
				document.frmAddArticle.action   = "upload.asp?task=UPLOAD";
			}
			else 
			{
				// This is the Link pick
				$("div#direct").hide("slow");
				$("div#upload").hide("slow");
				$("div#link").show("slow");
				document.frmAddArticle.encoding = "application/x-www-form-urlencoded";
				document.frmAddArticle.action   = "adddocumentdo.asp?method=link";
			} 
		}

		function ValidateFileName()
		{
			var rege;
			var Ok;
			if (document.frmAddArticle.txtMethod.value == "upload")
			{
				if (document.frmAddArticle.binFile.value == "")
				{

					// If they did not enter a filename, take them back.
					alert("Please select a file to upload");
					document.frmAddArticle.binFile.focus();
				}
				else
				{
					var firstpos;
					var filename = document.frmAddArticle.binFile.value;
					var tempname = document.frmAddArticle.binFile.value;

					if (tempname.lastIndexOf('\\') != -1)
					{
						firstpos = tempname.lastIndexOf('\\')+1;
						filename = tempname.substring(firstpos);
					}

					rege = /^[\w- :\\]+\.{1}[A-Za-z0-9]{2}[A-Za-z0-9]{0,2}$/;
					Ok = rege.test(filename);

					if (! Ok)
					{
						alert ("The filename has characters that are not allowed. Allowed characters include [A-Za-z0-9_-], spaces and one '.' \n\n Example: C:\\Documents and Settings\\My Doc_1-2006.txt \n\n Please rename the file on your PC prior to uploading it.  Remove any special characters from the file name.");
						document.frmAddArticle.binFile.focus();
					}
					else
					{
						if (document.frmAddArticle.txtTopic.value == "")
						{
							document.getElementById("screenMsg").innerHTML = "<strong style=\"color: #ff0000\">*** Sorry, your Destination is invalid.  Please click the \"Back to Documents\" link and try again. ***</strong><br />";
							document.getElementById("screenMsg").style.styleFloat = 'none';
							document.getElementById("screenMsg").style.cssFloat = 'none';
						}
						else
						{
							document.getElementById("cancelbutton").disabled = true;
							document.getElementById("createbutton").disabled = true;
							document.getElementById("screenMsg").innerHTML = "<strong style=\"color: #ff0000\">*** Uploading File... ***</strong>";
							//alert(document.frmAddArticle.action);
							document.frmAddArticle.submit();
						}
					}
				}
			}
			else
			{
				if (document.frmAddArticle.txtMethod.value == "link")
				{
					if ($("#txtURL").val() == "")
					{
						// If they did not enter a URL, so take them back.
						alert("Please enter a URL");
						$("#txtURL").focus();
						return;
					}

					//alert( "#txtURLTitle " + $("#txtURLTitle").val() );
					var sLinkName = $("#txtURLTitle").val();
					if (sLinkName == "")
					{
						// If they did not enter a title, so take them back.
						alert("Please enter a title");
						$("#txtURLTitle").focus();
						return;
					}
					rege = /^[\w- ]+$/;
					Ok = rege.test(sLinkName);

					if (! Ok)
					{
						alert ("The title has characters that are not allowed. Allowed characters include [A-Za-z0-9_-] and spaces.\n\n Example: My Link\n\n Please remove any special characters from the title.");
						$("#txtURLTitle").focus();
						return;
					}
					else
					{
						document.getElementById("cancelbutton").disabled = true;
						document.getElementById("createbutton").disabled = true;
						document.getElementById("screenMsg").innerHTML = "<strong style=\"color: #ff0000\">*** Processing... ***</strong>";
						//alert(document.frmAddArticle.action);
						document.frmAddArticle.submit();
					}
				}
				else
				{
					// This is direct entry of text
					if ($("#txtContent").val() == "")
					{
						// If they did not enter any content, so take them back.
						alert("Please enter some content.");
						$("#txtContent").focus();
						return;
					}
					
					//alert( "#txtTitle " + $("#txtTitle").val() );
					var sTextName = $("#txtTitle").val();
					if (sTextName == "")
					{
						// If they did not enter a title, so take them back.
						alert("Please enter a title");
						$("#txtTitle").focus();
						return;
					}
					rege = /^[\w- ]+$/;
					Ok = rege.test(sTextName);

					if (! Ok)
					{
						alert ("The title has characters that are not allowed. Allowed characters include [A-Za-z0-9_-] and spaces.\n\n Example: My Content File\n\n Please remove any special characters from the title.");
						$("#txtTitle").focus();
						return;
					}
					else
					{
						document.getElementById("cancelbutton").disabled = true;
						document.getElementById("createbutton").disabled = true;
						document.getElementById("screenMsg").innerHTML = "<strong style=\"color: #ff0000\">*** Processing... ***</strong>";
						//alert(document.frmAddArticle.action);
						document.frmAddArticle.submit();
					}
				}
			}
		}


	//-->
	</script>

</head>

<body onload="loader();">

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Documents: Add Document</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<table id="screenMsgtable"><tr><td>
				<span id="screenMsg">&nbsp;</span>
				<img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='default.asp'>Back To Documents</a>
			</td></tr></table>

			<table border="0" cellspacing="0" class="start" width="100%">
			  <tr>
				<td width="100%">
				  <form name="frmAddArticle" action="upload.asp?task=UPLOAD" method="post" enctype="multipart/form-data">
					<input type="hidden" name="txtTopic" id="txtTopic" value="<%=sParentFolder%>" />
					<table width="100%" cellpadding="5" cellspacing="0" border="0" class="tableadmin">
					  <tr>
						  <th align="left" colspan="2">New Document</th>
					  </tr>
<%
						If lcl_show_continueupload_msg Then 
							response.write "<tr>" 
							response.write "<td colspan=""2"" style=""color:#ff0000"">NOTE: You can continue uploading files to the same ""Destination"" or expand the folder list and select a new folder to add a document to.</td>" 
							response.write "</tr>" 
						End If 
%>
					  <tr>
						  <td width="1%" valign="top">Destination:</td>
						  <td>
<%
							If sParentFolder = "/" Then 
								response.write "<select name=""selTopic"" style=""width:100%"" onchange=""$('#txtTopic').val(this.value);"">" 
								DrawListBoxContents Application("ECapture_ArticlesPath")
								response.write "</select>" 
							Else 
								response.write "<script language=""javascript"">$('#txtTopic').val('" & sParentFolder & "');</script>" 
								response.write Replace(sParentFolder, Application("eCapture_ArticlesPath") & "/", "") 
							End If 
%>
						  </td>
					  </tr>
					  <tr>
						<td valign="top">Method:</td>
						<td>
						  <select name="txtMethod" onchange="showMethod(this[this.selectedIndex].value);">
							<option value="upload">Upload</option>
							<option value="direct">Directly Add Content</option>
							<option value="link">Create Link Only</option>
						  </select>
						</td>
					  </tr>
					  <tr>
						<td valign="top" nowrap><%=langDefineDoc%>:&nbsp;</td>
						<td>

						  <div id="upload">
										   <!-- File upload field is here -->
							<input type="file" name="binFile" size="90" />
										   <!-- For Testing -->
										   <!--<input type="text" name="binFile" size="50"> -->
						  </div>

						  <div id="direct" style="display:none;">
							<table>
							  <tr>
								<td><font color="#666666"><%=langDocTitle%></font></td>
								<td>
									<input type="text" id="txtTitle" name="txtTitle" size="50" maxlength="30" value="" />
								</td>
							  </tr>
							  <tr>
								<td>&nbsp;</td>
								<td>
									<input type="checkbox" name="blnIsHTML" /> <%=langDocHTML%>
								</td>
							  <tr>
								<td colspan="2">
								  <textarea id="txtContent" name="txtContent" rows="15" cols="75"></textarea>
								</td>
							  </tr>
							</table>
						  </div>

						  <div id="link" style="display:none;">
							<table>
							  <tr>
								<td><font color="#666666"><%=langDocTitle%></font></td>
								<td><input type="text" id="txtURLTitle" name="txtURLTitle" size="50" maxlength="30" value="" /></td>
							  </tr>
							  <tr>
								<td><font color="#666666"><%=langDocURL%></font>&nbsp;&nbsp;</td>
								<td><input type="text" id="txtURL" name="txtURL" size="50" maxlength="255" value="" /></td>
							  </tr>
							  <tr>
								<td></td>
								<td><input type="checkbox" name="openNew" />&nbsp;<%=langDocNewWindow%></td>
							</table>
						  </div>
						</td>
					  </tr>
					  <tr>
						<td nowrap valign="top"></td>
						<td><input type="checkbox" class="listCheck" name="chkOverwrite" />&nbsp;
							Overwrite existing file
						</td>
					  </tr>
					</table>

<%					displayButtons  

					If lcl_show_continueupload_msg Then	%>
						<p>
							<table style="width:480px;" cellpadding="5" cellspacing="0" border="0" class="tableadmin" id="documentlinks">
							  <tr>
								  <th align="left" colspan="2">Uploaded Document</th>
							  </tr>
							   <tr>
								  <td style="width:80px;" valign="top" nowrap="nowrap">File Name:</td>
								  <td>
									<%=sFileName%>
								  </td>
							  </tr>
<%
								strLinkTitle = sFileName		'"/" & sFileName
								sLinkURL     = Replace(session("egovclientwebsiteurl"), "/" & session("virtualdirectory"),"") 
								sLinkURL     = sLinkURL & Replace(sParentFolder, "custom/pub/", "")
								'sLinkURL     = Replace(sLinkURL, "egovlink300_docs", "public_documents300")
								sLinkURL     = Replace(sLinkURL, Application("DocumentsRootDirectory"), "public_documents300")
								sLinkURL     = sLinkURL & strLinkTitle
								sTxtLinkURL  = Replace(sLinkURL, " ", "%20")
%>
							  <tr>
								  <td style="width:80px;" valign="top" nowrap="nowrap">Site Link:</td>
								  <td>
									<input type="text" name="SiteLink" style="width:400px; height:20px;" value="<a target='_EGOVLINK' href='<% = sLinkURL %>'><%=sFileName%></a>" />
								  </td>
							  </tr>
							  <tr>
								  <td style="width:80px;" valign="top" nowrap="nowrap">Savvy Link:</td>
								  <td>
									<input type="text" name="SiteURL" style="width:400px; height:20px;" value="<% = sLinkURL %>" />
								  </td>
							  </tr>
							  <tr>
								  <td style="width:80px;" valign="top" nowrap="nowrap">Text Link:</td>
								  <td>
									<input type="text" name="TextLink" style="width:400px; height:20px;" value="<% = sTxtLinkURL %>" />
								  </td>
							  </tr>
							</table>
						</p>
<%					End If		%>
				  </form>
				</td>
			  </tr>
			</table>

		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

</body>

</html>


<%
'--------------------------------------------------------------------------------------------------
' void DrawListBoxContents dirPath
'--------------------------------------------------------------------------------------------------
Sub DrawListBoxContents( ByVal dirPath )
	Dim newpath, name, sSql, oRs, i, iLevel, padding, optionpath

	sSql = "EXEC ListFolder " & Session("OrgID") & ", " & Session("UserID") & ", '" & dirPath & "'"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		Do While Not oRs.EOF
			optionpath = oRs("FolderPath")
			name = oRs("FolderName")
			padding = ""
			iLevel = oRs("FolderLevel")*2

			If name = "root" Then
				response.write vbcrlf & "<option value=""" & dirPath & "/" & name & "/"" selected>Main Category</option>" 
			Else
				For i = 0 To iLevel
					padding = padding & "&nbsp;"
				Next
				response.write vbcrlf & "<option value=""" & optionpath & "/"">" & padding & name & "</option>"
			End If
			oRs.MoveNext
		Loop
	End If

	oRs.Close
	Set oRs = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' void displayButtons
'--------------------------------------------------------------------------------------------------
Sub displayButtons() 

	response.write "<div style=""font-size:10px; padding-top:5px;"">" 
	response.write "<input type=""button"" name=""create"" id=""createbutton"" value=""Create"" class=""button"" onclick=""ValidateFileName();"" /> &nbsp; &nbsp; " 
	response.write "<input type=""button"" name=""cancel"" id=""cancelbutton"" value=""Cancel"" class=""button"" onclick=""location.href='default.asp'"" />" 
	response.write "</div>" 

End Sub



%>
