<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitcustomfieldtypeedit.asp
' AUTHOR: Steve Loar
' CREATED: 01/15/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Creates and edits permit Review types
'
' MODIFICATION HISTORY
' 1.0   01/15/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sTitle, iCustomFieldTypeid, iFieldTypeId, sFieldName, sPdfFieldName, sPrompt, sFieldSize
Dim sValueList, bHasValues, bCanSetSize, iMaximumSize, iMiniumSize, bOrgHasPermitTypeReport
Dim sSuccessFlag, sLoadMsg, sSuccessMsg, sReportTitle

sLevel = "../" ' Override of value from common.asp

'PageDisplayCheck "permit custom field types", sLevel	' In common.asp
PageDisplayCheck "permitv2 types", sLevel	' In common.asp

iCustomFieldTypeid = CLng(request("cft") )

bHasValues = False 
bHasSize = False 
sLoadMsg = ""

If CLng(iCustomFieldTypeid) > CLng(0) Then
	sTitle = "Edit"
	GetCustomFieldTypeValues iCustomFieldTypeid, iFieldTypeId, sFieldName, sPdfFieldName, sPrompt, sValueList, sFieldSize, sReportTitle
Else
	sTitle = "New"
	sFieldSize = "1"
	iMaximumSize = "1"
End If 

If request("fieldtypeid") <> "" Then
	iFieldTypeId = CLng(request("fieldtypeid"))
Else
	If CLng(iCustomFieldTypeid) = CLng(0) Then
		iFieldTypeId = GetFirstFieldTypeId()
	End If 
End If 

' Get if the field type has values or size
GetFieldTypeValues iFieldTypeId, bHasValues, bCanSetSize, iMaximumSize, iMiniumSize

If request("fieldname") <> "" Then
	sFieldName = request("fieldname")
End If 

If request("pdffieldname") <> "" Then
	sPdfFieldName = request("pdffieldname")
End If 

If request("prompt") <> "" Then
	sPrompt = request("prompt")
End If 

If request("fieldsize") <> "" Then 
	sFieldSize = request("fieldsize")
End If 

If request("valuelist") <> "" Then
	sValueList = request("valuelist")
End If 

If request("reporttitle") <> "" Then 
	sReportTitle = request("reporttitle")
End If 

If CLng(sFieldSize) > CLng(iMaximumSize) Then
	sFieldSize = iMaximumSize
End If 

bOrgHasPermitTypeReport = OrgHasFeature( "permit type report" )		' in common.asp

sSuccessFlag = request("success")
If sSuccessFlag <> "" Then
	sLoadMsg = "displayScreenMsg('" & sSuccessFlag & "');"
End If 

%>


<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<script language="JavaScript" src="../scripts/jquery-1.4.2.min.js"></script>

	<script language="JavaScript" src="../scripts/formatnumber.js"></script>
	<script language="JavaScript" src="../scripts/removespaces.js"></script>
	<script language="JavaScript" src="../scripts/removecommas.js"></script>
	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="javascript" src="../scripts/textareamaxlength.js"></script>
	<script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

	<script language="Javascript">
	<!--

		function Another()
		{
			location.href="permitcustomfieldtypeedit.asp?cft=0";
		}

		function Validate()
		{
			if ($("#fieldname").val() == '')
			{
				$("#fieldname").focus();
				inlineMsg($("#fieldname").attr('id'),'<strong>Missing Value: </strong>Please enter a field name.',8,$("#fieldname").attr('id'));
				return;
			}

			if ($("#prompt").val() == '')
			{
				$("#prompt").focus();
				inlineMsg($("#prompt").attr('id'),'<strong>Missing Value: </strong>Please enter a prompt.',8,$("#prompt").attr('id'));
				return;
			}

			if ($("#valuelist").length > 0)
			{
				if ($("#valuelist").val() == '')
				{
					$("#valuelist").focus();
					inlineMsg($("#valuelist").attr('id'),'<strong>Missing Value: </strong>Please enter some choices.',8,$("#valuelist").attr('id'));
					return;
				}
			}

			if ($("#fieldsize").length > 0)
			{
				if ($("#fieldsize").val() == '')
				{
					$("#fieldsize").focus();
					inlineMsg($("#fieldsize").attr('id'),'<strong>Missing Value: </strong>Please enter a field size.',8,$("#fieldsize").attr('id'));
					return;
				}
				else
				{
					var sFileSize = $("#fieldsize").val();
					//if ( parseInt($("#fieldsize").val()) != $("#fieldsize").val() )
					var rege = /^\d*$/
					if ( rege.test(sFileSize) == false )
					{
						$("#fieldsize").focus();
						inlineMsg($("#fieldsize").attr('id'),'<strong>Value Error: </strong>The field size must be an positive integer.',8,$("#fieldsize").attr('id'));
						return;
					}

					if ( parseInt($("#fieldsize").val()) > parseInt($("#maxsize").val()) )
					{
						$("#fieldsize").focus();
						inlineMsg($("#fieldsize").attr('id'),'<strong>Value Error: </strong>The field size cannot be greater that the maximum allowed size.',8,$("#fieldsize").attr('id'));
						return;
					}

					if ( parseInt($("#fieldsize").val()) == parseInt("0") )
					{
						$("#fieldsize").focus();
						inlineMsg($("#fieldsize").attr('id'),'<strong>Value Error: </strong>The field size must be greater than zero.',8,$("#fieldsize").attr('id'));
						return;
					}
				}
			}

			//alert("Ok.");
			// submit the form here
			document.frmFields.submit();
		}

		function Delete() 
		{
			if (confirm("Do you wish to delete this custom field type?"))
			{
				location.href="permitcustomfieldtypedelete.asp?cft=<%=iCustomFieldTypeid%>";
			}
		}

		function refreshPage()
		{
			document.frmFields.action = "permitcustomfieldtypeedit.asp";
			document.frmFields.submit();
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

		$(document).ready(function(){
			setMaxLength();
			<%=sLoadMsg%>
			$("#fieldname").focus(); 
		});

	//-->
	</script>
	<script>
		function commonIFrameUpdateFunction()
		{
			UpdateCustomFieldTypes();
		}
		function UpdateCustomFieldTypes()
		{
			//Get New Values
			var request = new XMLHttpRequest();
			request.open('GET', 'popselectbox.asp?type=customfieldtypes', false);  // `false` makes the request synchronous
			request.send();

			if (request.status === 200) {
  				newDDVals = request.responseText;

				//Get elements from parent
				var pfDD = parent.document.getElementsByClassName('permitcustomfieldtypeDD');
				for (var i = 0; i < pfDD.length; i++) {
					//Get Selected Value
  					//pfDD[i].style.display = 'inline-block';
					var selVal = pfDD[i].options[pfDD[i].selectedIndex].value;
					
					//Update The Values
					pfDD[i].innerHTML = newDDVals;
	
					//Select Previous Option
					pfDD[i].value = selVal;
				}
			}

		}
		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}
	</script>

</head>

<body>

	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent" class="widecontent">
		<div class="gutters">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong><%=sTitle%> Permit Custom Field Type</strong></font>
			</p>
			<p>
				<span id="screenMsg">&nbsp;</span>
				<a href="permitcustomfieldtypelist.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0" />&nbsp;<%=langBackToStart%></a>
			</p>
			<!--END: PAGE TITLE-->

			<!--BEGIN: EDIT FORM-->
			<div id="functionlinks">
<%		If CLng(iCustomFieldTypeid) = CLng(0) Then %>
			<input type="button" id="savebutton" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" value="Create" /><br />
<%		Else %>
			<input type="button" id="savebutton" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" value="Save Changes" /> &nbsp; &nbsp;
			<input type="button" class="showiniframe button ui-button ui-widget ui-corner-all" value="Close" onClick="doClose();" />&nbsp; &nbsp; 
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Delete();" value="Delete" /> &nbsp; &nbsp;
<%			If sSuccessFlag = "This Custom Field Type has been created." Then %>
				<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Another();" value="Create Another" />
<%			End If		%>
			<br />
<%		End If %>
			</div>

		<form name="frmFields" action="permitcustomfieldtypeupdate.asp" method="post">
			<input type="hidden" name="cft" value="<%=iCustomFieldTypeid%>" />
		
			<p>
				<span class="pagelabel">Field Name:</span>
				<input type="text" id="fieldname" name="fieldname" value="<%=sFieldName%>" size="50" maxlength="50" />
			</p>
			<p>
				<span class="pagelabel">PDF Field Name:</span>
				<input type="text" id="pdffieldname" name="pdffieldname" value="<%=sPdfFieldName%>" size="25" maxlength="25" />
				(optional &mdash; no spaces, use lowercase, letters and numbers ONLY)
			</p>
<%			If bOrgHasPermitTypeReport Then			%>
				<p>
					<span class="pagelabel"><%= GetFeatureName( "permit type report" )%> Column Title:</span>
					<input type="text" id="reporttitle" name="reporttitle" value="<%=sReportTitle%>" size="25" maxlength="25" />
					(optional)
				</p>
<%			Else			%>
				<input type="hidden" name="reporttitle" value="" />
<%			End If		%>
			<p>
				<span class="pagelabel">Prompt:</span>
				<input type="text" id="prompt" name="prompt" value="<%=sPrompt%>" size="150" maxlength="150" style="width:100%" />
			</p>

			<p>
				<span class="pagelabel">Field Type:</span>
<%				ShowPermitFieldTypePicks iFieldTypeId		%>				
			</p>

<%			If bHasValues Then								
				' These are radio buttons, check boxes and dropdowns %>
				<p>
					<span class="pagelabel">Available Choices:</span><br />
					(Put each value on a separate line. Each value should only contain letters, numbers, '(', ')', or spaces)<br />
					<textarea id="valuelist" name="valuelist" maxlength="250"><%=sValueList%></textarea>
				</p>
<%			Else   ' is probably a text box or textarea	%>
				<p>
<%					If bCanSetSize Then						%>
						<span class="pagelabel">Field Size:</span>
						<input type="text" id="fieldsize" name="fieldsize" value="<%=sFieldSize%>" size="4" maxlength="4" />
						(Max size is <%=iMaximumSize%> characters)
						<input type="hidden" id="maxsize" name="maxsize" value="<%=iMaximumSize%>" />
<%					End If									%>
				</p>
<%			End If											%>


		</form>
		<!--END: EDIT FORM-->
			<div id="functionlinks">
<%		If CLng(iCustomFieldTypeid) = CLng(0) Then %>
			<input type="button" id="savebutton" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" value="Create" /><br />
<%		Else %>
			<input type="button" id="savebutton" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" value="Save Changes" /> &nbsp; &nbsp;
			<input type="button" class="showiniframe button ui-button ui-widget ui-corner-all" value="Close" onClick="doClose();" />&nbsp; &nbsp; 
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Delete();" value="Delete" /> &nbsp; &nbsp;
<%			If sSuccessFlag = "This Custom Field Type has been created." Then %>
				<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Another();" value="Create Another" />
<%			End If		%>
			<br />
<%		End If %>
			</div>

		</div>
		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

<%	If request("success") <> "" Then 
		SetupMessagePopUp request("success")
	End If	
%>

</body>
</html>


<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void ShowPermitFieldTypePicks iFieldTypeId
'--------------------------------------------------------------------------------------------------
Sub ShowPermitFieldTypePicks( ByVal iFieldTypeId )
	Dim sSql, oRs

	sSql = "SELECT fieldtypeid, fieldtype FROM egov_permitfieldtypes ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select name=""fieldtypeid"" id=""fieldtypeid"" onchange=""refreshPage();"">"

		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("fieldtypeid") & """"
			If CLng(iFieldTypeId) = CLng(oRs("fieldtypeid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("fieldtype") & "</option>"
			oRs.MoveNext
		Loop

		response.write vbcrlf & "</select>"
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' integer GetFirstFieldTypeId( )
'--------------------------------------------------------------------------------------------------
Function GetFirstFieldTypeId()
	Dim sSql, oRs

	sSql = "SELECT fieldtypeid FROM egov_permitfieldtypes ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetFirstFieldTypeId = CLng(oRs("fieldtypeid"))
	Else
		GetFirstFieldTypeId = CLng(0)	' this would be bad
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void GetCustomFieldTypeValues iCustomFieldTypeid, iFieldTypeId, sFieldName, sPdfFieldName, sPrompt, sValueList, sFieldSize, sReportTitle
'--------------------------------------------------------------------------------------------------
Sub GetCustomFieldTypeValues( ByVal iCustomFieldTypeid, ByRef iFieldTypeId, ByRef sFieldName, ByRef sPdfFieldName, ByRef sPrompt, ByRef sValueList, ByRef sFieldSize, ByRef sReportTitle )
	Dim sSql, oRs

	sSql = "SELECT fieldtypeid, fieldname, pdffieldname, prompt, ISNULL(valuelist,'') AS valuelist, "
	sSql = sSql & "ISNULL(fieldsize,1) AS fieldsize, ISNULL(reporttitle,'') AS reporttitle "
	sSql = sSql & "FROM egov_permitcustomfieldtypes WHERE customfieldtypeid = " & iCustomFieldTypeid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		iFieldTypeId = oRs("fieldtypeid")
		sFieldName = oRs("fieldname")
		sPdfFieldName = oRs("pdffieldname")
		sPrompt = oRs("prompt")
		sValueList = oRs("valuelist")
		sFieldSize = oRs("fieldsize")
		sReportTitle = oRs("reporttitle")
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void GetFieldTypeValues iFieldTypeId, bHasValues, bHasSize, iMaximumSize, iMiniumSize
'--------------------------------------------------------------------------------------------------
Sub GetFieldTypeValues( ByVal iFieldTypeId, ByRef bHasValues, ByRef bCanSetSize, ByRef iMaximumSize, ByRef iMiniumSize )
	Dim sSql, oRs

	sSql = "SELECT hasvalues, cansetsize, ISNULL(maximumsize,1) AS maximumsize, ISNULL(miniumsize,1) AS miniumsize "
	sSql = sSql & "FROM egov_permitfieldtypes WHERE fieldtypeid = " & iFieldTypeid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If oRs("hasvalues") Then
			bHasValues = True 
		Else
			bHasValues = False 
		End If 
		If oRs("cansetsize") Then
			bCanSetSize = True 
		Else
			bCanSetSize = False 
		End If 
		iMaximumSize = oRs("maximumsize")
		iMiniumSize = oRs("miniumsize")
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Sub 



%>
