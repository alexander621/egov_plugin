<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalsguifunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalcategoryedit.asp
' AUTHOR: Steve Loar
' CREATED: 08/13/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Create and Edit rental categories.
'
' MODIFICATION HISTORY
' 1.0   08/13/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iRecreationCategoryId, sTitle, sButtonValue, sCategoryTitle, sLoadMsg, sHasRestrictedPeriod
Dim iRestrictedPeriodId, sImgUrl, sCategoryDescription, sHideFromPublic

iRecreationCategoryId = CLng(request("rc"))

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "edit rentals categories", sLevel	' In common.asp

sCategoryTitle = ""
sHideFromPublic = "" 

If iRecreationCategoryId = CLng(0) Then
	sTitle = "Create A Rental Category"
	sButtonValue = "Create Rental Category"
Else
	sTitle = "Edit Rental Category"
	sButtonValue = "Save Changes"
End If 

blnHasWP = hasWordPress()
sHomeWebsiteURL = getOrganization_WP_URL(session("orgid"), "OrgPublicWebsiteURL")

GetCategoryValues iRecreationCategoryId, sCategoryTitle

If request("s") <> "" Then
	If request("s") = "n" Then
		sLoadMsg = "displayScreenMsg('This Category Was Successfully Created.');"
	End If
	If request("s") = "u" Then
		sLoadMsg = "displayScreenMsg('Your Changes Were Successfully Saved.');"
	End If 
End If 

%>

<html lang="en">
<head>
	<meta charset="UTF-8">
	
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalsstyles.css" />


  	<script src="//code.jquery.com/jquery-1.12.4.js"></script>
   	<script src="//code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
	<!--#include file="../includes/wp-image-picker.asp"-->

	<script src="../scripts/modules.js"></script>
	<script src="../scripts/textareamaxlength.js"></script>
	<script src="../scripts/formvalidation_msgdisplay.js"></script>


	<script>
	<!--

		function SetUpPage()
		{
			setMaxLength();
			<%=sLoadMsg%>
			$("#categorytitle").focus();
		}

		function displayScreenMsg(iMsg) 
		{
			if(iMsg!="") 
			{
				$("#screenMsg").html( "*** " + iMsg + " ***" );
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			$("#screenMsg").html( "" );
		}

		function DeleteCategory()
		{
			if (confirm('Delete this category?'))
			{
				location.href='rentalcategorydelete.asp?rc=<%=iRecreationCategoryId%>';
			}
		}

		function doImagePicker(sFormField) 
		{
			var w = (screen.width - 350)/2;
			var h = (screen.height - 350)/2;
			eval('window.open("imagepicker/default.asp?name=frmRentalCategory.' + sFormField + '", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function storeCaret (textEl) 
		{
		   if (textEl.createTextRange)
			 textEl.caretPos = document.selection.createRange().duplicate();
		 }

		 function insertAtURL (textEl, text) 
		 {
		   if (textEl.createTextRange && textEl.caretPos) 
		   {
			 var caretPos = textEl.caretPos;
			 caretPos.text =
			   caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
				 text + ' ' : text;
		   }
		   else
			 textEl.value  = text;
			$("#" + textEl.name + "pic").attr("src",text);
			if (textEl.name.indexOf("document") >= 0)
			{
				$("#" + textEl.name + 'pic').html('<a href="' + text + '" target="_newwindow">View Document</a>&nbsp;&nbsp;');
			}
		 }

		function validate()
		{
			var okToSubmit = true;
			if ($("#categorytitle").val() == '')
			{
				inlineMsg("categorytitle","<strong>Missing Value: </strong>Please provide a category name.");
				okToSubmit = false;
			}

			if ($("#categorydescription").val() == '')
			{
				inlineMsg("categorydescription","<strong>Missing Value: </strong>Please provide a category description.");
				okToSubmit = false;
			}
			
			if ( okToSubmit ) {
				document.frmRentalCategory.submit();
			} else {
				return false;
			}
		}

	//-->
	</script>

</head>
<body onload="SetUpPage();">

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong><%=sTitle%></strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<div class="btn-group">
				<span id="screenMsg"></span><br />
				<input type="button" class="button" value="<< Back" onclick="location.href='rentalscategorieslist.asp';" /> &nbsp; 
				<% If iRecreationCategoryId > CLng(0) Then	%>
					<input type="button" class="button" value="Delete" onclick="DeleteCategory();" /> &nbsp; 
				<% End If	%>
				<input type="button" class="button" value="<%=sButtonValue%>" onclick="validate();" />
			</div>

			<form name="frmRentalCategory" action="rentalcategoryupdate.asp" method="post">
				<input type="hidden" id="recreationcategoryid" name="recreationcategoryid" value="<%=iRecreationCategoryId%>" />
				<p id="rentalnamecontainer">
					Category Name: <input type="text" id="categorytitle" name="categorytitle" size="90" maxlength="90" value="<%=sCategoryTitle%>" />
				</p>
				
					Category Image: <br /><input type="<% if blnHasWP then %>hidden<%else%>text<%end if%>" id="imgurl" class="imageurl" name="imgurl" size="100" maxlength="250" value="<%=sImgUrl%>" /> 
					<img src="<%=sImgUrl%>" id="imgurlpic" align="middle" width="240" height="180"  onerror="this.src = '../images/placeholder.png';" />
					<% if blnHasWP then %>
						<input type="button" class="button" value="Change" onclick="showModal('Pick Image',65,80,'imgurl');" /-->
					<% else %>
						<input type="button" class="button" value="Pick" onclick="doImagePicker('imgurl');" />
					<% end if %>
					<div id="categoryimgtag" class="helpmsg">
						<strong>* Images should be 100px width by 100px height and should be less than 20KB.</strong>
					</div>
				
				<p><br />
					Category Description (Use simple HTML to format):<br />
					<textarea id="categorydescription" name="categorydescription" maxlength="2000" wrap="soft"><%=sCategoryDescription%></textarea>
				</p>
				<p>
					<input type="checkbox" id="hasrestrictedperiod" name="hasrestrictedperiod" <%=sHasRestrictedPeriod%> /> 
					Restrict the public to 1 reservation per 
<%					ShowRestrictedPeriods iRestrictedPeriodId	%>
				</p>
				<p>
					<input type="checkbox" id="hidefrompublic" name="hidefrompublic" <% = sHideFromPublic %> /> Hide This Category From Public View
				</p>
				
				<div class="btn-group">
					<input type="button" class="button" value="<%=sButtonValue%>" onclick="validate();" />
				</div>
			</form>
		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


<%
'--------------------------------------------------------------------------------------------------
' GetCategoryValues iRecreationCategoryId, sCategoryTitle 
'--------------------------------------------------------------------------------------------------
Sub GetCategoryValues( ByVal iRecreationCategoryId, ByRef sCategoryTitle )
	Dim sSql, oRs

	sSql = "SELECT categorytitle, hasrestrictedperiod, ISNULL(restrictedperiodid,0) AS restrictedperiodid, "
	sSql = sSql & "ISNULL(imgurl,'') AS imgurl, ISNULL(categorydescription,'') AS categorydescription, hidefrompublic "
	sSql = sSql & "FROM egov_recreation_categories WHERE recreationcategoryid = " & iRecreationCategoryId
	sSql = sSql & " AND orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		sCategoryTitle = oRs("categorytitle")
		If oRs("hasrestrictedperiod") Then 
			sHasRestrictedPeriod = " checked=""checked"" "
		Else
			sHasRestrictedPeriod = ""
		End If 
		iRestrictedPeriodId = oRs("restrictedperiodid")
		sImgUrl = oRs("imgurl")
		sCategoryDescription = orS("categorydescription")
		If oRs("hidefrompublic") Then
			sHideFromPublic = " checked=""checked"" "
		Else
			sHideFromPublic = "" 
		End If 
	Else
		sCategoryTitle = ""
		sHasRestrictedPublicPeriod = ""
		sRestrictedQuantity = ""
		sRestrictedPeriodQuantity = ""
		iRestrictedPeriodId = 0
		sImgUrl = ""
		sCategoryDescription = ""
		sHideFromPublic = "" 
	End If 

	oRs.Close
	Set oRs = Nothing
	
End Sub



%>
