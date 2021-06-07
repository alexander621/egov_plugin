<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: regattaeventedit.asp
' AUTHOR: Steve Loar
' CREATED: 02/23/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page allows the creating and editing of regatta events
'
' MODIFICATION HISTORY
' 1.0  02/23/2009 Steve Loar - INITIAL VERSION
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim lcl_orghasfeature_gl_accounts

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "manage classes", sLevel	' In common.asp

lcl_orghasfeature_gl_accounts = orghasfeature("gl accounts")

blnHasWP = hasWordPress()
sHomeWebsiteURL = getOrganization_WP_URL(session("orgid"), "OrgPublicWebsiteURL")

%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />

	<title>E-Gov Administration Console</title>

	<style type="text/css">
	/*margin and padding on body element
	  can introduce errors in determining
	  element position and are not recommended;
	  we turn them off as a foundation for YUI
	  CSS treatments. */
	body {
		margin:0;
		padding:0;
	}
	</style>

	<link rel="stylesheet" type="text/css" href="../yui/build/tabview/assets/skins/sam/tabview.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />

	<!--
	<script type="text/javascript" src="../yui/build/yahoo-dom-event/yahoo-dom-event.js"></script>
	<script type="text/javascript" src="../yui/build/element/element-beta.js"></script>
	<script type="text/javascript" src="../yui/build/tabview/tabview.js"></script>
	-->
	<script type="text/javascript" src="../yui/yahoo-dom-event.js"></script>  
	<script type="text/javascript" src="../yui/element-min.js"></script>  
	<script type="text/javascript" src="../yui/tabview-min.js"></script>

  	<script src="//code.jquery.com/jquery-1.12.4.js"></script>
   	<script src="//code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
	<!--#include file="../includes/wp-image-picker.asp"-->
	<script language="JavaScript" src="../scriptaculous/src/scriptaculous.js"></script>

	<script language="javascript" src="../scripts/ajaxLib.js"></script>
	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="javascript" src="../scripts/formatnumber.js"></script>
	<script language="javascript" src="../scripts/removespaces.js"></script>
	<script language="javascript" src="../scripts/removecommas.js"></script>
	<script language="JavaScript" src="../scripts/layers.js"></script>
	<script language="javascript" src="../scripts/setfocus.js"></script>
	<script language="javascript" src="../scripts/isvaliddate.js"></script>
	<script language="javascript" src="../scripts/textareamaxlength.js"></script>

	<script language="javascript">
	<!--
		var tabView;

		(function() {
			tabView = new YAHOO.widget.TabView('demo');
			tabView.set('activeIndex', 0); 

		})();

		function ClassChange() 
		{
			// Try to get a drop down of names
			//alert(document.getElementById("earlyregistrationclassseasonid").options[document.getElementById("earlyregistrationclassseasonid").selectedIndex].value);
			doAjax('getseasonclasses.asp', 'classseasonid=' + document.getElementById("earlyregistrationclassseasonid").options[document.getElementById("earlyregistrationclassseasonid").selectedIndex].value, 'UpdateClasses', 'get', '0');
		}

		function UpdateClasses( sResult )
		{
			document.getElementById('earlyclass').innerHTML = sResult;
		}

		function selectAll(selectBox,selectAll) 
		{
			// have we been passed an ID
			if (typeof selectBox == "string") {
				selectBox = document.getElementById(selectBox);
			}
			// is the select box a multiple select box?
			if (selectBox.type == "select-multiple") {
				for (var i = 0; i < selectBox.options.length; i++) {
					selectBox.options[i].selected = selectAll;
				}
			}
		}

		function ValidatePrice( oPrice )
		{
			var bValid = true;
			var total = 0.00;

			// Remove any extra spaces
			oPrice.value = removeSpaces(oPrice.value);
			//Remove commas that would cause problems in validation
			oPrice.value = removeCommas(oPrice.value);

			// Validate the format of the price
			if (oPrice.value != "")
			{
				var rege = /^\d*\.?\d{0,2}$/
				var Ok = rege.exec(oPrice.value);
				if ( Ok )
				{
					oPrice.value = format_number(Number(oPrice.value),2);
				}
				else 
				{
					oPrice.value = "";
					alert("Prices must be numbers in currency format or blank.\nPlease correct to continue.");
					tabView.set('activeIndex',4);
					setfocus(oPrice);
					return false;
				}
			}
		}

		function doCalendar(sField) 
		{
		  var w = (screen.width - 350)/2;
		  var h = (screen.height - 350)/2;
		  eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=ClassForm", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function doPicker(sFormField) 
		{
		  w = (screen.width - 350)/2;
		  h = (screen.height - 350)/2;
		  eval('window.open("imagepicker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
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
		}

		function openWin1(url, name) 
		{
			var w = (screen.width - 350)/2;
			var h = (screen.height - 350)/2;
			popupWin = eval('window.open(url, name,"resizable,width=800,height=400,left=' + 100 + ',top=' + h + '")');
		}

		function openWin2(url, name) 
		{
			var w = (screen.width - 350)/2;
			var h = (screen.height - 550)/2;
			popupWin = eval('window.open(url, name,"resizable,width=820,height=600,scrollbars=1,left=' + 80 + ',top=' + h + '")');
		}

		function CopyClass(iClassid)
		{
			if (confirm('Copy this to a new Class/Event?'))
			{
				//location.href='class_copyclass.asp?classid=' + iClassid + '&seasonid=' + something + '&bring=' + bring;
				document.copyForm.submit();
			}
		}

		function CopyToChild(iClassid)
		{
			if (confirm('Create a new Individual for this Series?'))
			{
				location.href='class_copyaschild.asp?classid=' + iClassid;
			}
		}

		function AddSingle( )
		{
			if (confirm('Add \'' + document.SingleForm.classid.options[document.SingleForm.classid.selectedIndex].text + '\' to the Series?'))
			{
				document.SingleForm.submit();
			}
		}

		function ValidateForm()
		{
			var rege;
			var Ok;
			var iRow;
			var x;
			var regOK;
			var TimeRequired;

			// check the class name
			if (document.ClassForm.classname.value == "")
			{
				alert('Please enter a name.');
				document.ClassForm.classname.focus();
				return;
			}
			if (document.ClassForm.classname.value.length > 255)
			{
				alert('The classname is limited to 255 characters.\nPlease shorten this name.');
				document.ClassForm.classname.focus();
				return;
			}
			// check the class description
			if (document.ClassForm.classdescription.value == "")
			{
				tabView.set('activeIndex',0);
				alert('Please enter a description.');
				document.ClassForm.classdescription.focus();
				return;
			}

			// check the length of the search key words
			if (document.ClassForm.searchkeywords.value.length > 1024)
			{
				tabView.set('activeIndex',0);
				alert("The maximum length of search keywords is 1024 characters. \nPlease make this smaller.");
				document.ClassForm.searchkeywords.focus();
				return;
			}

			// check the startdate
			if (document.ClassForm.startdate.value == "")
			{
				tabView.set('activeIndex',3);
				alert("Please enter a start date");
				document.ClassForm.startdate.focus();
				return;
			}
			else
			{
				//rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				//Ok = rege.test(document.ClassForm.startdate.value);
				//if (! Ok)
				if (! isValidDate(document.ClassForm.startdate.value))
				{
					tabView.set('activeIndex',3);
					alert("Start date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.ClassForm.startdate.focus();
					return;
				}
			}

			// check the enddate
			if (document.ClassForm.enddate.value == "")
			{
				tabView.set('activeIndex',3);
				alert("Please enter an end date");
				document.ClassForm.enddate.focus();
				return;
			}
			else
			{
				//rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				//Ok = rege.test(document.ClassForm.enddate.value);
				//if (! Ok)
				if (! isValidDate(document.ClassForm.enddate.value))
				{
					tabView.set('activeIndex',3);
					alert("End date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.ClassForm.enddate.focus();
					return;
				}
			}

			// check publish start date
			if (document.ClassForm.publishstartdate.value != "")
			{
				//rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				//Ok = rege.test(document.ClassForm.publishstartdate.value);
				//if (! Ok)
				if (! isValidDate(document.ClassForm.publishstartdate.value))
				{
					tabView.set('activeIndex',3);
					alert("Publication start date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.ClassForm.publishstartdate.focus();
					return;
				}
			}

			// check publish end date
			if (document.ClassForm.publishenddate.value != "")
			{
				//rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				//Ok = rege.test(document.ClassForm.publishenddate.value);
				//if (! Ok)
				if (! isValidDate(document.ClassForm.publishenddate.value))
				{
					tabView.set('activeIndex',3);
					alert("Publication end date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.ClassForm.publishenddate.focus();
					return;
				}
			}

			// check registration end date
			if (document.ClassForm.registrationenddate.value != "")
			{
				//rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				//Ok = rege.test(document.ClassForm.registrationenddate.value);
				//if (! Ok)
				if (! isValidDate(document.ClassForm.registrationenddate.value))
				{
					tabView.set('activeIndex',3);
					alert("Registration end date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.ClassForm.registrationenddate.focus();
					return;
				}
			}

			// Check pricing on the ones they checked
			for (var p = parseInt(document.ClassForm.minpricetypeid.value); p <= parseInt(document.ClassForm.maxpricetypeid.value); p++)
			{
				//alert(document.getElementById("pricetypeid" + p).value);
				// Does it exist
				if (document.getElementById("pricetypeid" + p))
				{
					//alert(document.getElementById("pricetypeid" + p).checked);
					//Is is checked
					if(document.getElementById("pricetypeid" + p).checked)
					{
						// Remove any extra spaces
						document.getElementById("amount" + p).value = removeSpaces(document.getElementById("amount" + p).value);
						//Remove commas that would cause problems in validation
						document.getElementById("amount" + p).value = removeCommas(document.getElementById("amount" + p).value);

						// Is the price formated correctly
						rege = /^\d+\.\d{2}$/;
						Ok = rege.test(document.getElementById("amount" + p).value);
						if (! Ok)
						{
							tabView.set('activeIndex',4);
							alert("Selected prices cannot be blank and must be in currency format.");
							document.getElementById("amount" + p).focus();
							return;
						}
						// check that there is a registration start date
						if (document.getElementById("registrationstartdate" + p).getAttribute("type") != 'hidden')
						{
							//alert(document.getElementById("registrationstartdate" + p).value);
							//rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
							//Ok = rege.test(document.getElementById("registrationstartdate" + p).value);
							//if (! Ok)
							if (! isValidDate(document.getElementById("registrationstartdate" + p).value))
							{
								tabView.set('activeIndex',4);
								alert("Selected registration start dates cannot be blank and should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
								document.getElementById("registrationstartdate" + p).focus();
								return;
							}
						}
					}
				}
			}

			//alert('OK');
			//return;

			// all is ok, so submit and save the changes
			document.ClassForm.submit();
		}

		function SetUpPage()
		{
			setMaxLength();
			$("#classname").focus();
		}

	<%		If request("success") <> "" Then 
				DisplayMessagePopUp 
			End If 
	%>

	//-->
	</script>
</head>
<body class="yui-skin-sam" onload="SetUpPage();">
 
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

	<p>
  		<a href="javascript:location.href='class_list.asp'"><img src="../images/arrow_2back.gif" align="absmiddle" border="0" />&nbsp;Return to Class/Event List</a>
	</p>
	
<%
	Dim sSql, oClass, iClassId, bIsSeriesParent, sClassName, iClassSeasonId, sClassDesc, sKeyWords, sNotes, iRegattaSignupTypeId
	Dim sStartDate, sEndDate, sPublishStartDate, sPublishEndDate, sRegistrationStartDate, sRegistrationEndDate
	Dim sExternalurl, sExternallinktext, sImageURL, sImageAltText

	iClassId = CLng(request("classid"))

	If iClassId > CLng(0) Then 

		sSQL = "SELECT classname, isnull(classseasonid,0) as classseasonid, ISNULL(classdescription,'') AS classdescription, "
		sSQL = sSQL & " startdate, enddate, publishstartdate, publishenddate, registrationstartdate, registrationenddate, "
		sSQL = sSQL & " ISNULL(searchkeywords,'') AS searchkeywords, "
		sSQL = sSQL & " isparent, classtypeid, statusid, cancelreason, "
		sSQL = sSQL & " ISNULL(imgurl,'') AS imgurl, ISNULL(imgalttag,'') AS imgalttag, "
		sSQL = sSQL & " ISNULL(notes,'') AS notes, ISNULL(externalurl,'') AS externalurl, ISNULL(externallinktext,'') AS externallinktext, "
		sSQL = sSQL & " isregatta, ISNULL(regattasignuptypeid, 0) AS regattasignuptypeid "
		sSQL = sSQL & " FROM egov_class "
		sSQL = sSQL & " WHERE classid = " & iClassId
		sSQL = sSQL & " AND orgid = "   & session("orgid")

		Set oClass = Server.CreateObject("ADODB.Recordset")
		oClass.Open sSQL, Application("DSN"), 3, 1
		

		If Not oClass.EOF Then
			sClassName = oClass("classname")
			iClassSeasonId = CLng(oClass("classseasonid"))
			sClassDesc = oClass("classdescription")
			sKeyWords = oClass("searchkeywords")
			sNotes = oClass("notes")
			sExternalurl = oClass("externalurl")
			sExternallinktext = oClass("externallinktext")
			sImageURL = oClass("imgurl")
			sImageAltText = oClass("imgalttag")

			If CLng(oClass("regattasignuptypeid")) > CLng(0) Then 
				iRegattaSignupTypeId = oClass("regattasignuptypeid")
			Else
				iRegattaSignupTypeId = 0
			End If 

			If IsNull(oClass("startdate")) Then
				sStartDate = ""
			Else
				sStartDate = FormatDateTime(oClass("startdate"),2)
			End If 
			If IsNull(oClass("enddate")) Then
				sEndDate = ""
			Else
				sEndDate = FormatDateTime(oClass("enddate"),2)
			End If 
			If IsNull(oClass("publishstartdate")) Then
				sPublishStartDate = ""
			Else
				sPublishStartDate = FormatDateTime(oClass("publishstartdate"),2)
			End If 
			If IsNull(oClass("publishenddate")) Then
				sPublishEndDate = ""
			Else
				sPublishEndDate = FormatDateTime(oClass("publishenddate"),2)
			End If 
			If IsNull(oClass("registrationstartdate")) Then
				sRegistrationStartDate = ""
			Else
				sRegistrationStartDate = FormatDateTime(oClass("registrationstartdate"),2)
			End If 
			If IsNull(oClass("registrationenddate")) Then
				sRegistrationEndDate = ""
			Else
				sRegistrationEndDate = FormatDateTime(oClass("registrationenddate"),2)
			End If 
			sSaveText = "Save Updates"

		End If 

		oClass.Close
		Set oClass = Nothing 
	Else
		sClassName = ""
		iClassSeasonId = 0
		sClassDesc = ""
		sKeyWords = ""
		iRegattaSignupTypeId = 0
		sNotes = ""
		sStartDate = ""
		sEndDate = ""
		sPublishStartDate = ""
		sPublishEndDate = ""
		sRegistrationStartDate = ""
		sRegistrationEndDate = ""
		sSaveText = "Create Regatta Event"
		sExternalurl = ""
		sExternallinktext = ""
		sImageURL = ""
		sImageAltText = ""
	End If 

%>
	<p>
		  <input type="button" class="button" name="update1" id="update1" value="<%=sSaveText%>" onclick="ValidateForm();" />
	</p>
<%

		response.write "<form name=""ClassForm"" id=""ClassForm"" method=""post"" action=""regattaeventupdate.asp"">"
		response.write "<input type=""hidden"" name=""classid"" id=""classid"" value=""" & iClassId & """ />"
		response.write "<p>"
		response.write "Name: <input type=""text"" name=""classname"" id=""classname"" value=""" & sClassName & """ size=""50"" maxlength=""50"" />"
		response.write "</p>"

		response.write "<p>" & vbcrlf
		response.write "   Season: " & vbcrlf

		ShowClassSeasonFilterPicks iClassSeasonId 'In class_global_functions.asp

		response.write "</p>" & vbcrlf
%>

	<div id="demo" class="yui-navset">
		 <ul class="yui-nav">
			  <li><a href="#tab1"><em>General Information</em></a></li>
			  <li><a href="#tab2"><em>Categories</em></a></li>
			  <li><a href="#tab3"><em>Waivers</em></a></li>
			  <li><a href="#tab4"><em>Key Dates</em></a></li>
			  <li><a href="#tab5"><em>Purchasing</em></a></li>
		 </ul>            
		<div class="yui-content">

  			<div id="tab1"> <!-- General Information -->
				<p><br />
					Regatta Signup Type:&nbsp; <%ShowRegattaSignupTypes iRegattaSignupTypeId %>
				</p>
				<p>
					Description:<br />
					<textarea name="classdescription" id="classdescription" maxlength="6000" wrap="soft"><%=sClassDesc%></textarea>
 				</p>
 				<p>
				Image:

					<img src="<%=sImageURL%>" id="imgurlpic" name="imgurlpic" border="0" alt="<%=sImageAtlText%>" align="middle" width="180" height="180"  onerror="this.src = '../images/placeholder.png';" />
					<input type="<%if blnHasWP then %>hidden<%else%>text<%end if%>" name="imgurl" class="imageurl" id="imgurl" value="<%=sImageURL%>" size="50" maxlength="255" style="display:block" />
					<% if blnHasWP then %>
						<input type="button" class="button" value="Change" onclick="showModal('Pick Image',65,80,'imgurl');" />
					<% else%>
						<input type="button" class="button" value="Browse..." onclick="javascript:doPicker('ClassForm.imgurl');" />
						&nbsp; &nbsp; <input type="button" class="button" name="upload" value="Upload" onclick="openWin2('../docs/default.asp','_blank')" />
					<% end if%>
					<br /><span id="imgalttag">Image Alt Tag:</span> <input type="text" name="imgalttag" value="<%=sImageAltText%>" size="50" maxlength="255" />
					<br /><span id="imgalttabdesc"> The Image Alt Tag is a description used for ADA compliance.</span>
 				</p>

 				<p>
				   Search Keywords:<br />
				   <textarea name="searchkeywords" id="searchkeywords" maxlength="1024" wrap="soft"><%=sKeyWords%></textarea>
 				</p>

<%
				'External URL and Text -------------------------------------------------
				response.write vbcrlf & "<p>"
				response.write "External URL: "
				response.write "<input type=""text"" name=""externalurl"" id=""externalurl"" value=""" & sExternalurl & """ size=""50"" maxlength=""255"" />"
				response.write "<br />" & vbcrlf
				response.write "External URL Text: "
				response.write "<input type=""text"" name=""externallinktext"" id=""externallinktext"" value=""" & sExternallinktext & """ size=""50"" maxlength=""255"" />"
				response.write "</p>" & vbcrlf


			   'Receipt Notes ---------------------------------------------------------
				response.write "<p>"
				response.write "Activity Notes: (These show on the receipt)<br />" & vbcrlf
				response.write "<textarea name=""notes"" id=""receiptnotes"" maxlength=""1024"" wrap=""soft"">" & sNotes & "</textarea>"
				response.write "</p>" & vbcrlf
%>
			</div>

			<div id="tab2"> <!-- Categories -->
				<div style="display: block;">
					<div>
						<div class="rightbuttons">
							<input type="button" class="assignbuttons" name="categorymgr" value="Manage Categories" onclick="openWin2('category_mgmt.asp','_blank')" /> 
						</div>
						<% ShowCategories iClassId, session("orgid") %>
					</div>
				</div>
			</div>

			<div id="tab3"> <!-- Waivers -->
				<div>
					<div class="rightbuttons">
						<input type="button" class="assignbuttons" name="waivermgr" value="Manage Waivers" onclick="openWin2('class_waivers.asp','_blank')" /> 
					</div>
					<% ShowWaiverPicks iClassId %>
					<br /><input type="button" class="assignbuttons" name="NoWaivers" value="Clear Selection" onclick="selectAll(document.getElementById('waiverid'),false)" />
				</div>
				
				<div id="waivernote">
					Note: To add new waivers, click on Manage Waivers, and create the waiver.
					The new waiver will not appear in this list until after you save your changes to this page.
				</div>
			</div>

			<div id="tab4"> <!-- Key Dates -->
			  <p><br />
				 <table id="criticaldates" border="0" cellpadding="1" cellspacing="3">
					  <tr>
      					<td align="right">Event Starts:</td>
						<td><input type="text" maxlength="10" class="datefield" id="startdate" name="startdate" value="<%=sStartDate%>" />&nbsp;
							<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('startdate');" /></span>
					    </td>																							
      					<td align="right">Event Ends:</td>
						<td><input type="text" maxlength="10" class="datefield" id="enddate" name="enddate" value="<%=sEndDate%>" />&nbsp;
							<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('enddate');" /></span>
      					</td>
   					</tr>
					<tr>
						<td align="right">Publication Starts:</td>
						<td>
							<input type="text" maxlength="10" class="datefield" id="publishstartdate" name="publishstartdate" value="<%=sPublishStartDate%>" />
							<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('publishstartdate');" /></span>
						</td>
						<td align="right">Publication Ends:</td>
						<td>
							<input type="text" maxlength="10" class="datefield" id="publishenddate" name="publishenddate" value="<%=sPublishEndDate%>" />
							<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('publishenddate');" /></span>
						</td>
					</tr>
   					<tr>
			      		<td align="right">Registration Starts:</td>
						<td><input type="text" maxlength="10" class="datefield" name="registrationstartdate" value="<%=sRegistrationStartDate%>" />&nbsp;
							<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('registrationstartdate');" /></span>
						</td>
      					<td align="right">Registration Ends:</td>
						<td><input type="text" maxlength="10" class="datefield" name="registrationenddate" value="<%=sRegistrationEndDate%>" />&nbsp;
							<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('registrationenddate');" /></span>
        				</td>
					</tr>
			 	</table>
			</p>
		</div>

		<div id="tab5"> <!-- Purchasing -->
			<p>
				<strong>Pricing:</strong><br /><br />
				<table id="pricingtable" border="0" cellpadding="0" cellspacing="0">
					<tr><th>Type</th>
           <%
					If lcl_orghasfeature_gl_accounts Then 
						response.write "<th>Account</th>" & vbcrlf
					End If 
           %>
					<th>Price</th><th>Registration Starts</th></tr>

  					<% GetPricing iClassId, iClassSeasonId, iMin, iMax %>
					</table>
<%
  					response.write vbcrlf & "<input type=""hidden"" name=""minpricetypeid"" value=""" & iMin & """ />"
					response.write vbcrlf & "<input type=""hidden"" name=""maxpricetypeid"" value=""" & iMax & """ />"
%>
				</p>
				<input type="hidden" name="pricediscountid" value="0" />
		</div>

		</div>
	</div>

		<p>
			<input type="button" class="button" name="update2" value="<%=sSaveText%>" onclick="ValidateForm();" />
		</p>
		</form>

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
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub GetPricing( iClassId, iClassSeasonId, ByRef iMin, ByRef iMax )
'--------------------------------------------------------------------------------------------------
Sub GetPricing( iClassId, iClassSeasonId, ByRef iMin, ByRef iMax )
	Dim sSql, oPrice, oRs, iRegattaPriceCount

	iMax = 0
	iMin = 10000

	iRegattaPriceCount = GetCountOfRegattaPrices( )

	sSql = "SELECT pricetypeid, pricetypename, needsregistrationstartdate "
	sSql = sSql & " FROM egov_price_types "
	sSql = sSql & " WHERE isregattaprice = 1 AND orgid = " & Session("OrgID")
	sSql = sSql & " ORDER BY displayorder "

	Set oPrice = Server.CreateObject("ADODB.Recordset")
	oPrice.Open sSQL, Application("DSN"), 0, 1

	Do While Not oPrice.EOF
		If CLng(oPrice("pricetypeid")) > CLng(iMax) Then 
		   iMax = oPrice("pricetypeid")
		End If 
		If CLng(oPrice("pricetypeid")) < CLng(iMin) Then 
		   iMin = oPrice("pricetypeid")
		End If 
		response.write "<tr>" & vbcrlf
		response.write "<td class=""type"">"
		response.write "<input type=""checkbox"" name=""pricetypeid"" id=""pricetypeid" & oPrice("pricetypeid") & """ value=""" & oPrice("pricetypeid") & """"
		If CLng(iClassid) = CLng(0) And CLng(iRegattaPriceCount) = CLng(1) Then
			response.write " checked=""checked"" "
		Else 
			If ClassHasPriceType( iClassId, clng(oPrice("pricetypeid")) ) Then
				response.write " checked=""checked"" "
			End If 
		End If 

		response.write " />&nbsp;" & oPrice("pricetypename") 

		If lcl_orghasfeature_gl_accounts Then 
  			response.write "    </td>"
		  	iAccountId = GetClassPricetypeAccount( iClassId, CLng(oPrice("pricetypeid")) )
  			response.write "    <td align=""center"">"
		  	ShowAccountPicks iAccountId, oPrice("pricetypeid")  ' In common.asp
		Else
  			response.write vbcrlf & "<input type=""hidden"" name=""accountid" & oPrice("pricetypeid") & """ value=""0"" />"
		End If 
		response.write "    </td>"

		response.write "    <td align=""center""><input type=""text"" name=""amount" & oPrice("pricetypeid") & """ id=""amount" & oPrice("pricetypeid") & """ value=""" & GetPriceAmount( oPrice("pricetypeid"), iClassId ) & """ size=""10"" maxlength=""9"" onchange=""ValidatePrice(this);"" /></td>"
		
		If oPrice("needsregistrationstartdate") Then 
			response.write "<td align=""center"">" & vbcrlf
	'		response.write "    <input type=""text"" maxlength=""10"" class=""datefield"" name=""registrationstartdate" & oPrice("pricetypeid") & """ id=""registrationstartdate" & oPrice("pricetypeid") & """ value=""" & GetClassPricetypeRegistrationStart( iClassId, oPrice("pricetypeid"), GetDefaultSeasonDate( iClassSeasonId, "registrationstartdate" ) ) & """ />&nbsp;" & vbcrlf

			'Find the seasonal start date for the pricetype
			 sSql = "SELECT registrationstartdate "
			 sSql = sSql & " FROM egov_class_seasons_to_pricetypes_dates "
			 sSql = sSql & " WHERE classseasonid = " & iClassSeasonId
			 sSql = sSql & " AND pricetypeid = " & oPrice("pricetypeid")

			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.Open sSql, Application("DSN"), 0, 1

			 If Not oRs.eof Then
				If Not IsNull() Then 
					lcl_registration_start_date = FormatDateTime(oRs("registrationstartdate"),2)
				Else
					lcl_registration_start_date = ""
				End If  
			 Else 
				lcl_registration_start_date = ""
			 End If 

			 If ClassHasPriceType( iClassId, clng(oPrice("pricetypeid")) ) Then 
				lcl_registration_start_date = GetClassPricetypeRegistrationStart( iClassId, oPrice("pricetypeid"), lcl_registration_start_date )
			 End If 

			 response.write "    <input type=""text"" maxlength=""10"" class=""datefield"" name=""registrationstartdate" & oPrice("pricetypeid") & """ id=""registrationstartdate" & oPrice("pricetypeid") & """ value=""" & lcl_registration_start_date & """ />&nbsp;" & vbcrlf
			 response.write "    <span class=""calendarimg"" style=""cursor:hand;"">" & vbcrlf
			 response.write "      <img src=""../images/calendar.gif"" height=""16"" width=""16"" border=""0"" onclick=""javascript:void doCalendar('registrationstartdate" & oPrice("pricetypeid") & "');"" />" & vbcrlf
			 response.write "    </span>" & vbcrlf
			 response.write "</td>"
		Else 
			 response.write "<td>&nbsp;<input type=""hidden"" name=""registrationstartdate" & oPrice("pricetypeid") & """ id=""registrationstartdate" & oPrice("pricetypeid") & """ value="""" /></td>"
		End If 
		response.write "</tr>"	

		oPrice.MoveNext
	Loop 

	oPrice.Close
	Set oPrice = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function ClassHasPriceType( iClassId, iPriceTypeId )
'--------------------------------------------------------------------------------------------------
Function ClassHasPriceType( iClassId, iPriceTypeId )
	Dim sSql, oPrice

	sSQL = "SELECT count(pricetypeid) as hits "
	sSQL = sSQL & " FROM egov_class_pricetype_price "
	sSQL = sSQL & " WHERE pricetypeid = " & iPriceTypeId
	sSQL = sSQL & " AND classid = " & iClassId

	Set oPrice = Server.CreateObject("ADODB.Recordset")
	oPrice.Open sSQL, Application("DSN"), 0, 1

	if clng(oPrice("hits")) > clng(0) then
    ClassHasPriceType = True
	else
  		ClassHasPriceType = False
	end if

	oPrice.close
	Set oPrice = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetClassPricetypeAccount( iClassId, iPriceTypeId )
'--------------------------------------------------------------------------------------------------
Function GetClassPricetypeAccount( iClassId, iPriceTypeId )
	Dim sSql, oPrice

	sSQL = "SELECT isnull(accountid,0) as accountid "
	sSQL = sSQL & " FROM egov_class_pricetype_price "
	sSQL = sSQL & " WHERE pricetypeid = " & iPriceTypeId
	sSQL = sSQL & " AND classid = " & iClassId

	Set oPrice = Server.CreateObject("ADODB.Recordset")
	oPrice.Open sSQL, Application("DSN"), 0, 1

	If Not oPrice.EOF Then 
		GetClassPricetypeAccount = oPrice("accountid")
	Else 
		GetClassPricetypeAccount = 0
	End If 

	oPrice.close
	Set oPrice = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetClassPricetypeInstructorPercent( iClassId, iPriceTypeId, iDefaultPercent )
'--------------------------------------------------------------------------------------------------
Function GetClassPricetypeInstructorPercent( iClassId, iPriceTypeId, iDefaultPercent )
	Dim sSql, oPrice

	sSql = "Select isnull(instructorpercent,0) as instructorpercent from egov_class_pricetype_price where pricetypeid = " & iPriceTypeId & " and classid = " & iClassId

	Set oPrice = Server.CreateObject("ADODB.Recordset")
	oPrice.Open sSQL, Application("DSN"), 0, 1

	If Not oPrice.EOF Then
		GetClassPricetypeInstructorPercent = oPrice("instructorpercent")
	Else 
		GetClassPricetypeInstructorPercent = iDefaultPercent
	End If 

	oPrice.close
	Set oPrice = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetClassPricetypeRegistrationStart( iClassId, iPriceTypeId, sDefaultDate )
'--------------------------------------------------------------------------------------------------
Function GetClassPricetypeRegistrationStart( iClassId, iPriceTypeId, sDefaultDate )
	Dim sSql, oPrice

	sSQL = "SELECT registrationstartdate "
	sSQL = sSQL & " FROM egov_class_pricetype_price "
	sSQL = sSQL & " WHERE pricetypeid = " & iPriceTypeId
	sSQL = sSQL & " AND classid = " & iClassId

	Set oPrice = Server.CreateObject("ADODB.Recordset")
	oPrice.Open sSQL, Application("DSN"), 0, 1

	If Not oPrice.EOF Then
		  GetClassPricetypeRegistrationStart = oPrice("registrationstartdate")
	Else 
		  GetClassPricetypeRegistrationStart = sDefaultDate
	End If 

	oPrice.close
	Set oPrice = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetPriceAmount( iPriceTypeId, iClassId )
'--------------------------------------------------------------------------------------------------
Function GetPriceAmount( iPriceTypeId, iClassId )
	Dim sSql, oAmount, sAmount

	sSql = "Select amount from egov_class_pricetype_price where pricetypeid = " & iPriceTypeId & " and classid = " & iClassId

	Set oAmount = Server.CreateObject("ADODB.Recordset")
	oAmount.Open sSQL, Application("DSN"), 0, 1

	If Not oAmount.EOF Then
		sAmount = FormatNumber(oAmount("amount"),2,,,0)
		GetPriceAmount = Replace(sAmount,",","") ' take out commas
	Else 
		GetPriceAmount = ""
	End If 

	oAmount.close
	Set oAmount = Nothing

End Function


'--------------------------------------------------------------------------------------------------
' Function GetDefaultSeasonDate( iClassSeasonId, sDateField )
'--------------------------------------------------------------------------------------------------
Function GetDefaultSeasonDate( iClassSeasonId, sDateField )
	Dim sSql, oSeason

	sSql = "SELECT " & sDateField & " FROM egov_class_seasons WHERE classseasonid = " & iClassSeasonId

	Set oSeason = Server.CreateObject("ADODB.Recordset")
	oSeason.Open sSQL, Application("DSN"), 0, 1

	if NOT oSeason.EOF then
		  if IsNull(oSeason(sDateField)) then
			    GetDefaultSeasonDate = ""
    else
		    	GetDefaultSeasonDate = oSeason(sDateField)
  		end if
	else
		  GetDefaultSeasonDate = ""
	end if

	oSeason.close
	Set oSeason = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Sub GetPricing1( iClassId, iOrgID, iMembershipId )
'--------------------------------------------------------------------------------------------------
Sub GetPricing1( iClassId, iOrgID, iMembershipId )
	Dim sSql, oPrice, iRow

	iRow = 0
	sSql = "Select pricetypeid, pricetypename, ismember from egov_price_types where orgid = " & iOrgId & " Order By displayorder"

	Set oPrice = Server.CreateObject("ADODB.Recordset")
	oPrice.Open sSQL, Application("DSN"), 0, 1

	Do While Not oPrice.EOF
		iRow = iRow + 1
		response.write vbcrlf & "<tr><td><input type=""hidden"" name=""pricetypeid" & iRow & """ value=""" & oPrice("pricetypeid") & """ />" & oPrice("pricetypename") & "</td>"
		response.write "<td><input type=""text"" name=""amount" & iRow & """ value=""" & GetPriceAmount( oPrice("pricetypeid"), iClassId ) & """ size=""10"" maxlength=""9"" /></td><td>"
		If oPrice("ismember") Then
			' Show the membership picks for the one that requires membership
			response.write "Membership: &nbsp; "
			ShowMembershipPicks iMembershipId
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td></tr>"	
		oPrice.movenext
	Loop 
	response.write vbcrlf & "<input type=""hidden"" name=""pricetypeidcount"" value=""" & iRow & """ />"

	oPrice.close
	Set oPrice = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowSeriesChildren( iParentClassId )
'--------------------------------------------------------------------------------------------------
Sub ShowSeriesChildren( iParentClassId )
	Dim sSql, oChildren, iRowCOunt

	sSql = "Select C.classid, C.classname, C.startdate, C.enddate, S.statusname from egov_class C, egov_class_status S "
	sSql = sSql & " where C.parentclassid = " & iParentClassId & " and C.statusid = S.statusid Order By C.startdate, C.classname"
	iRowCOunt = 1

	Set oChildren = Server.CreateObject("ADODB.Recordset")
	oChildren.Open sSQL, Application("DSN"), 0, 1

	Do While Not oChildren.EOF
		iRowCOunt = iRowCOunt + 1
		If ( iRowCOunt Mod 2 ) = 0 Then 
			response.write vbcrlf & "<tr class=""alt_row"">"
		Else 
			response.write vbcrlf & "<tr>"
		End If 
		response.write vbcrlf & "<td>" & oChildren("classname") & "</td><td align=""center"">" & oChildren("statusname") & "</td><td align=""center"">" & oChildren("startdate") & "</td><td align=""center"">" & oChildren("enddate") & "</td>"
		response.write "<td align=""center""><input type=""button"" class=""button"" name=""editchild"" value=""Edit"" onclick=""javascript:location.href='edit_class.asp?classid=" & oChildren("classid") & "';"" /></td></tr>"
		oChildren.movenext
	Loop 

	oChildren.close
	Set oChildren = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowRegattaSignupTypes( iRegattaSignupTypeId )
'--------------------------------------------------------------------------------------------------
Sub ShowRegattaSignupTypes( ByVal iRegattaSignupTypeId )
	Dim sSql, oRs

	sSql = "SELECT regattasignuptypeid, regattasignuptype FROM egov_regattasignuptype "
	sSql = sSql & " WHERE orgid = " & SESSION("orgid") & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""regattasignuptypeid"">"
		Do While Not oRs.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oRs("regattasignuptypeid") & """ "
				If clng(oRs("regattasignuptypeid")) = clng(iRegattaSignupTypeId) Then 
					response.write " selected=""selected"" "
				End If 
			response.write ">" & oRs("regattasignuptype") & "</option>"
			oRs.MoveNext
		Loop 
		response.write "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowLocationPicks( iLocationId, iOrgId )
'--------------------------------------------------------------------------------------------------
Sub ShowLocationPicks( iLocationId, iOrgId )
	Dim sSql, oLocation

	sSql = "Select locationid, name from egov_class_location where orgid = " & iOrgId & " Order By name"

	Set oLocation = Server.CreateObject("ADODB.Recordset")
	oLocation.Open sSQL, Application("DSN"), 0, 1

	If Not oLocation.EOF Then
		response.write vbcrlf & "<select name=""locationid"" >"
		'response.write vbcrlf & vbtab & "<option value=""0"">Select a Location</option>"
		Do While Not oLocation.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oLocation("locationid") & """ "
				If clng(oLocation("locationid")) = clng(iLocationId) Then 
					response.write " selected=""selected"" "
				End If 
			response.write ">" & oLocation("name") & "</option>"
			oLocation.movenext
		Loop 
		response.write "</select>"
	End If 

	oLocation.close
	Set oLocation = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowPOCPicks( iPocId, iOrgId )
'--------------------------------------------------------------------------------------------------
Sub ShowPOCPicks( iPocId, iOrgId )
	Dim sSql, oPOC

	sSql = "Select pocid, name from egov_class_pointofcontact where orgid = " & iOrgId & " Order By name"

	Set oPOC = Server.CreateObject("ADODB.Recordset")
	oPOC.Open sSQL, Application("DSN"), 0, 1

	If Not oPOC.EOF Then
		response.write vbcrlf & "<select name=""pocid"" >"
		'response.write vbcrlf & vbtab & "<option value=""0"">Select a POC</option>"
		Do While Not oPOC.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oPOC("pocid") & """ "
				If clng(oPOC("pocid")) = clng(iPocId) Then 
					response.write " selected=""selected"" "
				End If 
			response.write ">" & oPOC("name") & "</option>"
			oPOC.movenext
		Loop 
		response.write "</select>"
	End If 

	oPOC.close
	Set oPOC = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowInstructors( iClassId )
'--------------------------------------------------------------------------------------------------
Sub ShowInstructors( iClassId )
	Dim sSql, oInstructor, iRow

	iRow = 0
	sSql = "Select L.lastname, L.firstname from egov_class_instructor L, egov_class_to_instructor C"
	sSql = sSql & " where L.instructorid = C.instructorid"
	sSql = sSql & " and C.classid = " & iClassId & " Order By L.lastname, L.firstname"

	Set oInstructor = Server.CreateObject("ADODB.Recordset")
	oInstructor.Open sSQL, Application("DSN"), 0, 1

	If Not oInstructor.EOF Then
		Do While Not oInstructor.EOF
			If iRow > 0 Then
				response.write ", "
			End If 
			response.write oInstructor("firstname") & " " & oInstructor("lastname")
			oInstructor.movenext
			iRow = iRow + 1
		Loop 
	Else 
		response.write " No Instructor Assigned"
	End If 

	oInstructor.close
	Set oInstructor = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowWaivers( iClassId )
'--------------------------------------------------------------------------------------------------
Sub ShowWaivers( iClassId )
	Dim sSql, oWaiver, iRow

	iRow = 0
	sSql = "Select waivername from egov_class_waivers W, egov_class_to_waivers C"
	sSql = sSql & " where W.waiverid = C.waiverid"
	sSql = sSql & " and C.classid = " & iClassId & " Order By waivername"

	Set oWaiver = Server.CreateObject("ADODB.Recordset")
	oWaiver.Open sSQL, Application("DSN"), 0, 1

	If Not oWaiver.EOF Then
		Do While Not oWaiver.EOF
			If iRow > 0 Then
				response.write ", "
			End If 
			response.write oWaiver("waivername") 
			oWaiver.movenext
			iRow = iRow + 1
		Loop 
	Else 
		response.write " No 3rd Party Waivers Assigned"
	End If 

	oWaiver.close
	Set oWaiver = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowRegistrationPicks( iOptionId )
'--------------------------------------------------------------------------------------------------
Sub ShowRegistrationPicks( iOptionId )
	Dim sSql, oOption

	sSql = "Select optionid, optionname, optiondescription, requirestime from egov_registration_option Order By optionid"

	Set oOption = Server.CreateObject("ADODB.Recordset")
	oOption.Open sSQL, Application("DSN"), 1, 1

	If Not oOption.EOF Then
		response.write vbcrlf & "<select name=""optionid"" >"
		'response.write vbcrlf & vbtab & "<option value=""0"">Select an Option</option>"
		Do While Not oOption.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oOption("optionid") & """ "
				If clng(oOption("optionid")) = clng(iOptionId) Then 
					response.write " selected=""selected"" "
				End If 
			response.write ">" & oOption("optionname") & " &ndash; " & oOption("optiondescription") & "</option>"
			oOption.movenext
		Loop 
		response.write vbcrlf & "</select>"

		' Put in the hidden fields to check for time requirements
		oOption.movefirst
		Do While Not oOption.EOF
			response.write vbcrlf & "<input type=""hidden"" name=""requirestime" & oOption("optionid") & """ value=""" & oOption("requirestime") & """ />"
			oOption.movenext
		Loop 

	End If 

	oOption.close
	Set oOption = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function CheckDayOfWeek(iClassId, iDayOfWeek)
'--------------------------------------------------------------------------------------------------
Function CheckDayOfWeek(iClassId, iDayOfWeek)
	Dim sSql, oDOW

	sSql = "Select rowid from egov_class_dayofweek where dayofweek = " & iDayOfWeek & " and classid = " & iClassId

	Set oDOW = Server.CreateObject("ADODB.Recordset")
	oDOW.Open sSQL, Application("DSN"), 0, 1

	If Not oDOW.EOF Then
		CheckDayOfWeek = " checked=""checked"" "
	Else 
		CheckDayOfWeek = ""
	End If 

	oDOW.close
	Set oDOW = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function CheckClassInCategory( iClassId, iCategoryId )
'--------------------------------------------------------------------------------------------------
Function CheckClassInCategory( iClassId, iCategoryId )
	Dim sSql, oCat

	sSql = "Select rowid from egov_class_category_to_class where categoryid = " & iCategoryId & " and classid = " & iClassId

	Set oCat = Server.CreateObject("ADODB.Recordset")
	oCat.Open sSQL, Application("DSN"), 0, 1

	If Not oCat.EOF Then
		CheckClassInCategory = " checked=""checked"" "
	Else 
		CheckClassInCategory = ""
	End If 

	oCat.close
	Set oCat = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowCategories( iClassId, iOrgId )
'--------------------------------------------------------------------------------------------------
Sub ShowCategories( iClassId, iOrgId )
	Dim sSql, oCategory, iRow

	iRow = 0
	sSql = "Select categoryid, categorytitle from egov_class_categories where isroot = 0 and orgid = " & iOrgId & " Order By categorytitle"

	Set oCategory = Server.CreateObject("ADODB.Recordset")
	oCategory.Open sSQL, Application("DSN"), 0, 1

	Do While Not oCategory.EOF
		'<input type="checkbox" name="dayofweek" value="1" CheckDayOfWeek(iClassId, 1) /> Sunday
		response.write vbcrlf & "<input type=""checkbox"" name=""categoryid"" value=""" & oCategory("categoryid") & """ "
		response.write CheckClassInCategory( iClassId, oCategory("categoryid") )
		response.write " /> " & oCategory("categorytitle") & " <br />"
		oCategory.MoveNext 
	Loop 

	oCategory.Close
	Set oCategory = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowClassCategories( iClassId )
'--------------------------------------------------------------------------------------------------
Sub ShowClassCategories( iClassId )
	Dim sSql, oCategory, iRow

	iRow = 0
	sSql = "Select categorytitle from egov_class_categories C, egov_class_category_to_class G where C.categoryid = G.categoryid"
	sSql = sSql & " and classid = " & iClassId & " Order By categorytitle"

	Set oCategory = Server.CreateObject("ADODB.Recordset")
	oCategory.Open sSQL, Application("DSN"), 0, 1

	Do While Not oCategory.EOF
		If iRow > 0 Then
			response.write ", "
		End If 
		response.write oCategory("categorytitle") 
		oCategory.movenext
		iRow = iRow + 1
	Loop 

	oCategory.close
	Set oCategory = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowPriceDiscountPicks( iPriceDiscountId, iOrgId )
'--------------------------------------------------------------------------------------------------
Sub ShowPriceDiscountPicks( iPriceDiscountId, iOrgId )
	Dim sSql, oDiscounts

'	iDiscountId = GetClassPriceDiscount( iClassId )

	sSql = "Select pricediscountid, discountname from egov_price_discount Where orgid = " & iOrgId & " Order By discountname"

	Set oDiscounts = Server.CreateObject("ADODB.Recordset")
	oDiscounts.Open sSQL, Application("DSN"), 0, 1

	response.write vbcrlf & "<select name=""pricediscountid"" >"
	response.write vbcrlf & vbtab & "<option value=""0"">No Discount Applied</option>"
	Do While Not oDiscounts.EOF
		response.write vbcrlf & vbtab & "<option value=""" & oDiscounts("pricediscountid") & """ "
			If clng(oDiscounts("pricediscountid")) = clng(iPriceDiscountId) Then 
				response.write " selected=""selected"" "
			End If 
		response.write ">" & oDiscounts("discountname") & "</option>"
		oDiscounts.movenext
	Loop 
	response.write vbcrlf & "</select>"

	oDiscounts.close
	Set oDiscounts = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowClassTimesold( iClassId )
'--------------------------------------------------------------------------------------------------
Sub ShowClassTimesold( iClassId )
	Dim sSql, oTimes, iRow, sMin, sMax, sWaitlistmax

	iRow = 0
	sSql = "Select timeid, starttime, endtime, isnull(min,0) as min, isnull(max, 0) as max, isnull(waitlistmax,0) as waitlistmax from egov_class_time Where classid = " & iClassId & " Order By timeid"

	Set oTimes = Server.CreateObject("ADODB.Recordset")
	oTimes.Open sSQL, Application("DSN"), 0, 1

	Do While Not oTimes.EOF
		iRow = iRow + 1
		If clng(oTimes("min")) = 0 Then
			sMin = ""
		Else
			sMin = oTimes("min")
		End If 
		If clng(oTimes("max")) = 0 Then
			sMax = ""
		Else
			sMax = oTimes("max")
		End If 
		If clng(oTimes("waitlistmax")) = 0 Then
			sWaitlistmax = ""
		Else
			sWaitlistmax = oTimes("waitlistmax")
		End If 
		response.write vbcrlf & "<tr>"
		response.write "<input type=""hidden"" name=""timeid" & iRow & """ value=""" & oTimes("timeid") & """ />"
		response.write "<td align=""center""><input type=""checkbox"" name=""removetime" & iRow & """ /></td>"
		response.write "<td><input type=""text"" name=""starttime" & iRow & """ value=""" & oTimes("starttime") & """ size=""8"" maxlength=""7"" /></td>"
		response.write "<td><input type=""text"" name=""endtime" & iRow & """ value=""" & oTimes("endtime") & """ size=""8"" maxlength=""7"" /></td>"
		response.write "<td><input type=""text"" name=""min" & iRow & """ value=""" & sMin & """ size=""4"" maxlength=""5"" /></td>"
		response.write "<td><input type=""text"" name=""max" & iRow & """ value=""" &sMax & """ size=""4"" maxlength=""5"" /></td>"
		response.write "<td><input type=""text"" name=""waitlistmax" & iRow & """ value=""" & sWaitlistmax & """ size=""4"" maxlength=""5"" /></td>"
		response.write vbcrlf & "</tr>"
		oTimes.movenext
	Loop 

	oTimes.close
	Set oTimes = Nothing
	response.write vbcrlf & "<input type=""hidden"" name=""timecount"" value=""" & iRow & """ />"

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowAvailableClasses( )
'--------------------------------------------------------------------------------------------------
Sub ShowAvailableClasses( )
	Dim sSql, oList

	sSql = "Select classid, classname from egov_class "
	sSql = sSql & " where classtypeid = 3 and orgid = " & Session("OrgID") & " Order By classname"
	'response.write sSql & "<br />"

	Set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSQL, Application("DSN"), 0, 1

	response.write vbcrlf & "<select border=""0"" name=""classid"">"
	Do While Not oList.EOF
		response.write "<option value=""" & oList("classid") & """>" & oList("classname") & "</option>"
		oList.movenext
	Loop 
	response.write vbcrlf & "</select>"

	oList.close
	Set oList = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function ShowEarlyRegistrationClassSeasons( iClassSeasonId )
'--------------------------------------------------------------------------------------------------
Function ShowEarlyRegistrationClassSeasons( iClassSeasonId )
	Dim sSql, oSeasons, iSelectedSeasonId

	iSelectedSeasonId = CLng(iClassSeasonId)

	sSQL = "SELECT C.classseasonid, C.seasonname FROM egov_class_seasons C, egov_seasons S  "
	sSql = sSql & " WHERE C.seasonid = S.seasonid AND orgid = " & SESSION("orgid") & " ORDER BY C.seasonyear desc, S.displayorder DESC, C.seasonname"

	Set oSeasons = Server.CreateObject("ADODB.Recordset")
	oSeasons.Open sSQL, Application("DSN"), 0, 1
	
	If Not oSeasons.EOF Then
		Do While NOT oSeasons.EOF
			If iSelectedSeasonId = CLng(0) Then
				iSelectedSeasonId = CLng(oSeasons("classseasonid"))
			End If 
			response.write vbcrlf & "<option value=""" & oSeasons("classseasonid") & """"  
			If CLng(iClassSeasonId) = CLng(oSeasons("classseasonid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oSeasons("seasonname") & "</option>"
			oSeasons.MoveNext
		Loop
	End If

	oSeasons.Close
	Set oSeasons = Nothing

	ShowEarlyRegistrationClassSeasons = iSelectedSeasonId

End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowEarlRegistrationClasses( iClassSeasonId, iClassId )
'--------------------------------------------------------------------------------------------------
Sub ShowEarlRegistrationClasses( iClassSeasonId, iClassId )
	Dim sSql, oRs

	sSQL = "SELECT C.classid, C.classname "
	sSql = sSql & " FROM egov_class C, egov_class_status S, egov_registration_option RO "
	sSql = sSql & " WHERE C.statusid = S.statusid AND S.statusname = 'ACTIVE' AND C.classseasonid = " & iClassSeasonId
	sSql = sSql & " AND RO.optionid = C.optionid AND RO.canpurchase = 1 "
	sSql = sSql & " AND C.orgid = " & SESSION("orgid") & " ORDER BY C.classname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write "<select id=""earlyregistrationclassid"" name=""earlyregistrationclassid"" multiple=""multiple"" size=""20"">"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("classid") & """"
			If ClassIsInEarlyRegistration( iClassId, CLng(oRs("classid")) ) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("classname") & "</option>"
			oRs.MoveNext
		Loop
		response.write "</select>"
	End If 
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Function ClassIsInEarlyRegistration( iClassId, iPotentialClassId )
'--------------------------------------------------------------------------------------------------
Function ClassIsInEarlyRegistration( iClassId, iPotentialClassId )
	Dim sSql, oRs

	sSQL = "SELECT COUNT(earlyregistrationclassid) AS hits FROM egov_class_earlyregistrations "
	sSQl = sSql & " WHERE classid = " & iClassId & " AND earlyregistrationclassid = " & iPotentialClassId

	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then
			ClassIsInEarlyRegistration = True 
		Else
			ClassIsInEarlyRegistration = False 
		End If 
	Else
		ClassIsInEarlyRegistration = False 
	End If 
	
	oRs.Close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetCountOfRegattaPrices( )
'--------------------------------------------------------------------------------------------------
Function GetCountOfRegattaPrices( )
	Dim sSql, oRs

	sSql = "SELECT COUNT(pricetypeid) AS hits "
	sSql = sSql & " FROM egov_price_types "
	sSql = sSql & " WHERE isregattaprice = 1 AND orgid = " & Session("OrgID")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then
			GetCountOfRegattaPrices = CLng(oRs("hits")) 
		Else
			GetCountOfRegattaPrices = 0 
		End If 
	Else
		GetCountOfRegattaPrices = 0 
	End If 
	
	oRs.Close
	Set oRs = Nothing 
End Function



%>
