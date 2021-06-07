<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitcontacttypeedit.asp
' AUTHOR: Steve Loar
' CREATED: 01/23/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates and edits permit contact type information.
'
' MODIFICATION HISTORY
' 1.0   01/23/2008	Steve Loar - INITIAL VERSION
' 1.1	06/05/2008	Steve Loar - License Date dropped, Licensee addded
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitContactTypeid, oRs, sSql, sDisabled, sTitle, sFirstName, sLastName, sCompany, sCity
Dim sAddress, sState, sZip, sEmail, sPhone, sCell, sFax, sIsArchitect, sIsContractor, sIsOwner
Dim iMaxLicenseRows, iUserId, sSearchName, iActiveTabId, iContractorTypeId, iMaxUsers
Dim iIsOrganization, sBackUrl, iBusinessTypeId, sStateLicense, sEmployeeCount, sReference1
Dim sReference2, sReference3, sOtherLicensedCity1, sOtherLicensedCity2, sGeneralLiabilityAgent
Dim sGeneralLiabilityPhone, sWorkersCompAgent, sWorkersCompPhone, sAutoInsuranceAgent
Dim sAutoInsurancePhone, sBondAgent, sBondAgentPhone

sLevel = "../" ' Override of value from common.asp

PageDisplayCheck "permit contacts", sLevel	' In common.asp

If request("activetab") <> "" Then 
	If IsNumeric(request("activetab")) Then 
		iActiveTabId = clng(request("activetab"))
	Else
		iActiveTabId = clng(0)
	End If 
Else
	iActiveTabId = clng(0)
End If 

iMaxUsers = CLng(0)

' GET contact ID
If CLng(request("permitcontacttypeid")) = 0 Then
	' CREATE NEW contact
	iPermitContactTypeid = 0
	sTitle = "New"
	iIsOrganization = clng(request("isorganization"))
Else
	' EDIT EXISTING contact
	iPermitContactTypeid = request("permitcontacttypeid")
	sTitle = "Edit"
	sDisabled = GetDisabledText( iPermitContactTypeid )
	iIsOrganization = GetIsOrganizationFlag( iPermitContactTypeid )
End If

If clng(iIsOrganization) = clng(0) Then
	sTitle = sTitle & " Permit Contractor"
	sBackUrl = "permitcontactlist.asp"
Else
	sTitle = sTitle & " Permit Organization"
	sBackUrl = "permitorganizationlist.asp"
End If 


sSQL = "SELECT ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, ISNULL(company,'') AS company, "
sSql = sSql & " address, city, state, zip, email, phone, fax, cell, ISNULL(userid,0) AS userid, "
sSql = sSql & " ISNULL(contractortypeid,0) AS contractortypeid, ISNULL(businesstypeid,0) AS businesstypeid, "
sSql = sSql & " ISNULL(statelicense,'') AS statelicense, ISNULL(employeecount,'') AS employeecount, ISNULL(reference1,'') AS reference1, "
sSql = sSql & " ISNULL(reference2,'') AS reference2, ISNULL(reference3,'') AS reference3, ISNULL(otherlicensedcity1,'') AS otherlicensedcity1, "
sSql = sSql & " ISNULL(otherlicensedcity2,'') AS otherlicensedcity2, ISNULL(generalliabilityagent,'') AS generalliabilityagent, ISNULL(generalliabilityphone,'') AS generalliabilityphone, "
sSql = sSql & " ISNULL(workerscompagent,'') AS workerscompagent, ISNULL(workerscompphone,'') AS workerscompphone, ISNULL(autoinsuranceagent,'') AS autoinsuranceagent, "
sSql = sSql & " ISNULL(autoinsurancephone,'') AS autoinsurancephone, ISNULL(bondagent,'') AS bondagent, ISNULL(bondagentphone,'') AS bondagentphone "
sSql = sSql & " FROM egov_permitcontacttypes WHERE orgid = " & session("orgid") & " AND permitcontacttypeid = " & iPermitContactTypeid 

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 3, 1

If NOT oRs.EOF Then
	sFirstName = oRs("firstname")
	sLastName = oRs("lastname")
	sCompany = Replace(oRs("company"),Chr(34),"&quot;")
	sAddress = oRs("address")
	sCity = oRs("city")
	sState = oRs("state")
	sZip = oRs("zip")
	sEmail = oRs("email")
	sPhone = oRs("phone")
	sCell = oRs("cell")
	sFax = oRs("fax")
	iUserId = oRs("userid")
	iContractorTypeId = clng(oRs("contractortypeid"))
	iBusinessTypeId = clng(oRs("businesstypeid"))
	sStateLicense = oRs("statelicense")
	sEmployeeCount = oRs("employeecount")
	sReference1 = oRs("reference1")
	sReference2 = oRs("reference2")
	sReference3 = oRs("reference3")
	sOtherLicensedCity1 = oRs("otherlicensedcity1")
	sOtherLicensedCity2 = oRs("otherlicensedcity2")
	sGeneralLiabilityAgent = oRs("generalliabilityagent")
	sGeneralLiabilityPhone = oRs("generalliabilityphone")
	sWorkersCompAgent = oRs("workerscompagent")
	sWorkersCompPhone = oRs("workerscompphone")
	sAutoInsuranceAgent = oRs("autoinsuranceagent")
	sAutoInsurancePhone = oRs("autoinsurancephone")
	sBondAgent = oRs("bondagent")
	sBondAgentPhone = oRs("bondagentphone")
Else
	sFirstName = ""
	sLastName = ""
	sCompany = ""
	sAddress = ""
	sCity = ""
	sState = ""
	sZip = ""
	sEmail = "" 
	sPhone = ""
	sCell = ""
	sFax = ""
	iUserId = 0
	iContractorTypeId = clng(0)
	iBusinessTypeId = clng(0)
	sStateLicense = ""
	sEmployeeCount = ""
	sReference1 = ""
	sReference2 = ""
	sReference3 = ""
	sOtherLicensedCity1 = ""
	sOtherLicensedCity2 = ""
	sGeneralLiabilityAgent = ""
	sGeneralLiabilityPhone = ""
	sWorkersCompAgent = ""
	sWorkersCompPhone = ""
	sAutoInsuranceAgent = ""
	sAutoInsurancePhone = ""
	sBondAgent = ""
	sBondAgentPhone = ""
End If

oRs.Close
Set oRs = Nothing 

%>


<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../yui/build/tabview/assets/skins/sam/tabview.css" />
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<!--
	<script type="text/javascript" src="../yui/build/yahoo-dom-event/yahoo-dom-event.js"></script>
	<script type="text/javascript" src="../yui/build/element/element-beta.js"></script>
	<script type="text/javascript" src="../yui/build/tabview/tabview.js"></script>
	-->
	<script type="text/javascript" src="../yui/yahoo-dom-event.js"></script>  
	<script type="text/javascript" src="../yui/element-min.js"></script>  
	<script type="text/javascript" src="../yui/tabview-min.js"></script>

	<script language="JavaScript" src="../scripts/isvaliddate.js"></script>
	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
	<script language="JavaScript" src="../scriptaculous/src/scriptaculous.js"></script>
  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

	<script language="Javascript">
	<!--
		var tabView;
		var winHandle;
		var w = (screen.width - 640)/2;
		var h = (screen.height - 480)/2;

		(function() {
			tabView = new YAHOO.widget.TabView('demo');
			tabView.set('activeIndex', <%=iActiveTabId%>); 

		})();

		var isNN = (navigator.appName.indexOf("Netscape")!=-1);

		function Another()
		{
			location.href="permitcontacttypeedit.asp?permitcontacttypeid=0";
		}

		function SelectUser( )
		{
			//winHandle = eval('window.open("contactuserpicker.asp?permitcontacttypeid=<%=iPermitContactTypeid%>", "_contact", "width=800,height=500,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('contactuserpicker.asp?permitcontacttypeid=<%=iPermitContactTypeid%>', 'Registered User Selection', 30, 30);
		}

		function RemoveUserRows()
		{
			if (confirm("Remove the selected users?"))
			{
				var tbl = document.getElementById("contractoruserlist");
				var iMaxUsers = parseInt($("#maxusers").val());
				var iTableRow = 1;

				// Check the User rows for any selected for removal
				for (var t = 1; t <= iMaxUsers; t++)
				{
					// See if a row exists for this one
					if ($("#removeuser" + t).length > 0)
					{
						// If it is marked for removal, remove it
						if ($("#removeuser" + t).is(":checked"))
						{
							// Fire off an Ajax Job to remove them
							//alert($("removeuser" + t).value);
							doAjax('removepermitcontactuser.asp', 'permitcontacttypeid=<%=iPermitContactTypeid%>&userid=' + $("#removeuser" + t).val(), '', 'get', '0');
							tbl.deleteRow(iTableRow);
							//alert('Deleted: ' + iTableRow);
						}
						else
						{
							iTableRow = iTableRow + 1;
						}
					}
				}
			}
		}

		function autoTab(input,len, e) 
		{
			var keyCode = (isNN) ? e.which : e.keyCode; 
			var filter = (isNN) ? [0,8,9] : [0,8,9,16,17,18,37,38,39,40,46];

			if(input.value.length >= len && !containsElement(filter,keyCode)) {
				input.value = input.value.slice(0, len);
			var addNdx = 1;

			while(input.form[(getIndex(input)+addNdx) % input.form.length].type == "hidden") 
			{
				addNdx++;
				//alert(input.form[(getIndex(input)+addNdx) % input.form.length].type);
			}

			input.form[(getIndex(input)+addNdx) % input.form.length].focus();
		}

		function containsElement(arr, ele) 
		{
			var found = false, index = 0;

			while(!found && index < arr.length)
				if(arr[index] == ele)
					found = true;
				else
					index++;
			return found;
		}

		function getIndex(input) 
		{
			var index = -1, i = 0, found = false;

			while (i < input.form.length && index == -1)
				if (input.form[i] == input)index = i;
				else i++;
					return index;
		}
			return true;
		}

		function NewLicenseRow()
		{
			document.frmContact.maxlicenserows.value = parseInt(document.frmContact.maxlicenserows.value) + 1;
			var tbl = document.getElementById("licensetable");
			var lastRow = tbl.rows.length;
			var newRow = parseInt(document.frmContact.maxlicenserows.value);
			var row = tbl.insertRow(lastRow);

			// Remove Row checkbox
			var cellZero = row.insertCell(0);
			cellZero.className = 'firstcell';
			var e = document.createElement('input');
			e.type = 'checkbox';
			e.name = 'removelicense' + newRow;
			e.id = 'removelicense' + newRow;
			cellZero.appendChild(e);

			// Number text
			cellOne = row.insertCell(1);
			cellOne.align = 'center';
			e1 = document.createElement('input');
			e1.type = 'text';
			e1.name = 'licensenumber' + newRow;
			e1.id = 'licensenumber' + newRow;
			e1.size = '10';
			e1.maxLength = '20';
			cellOne.appendChild(e1);

			// Class text
			cellOne = row.insertCell(2);
			cellOne.align = 'center';
			e1 = document.createElement('input');
			e1.type = 'text';
			e1.name = 'licenseclass' + newRow;
			e1.id = 'licenseclass' + newRow;
			e1.size = '20';
			e1.maxLength = '25';
			cellOne.appendChild(e1);

			// Type text
			//var cellOne = row.insertCell(3);
			//cellOne.align = 'center';
			//var e1 = document.createElement('input');
			//e1.type = 'text';
			//e1.name = 'licensetype' + newRow;
			//e1.id = 'licensetype' + newRow;
			//e1.size = '25';
			//e1.maxLength = '25';
			//cellOne.appendChild(e1);

			// Type Dropdown
			// Find the first row that exists
			for (var t = 0; t <= parseInt(document.frmContact.maxlicenserows.value); t++ )
			{
				if (document.getElementById("licensetypeid" + t))
				{
					break;
				}
			}
			var cellOne = row.insertCell(3);
			cellOne.align = 'center';
			e1 = document.createElement('select');
			e1.name = 'licensetypeid' + newRow;
			e1.id = 'licensetypeid' + newRow;
			cellOne.appendChild(e1);
			slength = document.getElementById("licensetypeid" + t).length;
			for ( s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("licensetypeid" + t).options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("licensetypeid" + t).options[s].value );
				e1.appendChild(op);
			}

			// Licensee text
			cellOne = row.insertCell(4);
			cellOne.align = 'center';
			e1 = document.createElement('input');
			e1.type = 'text';
			e1.name = 'licensee' + newRow;
			e1.id = 'licensee' + newRow;
			e1.size = '30';
			e1.maxLength = '100';
			cellOne.appendChild(e1);

			// License End date
			cellOne = row.insertCell(5);
			cellOne.align = 'center';
			e1 = document.createElement('input');
			e1.type = 'text';
			e1.name = 'licenseenddate' + newRow;
			e1.id = 'licenseenddate' + newRow;
			e1.size = '10';
			e1.maxLength = '10';
			cellOne.appendChild(e1);

		}

		function RemoveLicenseRows()
		{
			var iRow = 0;
			var tbl = document.getElementById("licensetable");
			var iMaxLicenseRows = document.frmContact.maxlicenserows.value;
			// Check the License rows for any selected for removal
			for (var t = 0; t <= parseInt(iMaxLicenseRows); t++)
			{
				// See if a row exists for this one
				if (document.getElementById("removelicense" + t))
				{
					// The row exists so increment the row counter
					iRow++;
					// If it is marked for removal, remove it
					if (document.getElementById("removelicense" + t).checked == true)
					{
						if (tbl.rows.length > 2)
						{
							// Remove the unwanted row
							tbl.deleteRow(iRow);
							// Decrement the row counter as we have one less row now
							iRow--;
							// Decrement the maxlicenserows
							document.frmContact.maxlicenserows.value = parseInt(document.frmContact.maxlicenserows.value) - 1;
						}
						else
						{
							// Down to one row, so just reset it to it's initial defaults
							document.getElementById("removelicense" + t).checked = false;
							document.getElementById("licensenumber" + t).value= '';
							document.getElementById("licenseclass" + t).value= '';
							//document.getElementById("licensetype" + t).value= '';
							document.getElementById("licensetypeid" + t).options[0].selected = true;
							document.getElementById("licensee" + t).value= '';
							document.getElementById("licenseenddate" + t).value= '';
						}
					}
				}
			}
		}

		function SearchCitizens( iSearchStart )
		{
			var optiontext;
			var optionchanged;
			//alert(document.BuyerForm.searchname.value);
			var searchtext = document.frmContact.searchname.value;
			var searchchanged = searchtext.toLowerCase();

			iSearchStart = parseInt(iSearchStart) + 1;
			
			for (x=iSearchStart; x < document.frmContact.userid.length ; x++)
			{
				optiontext = document.frmContact.userid.options[x].text;
				optionchanged = optiontext.toLowerCase();
				if (optionchanged.indexOf(searchchanged) != -1)
				{
					document.frmContact.userid.selectedIndex = x;
					document.frmContact.results.value = 'Possible Match Found.';
					document.getElementById('searchresults').innerHTML = 'Possible Match Found.';
					document.frmContact.searchstart.value = x;
					return;
				}
			}
			document.frmContact.results.value = 'No Match Found.';
			document.getElementById('searchresults').innerHTML = 'No Match Found.';
			document.frmContact.searchstart.value = -1;
		}

		function ClearSearch()
		{
			document.frmContact.searchstart.value = -1;
		}

		function UserPick()
		{
			document.frmContact.searchname.value = '';
			document.frmContact.results.value = '';
			document.getElementById('searchresults').innerHTML = '';
			document.frmContact.searchstart.value = -1;
		}

		function Delete() 
		{
			if (confirm("Do you wish to delete this contact?"))
			{
				location.href="permitcontacttypedelete.asp?permitcontacttypeid=<%=iPermitContactTypeid%>";
			}
		}

		function Validate()
		{
			document.getElementById("activetab").value = tabView.get("activeIndex");

			var rege; 
			var Ok;
			var sPhone;
			//var daterege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;

			// Check for a contact name or company
			if (document.frmContact.firstname.value == '' && document.frmContact.lastname.value == '' && document.frmContact.company.value == '')
			{
				alert('Please provide either a first name and last name or a company name.\nThen try saving again.');
				document.frmContact.firstname.focus();
				return;
			}
			if (document.frmContact.firstname.value == '' && document.frmContact.lastname.value != '')
			{
				alert('Please provide either a first name for this contact.\nThen try saving again.');
				document.frmContact.firstname.focus();
				return;
			}
			if (document.frmContact.lastname.value == '' && document.frmContact.firstname.value != '')
			{
				alert('Please provide either a lastname name for this contact.\nThen try saving again.');
				document.frmContact.lastname.focus();
				return;
			}

			// Check the email
			if (document.frmContact.email.value != "" )
			{
				//rege = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*\.(\w{2}|(com|net|org|edu|mil|gov|biz|us))$/;
				rege = /.+@.+\..+/i;
				Ok = rege.test(document.frmContact.email.value);
				if (! Ok)
				{
					tabView.set('activeIndex',0);
					alert("The email must be in a valid format or blank.\nPlease correct this and try saving again.");
					document.frmContact.email.focus();
					return;
				}
			}

			// check the phone
			if (document.frmContact.phone1.value != "" || document.frmContact.phone2.value != "" || document.frmContact.phone3.value != "")
			{
				sPhone = document.frmContact.phone1.value + document.frmContact.phone2.value + document.frmContact.phone3.value;
				if (sPhone.length < 10)
				{
					tabView.set('activeIndex',0);
					alert("The phone number must be a valid phone number, including area code, or blank.\nPlease correct this and try saving again.");
					document.frmContact.phone1.focus();
					return;
				}
				else
				{
					document.frmContact.phone.value = document.frmContact.phone1.value + document.frmContact.phone2.value + document.frmContact.phone3.value;
					var rege = /^\d+$/;
					var Ok = rege.exec(document.frmContact.phone.value);
					if ( ! Ok )
					{
						tabView.set('activeIndex',0);
						alert("The phone number must be a valid phone number, including area code, or blank.\nPlease correct this and try saving again.");
						document.frmContact.phone1.focus();
						return;
					}
				}
			}
			else
			{
				document.frmContact.phone.value = '';
			}

			// check the cell
			if (document.frmContact.cell1.value != "" || document.frmContact.cell2.value != "" || document.frmContact.cell3.value != "")
			{
				sPhone = document.frmContact.cell1.value + document.frmContact.cell2.value + document.frmContact.cell3.value;
				if (sPhone.length < 10)
				{
					tabView.set('activeIndex',0);
					alert("The cell number must be a valid phone number, including area code, or blank.\nPlease correct this and try saving again.");
					document.frmContact.cell1.focus();
					return;
				}
				else
				{
					document.frmContact.cell.value = document.frmContact.cell1.value + document.frmContact.cell2.value + document.frmContact.cell3.value;
					var rege = /^\d+$/;
					var Ok = rege.exec(document.frmContact.cell.value);
					if ( ! Ok )
					{
						tabView.set('activeIndex',0);
						alert("The cell number must be a valid phone number, including area code, or blank.\nPlease correct this and try saving again.");
						document.frmContact.cell1.focus();
						return;
					}
				}
			}
			else
			{
				document.frmContact.cell.value = '';
			}

			// check the fax
			if (document.frmContact.fax1.value != "" || document.frmContact.fax2.value != "" || document.frmContact.fax3.value != "")
			{
				sPhone = document.frmContact.fax1.value + document.frmContact.fax2.value + document.frmContact.fax3.value;
				if (sPhone.length < 10)
				{
					tabView.set('activeIndex',0);
					alert("The fax number must be a valid phone number, including area code, or blank.\nPlease correct this and try saving again.");
					document.frmContact.fax1.focus();
					return;
				}
				else
				{
					document.frmContact.fax.value = document.frmContact.fax1.value + document.frmContact.fax2.value + document.frmContact.fax3.value;
					var rege = /^\d+$/;
					var Ok = rege.exec(document.frmContact.fax.value);
					if ( ! Ok )
					{
						tabView.set('activeIndex',0);
						alert("The fax number must be a valid phone number, including area code, or blank.\nPlease correct this and try saving again.");
						document.frmContact.fax1.focus();
						return;
					}
				}
			}
			else
			{
				document.frmContact.fax.value = '';
			}

			// Check the licenses - want a type if there is anything else on the row
			for (var t = 0; t <= parseInt(document.frmContact.maxlicenserows.value); t++)
			{
				if (document.getElementById("licenseenddate" + t))
				{
					// Validate the format of the license end date 
					if (document.getElementById("licenseenddate" + t).value != "")
					{
						if (isValidDate(document.getElementById("licenseenddate" + t).value) == false)
						{
							tabView.set('activeIndex',2);
							alert("Invalid license expiration date value or format.\nIt should be a valid date in the format of MM/DD/YYYY.");
							document.getElementById("licenseenddate" + t).focus();
							return;
						}
					}
					else
					{
						if (document.getElementById("licensenumber" + t).value != "")
						{
							tabView.set('activeIndex',2);
							alert("All licenses require an expiration date.\nIt should be a valid date in the format of MM/DD/YYYY.");
							document.getElementById("licenseenddate" + t).focus();
							return;
						}
					}
				}
			}

			// All is ok so submit
			//alert("OK");
			document.frmContact.submit();
		}

		function doCalendar( sField ) 
		{
		  var w = (screen.width - 350)/2;
		  var h = (screen.height - 350)/2;
		  eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=frmContact", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

  $( function() {
    $( ".datepicker" ).datepicker({
      changeMonth: true,
      showOn: "both",
      buttonText: "<i class=\"fa fa-calendar\"></i>",
      changeYear: true
    });
  } );
<%		If request("success") <> "" Then 
			DisplayMessagePopUp 
		End If 
%>

	//-->
	</script>

</head>

<body class="yui-skin-sam">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	<div class="gutters">
	
<!--BEGIN: PAGE TITLE-->
<p>
	<font size="+1"><strong><%=sTitle%></strong></font><br /><br />
	<a href="<%=sBackUrl%>"><img src="../images/arrow_2back.gif" align="absmiddle" border="0" />&nbsp;<%=langBackToStart%></a>
</p>
<!--END: PAGE TITLE-->


<!--BEGIN: EDIT FORM-->
		<div id="functionlinks">
<%		If CLng(iPermitContactTypeid) = CLng(0) Then %>
			<input type="button" class="button ui-button ui-widget ui-corner-all" id="savebutton" onclick="javascript:Validate();" value="Create" /><br />
<%		Else %>
			<input type="button" class="button ui-button ui-widget ui-corner-all" id="savebutton" onclick="javascript:Validate();" value="Save Changes" /> &nbsp; &nbsp;
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Delete();" value="Delete" <%=sDisabled%> /> &nbsp; &nbsp;
<%			If request("success") <> "" Then %>
				<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Another();" value="Create Another" />
<%			End If		%>
			<br />
<%		End If %>
		</div>

<form name="frmContact" action="permitcontacttypeupdate.asp" method="post">
	<input type="hidden" name="permitcontacttypeid" value="<%=iPermitContactTypeid%>" />
	<input type="hidden" name="isorganization" value="<%=iIsOrganization%>" />
<%	If clng(iIsOrganization) = clng(1) Then %>
	<input type="hidden" name="maxlicenserows" value="0" />
<%	End If		%>

	<input type="hidden" name="activetab" id="activetab" value="<%=iActiveTabId%>" />

	<p>
		First Name: &nbsp; <input type="text" name="firstname" value="<%=sFirstName%>" size="25" maxlength="25" />
		&nbsp;&nbsp;
		Last Name: &nbsp; <input type="text" name="lastname" value="<%=sLastName%>" size="25" maxlength="25" />
	</p>
	<p>
		Company: &nbsp; <input type="text" name="company" value="<%=sCompany%>" size="50" maxlength="50" />
	</p>
	<% If iIsOrganization = clng(0) Then %>
	<p>
		Contractor Type: &nbsp; <% ShowContractorTypes iContractorTypeId %>
	</p>
	<% End If %>

	<div id="demo" class="yui-navset">
		<ul class="yui-nav">
			<li><a href="#tab1"><em>Contractor Information</em></a></li>
			<li><a href="#tab2"><em>Associate With Registered Users</em></a></li>
			<% If clng(iIsOrganization) = clng(0) Then %>
			<li><a href="#tab3"><em>Licenses and Certifications</em></a></li>
			<% End If %>
		</ul>            
		<div class="yui-content">
			<div id="tab1"> <!-- Contractor Information -->
				<p><br />
				<table id="permitcontactinfo" cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td align="right" class="labelcolumn">Address:</td><td class="datacolumn"><input type="text" name="address" value="<%=sAddress%>" size="50" maxlength="50" /></td>
					</tr>
					<tr>
						<td align="right" class="labelcolumn">City:</td><td class="datacolumn"><input type="text" name="city" value="<%=sCity%>" size="25" maxlength="25" /></td>
					</tr>
					<tr>
						<td align="right" class="labelcolumn">State:</td><td class="datacolumn"><input type="text" name="state" value="<%=sState%>" size="2" maxlength="2" /></td>
					</tr>
					<tr>
						<td align="right" class="labelcolumn">Zip:</td><td class="datacolumn"><input type="text" name="zip" value="<%=sZip%>" size="10" maxlength="10" /></td>
					</tr>
					<tr>
						<td align="right" class="labelcolumn">Email:</td><td class="datacolumn"><input type="text" name="email" value="<%=sEmail%>" size="75" maxlength="100" /></td>
					</tr>
					<tr>
						<td align="right" class="labelcolumn">Phone:<input type="hidden" name="phone" value="<%=sPhone%>" /></td>
						<td class="datacolumn">(<input type="text" name="phone1" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3" value="<%=Mid(sPhone,1,3)%>">) <input value="<%=Mid(sPhone,4,3)%>" type="text" name="phone2" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3" /> <input value="<%=Mid(sPhone,7,4)%>" type="text" name="phone3" onKeyUp="return autoTab(this, 4, event);" size="4" maxlength="4" /></td>
					</tr>
					<tr>
						<td align="right" class="labelcolumn">Cell:<input type="hidden" name="cell" value="<%=sCell%>" /></td>
						<td class="datacolumn">(<input type="text" name="cell1" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3" value="<%=Mid(sCell,1,3)%>">) <input value="<%=Mid(sCell,4,3)%>" type="text" name="cell2" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3" /> <input value="<%=Mid(sCell,7,4)%>" type="text" name="cell3" onKeyUp="return autoTab(this, 4, event);" size="4" maxlength="4" /></td>
					</tr>
					<tr>
						<td align="right" class="labelcolumn">Fax:<input type="hidden" name="fax" value="<%=sFax%>" /></td>
						<td class="datacolumn">(<input type="text" name="fax1" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3" value="<%=Mid(sFax,1,3)%>">) <input value="<%=Mid(sFax,4,3)%>" type="text" name="fax2" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3" /> <input value="<%=Mid(sFax,7,4)%>" type="text" name="fax3" onKeyUp="return autoTab(this, 4, event);" size="4" maxlength="4" /></td>
					</tr>
				<% If clng(iIsOrganization) = clng(0) Then %>
						<tr>
							<td align="right" class="labelcolumn">Business Type:</td><td class="datacolumn"><% ShowBusinessTypes iBusinessTypeId %></td>
						</tr>
						<tr>
							<td align="right" class="labelcolumn">State License:</td><td class="datacolumn"><input type="text" name="statelicense" value="<%=sStateLicense%>" size="30" maxlength="30" /></td>
						</tr>
						<tr>
							<td align="right" class="labelcolumn" nowrap="nowrap">Number of Employees:</td><td class="datacolumn"><input type="text" name="employeecount" value="<%=sEmployeeCount%>" size="30" maxlength="30" /></td>
						</tr>
						<tr>
							<td align="right" class="labelcolumn" nowrap="nowrap">Mandatory References:</td><td class="datacolumn"><input type="text" name="reference1" value="<%=sReference1%>" size="30" maxlength="30" /></td>
						</tr>
						<tr>
							<td class="labelcolumn">&nbsp;</td><td class="datacolumn"><input type="text" name="reference2" value="<%=sReference2%>" size="30" maxlength="30" /></td>
						</tr>
						<tr>
							<td class="labelcolumn">&nbsp;</td><td class="datacolumn"><input type="text" name="reference3" value="<%=sReference3%>" size="30" maxlength="30" /></td>
						</tr>
						<tr>
							<td align="right" class="labelcolumn" nowrap="nowrap">Other Cities Licensed In:</td><td class="datacolumn"><input type="text" name="otherlicensedcity1" value="<%=sOtherLicensedCity1%>" size="30" maxlength="30" /></td>
						</tr>
						<tr>
							<td class="labelcolumn">&nbsp;</td><td class="datacolumn"><input type="text" name="otherlicensedcity2" value="<%=sOtherLicensedCity2%>" size="30" maxlength="30" /></td>
						</tr>
						<tr>
							<td colspan="2">Insurance Agents</td></td>
						</tr>
						<tr>
							<td align="right" class="labelcolumn">General Liability:</td><td class="datacolumn"><input type="text" name="generalliabilityagent" value="<%=sGeneralLiabilityAgent%>" size="30" maxlength="30" /> &nbsp; Phone: <input type="text" name="generalliabilityphone" value="<%=sGeneralLiabilityPhone%>" size="20" maxlength="20" /></td>
						</tr>
						<tr>
							<td align="right" class="labelcolumn">Worker's Compensation:</td><td class="datacolumn"><input type="text" name="workerscompagent" value="<%=sWorkersCompAgent%>" size="30" maxlength="30" /> &nbsp; Phone: <input type="text" name="workerscompphone" value="<%=sWorkersCompPhone%>" size="20" maxlength="20" /></td>
						</tr>
						<tr>
							<td align="right" class="labelcolumn">Auto Insurance:</td><td class="datacolumn"><input type="text" name="autoinsuranceagent" value="<%=sAutoInsuranceAgent%>" size="30" maxlength="30" /> &nbsp; Phone: <input type="text" name="autoinsurancephone" value="<%=sAutoInsurancePhone%>" size="20" maxlength="20" /></td>
						</tr>
						<tr>
							<td align="right" class="labelcolumn">Bond Agent:</td><td class="datacolumn"><input type="text" name="bondagent" value="<%=sBondAgent%>" size="30" maxlength="30" /> &nbsp; Phone: <input type="text" name="bondagentphone" value="<%=sBondAgentPhone%>" size="20" maxlength="20" /></td>
						</tr>
				<% End If %>
				</table>
				</p>
			</div>
			<div id="tab2"> <!-- Associate With Registered Users -->
					<p><br />
						<!-- Select Name: <select name="userid" onchange="javascript:UserPick();">
										<option value="0">Select a registered user from the list</option>
										<% 'ShowUserDropDown( iUserId )%>
									</select>
						&nbsp; <input class="button ui-button ui-widget ui-corner-all" type="button" value="Edit/View" onClick="location.href='../dirs/update_citizen.asp?userid=' + document.frmContact.userid.options[document.frmContact.userid.selectedIndex].value;" /> 
						&nbsp; <input onClick="location.href='../dirs/register_citizen.asp';" class="button ui-button ui-widget ui-corner-all" type="button" value="New User" />
						<br /><br />Name Search: <input type="text" name="searchname" value="<%'=sSearchName%>" size="25" maxlength="50" onchange="javascript:ClearSearch();" />
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Search" onclick="javascript:SearchCitizens(document.frmContact.searchstart.value);" />
						<input type="hidden" name="results" value="" /><input type="hidden" name="searchstart" value="<%'=sSearchStart%>" />
						<span id="searchresults"><%'=sResults%></span> -->

						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Add A User" onclick="SelectUser( );" /> &nbsp; <input type="button" class="button ui-button ui-widget ui-corner-all" value="Remove The Selected Users" onClick="RemoveUserRows()"  />
						<br /><br />
						<table cellpadding="0" cellspacing="0" border="0" id="contractoruserlist">
							<tr><th class="selectcol">&nbsp;</th><th>Name</th><th class="pickcol">Can Add<br />Others</th><th class="pickcol">Primary<br />Contact</th></tr>
<%							iMaxUsers = ShowContractorUsers( iPermitContactTypeid )		%>									
						</table>
						<input type="hidden" id="maxusers" name="maxusers" value="<%=iMaxUsers%>" />

						<br />								
					</p>
			</div>
			<% If clng(iIsOrganization) = clng(0) Then %>
			<div id="tab3"> <!-- Licenses and Certifications -->
				<p><br />
					<input type="button" class="button ui-button ui-widget ui-corner-all" value="Add Row" id="addlicensebutton" onClick="NewLicenseRow()" /> &nbsp;&nbsp; 
					<input type="button" class="button ui-button ui-widget ui-corner-all" value="Remove Selected" id="removelicensebutton" onClick="RemoveLicenseRows()" />
					<table id="licensetable" border="0" cellpadding="0" cellspacing="0">
						<tr><th>&nbsp;</th><th>Number</th><th>Class</th><th>Type</th><th>Licensee Name</th><th>Expires</th></tr>
<%													iMaxLicenseRows = ShowLicenseTable( iPermitContactTypeid ) %>
					</table>
					<input type="hidden" name="maxlicenserows" value="<%=iMaxLicenseRows%>" />
				</p>
			</div>
			<% End If %>
		</div>
	</div>

</form>
		<div id="functionlinks">
<%		If CLng(iPermitContactTypeid) = CLng(0) Then %>
			<input type="button" class="button ui-button ui-widget ui-corner-all" id="savebutton" onclick="javascript:Validate();" value="Create" /><br />
<%		Else %>
			<input type="button" class="button ui-button ui-widget ui-corner-all" id="savebutton" onclick="javascript:Validate();" value="Save Changes" /> &nbsp; &nbsp;
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Delete();" value="Delete" <%=sDisabled%> /> &nbsp; &nbsp;
<%			If request("success") <> "" Then %>
				<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Another();" value="Create Another" />
<%			End If		%>
			<br />
<%		End If %>
		</div>
<!--END: EDIT FORM-->

	</div>
	</div>
</div>

<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  
<!--#Include file="modal.asp"-->  

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
' Function ShowLicenseTable( iPermitContactTypeid )
'--------------------------------------------------------------------------------------------------
Function ShowLicenseTable( iPermitContactTypeid )
	Dim oRs, sSql, iRowCount, sRowClass

	iRowCount = -1

	sSql = "SELECT licensetype, licensenumber, licenseexpiration, licensee, licenseenddate, "
	sSql = sSql & " ISNULL(licenseclass,'') AS licenseclass, ISNULL(licensetypeid,0) AS licensetypeid "
	sSql = sSql & " FROM egov_permitcontacttype_licenses "
	sSql = sSql & " WHERE permitcontacttypeid = " & iPermitContactTypeid
	sSql = sSql & " ORDER BY licenseenddate DESC, licensenumber DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			iRowCount = iRowCount + 1
			If iRowCount Mod 2 = 0 Then 
				sRowClass = ""
			Else
				sRowClass = " class=""altrow"" "
			End If 
			response.write vbcrlf & "<tr" & sRowClass & "><td class=""firstcell"">"
			response.write "<input type=""checkbox"" id=""removelicense" & iRowCount & """ name=""removelicense" & iRowCount & """ /></td>"
			response.write "<td align=""center""><input type=""text"" id=""licensenumber" & iRowCount & """ name=""licensenumber" & iRowCount & """ value=""" & Replace(oRs("licensenumber"),Chr(34),"&quot;") & """ size=""10"" maxlength=""20"" /></td>"
			response.write "<td align=""center""><input type=""text"" id=""licenseclass" & iRowCount & """ name=""licenseclass" & iRowCount & """ value=""" & Replace(oRs("licenseclass"),Chr(34),"&quot;") & """ size=""20"" maxlength=""25"" /></td>"
			response.write "<td align=""center"">"
			ShowLicenseTypePicks oRs("licensetypeid"), iRowCount
			response.write "</td>"
			'response.write "<td align=""center""><input type=""text"" id=""licensetype" & iRowCount & """ name=""licensetype" & iRowCount & """ value=""" & Replace(oRs("licensetype"),Chr(34),"&quot;") & """ size=""25"" maxlength=""25"" /></td>"
			response.write "<td align=""center""><input type=""text"" id=""licensee" & iRowCount & """ name=""licensee" & iRowCount & """ value=""" & oRs("licensee") & """ size=""30"" maxlength=""100"" /></td>"
			response.write "<td align=""center""><input type=""text"" id=""licenseenddate" & iRowCount & """ name=""licenseenddate" & iRowCount & """ value=""" & FormatDateTime(oRs("licenseenddate"),2) & """ size=""10"" maxlength=""10"" class=""datepicker"" />"
			response.write "</td>"
			response.write "</tr>"
			oRs.MoveNext 
		Loop 
	Else
		' put in a starter row.
		iRowCount = 0
		response.write vbcrlf & "<tr><td class=""firstcell"">"
		response.write "<input type=""checkbox"" id=""removelicense" & iRowCount & """ name=""removelicense" & iRowCount & """ /></td>"
		response.write "<td align=""center""><input type=""text"" id=""licensenumber" & iRowCount & """ name=""licensenumber" & iRowCount & """ value="""" size=""20"" maxlength=""25"" /></td>"
		response.write "<td align=""center""><input type=""text"" id=""licenseclass" & iRowCount & """ name=""licenseclass" & iRowCount & """ value="""" size=""20"" maxlength=""25"" /></td>"
		response.write "<td align=""center"">"
		ShowLicenseTypePicks 0, iRowCount
		response.write "</td>"
		'response.write "<td align=""center""><input type=""text"" id=""licensetype" & iRowCount & """ name=""licensetype" & iRowCount & """ value="""" size=""25"" maxlength=""25"" /></td>"
		response.write "<td align=""center""><input type=""text"" id=""licensee" & iRowCount & """ name=""licensee" & iRowCount & """ value="""" size=""30"" maxlength=""100"" /></td>"
		response.write "<td align=""center""><input type=""text"" id=""licenseenddate" & iRowCount & """ name=""licenseenddate" & iRowCount & """ value="""" size=""10"" maxlength=""10"" class=""datepicker"" />"
		response.write "</td>"
		response.write "</tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 
	ShowLicenseTable = iRowCount
End Function


'--------------------------------------------------------------------------------------------------
' Function GetDisabledText( iPermitContactTypeId )
'--------------------------------------------------------------------------------------------------
Function GetDisabledText( iPermitContactTypeId )
	Dim sSql, oRs

	'If this contact is used, keep it from being deleted

	sSql = "SELECT COUNT(permitcontacttypeid) AS hits FROM egov_permitcontacts "
	sSql = sSql & " WHERE permitcontacttypeid = " & iPermitContactTypeId
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then 
			GetDisabledText = " disabled=""disabled"" " 
		Else
			GetDisabledText = "" 
		End If 
	Else
		GetDisabledText = "" 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'--------------------------------------------------------------------------------------------------
' Sub ShowUserDropDown( iUserId )
'--------------------------------------------------------------------------------------------------
Sub ShowUserDropDown( iUserId )
	Dim oCmd, oResident

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetEgovUserWithAddressList"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iOrgid", 3, 1, 4, Session("OrgID"))
	    Set oResident = .Execute
	End With

	Do While Not oResident.eof 
		response.write vbcrlf & "<option value=""" & oResident("userid") & """"
		If CLng(iUserId) = CLng(oResident("userid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oResident("userlname") & ", " & oResident("userfname") & " &ndash; " & oResident("useraddress") & "</option>"
		oResident.movenext
	Loop 
		
	oResident.close
	Set oResident = Nothing
	Set oCmd = Nothing
End Sub  


%>


