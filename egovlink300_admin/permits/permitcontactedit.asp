<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitcontactedit.asp
' AUTHOR: Steve Loar
' CREATED: 03/12/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Edits permit contacts
'
' MODIFICATION HISTORY
' 1.0   03/12/2008	Steve Loar - INITIAL VERSION
' 1.1	06/05/2008	Steve Loar - License Date dropped, Licensee addded
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitContactId, sSQL, oRs, sTitle, iUserId, sSearchName, sFirstName, sLastName, sCompany, sCity
Dim sAddress, sState, sZip, sEmail, sPhone, sCell, sFax, sSaveButton, sType, iPermitcontacttypeid
Dim iPermitId, iPermitStatusId, bCanSaveChanges, iActiveTabId, iContractorTypeId, bIsOtherContractor
Dim iMaxUsers, iIsOrganization, iBusinessTypeId, sStateLicense, sEmployeeCount, sReference1
Dim sReference2, sReference3, sOtherLicensedCity1, sOtherLicensedCity2, sGeneralLiabilityAgent
Dim sGeneralLiabilityPhone, sWorkersCompAgent, sWorkersCompPhone, sAutoInsuranceAgent
Dim sAutoInsurancePhone, sBondAgent, sBondAgentPhone

iPermitContactId = CLng(request("permitcontactid"))

If request("permitid") <> "" Then 
	iPermitId = CLng(request("permitid"))
	iPermitStatusId = CLng(request("permitstatusid"))
	If iPermitContactId <> CLng(0) Then 
		bCanSaveChanges = StatusAllowsSaveChanges( iPermitStatusId ) 	' in permitcommonfunctions.asp
		'response.write "iPermitStatusId = " & iPermitStatusId & ", bCanSaveChanges = " & bCanSaveChanges & "<br />"
	Else
		bCanSaveChanges = True 
	End If 
Else
	iPermitId = CLng(0)
	iPermitStatusId = CLng(0)
	bCanSaveChanges = True 
End If 
sType = request("type")
'response.write "<h1>" & sType & "</h1>"

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

If request("updatetitle") <> "" Then
	bUpdateParent = True 
Else 
	bUpdateParent = False 
End If 


sTitle = "New Contactor"
bIsOtherContractor = False 

sSQL = "SELECT ISNULL(permitcontacttypeid,0) AS permitcontacttypeid, ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, "
sSql = sSql & " ISNULL(company,'') AS company, ISNULL(address,'') AS address, ISNULL(city,'') AS city, ISNULL(state,'') AS state, "
sSql = sSql & " ISNULL(zip,'') AS zip, ISNULL(email,'') AS email, ISNULL(phone,'') AS phone, ISNULL(fax,'') AS fax, ISNULL(cell,'') AS cell, isorganization, "
sSql = sSql & " isapplicant, isprimarycontractor, isarchitect, iscontractor, isowner, isbillingcontact, ISNULL(userid,0) AS userid, "
sSql = sSql & " ISNULL(contractortypeid,0) AS contractortypeid, ISNULL(businesstypeid,0) AS businesstypeid, "
sSql = sSql & " ISNULL(statelicense,'') AS statelicense, ISNULL(employeecount,'') AS employeecount, ISNULL(reference1,'') AS reference1, "
sSql = sSql & " ISNULL(reference2,'') AS reference2, ISNULL(reference3,'') AS reference3, ISNULL(otherlicensedcity1,'') AS otherlicensedcity1, "
sSql = sSql & " ISNULL(otherlicensedcity2,'') AS otherlicensedcity2, ISNULL(generalliabilityagent,'') AS generalliabilityagent, ISNULL(generalliabilityphone,'') AS generalliabilityphone, "
sSql = sSql & " ISNULL(workerscompagent,'') AS workerscompagent, ISNULL(workerscompphone,'') AS workerscompphone, ISNULL(autoinsuranceagent,'') AS autoinsuranceagent, "
sSql = sSql & " ISNULL(autoinsurancephone,'') AS autoinsurancephone, ISNULL(bondagent,'') AS bondagent, ISNULL(bondagentphone,'') AS bondagentphone "
sSql = sSql & " FROM egov_permitcontacts WHERE orgid = "& session("orgid") & " AND permitcontactid = " & iPermitContactid 
'response.write sSql & "<br />"

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
	If oRs("isprimarycontractor") Then
		sTitle = "Primary Contractor"
	End If 
	If oRs("isarchitect") Then
		sTitle = "Architect/Engineer"
	End If 
	If oRs("iscontractor") Then
		sTitle = "Other Contractor"
		bIsOtherContractor = True 
	End If 
	If oRs("isowner") Then
		sTitle = "Owner"
	End If 
	If oRs("isapplicant") Then
		sTitle = "Applicant"
	End If 
	If oRs("isbillingcontact") Then
		sTitle = "Billing Contact"
	End If 
	iUserId = oRs("userid")
	sSaveButton = "Save Changes"
	iPermitcontacttypeid = CLng(oRs("permitcontacttypeid"))
	iContractorTypeId = clng(oRs("contractortypeid"))
	If oRs("isorganization") Then 
		iIsOrganization = 1
		'sTitle = "Organization"
	Else
		iIsOrganization = 0
	End If 
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
	sSaveButton = "Create"
	iPermitcontacttypeid = CLng(0)
	iContractorTypeId = clng(0)
	If request("isorganization") <> "" Then 
		iIsOrganization = clng(request("isorganization"))
		If iIsOrganization <> CLng(0) Then
			sTitle = "New Organization"
		End If 
	Else 
		iIsOrganization = CLng(0)
	End If 
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

oRs.close
Set oRs = Nothing 

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../yui/build/tabview/assets/skins/sam/tabview.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>
	<script type="text/javascript" src="../yui/yahoo-dom-event.js"></script>  
	<script type="text/javascript" src="../yui/element-min.js"></script>  
	<script type="text/javascript" src="../yui/tabview-min.js"></script>

	<script language="JavaScript" src="../scripts/isvaliddate.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

	<script language="Javascript">
	<!--
		parent.document.getElementById("modal"+window.frameElement.getAttribute("data-close")).style.width = "55%";
		parent.document.getElementById("modal"+window.frameElement.getAttribute("data-close")).style.height = "90%";
		parent.document.getElementById("modal"+window.frameElement.getAttribute("data-close")).style.left = "25%";
		parent.document.getElementById("modal"+window.frameElement.getAttribute("data-close")).style.top = "5%";
		var tabView;
		var winHandle;
		var w = (screen.width - 640)/2;
		var h = (screen.height - 480)/2;

		(function() {
			tabView = new YAHOO.widget.TabView('demo');
			tabView.set('activeIndex', <%=iActiveTabId%>); 

		})();

		var isNN = (navigator.appName.indexOf("Netscape")!=-1);

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
			var cellOne = row.insertCell(2);
			cellOne.align = 'center';
			var e1 = document.createElement('input');
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

			// License End Date
			cellOne = row.insertCell(5);
			cellOne.align = 'center';
			e1 = document.createElement('input');
			e1.type = 'text';
			e1.name = 'licenseenddate' + newRow;
			e1.id = 'licenseenddate' + newRow;
			e1.size = '10';
			e1.maxLength = '10';
			cellOne.appendChild(e1);

			var space = document.createTextNode( '\u00A0' );
			cellOne.appendChild(space);
  			$( "#licenseenddate" + newRow ).datepicker({ changeMonth: true, showOn: "both", buttonText: "<i class=\"fa fa-calendar\"></i>", changeYear: true });

/*
			// Add span
			var spanTag = document.createElement('span');
			spanTag.Class = "calendarimg";
			spanTag.style.cursor = "pointer";
			var newImage = document.createElement('img');
			newImage.id = "calendarimg" + newRow;
			newImage.src = "../images/calendar.gif";
			newImage.height = "16";
			newImage.width = "16";
			spanTag.appendChild(newImage);

			var onC='doCalendar("licenseenddate' + newRow +'")'; 
			newImage.onclick=new Function(onC); 
			//spanTag.innerHTML = '<img src="../images/calendar.gif" height="16" width="16" border="0" />';
			cellOne.appendChild(spanTag);
			//"&nbsp;<span class=""calendarimg"" style=""cursor:hand;""><img src=""../images/calendar.gif"" height=""16"" width=""16"" border=""0"" onclick=""javascript:void doCalendar('licenseenddate" & iRowCount & "');"" /></span>"
			*/

		}

		function RemoveLicenseRows()
		{
			var iRow = 0;
			var tbl = document.getElementById("licensetable");
			// Check the License rows for any selected for removal
			var iMaxLicenseRows = document.frmContact.maxlicenserows.value; 
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

		function doUpdate()
		{
			document.getElementById("activetab").value = tabView.get("activeIndex");

			// valicate the input data
			var rege; 
			var Ok;
			var sPhone;

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

			//alert('OK');
			//return;

			// build the parameter string
			//var sParameter = 'permitcontactid=' + encodeURIComponent(document.frmContact.permitcontactid.value);
			var sParameter = 'permitcontacttypeid=' + encodeURIComponent($("#permitcontacttypeid").val());
			sParameter += '&permitcontactid=' + encodeURIComponent($("#permitcontactid").val());
			sParameter += '&firstname=' + encodeURIComponent(document.frmContact.firstname.value);
			sParameter += '&lastname=' + encodeURIComponent(document.frmContact.lastname.value);
			sParameter += '&company=' + encodeURIComponent(document.frmContact.company.value);
			sParameter += '&address=' + encodeURIComponent(document.frmContact.address.value);
			sParameter += '&city=' + encodeURIComponent(document.frmContact.city.value);
			sParameter += '&state=' + encodeURIComponent(document.frmContact.state.value);
			sParameter += '&zip=' + encodeURIComponent(document.frmContact.zip.value);
			sParameter += '&email=' + encodeURIComponent(document.frmContact.email.value);
			sParameter += '&phone=' + encodeURIComponent(document.frmContact.phone.value);
			sParameter += '&cell=' + encodeURIComponent(document.frmContact.cell.value);
			sParameter += '&fax=' + encodeURIComponent(document.frmContact.fax.value);
			//sParameter += '&userid=' + encodeURIComponent(document.frmContact.userid.value);
			sParameter += '&maxusers=' + encodeURIComponent(document.frmContact.maxusers.value);
			sParameter += '&maxlicenserows=' + encodeURIComponent(document.frmContact.maxlicenserows.value);
			sParameter += '&permitid=' + encodeURIComponent(document.frmContact.permitid.value);
			sParameter += '&permitstatusid=' + encodeURIComponent(document.frmContact.permitstatusid.value);
			sParameter += '&contractortypeid=' + encodeURIComponent(document.frmContact.contractortypeid.value);
			sParameter += '&isorganization=' + encodeURIComponent(document.frmContact.isorganization.value);
			<% If clng(iIsOrganization) = clng(0) Then %>
				sParameter += '&businesstypeid=' + encodeURIComponent($("#businesstypeid").val());
				sParameter += '&statelicense=' + encodeURIComponent($("#statelicense").val());
				sParameter += '&employeecount=' + encodeURIComponent($("#employeecount").val());
				sParameter += '&reference1=' + encodeURIComponent($("#reference1").val());
				sParameter += '&reference2=' + encodeURIComponent($("#reference2").val());
				sParameter += '&reference3=' + encodeURIComponent($("#reference3").val());
				sParameter += '&otherlicensedcity1=' + encodeURIComponent($("#otherlicensedcity1").val());
				sParameter += '&otherlicensedcity2=' + encodeURIComponent($("#otherlicensedcity2").val());
				sParameter += '&generalliabilityagent=' + encodeURIComponent($("#generalliabilityagent").val());
				sParameter += '&generalliabilityphone=' + encodeURIComponent($("#generalliabilityphone").val());
				sParameter += '&workerscompagent=' + encodeURIComponent($("#workerscompagent").val());
				sParameter += '&workerscompphone=' + encodeURIComponent($("#workerscompphone").val());
				sParameter += '&autoinsuranceagent=' + encodeURIComponent($("#autoinsuranceagent").val());
				sParameter += '&autoinsurancephone=' + encodeURIComponent($("#autoinsurancephone").val());
				sParameter += '&bondagent=' + encodeURIComponent($("#bondagent").val());
				sParameter += '&bondagentphone=' + encodeURIComponent($("#bondagentphone").val());
			<%	End If		%>

			for (var a = 0; a <= parseInt(document.frmContact.maxlicenserows.value); a++)
			{
				if (document.getElementById("licensetypeid" + a))
				{
					sParameter += '&licensetypeid' + a + '=' + encodeURIComponent(document.getElementById("licensetypeid" + a).value);
					sParameter += '&licensenumber' + a + '=' + encodeURIComponent(document.getElementById("licensenumber" + a).value);
					sParameter += '&licenseclass' + a + '=' + encodeURIComponent(document.getElementById("licenseclass" + a).value);
					sParameter += '&licensee' + a + '=' + encodeURIComponent(document.getElementById("licensee" + a).value);
					sParameter += '&licenseenddate' + a + '=' + encodeURIComponent(document.getElementById("licenseenddate" + a).value);
				}
			}
			if (parseInt(document.frmContact.maxusers.value) > 0)
			{
				for (a = 1; a <= parseInt(document.frmContact.maxusers.value); a++)
				{
					if (document.getElementById("canaddothers" + a))
					{
						if (document.getElementById("canaddothers" + a).checked)
						{
							sParameter += '&canaddothers' + a + '=' + encodeURIComponent(document.getElementById("canaddothers" + a).value);
						}
					}
				}
				if (document.frmContact.isprimarycontact)
				{
					// if there is only one row then the length is undefined
					if (! document.frmContact.isprimarycontact.length)
					{
						if (document.frmContact.isprimarycontact.checked)
						{
							sParameter += '&isprimarycontact=' + encodeURIComponent(document.frmContact.isprimarycontact.value);
						}
					}
					else 
					{
						// There are multiple rows and we can loop through them
						for (i=0; i < document.frmContact.isprimarycontact.length; i++) 
						{ 
							if (document.frmContact.isprimarycontact[i].checked == true) 
							{ 
								sParameter += '&isprimarycontact=' + encodeURIComponent(document.frmContact.isprimarycontact[i].value);
								break;
							}
					   }
				   }
			   }
		   }
			
			//alert( sParameter );
			//document.frmContact.submit();
			// Do the ajax call
			//doAjax('permitcontactupdate.asp', sParameter, '', 'post', '0');

<%			If bUpdateParent Then	%>
				doAjax('permitcontactupdate.asp', sParameter, 'CloseThisSaved', 'post', '0');
				//CloseThisSaved( 'sResult' );
			<% elseif sType <> "" then %>
				doAjax('permitcontactupdate.asp', sParameter, 'CloseThisSaved', 'post', '0');
<%			Else		%>
				//alert("closing");
				doAjax('permitcontactupdate.asp', sParameter, 'doClose', 'post', '0');
				//doClose();
<%			End If		%>
			//document.frmContact.submit();
		}

		function ShowResults1( sResult )
		{
			alert( sResult );
		}

		function CloseThisSaved( sResult )
		{
			var sDetailText;
			var sToolTip = '';
			//alert( sResult ); 
			//return;
			if (parseInt(document.frmContact.permitcontactid.value) > 0)
			{
				sDetailText = '<a href="javascript:EditContact(\'' + document.frmContact.permitcontactid.value + '\', \'<%=sType%>\');" ';
				if (document.frmContact.firstname.value != "")
				{
					sToolTip += '<strong>' + document.frmContact.firstname.value + ' ' + document.frmContact.lastname.value + '</strong><br />';
				}
				if (document.frmContact.company.value != "")
				{
					if (document.frmContact.firstname.value != "")
					{
						sToolTip += document.frmContact.company.value + '<br />';
					}
					else
					{
						sToolTip += '<strong>' + document.frmContact.company.value + '</strong><br />';
					}
				}
				
				if (document.frmContact.address.value != "")
				{
					sToolTip += document.frmContact.address.value + '<br />';
				}
				if (document.frmContact.city.value != "")
				{
					sToolTip += document.frmContact.city.value + ', ' + document.frmContact.state.value + ' ' + document.frmContact.zip.value + '<br />';
				}
				if (document.frmContact.phone1.value != "")
				{
					sToolTip += '(' + document.frmContact.phone1.value + ') ' + document.frmContact.phone2.value + '-' + document.frmContact.phone3.value;
				}
				var myRegExp = /[\']/g;
				sToolTip = sToolTip.replace(myRegExp, '\\&#39;');
				sDetailText += ' onMouseover="ddrivetip(\'' + sToolTip + '\', 300)"; onMouseout="hideddrivetip()"; '
				sDetailText += '>';
				if (document.frmContact.firstname.value != "")
				{
					sDetailText += document.frmContact.firstname.value + ' ' + document.frmContact.lastname.value ;
					if (document.frmContact.company.value != "")
					{
						sDetailText += ' (' + document.frmContact.company.value + ')';
					}
				}
				else
				{
					sDetailText += document.frmContact.company.value;
				}
				sDetailText += '</a>';
				//alert( sDetailText );
<%				If sType = "iscontractor" Then   %>
					parent.document.getElementById("contractor" + document.frmContact.permitcontactid.value).innerHTML = sDetailText;
<%				Else	%>
					parent.document.getElementById("<%=sType%>details").innerHTML = sDetailText;
<%				End If	%> 
				//parent.NiceTitles.autoCreated.anchors.addElements(parent.document.getElementsByTagName("a"), "title");

				doClose();
				parent.hideddrivetip();
			}
			else
			{
				var temp = new Array();
				temp = sResult.split('$');
				if (parseInt(temp[0]) > 0)
				{
<%					If sType = "iscontractor" Then   %>
						parent.document.getElementById("maxcontractors").value = parseInt(parent.document.getElementById("maxcontractors").value) + 1;
						var tbl = parent.document.getElementById("contractorlist");
						var lastRow = tbl.rows.length;
						var newRow = parseInt(parent.document.getElementById("maxcontractors").value);
						var row = tbl.insertRow(lastRow);

						// Add the Remove Row checkbox
						var cellZero = row.insertCell(0);
						cellZero.align = 'center';
						cellZero.className = 'contactpick';
						var e = parent.document.createElement('input');
						e.type = 'checkbox';
						e.name = 'removepermitcontactid' + newRow;
						e.id = 'removepermitcontactid' + newRow;
						cellZero.appendChild(e);

						// Add the Contact name text
						var cellOne = row.insertCell(1);
						var e1 = parent.document.createElement('input');
						e1.type = 'hidden';
						e1.name = 'contractor' + newRow;
						e1.id = 'contractor' + newRow;
						e1.value = temp[0];
						cellOne.appendChild(e1);
						var cellText = parent.document.createTextNode(temp[1]);
						cellOne.appendChild(cellText);
						//parent.document.getElementById("contractor" + document.frmContact.permitcontactid.value).innerHTML = temp[1];
<%					Elseif request("createpermit") <> "yes" then	%>
						parent.document.getElementById("<%=sType%>display").innerHTML = temp[1];
						parent.document.getElementById("<%=sType%>permitcontacttypeid").value = temp[0];
					<% else %>
						parent.document.getElementById("applicant").innerHTML = '<select name="userid" id="userid" onchange="toggleAddressSearch();"><option value="C' + temp[0] + '">' + temp[1] + '</option></select>';
<%					End If	%> 
				}
				else
				{
<%					If sType <> "iscontractor" Then   %>
						parent.document.getElementById("<%=sType%>display").innerHTML = 'None Selected ';
						parent.document.getElementById("<%=sType%>permitcontacttypeid").value = temp[0];
<%					End If	%> 
				}
				
				doClose();
				parent.hideddrivetip();
				//parent.focus();
			}
		}

		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

//		function SearchCitizens( iSearchStart )
//		{
//			var optiontext;
//			var optionchanged;
//			//alert(document.BuyerForm.searchname.value);
//			var searchtext = document.frmContact.searchname.value;
//			var searchchanged = searchtext.toLowerCase();
//
//			iSearchStart = parseInt(iSearchStart) + 1;
//			
//			for (x=iSearchStart; x < document.frmContact.userid.length ; x++)
//			{
//				optiontext = document.frmContact.userid.options[x].text;
//				optionchanged = optiontext.toLowerCase();
//				if (optionchanged.indexOf(searchchanged) != -1)
//				{
//					document.frmContact.userid.selectedIndex = x;
//					document.frmContact.results.value = 'Possible Match Found.';
//					document.getElementById('searchresults').innerHTML = 'Possible Match Found.';
//					document.frmContact.searchstart.value = x;
//					return;
//				}
//			}
//			document.frmContact.results.value = 'No Match Found.';
//			document.getElementById('searchresults').innerHTML = 'No Match Found.';
//			document.frmContact.searchstart.value = -1;
//		}
//
//		function ClearSearch()
//		{
//			document.frmContact.searchstart.value = -1;
//		}

		function UserPick()
		{
			document.frmContact.searchname.value = '';
			document.frmContact.results.value = '';
			document.getElementById('searchresults').innerHTML = '';
			document.frmContact.searchstart.value = -1;
		}

		function doCalendar( sField ) 
		{
		  var w = (screen.width - 350)/2;
		  var h = (screen.height - 350)/2;
		  eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=frmContact", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function SelectUser( )
		{
			//winHandle = eval('window.open("contactuserpicker.asp?permitcontacttypeid=<%=iPermitContactTypeid%>", "_picker", "width=800,height=500,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('contactuserpicker.asp?permitcontacttypeid=<%=iPermitContactTypeid%>', 'Registered User Selection', 50, 30);
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
					if ($("#removeuser" + t).length)
					{
						// If it is marked for removal, remove it
						if ($("#removeuser" + t).is(':checked') == true)
						{
							// Fire off an Ajax Job to remove them
							//alert($("removeuser" + t).value);
							doAjax('removepermitcontactuser.asp', 'permitcontacttypeid=<%=iPermitContactTypeid%>&userid=' + $("#removeuser" + t).val(), '', 'get', '0');
							tbl.deleteRow(iTableRow);
						}
						else
						{
							iTableRow = iTableRow + 1;
						}
					}
				}
			}
		}


	//-->
	</script>

</head>

<body class="yui-skin-sam">

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	
				<script>parent.document.getElementById('modaltitle'+window.frameElement.getAttribute("data-close")).innerHTML = '<%=sTitle%>';</script>
<!--BEGIN: EDIT FORM-->

	<form name="frmContact" action="permitcontactupdate.asp" method="post">
	<input type="hidden" id="permitcontactid" name="permitcontactid" value="<%=iPermitContactId%>" />
	<input type="hidden" id="permitcontacttypeid" name="permitcontacttypeid" value="<%=iPermitcontacttypeid%>" />
	<input type="hidden" id="permitid" name="permitid" value="<%=iPermitId%>" />
	<input type="hidden" id="permitstatusid" name="permitstatusid" value="<%=iPermitStatusId%>" />
	<input type="hidden" name="activetab" id="activetab" value="<%=iActiveTabId%>" />
	<input type="hidden" name="isorganization" id="isorganization" value="<%=iIsOrganization%>" />
<%	If clng(iIsOrganization) = clng(1) Then %>
	<input type="hidden" name="maxlicenserows" id="maxlicenserows" value="0" />
	<input type="hidden" name="contractortypeid" id="contractortypeid" value="0" />
<%	End If		%>

	<p>
		First Name: &nbsp; <input type="text" id="firstname" name="firstname" value="<%=sFirstName%>" size="25" maxlength="25" />
		&nbsp;&nbsp;
		Last Name: &nbsp; <input type="text" id="lastname" name="lastname" value="<%=sLastName%>" size="25" maxlength="25" />
	</p>
	<p>
		Company: &nbsp; <input type="text" id="company" name="company" value="<%=sCompany%>" size="50" maxlength="50" />
	</p>
	
	<% If iIsOrganization = clng(0) Then %>
	<p>
		Contractor Type: &nbsp; <% ShowContractorTypes iContractorTypeId %>
	</p>
	<% End If %>

	<div id="demo" class="yui-navset">
		<ul class="yui-nav">
			<li><a href="#tab1"><em>Contact Information</em></a></li>
			<li><a href="#tab2"><em>Associate With a Registered User</em></a></li>
			<% If clng(iIsOrganization) = clng(0) Then %>
			<li><a href="#tab3"><em>Licenses and Certifications</em></a></li>
			<% End If %>
		</ul>            
		<div class="yui-content">
			<div id="tab1">
				<p><br />
				<table id="permitcontactinfo" cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td align="right" class="labelcolumn">Address:</td><td class="datacolumn"><input type="text" id="address" name="address" value="<%=sAddress%>" size="50" maxlength="50" /></td>
					</tr>
					<tr>
						<td align="right" class="labelcolumn">City:</td><td class="datacolumn"><input type="text" id="city" name="city" value="<%=sCity%>" size="25" maxlength="25" /></td>
					</tr>
					<tr>
						<td align="right" class="labelcolumn">State:</td><td class="datacolumn"><input type="text" id="state" name="state" value="<%=sState%>" size="2" maxlength="2" /></td>
					</tr>
					<tr>
						<td align="right" class="labelcolumn">Zip:</td><td class="datacolumn"><input type="text" id="zip" name="zip" value="<%=sZip%>" size="10" maxlength="10" /></td>
					</tr>
					<tr>
						<td align="right" class="labelcolumn">Email:</td><td class="datacolumn"><input type="text" id="email" name="email" value="<%=sEmail%>" size="75" maxlength="100" /></td>
					</tr>
					<tr>
						<td align="right" class="labelcolumn">Phone:<input type="hidden" id="phone" name="phone" value="<%=sPhone%>" /></td>
						<td class="datacolumn">(<input type="text" name="phone1" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3" value="<%=Mid(sPhone,1,3)%>">) <input value="<%=Mid(sPhone,4,3)%>" type="text" name="phone2" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3" /> <input value="<%=Mid(sPhone,7,4)%>" type="text" name="phone3" onKeyUp="return autoTab(this, 4, event);" size="4" maxlength="4" /></td>
					</tr>
					<tr>
						<td align="right" class="labelcolumn">Cell:<input type="hidden" id="cell" name="cell" value="<%=sCell%>" /></td>
						<td class="datacolumn">(<input type="text" name="cell1" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3" value="<%=Mid(sCell,1,3)%>">) <input value="<%=Mid(sCell,4,3)%>" type="text" name="cell2" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3" /> <input value="<%=Mid(sCell,7,4)%>" type="text" name="cell3" onKeyUp="return autoTab(this, 4, event);" size="4" maxlength="4" /></td>
					</tr>
					<tr>
						<td align="right" class="labelcolumn">Fax:<input type="hidden" id="fax" name="fax" value="<%=sFax%>" /></td>
						<td class="datacolumn">(<input type="text" name="fax1" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3" value="<%=Mid(sFax,1,3)%>">) <input value="<%=Mid(sFax,4,3)%>" type="text" name="fax2" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3" /> <input value="<%=Mid(sFax,7,4)%>" type="text" name="fax3" onKeyUp="return autoTab(this, 4, event);" size="4" maxlength="4" /></td>
					</tr>
				<% If clng(iIsOrganization) = clng(0) Then %>
						<tr>
							<td align="right" class="labelcolumn">Business Type:</td><td class="datacolumn"><% ShowBusinessTypes iBusinessTypeId %></td>
						</tr>
						<tr>
							<td align="right" class="labelcolumn">State License:</td><td class="datacolumn"><input type="text" id="statelicense" name="statelicense" value="<%=sStateLicense%>" size="30" maxlength="30" /></td>
						</tr>
						<tr>
							<td align="right" class="labelcolumn" nowrap="nowrap">Number of Employees:</td><td class="datacolumn"><input type="text" id="employeecount" name="employeecount" value="<%=sEmployeeCount%>" size="30" maxlength="30" /></td>
						</tr>
						<tr>
							<td align="right" class="labelcolumn" nowrap="nowrap">Mandatory References:</td><td class="datacolumn"><input type="text" id="reference1" name="reference1" value="<%=sReference1%>" size="30" maxlength="30" /></td>
						</tr>
						<tr>
							<td class="labelcolumn">&nbsp;</td><td class="datacolumn"><input type="text" id="reference2" name="reference2" value="<%=sReference2%>" size="30" maxlength="30" /></td>
						</tr>
						<tr>
							<td class="labelcolumn">&nbsp;</td><td class="datacolumn"><input type="text" id="reference3" name="reference3" value="<%=sReference3%>" size="30" maxlength="30" /></td>
						</tr>
						<tr>
							<td align="right" class="labelcolumn" nowrap="nowrap">Other Cities Licensed In:</td><td class="datacolumn"><input type="text" id="otherlicensedcity1" name="otherlicensedcity1" value="<%=sOtherLicensedCity1%>" size="30" maxlength="30" /></td>
						</tr>
						<tr>
							<td class="labelcolumn">&nbsp;</td><td class="datacolumn"><input type="text" id="otherlicensedcity2" name="otherlicensedcity2" value="<%=sOtherLicensedCity2%>" size="30" maxlength="30" /></td>
						</tr>
						<tr>
							<td colspan="2">Insurance Agents</td></td>
						</tr>
						<tr>
							<td align="right" class="labelcolumn">General Liability:</td><td class="datacolumn"><input type="text" id="generalliabilityagent" name="generalliabilityagent" value="<%=sGeneralLiabilityAgent%>" size="30" maxlength="30" /> &nbsp; Phone: <input type="text" id="generalliabilityphone" name="generalliabilityphone" value="<%=sGeneralLiabilityPhone%>" size="20" maxlength="20" /></td>
						</tr>
						<tr>
							<td align="right" class="labelcolumn">Worker's Compensation:</td><td class="datacolumn"><input type="text" id="workerscompagent" name="workerscompagent" value="<%=sWorkersCompAgent%>" size="30" maxlength="30" /> &nbsp; Phone: <input type="text" id="workerscompphone" name="workerscompphone" value="<%=sWorkersCompPhone%>" size="20" maxlength="20" /></td>
						</tr>
						<tr>
							<td align="right" class="labelcolumn">Auto Insurance:</td><td class="datacolumn"><input type="text" id="autoinsuranceagent" name="autoinsuranceagent" value="<%=sAutoInsuranceAgent%>" size="30" maxlength="30" /> &nbsp; Phone: <input type="text" id="autoinsurancephone" name="autoinsurancephone" value="<%=sAutoInsurancePhone%>" size="20" maxlength="20" /></td>
						</tr>
						<tr>
							<td align="right" class="labelcolumn">Bond Agent:</td><td class="datacolumn"><input type="text" id="bondagent" name="bondagent" value="<%=sBondAgent%>" size="30" maxlength="30" /> &nbsp; Phone: <input type="text" id="bondagentphone" name="bondagentphone" value="<%=sBondAgentPhone%>" size="20" maxlength="20" /></td>
						</tr>
				<% End If %>
				</table>
				</p>
			</div>

			<div id="tab2"> <!-- Associate With a Registered User -->
				<p><br />
						<!--Select Name: <select name="userid" onchange="javascript:UserPick();">
										<option value="0">Select a registered user from the list</option>
										<% ShowUserDropDown( iUserId )%>
									</select>
						<br /><br />Name Search: <input type="text" name="searchname" value="<%=sSearchName%>" size="25" maxlength="50" onchange="ClearSearch();" />
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Search" onclick="SearchCitizens(document.frmContact.searchstart.value);" />
						<input type="hidden" name="results" value="" /><input type="hidden" name="searchstart" value="<%=sSearchStart%>" />
						<span id="searchresults"><%=sResults%></span>
						<br /><div id="searchtip">(last name, first name)</div>		-->
						
					<input type="button" class="button ui-button ui-widget ui-corner-all" value="Add A User" onclick="SelectUser( );" /> &nbsp; <input type="button" class="button ui-button ui-widget ui-corner-all" value="Remove The Selected Users" onClick="RemoveUserRows()"  />
					<br /><br />
					<table cellpadding="0" cellspacing="0" border="0" id="contractoruserlist">
						<tr><th class="selectcol">&nbsp;</th><th>Name</th><th class="pickcol">Can Add<br />Others</th><th class="pickcol">Primary<br />Contact</th></tr>
<%						iMaxUsers = ShowContractorUsers( iPermitcontacttypeid )		%>									
					</table>
					<input type="hidden" id="maxusers" name="maxusers" value="<%=iMaxUsers%>" />
					<br />								

				</p> 
			</div>

			<% If clng(iIsOrganization) = clng(0) Then %>
			<div id="tab3"> <!-- Licenses -->
				<p><br />
					<%					
					tooltipclass=""
					tooltip = ""
					disabled = ""
					If not bCanSaveChanges Then
						tooltipclass="tooltip"
						disabled = " disabled "
						tooltip = "<span class=""tooltiptext"">You don't have permission to save changes.</span>"
					end if
					%>
					<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" id="addlicensebutton" onclick="NewLicenseRow();">Add Row<%=tooltip%></button> &nbsp;&nbsp;
					<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" id="removelicensebutton" onclick="RemoveLicenseRows();">Remove Selected<%=tooltip%></button>
					<table id="licensetable" border="0" cellpadding="0" cellspacing="0">
						<tr><th>&nbsp;</th><th>Number</th><th>Class</th><th>Type</th><th>Licensee Name</th><th>Expires</th></tr>
<%						iMaxLicenseRows = ShowLicenseTable( iPermitContactId ) %>
					</table>
					<input type="hidden" name="maxlicenserows" value="<%=iMaxLicenseRows%>" />
				</p>
			</div>
			<% End If %>

		</div>
	</div>
	<br />
	<%					
	tooltipclass=""
	tooltip = ""
	disabled = ""
	If not bCanSaveChanges Then
		tooltipclass="tooltip"
		disabled = " disabled "
		tooltip = "<span class=""tooltiptext"">You don't have permission to save changes.</span>"
	end if
	%>
	<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" id="savebutton" onclick="doUpdate();"><%=sSaveButton%><%=tooltip%></button>&nbsp;&nbsp;&nbsp;&nbsp;
	<input type="button" class="button ui-button ui-widget ui-corner-all" value="Close" onclick="doClose();" />

</form>
<!--END: EDIT FORM-->

	</div>
</div>
<!--END: PAGE CONTENT-->


<!--#Include file="../admin_footer.asp"-->  
	<!--#Include file="modal.asp"-->  

</body>

</html>


<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void ShowUserDropDown iUserId 
'--------------------------------------------------------------------------------------------------
Sub ShowUserDropDown( ByVal iUserId )
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


'--------------------------------------------------------------------------------------------------
' integer ShowLicenseTable( iPermitContactid )
'--------------------------------------------------------------------------------------------------
Function ShowLicenseTable( ByVal iPermitContactid )
	Dim oRs, sSql, iRowCount, sRowClass

	iRowCount = -1

	sSql = "SELECT licensetype, licensenumber, licensee, licenseexpiration, licenseenddate, "
	sSql = sSql & " ISNULL(licenseclass,'') AS licenseclass, ISNULL(licensetypeid,0) AS licensetypeid "
	sSql = sSql & " FROM egov_permitcontacts_licenses "
	sSql = sSql & " WHERE permitcontactid = " & iPermitContactid
	sSql = sSql & " ORDER BY licenseenddate DESC, licensenumber DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

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
			'response.write "<td align=""center""><input type=""text"" id=""licensetype" & iRowCount & """ name=""licensetype" & iRowCount & """ value=""" & Replace(oRs("licensetype"),Chr(34),"&quot;") & """ size=""20"" maxlength=""25"" /></td>"
			response.write "<td align=""center""><input type=""text"" id=""licensee" & iRowCount & """ name=""licensee" & iRowCount & """ value=""" & oRs("licensee") & """ size=""30"" maxlength=""100"" /></td>"
			response.write "<td align=""center""><input type=""text"" id=""licenseenddate" & iRowCount & """ name=""licenseenddate" & iRowCount & """ value=""" & FormatDateTime(oRs("licenseenddate"),2) & """ size=""10"" maxlength=""10"" />"
  			response.write "<script>$( function() { $( ""#licenseenddate" & iRowCount & """ ).datepicker({ changeMonth: true, showOn: ""both"", buttonText: ""<i class=\""fa fa-calendar\""></i>"", changeYear: true }); });</script>"
			response.write "</td>"
			response.write "</tr>"
			oRs.MoveNext 
		Loop 
	Else
		' put in a starter row.
		iRowCount = 0
		response.write vbcrlf & "<tr><td class=""firstcell"">"
		response.write "<input type=""checkbox"" id=""removelicense" & iRowCount & """ name=""removelicense" & iRowCount & """ /></td>"
		response.write "<td align=""center""><input type=""text"" id=""licensenumber" & iRowCount & """ name=""licensenumber" & iRowCount & """ value="""" size=""10"" maxlength=""20"" /></td>"
		response.write "<td align=""center""><input type=""text"" id=""licenseclass" & iRowCount & """ name=""licenseclass" & iRowCount & """ value="""" size=""20"" maxlength=""25"" /></td>"
		response.write "<td align=""center"">"
		ShowLicenseTypePicks 0, iRowCount
		response.write "</td>"
		'response.write "<td align=""center""><input type=""text"" id=""licensetype" & iRowCount & """ name=""licensetype" & iRowCount & """ value="""" size=""20"" maxlength=""25"" /></td>"
		response.write "<td align=""center""><input type=""text"" id=""licensee" & iRowCount & """ name=""licensee" & iRowCount & """ value="""" size=""30"" maxlength=""100"" /></td>"
		response.write "<td align=""center""><input type=""text"" id=""licenseenddate" & iRowCount & """ name=""licenseenddate" & iRowCount & """ value="""" size=""10"" maxlength=""10"" />"
  		response.write "<script>$( function() { $( ""#licenseenddate" & iRowCount & """ ).datepicker({ changeMonth: true, showOn: ""both"", buttonText: ""<i class=\""fa fa-calendar\""></i>"", changeYear: true }); });</script>"
		response.write "</td>"
		response.write "</tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowLicenseTable = iRowCount

End Function



%>
