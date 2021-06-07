<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permittypeedit.asp
' AUTHOR: Steve Loar
' CREATED: 01/16/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Creates and edits permit types
'
' MODIFICATION HISTORY
' 1.0   01/16/2008	Steve Loar - INITIAL VERSION
' 1.1   06/??/2008	Steve Loar - Added Review Alerts
' 1.2	07/16/2008	Steve Loar - Added Inspection Alerts
' 1.3	07/25/2008	Steve Loar - Inspectors of unassigned added
' 2.0	10/27/2010	Steve Loar - Changes to allow any type of permits
' 2.1	01/11/2011	Steve Loar - Added flag to notify reviewers of attachments
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sTitle, iPermitTypeid, sPermitType, sPermitTypeDesc, sIsBuildingPermitType, iMaxFeeRows, iMaxInspRows
Dim sExpirationDays, sPermitNumberPrefix, sPublicDescription, iMaxReviewAlertRows, iMaxInspectionAlertRows
Dim sAdditionalFooterInfo, sPermitTitle, sApprovingOfficial, sPermitSubTitle, sPermitRightTitle, sPermitTitleBottom
Dim sPermitFooter, sPermitSubFooter, sListFixtures, sShowConstructionType, sShowFeeTotal, sShowOccupancyType
Dim sShowJobValue, sShowWorkDesc, sShowFootages, sShowProposedUse, sShowOtherContacts, sPermitLogo
Dim sGroupByInvoiceCategories, sInvoiceLogo, sInvoiceHeader, sShowElectricalContractor, sShowMechanicalContractor
Dim sShowPlumbingContractor, sShowApplicantLicense, sShowCounty, sShowParcelid, sShowPlansBy, oAddressOrg
Dim sShowPrimaryContact, iUseTypeId, sHasTempCo, sHasCo, sShowApprovedAsOnTCO, sShowApprovedAsOnCO
Dim sShowConstTypeOnTCO, sShowConstTypeOnCO, sShowOccTypeonTCO, sShowOccTypeonCO, sShowOccupantsOnTCO, sShowOccupantsOnCO
Dim sTempCOLogo, sCOLogo, sTempCOTitle, sTempCOSubTitle, sCOTitle, sCOSubTitle, sTempCoTitleSize, sTempCoTitleStyle
Dim sTempCoTitleWeight, sTempCoTitleFamily, sTempCOAddress, sCOAddress, sTempCOTopText, sCOTopText
Dim sTempCOBottomText, sCOBottomText, sTempCOCodeRef, sCOCodeRef, sTempCOApproval, sCOApproval
Dim sTempCOFooter, sCOFooter, sTempCOSubFooter, sCOSubFooter, sShowTotalSqFt, sShowApprovedAs, sShowFeeTypeTotals
Dim sShowOccupancyUse, sShowPayments, iDocumentId, iMaxDocRows, iPermitCategoryId, iMaxCustomFieldRows
Dim iPermitLocationRequirementId, sAttachmentReviewerAlert

sLevel = "../" ' Override of value from common.asp
iMaxFeeRows = 0
iMaxInspRows = 0
sExpirationDays = ""
sPermitNumberPrefix = ""
iMaxReviewAlertRows = 0
iMaxInspectionAlertRows = 0
sTempCoTitleSize = "10pt"
sTempCoTitleStyle = "normal"
sTempCoTitleWeight = "normal"
sTempCoTitleFamily = "arial"

PageDisplayCheck "permit types", sLevel	' In common.asp

iPermitTypeid = CLng(request("permittypeid") )

If CLng(iPermitTypeid) > CLng(0) Then
	sTitle = "Edit"
	GetPermitType iPermitTypeid
Else
	sTitle = "New"
	sPermitType = ""
	sPermitTypeDesc = ""
	iPermitLocationRequirementId = 0
	'sIsBuildingPermitType = " checked=""checked"" "
	sIsFinal = ""
	sPublicDescription = ""
	sPermitTitle = ""
	sAdditionalFooterInfo = ""
	sApprovingOfficial = ""
	sPermitSubTitle = ""
	sPermitRightTitle = ""
	sPermitTitleBottom = ""
	sPermitFooter = ""
	sPermitSubFooter = ""
	sListFixtures = ""
	sShowConstructionType = ""
	sShowFeeTotal = ""
	sShowOccupancyType = ""
	sShowJobValue = ""
	sShowWorkDesc = ""
	sShowFootages = ""
	sShowProposedUse = ""
	sShowOtherContacts = ""
	sPermitLogo = ""
	sGroupByInvoiceCategories = ""
	sInvoiceLogo = ""
	sInvoiceHeader = ""
	sShowElectricalContractor = "" 
	sShowMechanicalContractor = ""
	sShowPlumbingContractor = ""
	sShowApplicantLicense = ""
	sShowCounty = ""
	sShowParcelid = ""
	sShowPlansBy = ""
	sShowPrimaryContact = ""
	iUseTypeId = 0 
	sHasTempCo = ""
	sHasCo = ""
	sShowApprovedAsOnTCO = ""
	sShowApprovedAsOnCO = ""
	sShowConstTypeOnTCO = ""
	sShowConstTypeOnCO = ""
	sShowOccTypeOnTCO = ""
	sShowOccTypeOnCO = ""
	sShowOccupantsOnTCO = ""
	sShowOccupantsOnCO = ""
	sTempCOLogo = ""
	sCOLogo = ""
	sTempCOTitle = ""
	sTempCOSubTitle = ""
	sCOTitle = ""
	sCOSubTitle = ""
	sTempCOAddress = ""
	sCOAddress = ""
	sTempCOTopText = ""
	sCOTopText = ""
	sTempCOCodeRef = ""
	sCOCodeRef = ""
	sTempCOApproval = ""
	sCOApproval = ""
	sTempCOFooter = ""
	sCOFooter = ""
	sTempCOSubFooter = ""
	sCOSubFooter = ""
	sShowTotalSqFt = ""
	sShowApprovedAs = ""
	sShowFeeTypeTotals = ""
	sShowOccupancyUse = ""
	sShowPayments = ""
	iDocumentId = 0
	sAttachmentReviewerAlert = ""
End If 

If request("activetab") <> "" Then 
	If IsNumeric(request("activetab")) Then 
		iActiveTabId = clng(request("activetab"))
	Else
		iActiveTabId = clng(0)
	End If 
Else
	iActiveTabId = clng(0)
End If 

Set oAddressOrg = New classOrganization

bOrgHasPermitTypeReport = OrgHasFeature( "permit type report" )		' in common.asp

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

	<script type="text/javascript" src="../yui/yahoo-dom-event.js"></script>  
	<script type="text/javascript" src="../yui/element-min.js"></script>  
	<script type="text/javascript" src="../yui/tabview-min.js"></script>

	<script language="JavaScript" src="../scripts/formatnumber.js"></script>
	<script language="JavaScript" src="../scripts/removespaces.js"></script>
	<script language="JavaScript" src="../scripts/removecommas.js"></script>
	<script language="JavaScript" src="../scripts/textareamaxlength.js"></script>
	<script language="javascript" src="../scripts/modules.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
	<script language="JavaScript" src="../scriptaculous/src/scriptaculous.js"></script>
  	<script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  	<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

	<script language="Javascript">
	<!--
		var tabView;

		(function() {
			tabView = new YAHOO.widget.TabView('demo');
			//tabView.set('activeIndex', 0); 
			tabView.set('activeIndex',<%=iActiveTabId%>);

		})();

		function Another()
		{
			location.href="permittypeedit.asp?permittypeid=0";
		}

		function EditFeeTypes()
		{
			showModal('permitfeetypelist.asp', 'Permit Fee Types', 65, 95);
		}
		function EditPermitReviewTypes()
		{
			showModal('permitreviewtypelist.asp', 'Permit Review Types', 65, 95);
		}
		function EditPermitInspectionTypes()
		{
			showModal('permitinspectiontypelist.asp', 'Permit Inspection Types', 65, 95);
		}
		function EditPermitCustomFieldTypes()
		{
			showModal('permitcustomfieldtypelist.asp', 'Permit Custom Field Types', 65, 55);
		}

		/*
		function doPreview( sField )
		{
			var winHandle;
			var w = (screen.width - 800)/2;
			var h = (screen.height - 600)/2;
			var sText = $F(sField);
			//sText = sText.gsub(/&/,'%26');
			//sText = sText.gsub(/"/,'%22');
			sText = sText.gsub(/\n/,'');
			sText = sText.gsub(/\r/,'');
			//alert(sText);
			winHandle = eval('window.open("fontpreview.asp?fontsize=' + $F(sField + "size") + '&fontstyle=' + $F(sField + "style") + '&fontweight=' + $F(sField + "weight") + '&fontfamily=' + $F(sField + "family") + '&displaytext=' + encodeURIComponent(sText) + '", "_contact", "width=800,height=500,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,resizable=1,left=' + w + ',top=' + h + '")');
			winHandle.focus();
		}
		*/

		function NewFeeRow()
		{
			document.frmPermit.maxfeerows.value = parseInt(document.frmPermit.maxfeerows.value) + 1;
			var tbl = document.getElementById("permittypefeetable");
			var lastRow = tbl.rows.length;
			var newRow = parseInt(document.frmPermit.maxfeerows.value);
			var row = tbl.insertRow(lastRow);

			// Remove Row checkbox
			var cellZero = row.insertCell(0);
			cellZero.className = 'firstcell';
			var e = document.createElement('input');
			e.type = 'checkbox';
			e.name = 'removefee' + newRow;
			e.id = 'removefee' + newRow;
			cellZero.appendChild(e);

			//fee type pick here
			cellZero = row.insertCell(1);
			//cellZero.className = 'firstcell';
			//cellZero.align = 'center';
			var e0 = document.createElement('select');
			e0.name = 'permitfeetypeid' + newRow;
			e0.id = 'permitfeetypeid' + newRow;
			e0.classList.add('permitfeetypeDD');
			cellZero.appendChild(e0);

			// Find the first row that exists
			for (var t = 0; t <= parseInt(document.frmPermit.maxfeerows.value); t++ )
			{
				if (document.getElementById("permitfeetypeid" + t))
				{
					break;
				}
			}
			var slength = document.getElementById("permitfeetypeid" + t).length;
			var op;
			var newText; 
			for ( var s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("permitfeetypeid" + t).options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("permitfeetypeid" + t).options[s].value );
				e0.appendChild(op);
			}

			// Is Required checkbox
			var cellOne = row.insertCell(2);
			cellOne.align = 'center';
			var e1 = document.createElement('input');
			e1.type = 'checkbox';
			e1.name = 'isrequired' + newRow;
			e1.id = 'isrequired' + newRow;
			cellOne.appendChild(e1);

			// Add the new display order to the existing display order picks
			var newDisplayOrder = parseInt(document.frmPermit.maxfeerows.value) + 1;
			for ( var o=0; o < newRow; o++)
			{
				if (document.getElementById("displayorder" + o))
				{
					op = document.createElement('OPTION');
					newText = document.createTextNode( newDisplayOrder );
					op.appendChild( newText );
					op.setAttribute( 'value', newDisplayOrder );
					document.getElementById('displayorder' + o).appendChild(op);
				}
			}
			// The display order pick
			var cellTwo = row.insertCell(3);
			cellTwo.align = 'center';
			var e2 = document.createElement('select');
			e2.name = 'displayorder' + newRow;
			e2.id = 'displayorder' + newRow;
			cellTwo.appendChild(e2);
			slength = document.getElementById("displayorder" + t).length;
			for ( s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("displayorder" + t).options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("displayorder" + t).options[s].value );
				op.selected = true;
				e2.appendChild(op);
			}
		}

		function NewDocumentRow()
		{
			document.frmPermit.maxdocumentrows.value = parseInt(document.frmPermit.maxdocumentrows.value) + 1;
			var tbl = document.getElementById("permittypedocumenttable");
			var lastRow = tbl.rows.length;
			var newRow = parseInt(document.frmPermit.maxdocumentrows.value);
			var row = tbl.insertRow(lastRow);

			// Remove Row checkbox
			var cellZero = row.insertCell(0);
			cellZero.className = 'firstcell';
			var e = document.createElement('input');
			e.type = 'checkbox';
			e.name = 'removedocument' + newRow;
			e.id = 'removedocument' + newRow;
			cellZero.appendChild(e);

			//hidden field
			e = document.createElement('input');
			e.type = 'hidden';
			e.name = 'permitdocumentid' + newRow;
			e.id = 'permitdocumentid' + newRow;
			e.value = '0';
			cellZero.appendChild(e);

			// Document label
			cellZero = row.insertCell(1);
			var e = document.createElement('input');
			e.type = 'text';
			e.name = 'documentlabel' + newRow;
			e.id = 'documentlabel' + newRow;
			e.size = '50';
			e.maxlength = '50';
			e.value = '';
			cellZero.appendChild(e);

			//inspection type pick here
			cellZero = row.insertCell(2);
			var e0 = document.createElement('select');
			e0.name = 'documentid' + newRow;
			e0.id = 'documentid' + newRow;
			cellZero.appendChild(e0);

			// Find the first row that exists
			for (var t = 0; t <= parseInt(document.frmPermit.maxdocumentrows.value); t++ )
			{
				if (document.getElementById("documentid" + t))
				{
					break;
				}
			}

			var slength = document.getElementById("documentid" + t).length;
			var op;
			var newText; 
			for ( var s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("documentid" + t).options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("documentid" + t).options[s].value );
				e0.appendChild(op);
			}
		}

		function NewInspectionRow()
		{
			document.frmPermit.maxinspectionrows.value = parseInt(document.frmPermit.maxinspectionrows.value) + 1;
			var tbl = document.getElementById("permittypeinspectiontable");
			var lastRow = tbl.rows.length;
			var newRow = parseInt(document.frmPermit.maxinspectionrows.value);
			var row = tbl.insertRow(lastRow);

			// Remove Row checkbox
			var cellZero = row.insertCell(0);
			cellZero.className = 'firstcell';
			var e = document.createElement('input');
			e.type = 'checkbox';
			e.name = 'removeinspection' + newRow;
			e.id = 'removeinspection' + newRow;
			cellZero.appendChild(e);

			//inspection type pick here
			cellZero = row.insertCell(1);
			//cellZero.className = 'firstcell';
			//cellZero.align = 'center';
			var e0 = document.createElement('select');
			e0.name = 'permitinspectiontypeid' + newRow;
			e0.id = 'permitinspectiontypeid' + newRow;
			e0.classList.add('permitinspectiontypeDD');
			cellZero.appendChild(e0);

			// Find the first row that exists
			for (var t = 0; t <= parseInt(document.frmPermit.maxinspectionrows.value); t++ )
			{
				if (document.getElementById("permitinspectiontypeid" + t))
				{
					break;
				}
			}

			var slength = document.getElementById("permitinspectiontypeid" + t).length;
			var op;
			var newText; 
			for ( var s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("permitinspectiontypeid" + t).options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("permitinspectiontypeid" + t).options[s].value );
				e0.appendChild(op);
			}

			//inspector pick here
			var cellOne = row.insertCell(2);
			//cellZero.className = 'firstcell';
			cellOne.align = 'center';
			e1 = document.createElement('select');
			e1.name = 'permitinspectorid' + newRow;
			e1.id = 'permitinspectorid' + newRow;
			cellOne.appendChild(e1);
			slength = document.getElementById("permitinspectorid" + t).length;
			for ( s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("permitinspectorid" + t).options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("permitinspectorid" + t).options[s].value );
				e1.appendChild(op);
			}

			// Is Required checkbox
			var cellTwo = row.insertCell(3);
			cellTwo.align = 'center';
			e1 = document.createElement('input');
			e1.type = 'checkbox';
			e1.name = 'inspectionisrequired' + newRow;
			e1.id = 'inspectionisrequired' + newRow;
			cellTwo.appendChild(e1);

			// Is Final checkbox
			var cellThree = row.insertCell(4);
			cellThree.align = 'center';
			cellThree.innerHTML = '<input type="radio" name="isfinal" value="' + newRow + '" />';

			// Add the new inspection order to the existing inspection order picks
			var newDisplayOrder = parseInt(document.frmPermit.maxinspectionrows.value) + 1;
			for ( var o=0; o < newRow; o++)
			{
				if (document.getElementById("inspectionorder" + o))
				{
					op = document.createElement('OPTION');
					newText = document.createTextNode( newDisplayOrder );
					op.appendChild( newText );
					op.setAttribute( 'value', newDisplayOrder );
					document.getElementById('inspectionorder' + o).appendChild(op);
				}
			}
			// The inspection order pick
			var cellFive = row.insertCell(5);
			cellFive.align = 'center';
			var e2 = document.createElement('select');
			e2.name = 'inspectionorder' + newRow;
			e2.id = 'inspectionorder' + newRow;
			cellFive.appendChild(e2);
			slength = document.getElementById("inspectionorder" + t).length;
			for ( s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("inspectionorder" + t).options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("inspectionorder" + t).options[s].value );
				op.selected = true;
				e2.appendChild(op);
			}
		}

		function NewReviewAlertRow()
		{
			document.frmPermit.maxreviewalertrows.value = parseInt(document.frmPermit.maxreviewalertrows.value) + 1;
			var tbl = document.getElementById("permittypereviewalerttable");
			var lastRow = tbl.rows.length;
			var newRow = parseInt(document.frmPermit.maxreviewalertrows.value);
			var row = tbl.insertRow(lastRow);

			// Remove Row checkbox
			var cellZero = row.insertCell(0);
			cellZero.className = 'firstcell';
			var e = document.createElement('input');
			e.type = 'checkbox';
			e.name = 'removereviewalert' + newRow;
			e.id = 'removereviewalert' + newRow;
			cellZero.appendChild(e);

			//review alert type pick here
			cellZero = row.insertCell(1);
			cellZero.align = 'center';
			var e0 = document.createElement('select');
			e0.name = 'permitalerttypeid' + newRow;
			e0.id = 'permitalerttypeid' + newRow;
			cellZero.appendChild(e0);

			// Find the first row that exists
			for (var t = 0; t <= parseInt(document.frmPermit.maxreviewalertrows.value); t++ )
			{
				if (document.getElementById("permitalerttypeid" + t))
				{
					break;
				}
			}

			var slength = document.getElementById("permitalerttypeid" + t).length;
			var op;
			var newText; 
			for ( var s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("permitalerttypeid" + t).options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("permitalerttypeid" + t).options[s].value );
				e0.appendChild(op);
			}

			//notify reviewer pick here
			var cellOne = row.insertCell(2);
			cellOne.align = 'center';
			e1 = document.createElement('select');
			e1.name = 'notifyuserid' + newRow;
			e1.id = 'notifyuserid' + newRow;
			cellOne.appendChild(e1);
			slength = document.getElementById("notifyuserid" + t).length;
			for ( s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("notifyuserid" + t).options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("notifyuserid" + t).options[s].value );
				e1.appendChild(op);
			}
		}

		function NewInspectionAlertRow()
		{
			document.frmPermit.maxinspectionalertrows.value = parseInt(document.frmPermit.maxinspectionalertrows.value) + 1;
			var tbl = document.getElementById("permittypeinspectionalerttable");
			var lastRow = tbl.rows.length;
			var newRow = parseInt(document.frmPermit.maxinspectionalertrows.value);
			var row = tbl.insertRow(lastRow);

			// Remove Row checkbox
			var cellZero = row.insertCell(0);
			cellZero.className = 'firstcell';
			var e = document.createElement('input');
			e.type = 'checkbox';
			e.name = 'removeinspectionalert' + newRow;
			e.id = 'removeinspectionalert' + newRow;
			cellZero.appendChild(e);

			//review alert type pick here
			cellZero = row.insertCell(1);
			cellZero.align = 'center';
			var e0 = document.createElement('select');
			e0.name = 'permitinspectionalerttypeid' + newRow;
			e0.id = 'permitinspectionalerttypeid' + newRow;
			cellZero.appendChild(e0);

			// Find the first row that exists
			for (var t = 0; t <= parseInt(document.frmPermit.maxinspectionalertrows.value); t++ )
			{
				if (document.getElementById("permitinspectionalerttypeid" + t))
				{
					break;
				}
			}

			var slength = document.getElementById("permitinspectionalerttypeid" + t).length;
			var op;
			var newText; 
			for ( var s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("permitinspectionalerttypeid" + t).options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("permitinspectionalerttypeid" + t).options[s].value );
				e0.appendChild(op);
			}

			//notify reviewer pick here
			var cellOne = row.insertCell(2);
			cellOne.align = 'center';
			e1 = document.createElement('select');
			e1.name = 'notifyinspectoruserid' + newRow;
			e1.id = 'notifyinspectoruserid' + newRow;
			cellOne.appendChild(e1);
			slength = document.getElementById("notifyinspectoruserid" + t).length;
			for ( s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("notifyinspectoruserid" + t).options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("notifyinspectoruserid" + t).options[s].value );
				e1.appendChild(op);
			}
		}

		function NewReviewRow()
		{
			document.frmPermit.maxreviewrows.value = parseInt(document.frmPermit.maxreviewrows.value) + 1;
			var tbl = document.getElementById("permittypereviewtable");
			var lastRow = tbl.rows.length;
			var newRow = parseInt(document.frmPermit.maxreviewrows.value);
			var row = tbl.insertRow(lastRow);

			// Remove Row checkbox
			var cellZero = row.insertCell(0);
			cellZero.className = 'firstcell';
			var e = document.createElement('input');
			e.type = 'checkbox';
			e.name = 'removereview' + newRow;
			e.id = 'removereview' + newRow;
			cellZero.appendChild(e);

			//review type pick here
			cellZero = row.insertCell(1);
			//cellZero.className = 'firstcell';
			//cellZero.align = 'center';
			var e0 = document.createElement('select');
			e0.name = 'permitreviewtypeid' + newRow;
			e0.id = 'permitreviewtypeid' + newRow;
			e0.classList.add('permitreviewtypeDD');
			cellZero.appendChild(e0);

			// Find the first row that exists
			for (var t = 0; t <= parseInt(document.frmPermit.maxreviewrows.value); t++ )
			{
				if (document.getElementById("permitreviewtypeid" + t))
				{
					break;
				}
			}

			var slength = document.getElementById("permitreviewtypeid" + t).length;
			var op;
			var newText; 
			for ( var s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("permitreviewtypeid" + t).options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("permitreviewtypeid" + t).options[s].value );
				e0.appendChild(op);
			}

			//reviewer pick here
			var cellOne = row.insertCell(2);
			cellOne.align = 'center';
			e1 = document.createElement('select');
			e1.name = 'permitreviewerid' + newRow;
			e1.id = 'permitreviewerid' + newRow;
			cellOne.appendChild(e1);
			slength = document.getElementById("permitreviewerid" + t).length;
			for ( s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("permitreviewerid" + t).options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("permitreviewerid" + t).options[s].value );
				e1.appendChild(op);
			}

			// Is Required checkbox
			var cellTwo = row.insertCell(3);
			cellTwo.align = 'center';
			e1 = document.createElement('input');
			e1.type = 'checkbox';
			e1.name = 'reviewisrequired' + newRow;
			e1.id = 'reviewisrequired' + newRow;
			cellTwo.appendChild(e1);

			// Notify on Release checkbox notifyonrelease
			var cellTwo = row.insertCell(4);
			cellTwo.align = 'center';
			e1 = document.createElement('input');
			e1.type = 'checkbox';
			e1.name = 'notifyonrelease' + newRow;
			e1.id = 'notifyonrelease' + newRow;
			cellTwo.appendChild(e1);

			// Add the new display order to the existing display order picks
			var newDisplayOrder = parseInt(document.frmPermit.maxreviewrows.value) + 1;
			for ( var o=0; o < newRow; o++)
			{
				if (document.getElementById("revieworder" + o))
				{
					op = document.createElement('OPTION');
					newText = document.createTextNode( newDisplayOrder );
					op.appendChild( newText );
					op.setAttribute( 'value', newDisplayOrder );
					document.getElementById('revieworder' + o).appendChild(op);
				}
			}
			// The review order pick
			var cellFive = row.insertCell(5);
			cellFive.align = 'center';
			var e2 = document.createElement('select');
			e2.name = 'revieworder' + newRow;
			e2.id = 'revieworder' + newRow;
			cellFive.appendChild(e2);
			slength = document.getElementById("revieworder" + t).length;
			for ( s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("revieworder" + t).options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("revieworder" + t).options[s].value );
				op.selected = true;
				e2.appendChild(op);
			}
		}

		function NewCustomFieldRow()
		{
			document.frmPermit.maxcustomfieldrows.value = parseInt(document.frmPermit.maxcustomfieldrows.value) + 1;
			var tbl = document.getElementById("permittypecustomfieldstable");
			var lastRow = tbl.rows.length;
			var newRow = parseInt(document.frmPermit.maxcustomfieldrows.value);
			var row = tbl.insertRow(lastRow);

			// Remove Row checkbox
			var cellZero = row.insertCell(0);
			cellZero.className = 'firstcell';
			var e = document.createElement('input');
			e.type = 'checkbox';
			e.name = 'removecustomfield' + newRow;
			e.id = 'removecustomfield' + newRow;
			cellZero.appendChild(e);

			//custom field pick here
			cellZero = row.insertCell(1);
			var e0 = document.createElement('select');
			e0.name = 'customfieldtypeid' + newRow;
			e0.id = 'customfieldtypeid' + newRow;
			e0.classList.add('permitcustomfieldtypeDD');
			cellZero.appendChild(e0);

			// Find the first row that exists
			for (var t = 0; t <= parseInt(document.frmPermit.maxcustomfieldrows.value); t++ )
			{
				if (document.getElementById("customfieldtypeid" + t))
				{
					break;
				}
			}

			var slength = document.getElementById("customfieldtypeid" + t).length;
			var op;
			var newText; 
			for ( var s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("customfieldtypeid" + t).options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("customfieldtypeid" + t).options[s].value );
				e0.appendChild(op);
			}

			// Add the new display order to the existing display order picks
			var newDisplayOrder = parseInt(document.frmPermit.maxcustomfieldrows.value) + 1;
			for ( var o=0; o < newRow; o++)
			{
				if (document.getElementById("customfieldorder" + o))
				{
					op = document.createElement('OPTION');
					newText = document.createTextNode( newDisplayOrder );
					op.appendChild( newText );
					op.setAttribute( 'value', newDisplayOrder );
					document.getElementById('customfieldorder' + o).appendChild(op);
				}
			}

			// The custom field order pick
			var cellFive = row.insertCell(2);
			cellFive.align = 'center';
			var e2 = document.createElement('select');
			e2.name = 'customfieldorder' + newRow;
			e2.id = 'customfieldorder' + newRow;
			cellFive.appendChild(e2);
			slength = document.getElementById("customfieldorder" + t).length;
			for ( s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("customfieldorder" + t).options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("customfieldorder" + t).options[s].value );
				op.selected = true;
				e2.appendChild(op);
			}

		}

		function Validate()
		{
			var rege = /^\d+$/;
			var Ok; 
			var bDelete = false;
			var iSelected;
			var iCount;
			var z;

			// set the active tab to return to
			document.getElementById("activetab").value = tabView.get("activeIndex");

			// Check for a permit type name
			if (document.frmPermit.permittype.value == '')
			{
				alert("Please provide a permit type, then try saving again.");
				document.frmPermit.permittype.focus();
				return;
			}

			// Remove any extra spaces
			document.frmPermit.expirationdays.value = removeSpaces(document.frmPermit.expirationdays.value);
			//Remove commas that would cause problems in validation
			document.frmPermit.expirationdays.value = removeCommas(document.frmPermit.expirationdays.value);

			// Check for an Expiration Days and that it is formatted properly
			if (document.frmPermit.expirationdays.value != '')
			{
				Ok = rege.test(document.frmPermit.expirationdays.value);
				if (! Ok)
				{
					alert("The Days Until Expiration must be blank or a whole number.");
					document.frmPermit.expirationdays.focus();
					return;
				}
			}

			// Check each Fee type for duplicates
			for (t = 0; t <= parseInt(document.frmPermit.maxfeerows.value); t++)
			{
				// If the feetypeid exists
				if (document.getElementById("permitfeetypeid" + t))
				{
					// Get the selected fee type
					iSelected = document.getElementById("permitfeetypeid" + t).selectedIndex;
					if (iSelected > 0)
					{
						iCount = 0;
						for (z = 0; z <= parseInt(document.frmPermit.maxfeerows.value); z++)
						{
							if (document.getElementById("permitfeetypeid" + z))
							{
								if (iSelected == document.getElementById("permitfeetypeid" + z).selectedIndex)
								{
									iCount++;
								}
							}
						}
						if (iCount > 1)
						{
							tabView.set('activeIndex',1);
							alert("You have set some of the Fee Types to the same Fee Type choice.\nThey must have different Fee Types choices.\nPlease correct this and try saving again.");
							return;
						}
					}
				}
			}

			// Check each Inspection type for duplicates
			for (t = 0; t <= parseInt(document.frmPermit.maxinspectionrows.value); t++)
			{
				// If the inspectiontypeid exists
				if (document.getElementById("permitinspectiontypeid" + t))
				{
					// Get the selected inspection type
					iSelected = document.getElementById("permitinspectiontypeid" + t).selectedIndex;
					if (iSelected > 0)
					{
						iCount = 0;
						for (z = 0; z <= parseInt(document.frmPermit.maxinspectionrows.value); z++)
						{
							if (document.getElementById("permitinspectiontypeid" + z))
							{
								if (iSelected == document.getElementById("permitinspectiontypeid" + z).selectedIndex)
								{
									iCount++;
								}
							}
						}
						if (iCount > 1)
						{
							tabView.set('activeIndex',3);
							alert("You have set some of the Inspection Types to the same Inspection Type choice.\nThey must have different Inspection Types choices.\nPlease correct this and try saving again.");
							return;
						}
					}
				}
			}

			// Check each Review type for duplicates
			for (t = 0; t <= parseInt(document.frmPermit.maxreviewrows.value); t++)
			{
				// If the reviewtypeid exists
				if (document.getElementById("permitreviewtypeid" + t))
				{
					// Get the selected review type
					iSelected = document.getElementById("permitreviewtypeid" + t).selectedIndex;
					if (iSelected > 0)
					{
						iCount = 0;
						for (z = 0; z <= parseInt(document.frmPermit.maxreviewrows.value); z++)
						{
							if (document.getElementById("permitreviewtypeid" + z))
							{
								if (iSelected == document.getElementById("permitreviewtypeid" + z).selectedIndex)
								{
									iCount++;
								}
							}
						}
						if (iCount > 1)
						{
							tabView.set('activeIndex',2);
							alert("You have set some of the Review Types to the same Review Type choice.\nThey must have different Review Types choices.\nPlease correct this and try saving again.");
							return;
						}
					}
				}
			}

			// Check each Custom Field Type for duplicates
			for (t = 0; t <= parseInt(document.frmPermit.maxcustomfieldrows.value); t++)
			{
				// If the reviewtypeid exists
				if (document.getElementById("customfieldtypeid" + t))
				{
					// Get the selected review type
					iSelected = document.getElementById("customfieldtypeid" + t).selectedIndex;
					if (iSelected > 0)
					{
						iCount = 0;
						for (z = 0; z <= parseInt(document.frmPermit.maxcustomfieldrows.value); z++)
						{
							if (document.getElementById("customfieldtypeid" + z))
							{
								if (iSelected == document.getElementById("customfieldtypeid" + z).selectedIndex)
								{
									iCount++;
								}
							}
						}
						if (iCount > 1)
						{
							tabView.set('activeIndex',8);
							alert("You have set some of the Custom Fields to the same choice.\nEach Custom Field added must be for a different choice.\nPlease correct this and try saving again.");
							return;
						}
					}
				}
			}


			//alert("All was OK");
			// All is OK so submit
			document.frmPermit.submit();
		}

		function Delete() 
		{
			if (confirm("Do you wish to delete this permit type?"))
			{
				location.href="permittypedelete.asp?permittypeid=<%=iPermitTypeid%>";
			}
		}

		function RemoveFeeRows()
		{
			var iRow = 0;
			var tbl = document.getElementById("permittypefeetable");
			// Check the Fee rows for any selected for removal
			for (var t = 0; t <= parseInt(document.frmPermit.maxfeerows.value); t++)
			{
				// See if a row exists for this one
				if (document.getElementById("removefee" + t))
				{
					// The row exists so increment the row counter
					iRow++;
					// If it is marked for removal, remove it
					if (document.getElementById("removefee" + t).checked == true)
					{
						if (tbl.rows.length > 2)
						{
							// Remove the unwanted row
							tbl.deleteRow(iRow);
							// Decrement the row counter as we have one less row now
							iRow--;
						}
						else
						{
							// Down to one row, so just reset it to it's initial defaults
							document.getElementById("permitfeetypeid" + t).options[0].selected = true;
							document.getElementById("removefee" + t).checked = false;
							document.getElementById("isrequired" + t).checked = false;
							document.getElementById("displayorder" + t).options[0].selected = true;
						}
					}
				}
			}
		}

		function RemoveDocumentRows()
		{
			var iRow = 0;
			var tbl = document.getElementById("permittypedocumenttable");
			// Check the Inspection rows for any selected for removal
			for (var t = 0; t <= parseInt(document.frmPermit.maxdocumentrows.value); t++)
			{
				// See if a row exists for this one
				if (document.getElementById("removedocument" + t))
				{
					// The row exists so increment the row counter
					iRow++;
					// If it is marked for removal, remove it
					if (document.getElementById("removedocument" + t).checked == true)
					{
						if (tbl.rows.length > 2)
						{
							// Remove the unwanted row
							tbl.deleteRow(iRow);
							// Decrement the row counter as we have one less row now
							iRow--;
						}
						else
						{
							// Down to one row, so just reset it to it's initial defaults
							document.getElementById("removedocument" + t).checked = false;
							document.getElementById("documentid" + t).options[0].selected = true;
							document.getElementById("documentlabel" + t).value = '';
							document.getElementById("permitdocumentid" + t).value = '0';
						}
					}
				}
			}
		}

		function RemoveInspectionRows()
		{
			var iRow = 0;
			var tbl = document.getElementById("permittypeinspectiontable");
			// Check the Inspection rows for any selected for removal
			for (var t = 0; t <= parseInt(document.frmPermit.maxinspectionrows.value); t++)
			{
				// See if a row exists for this one
				if (document.getElementById("removeinspection" + t))
				{
					// The row exists so increment the row counter
					iRow++;
					// If it is marked for removal, remove it
					if (document.getElementById("removeinspection" + t).checked == true)
					{
						if (tbl.rows.length > 2)
						{
							// Remove the unwanted row
							tbl.deleteRow(iRow);
							// Decrement the row counter as we have one less row now
							iRow--;
						}
						else
						{
							// Down to one row, so just reset it to it's initial defaults
							document.getElementById("removeinspection" + t).checked = false;
							document.getElementById("permitinspectiontypeid" + t).options[0].selected = true;
							document.getElementById("permitinspectorid" + t).options[0].selected = true;
							document.getElementById("isfinal" + t).checked = false;
							document.getElementById("inspectionisrequired" + t).checked = false;
							//document.getElementById("scheduleddaysout" + t).value = '';
							document.getElementById("inspectionorder" + t).options[0].selected = true;
						}
					}
				}
			}
		}

		function RemoveReviewRows()
		{
			var iRow = 0;
			var tbl = document.getElementById("permittypereviewtable");
			// Check the Review rows for any selected for removal
			var iMaxReviews = document.frmPermit.maxreviewrows.value; 
			for (var t = 0; t <= parseInt(iMaxReviews); t++)
			{
				// See if a row exists for this one
				if (document.getElementById("removereview" + t))
				{
					// The row exists so increment the row counter
					iRow++;
					// If it is marked for removal, remove it
					if (document.getElementById("removereview" + t).checked == true)
					{
						if (tbl.rows.length > 2)
						{
							// Remove the unwanted row
							tbl.deleteRow(iRow);
							// Decrement the row counter as we have one less row now
							iRow--;
							document.frmPermit.maxreviewrows.value = parseInt(document.frmPermit.maxreviewrows.value) - 1;
						}
						else
						{
							// Down to one row, so just reset it to it's initial defaults  
							document.getElementById("removereview" + t).checked = false;
							document.getElementById("permitreviewtypeid" + t).options[0].selected = true;
							document.getElementById("permitreviewerid" + t).options[0].selected = true;
							document.getElementById("reviewisrequired" + t).checked = false;
							document.getElementById("notifyonrelease" + t).checked = true;
							document.getElementById("revieworder" + t).options[0].selected = true;
							document.frmPermit.maxreviewrows.value = 0;
						}
					}
				}
			}
		}

		function RemoveCustomFieldRows()
		{
			var iRow = 0;
			var tbl = document.getElementById("permittypecustomfieldstable");
			// Check the Review rows for any selected for removal
			var iMaxCustomFields = document.frmPermit.maxcustomfieldrows.value; 
			for (var t = 0; t <= parseInt(iMaxCustomFields); t++)
			{
				// See if a row exists for this one
				if (document.getElementById("removecustomfield" + t))
				{
					// The row exists so increment the row counter
					iRow++;
					// If it is marked for removal, remove it
					if (document.getElementById("removecustomfield" + t).checked == true)
					{
						if (tbl.rows.length > 2)
						{
							// Remove the unwanted row
							tbl.deleteRow(iRow);
							// Decrement the row counter as we have one less row now
							iRow--;
							document.frmPermit.maxcustomfieldrows.value = parseInt(document.frmPermit.maxcustomfieldrows.value) - 1;
						}
						else
						{
							// Down to one row, so just reset it to it's initial defaults  
							document.getElementById("removecustomfield" + t).checked = false;
							document.getElementById("customfieldtypeid" + t).options[0].selected = true;
							document.getElementById("customfieldorder" + t).options[0].selected = true;
							document.frmPermit.maxcustomfieldrows.value = 0;
						}
					}
				}
			}
		}

		function RemoveReviewAlertRows()
		{
			var iRow = 0;
			var tbl = document.getElementById("permittypereviewalerttable");
			// Check the Review alert rows for any selected for removal
			for (var t = 0; t <= parseInt(document.frmPermit.maxreviewalertrows.value); t++)
			{
				// See if a row exists for this one
				if (document.getElementById("removereviewalert" + t))
				{
					// The row exists so increment the row counter
					iRow++;
					// If it is marked for removal, remove it
					if (document.getElementById("removereviewalert" + t).checked == true)
					{
						if (tbl.rows.length > 2)
						{
							// Remove the unwanted row
							tbl.deleteRow(iRow);
							// Decrement the row counter as we have one less row now
							iRow--;
						}
						else
						{
							// Down to one row, so just reset it to it's initial defaults
							document.getElementById("removereviewalert" + t).checked = false;
							document.getElementById("permitalerttypeid" + t).options[0].selected = true;
							document.getElementById("notifyuserid" + t).options[0].selected = true;
						}
					}
				}
			}
		}

		function RemoveInspectionAlertRows()
		{
			var iRow = 0;
			var tbl = document.getElementById("permittypeinspectionalerttable");
			// Check the Review alert rows for any selected for removal
			for (var t = 0; t <= parseInt(document.frmPermit.maxinspectionalertrows.value); t++)
			{
				// See if a row exists for this one
				if (document.getElementById("removeinspectionalert" + t))
				{
					// The row exists so increment the row counter
					iRow++;
					// If it is marked for removal, remove it
					if (document.getElementById("removeinspectionalert" + t).checked == true)
					{
						if (tbl.rows.length > 2)
						{
							// Remove the unwanted row
							tbl.deleteRow(iRow);
							// Decrement the row counter as we have one less row now
							iRow--;
						}
						else
						{
							// Down to one row, so just reset it to it's initial defaults
							document.getElementById("removeinspectionalert" + t).checked = false;
							document.getElementById("permitinspectionalerttypeid" + t).options[0].selected = true;
							document.getElementById("notifyinspectoruserid" + t).options[0].selected = true;
						}
					}
				}
			}
		}

		function ClearIsFinalPick()
		{
			// if there is only one row then the length is undefined
			if (! document.frmPermit.isfinal.length)
			{
				document.frmPermit.isfinal.checked = false;
			}
			else 
			{
				// There are multiple rows and we can loop through them
				for (i=0; i < document.frmPermit.isfinal.length; i++) 
				{ 
					if (document.frmPermit.isfinal[i].checked == true) 
					{ 
						document.frmPermit.isfinal[i].checked = false;
						break;
					}
			   }
		   }
		}

		function CopyPermitType( iPermitTypeId )
		{
			if (confirm("Make a copy of this permit Type?"))
			{
				location.href="permittypecopy.asp?permittypeid=" + iPermitTypeId;
			}
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
		}

		function init()
		{
			setMaxLength();
		}

<%		If request("success") <> "" Then 
			'DisplayMessagePopUp %>
  		$( function() {
			$("#successmessage").show();
			$("#successmessage").fadeOut(2000);
		});
			

		<%End If %>

	//-->
	</script>

</head>
<body class="yui-skin-sam" onload="init()">

	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">
		<div class="gutters">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong><%=sTitle%> Permit Type</strong></font><br /><br />
				<a href="permittypelist.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0" />&nbsp;<%=langBackToStart%></a>
			</p>
			<!--END: PAGE TITLE-->

		<!--BEGIN: EDIT FORM-->
		<div id="functionlinks">
<%		If CLng(iPermitTypeid) = CLng(0) Then %>
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" id="savebutton" value="Create" /><br />
<%		Else %>
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" id="savebutton" value="Save Changes" /> &nbsp; &nbsp;
     			<div class="dropdown right">
  				<button class="ui-button ui-widget ui-corner-all dd-green"><i class="fa fa-bars" aria-hidden="true"></i> Tools</button>
  				<div class="dropdown-content">
					<a href="javascript:Delete();">Delete</a>
					<a href="javascript:CopyPermitType(<%=iPermitTypeid%>;)">Copy Permit Type</a>
<%			If request("success") <> "" Then %>
					<a href="javascript:Another();">Create Another</a>
<%			End If		%>
				</div>
			</div>
			<br />
<%		End If %>
		</div>

		<form name="frmPermit" action="permittypeupdate.asp" method="post">
			<input type="hidden" name="permittypeid" value="<%=iPermitTypeid%>" />
			<input type="hidden" name="activetab" id="activetab" value="<%=iActiveTabId%>" />
		
		<p>
			Permit Type Prefix: &nbsp; <input type="text" id="permittype" name="permittype" value="<%=sPermitType%>" size="50" maxlength="50" />
			 &nbsp;&nbsp; 
			<!--<input type="checkbox" id="isbuildingpermittype" name="isbuildingpermittype" <%=sIsBuildingPermitType%> /> Building Permit Type -->
			Category Type: <% ShowPermitCategoryPicks iPermitCategoryId %>
		</p>
		<p>
			Permit Type Description: &nbsp; <input type="text" id="permittypedesc" name="permittypedesc" value="<%=sPermitTypeDesc%>" size="100" maxlength="150" />
		</p>
		<p>
			Address/Location Requirement: &nbsp; <% ShowLocationRequirements iPermitLocationRequirementId	%>
		</p>
		
		
		<div id="demo" class="yui-navset">
			<ul class="yui-nav">
				<li><a href="#tab1"><em>General</em></a></li>
				<li><a href="#tab2"><em>Fees</em></a></li>
				<li><a href="#tab3"><em>Reviews</em></a></li>
				<li><a href="#tab4"><em>Inspections</em></a></li>
				<li><a href="#tab5"><em>Permit Document</em></a></li>
				<li><a href="#tab6"><em>Invoice</em></a></li>
				<li><a href="#tab7"><em>Documents</em></a></li>
				<li><a href="#tab8"><em>Detail Fields</em></a></li>
				<!--li><a href="#tab9"><em>Custom Fields</em></a></li-->
			</ul>            
			<div class="yui-content">

				<div id="tab1"> <!-- General -->
					<p><br />
						<p>
							Days Until Expiration: &nbsp;&nbsp; <input type="text" id="expirationdays" name="expirationdays" value="<%=sExpirationDays%>" size="4" maxlength="4" />
						</p>
						<p>
							Permit Number Prefix: &nbsp;&nbsp; <select name="permitnumberprefix">
																	<% ShowPermitNumberPrefixes sPermitNumberPrefix %>
																</select>
						</p>

						<p>Use Type: &nbsp;&nbsp; <% ShowUseTypes iUseTypeId %></p>

						<p>Required Licenses: &nbsp;
							
								<% ShowlicenseTypes iPermitTypeId, "1"	%>
						</p>
						<p>Display Licenses: &nbsp;
							
								<% ShowlicenseTypes iPermitTypeId, "0"	%>
						</p>
						<p>
							<input type="checkbox" name="hasco" <%=sHasCo%> /> This Permit Type Has a Certificate of Occupancy
						</p>
						<p>
							<input type="checkbox" name="hastempco" <%=sHasTempCo%> /> This Permit Type Has a Temporary Certificate of Occupancy
						</p>
					</p>
				</div>

				<div id="tab2"> <!-- Fees -->
					<p><br />
						<!-- <strong>Fees</strong><br /><br /> -->
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Add Row" id="addfeebutton" onClick="NewFeeRow()" /> &nbsp;&nbsp; 
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Remove Selected" id="removefeebutton" onClick="RemoveFeeRows()" /> &nbsp;&nbsp;
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Edit Fee Types" onClick="EditFeeTypes()" />
						

						<table id="permittypefeetable" border="0" cellpadding="0" cellspacing="0">
							<tr><th>&nbsp;</th><th>Fee Type</th><th>Required</th><th>Display<br />Order</th></tr>
<%							iMaxFeeRows = ShowFeeTypeTable( iPermitTypeid ) %>
						</table>

						<input type="hidden" name="maxfeerows" value="<%=iMaxFeeRows%>" />
						* Select Fee Types only once
					</p>
				</div>

				<div id="tab3"> <!-- Reviews -->
					<p><br />
					<table><tr><td>
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Add Row" id="addreviewbutton" onClick="NewReviewRow()" /> &nbsp;&nbsp; 
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Remove Selected" id="removereviewbutton" onClick="RemoveReviewRows()" /> &nbsp;&nbsp;
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Edit Permit Review Types" onClick="EditPermitReviewTypes();" />
						<table id="permittypereviewtable" border="0" cellpadding="0" cellspacing="0">
							<tr><th>&nbsp;</th><th>Review Type*</th><th>Reviewer</th><th>Notify On<br />Release</th><th>Required</th><th>Order</th></tr>
<%							iMaxReviewRows = ShowReviewTypeTable( iPermitTypeid ) %>
						</table>
						<input type="hidden" name="maxreviewrows" value="<%=iMaxReviewRows%>" />
						* Select Review Types only once

						<br /><br />
						<input type="checkbox" id="attachmentrevieweralert" name="attachmentrevieweralert" <%=sAttachmentReviewerAlert%> /> 
						Notify all reviewers when a new attachment has been uploaded
						<br /><br />

						<fieldset class="permittype">
							<legend>Review Alerts</legend>
							<input type="button" class="button ui-button ui-widget ui-corner-all" value="Add Row" id="addreviewalertbutton" onClick="NewReviewAlertRow()" /> &nbsp;&nbsp; 
							<input type="button" class="button ui-button ui-widget ui-corner-all" value="Remove Selected" id="removereviewalertbutton" onClick="RemoveReviewAlertRows()" />
							<table id="permittypereviewalerttable" border="0" cellpadding="0" cellspacing="0">
								<tr><th>&nbsp;</th><th>Review Alert Type</th><th>Notify Reviewer</th></tr>
<%								iMaxReviewAlertRows = ShowReviewAlertTypeTable( iPermitTypeid ) %>
							</table>
							<input type="hidden" name="maxreviewalertrows" value="<%=iMaxReviewAlertRows%>" />
						</fieldset>
					</td></tr></table>
					</p>
				</div>

				<div id="tab4"> <!-- Inspections -->
					<p><br />
					<table><tr><td>
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Add Row" id="addinspectionbutton" onClick="NewInspectionRow()" /> &nbsp;&nbsp; 
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Remove Selected" id="removeinspectionbutton" onClick="RemoveInspectionRows()" /> &nbsp;&nbsp;
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Clear Final Inspection Pick" onClick="ClearIsFinalPick()" />
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Edit Permit Inspection Types" onClick="EditPermitInspectionTypes();" />
						<table id="permittypeinspectiontable" border="0" cellpadding="0" cellspacing="0">
							<tr><th>&nbsp;</th><th>Inspection Type</th><th>Inspector</th><th>Required</th><th>Final<br />Inspection</th><th>Order</th></tr>
<%							iMaxInspRows = ShowInspectionTypeTable( iPermitTypeid ) %>
						</table>
						<input type="hidden" name="maxinspectionrows" value="<%=iMaxInspRows%>" />
						* Select Inspection Types only once

						<br /><br />
						<fieldset class="permittype">
							<legend>Inspection Alerts</legend>
							<input type="button" class="button ui-button ui-widget ui-corner-all" value="Add Row" id="addinspectionalertbutton" onClick="NewInspectionAlertRow()" /> &nbsp;&nbsp; 
							<input type="button" class="button ui-button ui-widget ui-corner-all" value="Remove Selected" id="removeinspectionalertbutton" onClick="RemoveInspectionAlertRows()" />
							<table id="permittypeinspectionalerttable" border="0" cellpadding="0" cellspacing="0">
								<tr><th>&nbsp;</th><th>Inspection Alert Type</th><th>Notify Inspector</th></tr>
<%								iMaxInspectionAlertRows = ShowInspectionAlertTypeTable( iPermitTypeid ) %>
							</table>
							<input type="hidden" name="maxinspectionalertrows" value="<%=iMaxInspectionAlertRows%>" />
						</fieldset>

					</td></tr></table>
					</p>
				</div>

				<div id="tab5"> <!-- Permit Document -->
					<p><br />
						<strong>Select the <!--Method or--> Document that will be used for generating the permit document for this permit type.</strong><br /><br />
						<% ShowPermitDocuments iDocumentId	%>
					</p>
					<!--p><br />
						<strong>If you have not selected a custom document, the following options can be applied to the permit document for this specific permit type. <br />You can use simple HTML in the text fields for formatting the output.</strong>
					</p>
					<p>
						<span class="permittypeeditlabels">Permit Logo:</span><br />
						<input type="text" id="permitlogo" name="permitlogo" value="<%=sPermitLogo%>" size="140" maxlength="250" /><br />
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Select Logo" onclick="doPicker('frmPermit.permitlogo');" />
<%						If sPermitLogo <> "" Then	%>
							<br /><br /><img src="<%=sPermitLogo%>" alt="permit logo" border="0" align="absmiddle" />
<%						End If		%>
					</p>
					<p>
						<span class="permittypeeditlabels">Permit Title:</span><br />
						<textarea name="permittitle" id="permittitle" maxlength="200" wrap="soft"><%=sPermitTitle%></textarea>
					</p>

					<p>
						<span class="permittypeeditlabels">Permit Subtitle:</span><br />
						<textarea name="permitsubtitle" id="permitsubtitle" maxlength="200" wrap="soft"><%=sPermitSubTitle%></textarea>
					</p>
					
					
					<p>
						<span class="permittypeeditlabels">Right Title Info:</span><br />
						<textarea name="permitrighttitle" id="permitrighttitle" maxlength="200" wrap="soft"><%=sPermitRightTitle%></textarea>
					</p>
					<p>
						<span class="permittypeeditlabels">Bottom Title Info:</span><br />
						<textarea name="permittitlebottom" id="permittitlebottom" maxlength="500" wrap="soft"><%=sPermitTitleBottom%></textarea>
					</p>
					<p>
						<span class="permittypeeditlabels">Additional Information:</span><br />
						<textarea name="additionalfooterinfo" id="additionalfooterinfo" maxlength="500" wrap="soft"><%=sAdditionalFooterInfo%></textarea>
					</p>
					<p>
						<span class="permittypeeditlabels">Permit Footer:</span><br />
						<textarea name="permitfooter" id="permitfooter" maxlength="1300" wrap="soft"><%=sPermitFooter%></textarea>
					</p>
					<p>
						<span class="permittypeeditlabels">Permit Subfooter:</span><br />
						<textarea name="permitsubfooter" id="permitsubfooter" maxlength="500" wrap="soft"><%=sPermitSubFooter%></textarea>
					</p>
					<p>
						<span class="permittypeeditlabels">Permit Approving Official:</span><br />
						<textarea name="approvingofficial" id="approvingofficial" maxlength="200" wrap="soft"><%=sApprovingOfficial%></textarea>
					</p>
					<table cellspacing="0" cellpadding="2" border="0" id="showpermittypepicks">
					<tr>
						<td nowrap="nowrap">
							<input type="checkbox" name="showconstructiontype" <%=sShowConstructionType%> /> Show Type of Construction
						</td>
						<td nowrap="nowrap">
							<input type="checkbox" name="showoccupancytype" <%=sShowOccupancyType%> /> Show Occupancy Type
						</td>
						<td nowrap="nowrap">
							<input type="checkbox" name="showoccupancyuse" <%=sShowOccupancyUse%> /> Show Occupancy Use Group
						</td>
					</tr>
					<tr>
						<td nowrap="nowrap">
							<input type="checkbox" name="showfeetotal" <%=sShowFeeTotal%> /> Show Fee Total
						</td>
						<td nowrap="nowrap">
							<input type="checkbox" name="showfeetypetotals" <%=sShowFeeTypeTotals%> /> Show Fee Type Totals
						</td>
						<td nowrap="nowrap">
							<input type="checkbox" name="showpayments" <%=sShowPayments%> /> Show Payments
						</td>
					</tr>
					<tr>
						<td nowrap="nowrap">
							<input type="checkbox" name="showtotalsqft" <%=sShowTotalSqFt%> /> Show Total Sq Ft Only
						</td>
						<td nowrap="nowrap">
							<input type="checkbox" name="showfootages" <%=sShowFootages%> /> Show All Sq Footage Values
						</td>
						<td nowrap="nowrap">
							<input type="checkbox" name="showjobvalue" <%=sShowJobValue%> /> Show Job Value
						</td>
						
					</tr>
					<tr>
						<td nowrap="nowrap">
							<input type="checkbox" name="showapprovedas" <%=sShowApprovedAs%> /> Show Approved As
						</td>
						<td nowrap="nowrap">
							<input type="checkbox" name="showproposeduse" <%=sShowProposedUse%> /> Show Proposed Use
						</td>
						<td nowrap="nowrap">
							<input type="checkbox" name="showworkdesc" <%=sShowWorkDesc%> /> Show Description of Work
						</td>
					</tr>
					<tr>
						<td nowrap="nowrap">
							<input type="checkbox" name="showelectricalcontractor" <%=sShowElectricalContractor%> /> Show Electrical Contractor
						</td>
						<td nowrap="nowrap">
							<input type="checkbox" name="showmechanicalcontractor" <%=sShowMechanicalContractor%> /> Show Mechanical Contractor
						</td>
						<td nowrap="nowrap">
							<input type="checkbox" name="showplumbingcontractor" <%=sShowPlumbingContractor%> /> Show Plumbing Contractor
						</td>
					</tr>
					<tr>
						<td nowrap="nowrap">
							<input type="checkbox" name="showprimarycontact" <%=sShowPrimaryContact%> /> Show Primary Contact
						</td>
						<td nowrap="nowrap">
							<input type="checkbox" name="showplansby" <%=sShowPlansBy%> /> Show Plans By
						</td>
						<td nowrap="nowrap">
							<input type="checkbox" name="showapplicantlicense" <%=sShowApplicantLicense%> /> Show Applicant License
						</td>
					</tr>
					<tr>
						<td nowrap="nowrap">
							<input type="checkbox" name="showcounty" <%=sShowCounty%> /> Show <%=oAddressOrg.GetOrgDisplayName("address grouping field")%>
						</td>
						<td nowrap="nowrap">
							<input type="checkbox" name="showparcelid" <%=sShowParcelid%> /> Show Parcel Id
						</td>
						<td nowrap="nowrap">
							<input type="checkbox" name="listfixtures" <%=sListFixtures%> /> List Fixtures
						</td>
					</tr>
					</table-->
				</div>
				<div id="tab6"> <!-- Permit Invoice -->
					<p><br />
						<strong>The following options can be applied to the invoice for this specific permit type. You can use simple HTML in the text fields for formatting the output.</strong>
					</p>

					<table cellspacing="0" cellpadding="2" border="0" id="showpicks">
					<tr>
						<td>
							<input type="checkbox" name="groupbyinvoicecategories" <%=sGroupByInvoiceCategories%> /> Group By Invoice Categories with Subtotals
						</td>
					</tr>
					</table>

					<p>
						<span class="permittypeeditlabels">Invoice Logo:</span><br />
						<input type="text" id="invoicelogo" name="invoicelogo" value="<%=sInvoiceLogo%>" size="140" maxlength="250" /><br />
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Select Logo" onclick="doPicker('frmPermit.invoicelogo');" />
<%						If sInvoiceLogo <> "" Then	%>
							<br /><br /><img src="<%=sInvoiceLogo%>" alt="invoice logo" border="0" align="absmiddle" />
<%						End If		%>
					</p>

					<p>
						<span class="permittypeeditlabels">Invoice Header:</span><br />
						<textarea class="headertextarea" name="invoiceheader" maxlength="200" wrap="soft"><%=sInvoiceHeader%></textarea>
					</p>

				</div>
				
				<div id="tab7"> <!-- Other Documents -->
					<p><br />
					<table><tr><td>
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Add A Document" id="adddocumentbutton" onClick="NewDocumentRow()" /> &nbsp;&nbsp; 
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Remove Selected Documents" id="removedocumentbutton" onClick="RemoveDocumentRows()" /> &nbsp;&nbsp;
						<table id="permittypedocumenttable" border="0" cellpadding="0" cellspacing="0">
							<tr><th>&nbsp;</th><th>Label</th><th>Document</th></tr>
<%							iMaxDocRows = ShowDocumentTable( iPermitTypeid ) %>
						</table>
						<input type="hidden" id="maxdocumentrows" name="maxdocumentrows" value="<%=iMaxDocRows%>" />
					</td></tr></table>
					</p>
				</div>

				<div id="tab8"> <!-- Fields Included on details page -->
					<p><br />
					Select the fields that should be included on the permits details tab for this permit type.

					<table id="permittypedetailfieldstable" border="0" cellpadding="0" cellspacing="0">
						<tr><th>&nbsp;</th><th>Details Field</th></tr>
<%						ShowDetailFields iPermitTypeid  %>
					</table>
					<table><tr><td>
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Add A Field" id="addcustomfieldbutton" onClick="NewCustomFieldRow()" /> &nbsp;&nbsp; 
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Remove Selected Fields" id="removecustomfieldbutton" onClick="RemoveCustomFieldRows()" /> &nbsp;&nbsp;
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Edit Custom Field Types" onClick="EditPermitCustomFieldTypes()" /> &nbsp;&nbsp;
						<table id="permittypecustomfieldstable" border="0" cellpadding="0" cellspacing="0">
							<tr><th>&nbsp;</th><th>Custom Field</th>
<%							If bOrgHasPermitTypeReport Then		%>							
								<th>Include In<br /><%= GetFeatureName( "permit type report" )%></th>
<%							End If		%>
							<th>Order</th></tr>
<%								iMaxCustomFieldRows = ShowCustomFieldTypeTable( iPermitTypeid, bOrgHasPermitTypeReport ) %>
							</table>
						<input type="hidden" id="maxcustomfieldrows" name="maxcustomfieldrows" value="<%=iMaxCustomFieldRows%>" />
					</td></tr></table>
					</p>
				</div>

				<div id="tab9"> <!-- Custom Fields -->
					<p><br />
					</p>
				</div>

			</div>
		</div>

		<p>
<%		If CLng(iPermitTypeid) = CLng(0) Then %>
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" id="savebutton" value="Create" />
<%		Else %>
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" id="savebutton" value="Save Changes" /> 
<%		End If	%>
		</p>
		</form>
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
Set oAddressOrg = Nothing 

'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void GetPermitType iPermitTypeid 
'--------------------------------------------------------------------------------------------------
Sub GetPermitType( ByVal iPermitTypeid )
	Dim sSql, oRs

	sSql = "SELECT permittypeid, ISNULL(permittype,'') AS permittype, ISNULL(permittypedesc,'') AS permittypedesc, permitcategoryid, "
	sSql = sSql & " expirationdays, permitnumberprefix, publicdescription, ISNULL(permittitle,'') AS permittitle, "
	sSql = sSql & " ISNULL(additionalfooterinfo,'') AS additionalfooterinfo, ISNULL(approvingofficial,'') AS approvingofficial, "
	sSql = sSql & " ISNULL(permitsubtitle,'') AS permitsubtitle, ISNULL(permitrighttitle,'') AS permitrighttitle, "
	sSql = sSql & " ISNULL(permittitlebottom,'') AS permittitlebottom, ISNULL(permitfooter,'') AS permitfooter, "
	sSql = sSql & " ISNULL(permitsubfooter,'') AS permitsubfooter, listfixtures, showconstructiontype, showfeetotal, "
	sSql = sSql & " showoccupancytype, showjobvalue, showworkdesc, showfootages, showproposeduse, showothercontacts, "
	sSql = sSql & " ISNULL(permitlogo,'') AS permitlogo, groupbyinvoicecategories, ISNULL(invoicelogo,'') AS invoicelogo,  "
	sSql = sSql & " ISNULL(invoiceheader,'') AS invoiceheader, showelectricalcontractor, showmechanicalcontractor, "
	sSql = sSql & " showplumbingcontractor, showapplicantlicense, showcounty, showparcelid, showplansby, showprimarycontact, "
	sSql = sSql & " ISNULL(usetypeid,0) AS usetypeid, hastempco, hasco, showapprovedasontco, showapprovedasonco, "
	sSql = sSql & " showconsttypeontco, showconsttypeonco, showocctypeontco, showocctypeonco, showoccupantsontco, showoccupantsonco, "
	sSql = sSql & " ISNULL(tempcologo,'') AS tempcologo, ISNULL(cologo,'') AS cologo, ISNULL(tempcotitle,'') AS tempcotitle, "
	sSql = sSql & " ISNULL(tempcosubtitle,'') AS tempcosubtitle, ISNULL(cotitle,'') AS cotitle, ISNULL(cosubtitle,'') AS cosubtitle, "
	sSql = sSql & " ISNULL(tempcoaddress,'') AS tempcoaddress, ISNULL(coaddress,'') AS coaddress, ISNULL(tempcotoptext,'') AS tempcotoptext, "
	sSql = sSql & " ISNULL(cotoptext,'') AS cotoptext, ISNULL(tempcobottomtext,'') AS tempcobottomtext, ISNULL(cobottomtext,'') AS cobottomtext, "
	sSql = sSql & " ISNULL(tempcocoderef,'') AS tempcocoderef, ISNULL(cocoderef,'') AS cocoderef, ISNULL(tempcoapproval,'') AS tempcoapproval, "
	sSql = sSql & " ISNULL(coapproval,'') AS coapproval, ISNULL(tempcofooter,'') AS tempcofooter, ISNULL(cofooter,'') AS cofooter, "
	sSql = sSql & " ISNULL(tempcosubfooter,'') AS tempcosubfooter, ISNULL(cosubfooter,'') AS cosubfooter, showtotalsqft, showapprovedas, "
	sSql = sSql & " showfeetypetotals, showoccupancyuse, showpayments, ISNULL(documentid,0) AS documentid, permitlocationrequirementid, "
	sSql = sSql & " attachmentrevieweralert "
	sSql = sSql & " FROM egov_permittypes WHERE permittypeid = " & iPermitTypeid
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sPermitType = Replace(oRs("permittype"),"""","&quot;")
		sPermitTypeDesc = Replace(oRs("permittypedesc"),"""","&quot;")
		iPermitCategoryId = oRs("permitcategoryid")
		iPermitLocationRequirementId = oRs("permitlocationrequirementid")
'		If oRs("isbuildingpermittype") Then 
'			sIsBuildingPermitType = " checked=""checked"" "
'		Else
'			sIsBuildingPermitType = ""
'		End If 
		sExpirationDays = oRs("expirationdays")
		sPermitNumberPrefix = oRs("permitnumberprefix")
		sPublicDescription = oRs("publicdescription")
		sPermitTitle = oRs("permittitle")
		sAdditionalFooterInfo = oRs("additionalfooterinfo")
		sApprovingOfficial = oRs("approvingofficial")
		sPermitSubTitle = oRs("permitsubtitle")
		sPermitRightTitle = oRs("permitrighttitle")
		sPermitTitleBottom = oRs("permittitlebottom")
		sPermitFooter = oRs("permitfooter")
		sPermitSubFooter = oRs("permitsubfooter")
		sPermitLogo = oRs("permitlogo")
		sInvoiceLogo = oRs("invoicelogo")
		sInvoiceHeader = oRs("invoiceheader")
		If oRs("listfixtures") Then 
			sListFixtures = " checked=""checked"" "
		End If 
		If oRs("showconstructiontype") Then 
			sShowConstructionType = " checked=""checked"" "
		End If 
		If oRs("showfeetotal") Then 
			sShowFeeTotal = " checked=""checked"" "
		End If 
		If oRs("showoccupancytype") Then 
			sShowOccupancyType = " checked=""checked"" "
		End If 
		If oRs("showjobvalue") Then 
			sShowJobValue = " checked=""checked"" "
		End If 
		If oRs("showworkdesc") Then 
			sShowWorkDesc = " checked=""checked"" "
		End If 
		If oRs("showfootages") Then 
			sShowFootages = " checked=""checked"" "
		End If 
		If oRs("showproposeduse") Then 
			sShowProposedUse = " checked=""checked"" "
		End If 
		If oRs("showothercontacts") Then 
			sShowOtherContacts = " checked=""checked"" "
		End If 
		If oRs("groupbyinvoicecategories") Then 
			sGroupByInvoiceCategories = " checked=""checked"" "
		End If 
		If oRs("showelectricalcontractor") Then 
			sShowElectricalContractor = " checked=""checked"" "
		End If 
		If oRs("showmechanicalcontractor") Then 
			sShowMechanicalContractor = " checked=""checked"" "
		End If 
		If oRs("showplumbingcontractor") Then 
			sShowPlumbingContractor = " checked=""checked"" "
		End If 
		If oRs("showapplicantlicense") Then 
			sShowApplicantLicense = " checked=""checked"" "
		End If 
		If oRs("showcounty") Then 
			sShowCounty = " checked=""checked"" "
		End If
		If oRs("showparcelid") Then 
			sShowParcelid = " checked=""checked"" "
		End If
		If oRs("showplansby") Then
			sShowPlansBy = " checked=""checked"" "
		End If 
		If oRs("showprimarycontact") Then
			sShowPrimaryContact = " checked=""checked"" "
		End If 
		iUseTypeId = oRs("usetypeid")
		If oRs("hastempco") Then 
			sHasTempCO = " checked=""checked"" "
		End If 
		If oRs("hasco") Then 
			sHasCO = " checked=""checked"" "
		End If 
		If oRs("showapprovedasontco") Then 
			sShowApprovedAsOnTCO = " checked=""checked"" "
		End If 
		If oRs("showapprovedasonco") Then 
			sShowApprovedAsOnCO = " checked=""checked"" "
		End If 
		If oRs("showconsttypeontco") Then 
			sShowConstTypeOnTCO = " checked=""checked"" "
		End If 
		If oRs("showconsttypeonco") Then 
			sShowConstTypeOnCO = " checked=""checked"" "
		End If 
		If oRs("showocctypeontco") Then 
			sShowOccTypeOnTCO = " checked=""checked"" "
		End If 
		If oRs("showocctypeonco") Then 
			sShowOccTypeOnCO = " checked=""checked"" "
		End If 
		If oRs("showoccupantsontco") Then 
			sShowOccupantsOnTCO = " checked=""checked"" "
		End If 
		If oRs("showoccupantsonco") Then 
			sShowOccupantsOnCO = " checked=""checked"" "
		End If 
		sTempCOLogo = oRs("tempcologo")
		sCOLogo = oRs("cologo")
		sTempCOTitle = oRs("tempcotitle")
		sTempCOSubTitle = oRs("tempcosubtitle")
		sCOTitle = oRs("cotitle")
		sCOSubTitle = oRs("cosubtitle")
		sTempCOAddress = oRs("tempcoaddress")
		sCOAddress = oRs("coaddress")
		sTempCOTopText = oRs("tempcotoptext")
		sCOTopText = oRs("cotoptext")
		sTempCOBottomText = oRs("tempcobottomtext")
		sCOBottomText = oRs("cobottomtext")
		sTempCOCodeRef = oRs("tempcocoderef")
		sCOCodeRef = oRs("cocoderef")
		sTempCOApproval = oRs("tempcoapproval")
		sCOApproval = oRs("coapproval")
		sTempCOFooter = oRs("tempcofooter")
		sCOFooter = oRs("cofooter")
		sTempCOSubFooter = oRs("tempcosubfooter")
		sCOSubFooter = oRs("cosubfooter")
		If oRs("showtotalsqft") Then 
			sShowTotalSqFt = " checked=""checked"" "
		End If 
		If oRs("showapprovedas") Then 
			sShowApprovedAs = " checked=""checked"" "
		End If 
		If oRs("showfeetypetotals") Then 
			sShowFeeTypeTotals = " checked=""checked"" "
		End If 
		If oRs("showoccupancyuse") Then 
			sShowOccupancyUse = " checked=""checked"" "
		End If 
		If oRs("showpayments") Then
			sShowPayments = " checked=""checked"" "
		End If 
		iDocumentId = oRs("documentid")
		If oRs("attachmentrevieweralert") Then
			sAttachmentReviewerAlert = " checked=""checked"" "
		End If 

	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' integer ShowInspectionTypeTable( iPermitTypeid )
'--------------------------------------------------------------------------------------------------
Function ShowInspectionTypeTable( ByVal iPermitTypeid )
	Dim oRs, sSql, iRowCount, sRowClass, iMaxRows

	iRowCount = -1
	iMaxRows = GetMaxInspectionRows( iPermitTypeid ) 

	sSql = "SELECT permitinspectiontypeid, ISNULL(inspectoruserid,0) AS inspectoruserid, isrequired, isfinal, inspectionorder "
	sSql = sSql & " FROM egov_permittypes_to_permitinspectiontypes "
	sSql = sSql & " WHERE permittypeid = " & iPermitTypeid
	sSql = sSql & " ORDER BY inspectionorder"

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
			response.write "<input type=""checkbox"" id=""removeinspection" & iRowCount & """ name=""removeinspection" & iRowCount & """ /></td><td>"
			ShowInspectionTypePicks oRs("permitinspectiontypeid"), iRowCount
			response.write "</td>"
			response.write "<td align=""center"">"
			ShowInspectorPicks oRs("inspectoruserid"), iRowCount
			response.write "</td>"
			response.write "<td align=""center""><input type=""checkbox"" id=""inspectionisrequired" & iRowCount & """ name=""inspectionisrequired" & iRowCount & """"
			If oRs("isrequired") Then
				response.write " checked=""checked"" "
			End If 
			response.write " />"
			response.write "</td>"
			response.write "<td align=""center""><input type=""radio"" id=""isfinal" & iRowCount & """ name=""isfinal"" value=""" & iRowCount & """"
			If oRs("isfinal") Then
				response.write " checked=""checked"" "
			End If 
			response.write " />"
			response.write "</td>"
			response.write "<td align=""center"">"
			showDisplayOrder "inspectionorder", iMaxRows, oRs("inspectionorder"), iRowCount
			response.write "</td>"
			response.write "</tr>"
			oRs.MoveNext 
		Loop 
	Else
		' put in a starter row.
		iRowCount = 0
		response.write vbcrlf & "<tr><td class=""firstcell"">"
		response.write "<input type=""checkbox"" id=""removeinspection" & iRowCount & """ name=""removeinspection" & iRowCount & """ /></td><td>"
		ShowInspectionTypePicks 0, iRowCount
		response.write "</td>"
		response.write "<td align=""center"">"
		ShowInspectorPicks 0, iRowCount
		response.write "</td>"
		response.write "<td align=""center""><input type=""checkbox"" id=""inspectionisrequired" & iRowCount & """ name=""inspectionisrequired" & iRowCount & """ />"
		response.write "</td>"
		response.write "<td align=""center""><input type=""radio"" id=""isfinal" & iRowCount & """ name=""isfinal"" value=""" & iRowCount & """ />"
		response.write "</td>"
		response.write "<td align=""center"">"
		showDisplayOrder "inspectionorder", 1, 1, iRowCount
		response.write "</td>"
		response.write "</tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowInspectionTypeTable = iRowCount

End Function 


'--------------------------------------------------------------------------------------------------
' integer ShowFeeTypeTable( iPermitTypeid )
'--------------------------------------------------------------------------------------------------
Function ShowFeeTypeTable( ByVal iPermitTypeid )
	Dim oRs, sSql, iRowCount, sRowClass, iMaxRows

	iRowCount = -1
	iMaxRows = GetMaxFeeRows( iPermitTypeid ) 

	sSql = "SELECT permitfeetypeid, isrequired, displayorder "
	sSql = sSql & " FROM egov_permittypes_to_permitfeetypes "
	sSql = sSql & " WHERE permittypeid = " & iPermitTypeid
	sSql = sSql & " ORDER BY displayorder"

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
			response.write "<input type=""checkbox"" id=""removefee" & iRowCount & """ name=""removefee" & iRowCount & """ /></td><td>"
			ShowFeeTypePicks oRs("permitfeetypeid"), iRowCount
			response.write "</td>"
			response.write "<td align=""center""><input type=""checkbox"" id=""isrequired" & iRowCount & """ name=""isrequired" & iRowCount & """"
			If oRs("isrequired") Then
				response.write " checked=""checked"" "
			End If 
			response.write " /></td>"
			response.write "<td align=""center"">"
			showDisplayOrder "displayorder", iMaxRows, oRs("displayorder"), iRowCount
			response.write "</td>"
			response.write "</tr>"
			oRs.MoveNext 
		Loop 
	Else
		' put in a starter row.
		iRowCount = 0
		response.write vbcrlf & "<tr><td class=""firstcell"">"
		response.write "<input type=""checkbox"" id=""removefee" & iRowCount & """ name=""removefee" & iRowCount & """ /></td><td>"
		ShowFeeTypePicks 0, iRowCount
		response.write "</td>"
		response.write "<td align=""center""><input type=""checkbox"" id=""isrequired" & iRowCount & """ name=""isrequired" & iRowCount & """"
		response.write " /></td>"
		response.write "<td align=""center"">"
		showDisplayOrder "displayorder", 1, 1, iRowCount
		response.write "</td>"
		response.write "</tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowFeeTypeTable = iRowCount

End Function 


'--------------------------------------------------------------------------------------------------
' integer ShowReviewTypeTable( iPermitTypeid )
'--------------------------------------------------------------------------------------------------
Function ShowReviewTypeTable( ByVal iPermitTypeid )
	Dim oRs, sSql, iRowCount, sRowClass, iMaxRows

	iRowCount = -1
	iMaxRows = GetMaxReviewRows( iPermitTypeid ) 

	sSql = "SELECT permitreviewtypeid, revieweruserid, isrequired, revieworder, notifyonrelease "
	sSql = sSql & " FROM egov_permittypes_to_permitreviewtypes "
	sSql = sSql & " WHERE permittypeid = " & iPermitTypeid
	sSql = sSql & " ORDER BY revieworder"

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
			response.write "<input type=""checkbox"" id=""removereview" & iRowCount & """ name=""removereview" & iRowCount & """ /></td><td>"
			ShowReviewTypePicks oRs("permitreviewtypeid"), iRowCount
			response.write "</td>"
			response.write "<td align=""center"">"
			ShowReviewerPicks oRs("revieweruserid"), iRowCount
			response.write "</td>"
			response.write "<td align=""center""><input type=""checkbox"" id=""notifyonrelease" & iRowCount & """ name=""notifyonrelease" & iRowCount & """"
			If oRs("notifyonrelease") Then
				response.write " checked=""checked"" "
			End If 
			response.write " /></td>"
			response.write "<td align=""center""><input type=""checkbox"" id=""reviewisrequired" & iRowCount & """ name=""reviewisrequired" & iRowCount & """"
			If oRs("isrequired") Then
				response.write " checked=""checked"" "
			End If 
			response.write " /></td>"
			response.write "<td align=""center"">"
			showDisplayOrder "revieworder", iMaxRows, oRs("revieworder"), iRowCount
			response.write "</td>"
			response.write "</tr>"
			oRs.MoveNext 
		Loop 
	Else
		' put in a starter row.
		iRowCount = 0
		response.write vbcrlf & "<tr><td class=""firstcell"">"
		response.write "<input type=""checkbox"" id=""removereview" & iRowCount & """ name=""removereview" & iRowCount & """ /></td><td>"
		ShowReviewTypePicks 0, iRowCount
		response.write "</td>"
		response.write "<td align=""center"">"
		ShowReviewerPicks 0, iRowCount
		response.write "</td>"
		response.write "<td align=""center""><input type=""checkbox"" id=""notifyonrelease" & iRowCount & """ name=""notifyonrelease" & iRowCount & """"
		response.write " /></td>"
		response.write "<td align=""center""><input type=""checkbox"" id=""reviewisrequired" & iRowCount & """ name=""reviewisrequired" & iRowCount & """"
		response.write " /></td>"
		response.write "<td align=""center"">"
		showDisplayOrder "revieworder", 1, 1, iRowCount
		response.write "</td>"
		response.write "</tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowReviewTypeTable = iRowCount

End Function 


'--------------------------------------------------------------------------------------------------
' integer ShowReviewAlertTypeTable( iPermitTypeid )
'--------------------------------------------------------------------------------------------------
Function ShowReviewAlertTypeTable( ByVal iPermitTypeid )
	Dim oRs, sSql, iRowCount, sRowClass

	iRowCount = -1

	sSql = "SELECT permitalerttypeid, notifyuserid "
	sSql = sSql & " FROM egov_permittypes_to_permitalerttypes "
	sSql = sSql & " WHERE isforreviews = 1 AND permittypeid = " & iPermitTypeid
	sSql = sSql & " ORDER BY permitalertid"

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
			response.write "<input type=""checkbox"" id=""removereviewalert" & iRowCount & """ name=""removereviewalert" & iRowCount & """ /></td>"
			response.write "<td align=""center"">"
			ShowReviewAlertTypePicks oRs("permitalerttypeid"), iRowCount
			response.write "</td>"
			response.write "<td align=""center"">"
			ShowNotifyPicks oRs("notifyuserid"), iRowCount, "ispermitreviewer", "notifyuserid"
			response.write "</td>"
			response.write "</tr>"
			oRs.MoveNext 
		Loop 
	Else
		' put in a starter row.
		iRowCount = 0
		response.write vbcrlf & "<tr><td class=""firstcell"">"
		response.write "<input type=""checkbox"" id=""removereview" & iRowCount & """ name=""removereview" & iRowCount & """ /></td>"
		response.write "<td align=""center"">"
		ShowReviewAlertTypePicks 0, iRowCount
		response.write "</td>"
		response.write "<td align=""center"">"
		ShowNotifyPicks 0, iRowCount, "ispermitreviewer", "notifyuserid"
		response.write "</td>"
		response.write "</tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowReviewAlertTypeTable = iRowCount

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowReviewAlertTypePicks iPermitAlertTypeId, iRowCount 
'--------------------------------------------------------------------------------------------------
Sub ShowReviewAlertTypePicks( ByVal iPermitAlertTypeId, ByVal iRowCount )
	Dim sSql, oRs

	sSql = "SELECT permitalerttypeid, permitalert FROM egov_permitalerttypes "
	sSql = sSql & " WHERE isforbuildingpermits = 1 AND isforreviews = 1 AND orgid = " & SESSION("orgid")
	sSql = sSql & " ORDER BY permitalert"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If not oRs.EOF Then
		response.write vbcrlf & "<select id=""permitalerttypeid" & iRowCount & """ name=""permitalerttypeid" & iRowCount & """>"
		response.write vbcrfl & "<option value=""0"">Select a Review Alert Type</option>"
		Do While NOT oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("permitalerttypeid") & """"  
			If CLng(iPermitAlertTypeId) = CLng(oRs("permitalerttypeid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("permitalert")
			response.write "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If

	oRs.Close
	Set oRs = Nothing

End Sub  


'--------------------------------------------------------------------------------------------------
' void ShowNotifyPicks iNotifyUserId, iRowCount, sNotifyType, sSelectName
'--------------------------------------------------------------------------------------------------
Sub ShowNotifyPicks( ByVal iNotifyUserId, ByVal iRowCount, ByVal sNotifyType, ByVal sSelectName )
	Dim sSql, oRs

	sSql = "SELECT userid, firstname, lastname, isdeleted FROM users WHERE (" & sNotifyType & " = 1 AND orgid = " & SESSION("orgid") & " and isdeleted = 0) or userid = '" & iNotifyUserId & "'"
	sSql = sSql & " ORDER BY isdeleted, lastname, firstname"
	'response.write "<!-- " & sSql & " -->"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	response.write vbcrlf & "<select id=""" & sSelectName & iRowCount & """ name=""" & sSelectName & iRowCount & """>"
	response.write vbcrlf & "<option value=""0"">Unassigned</option>"

	Do While Not oRs.EOF 
		response.write vbcrlf & "<option value=""" & oRs("userid") & """"  
		If CLng(iNotifyUserId) = CLng(oRs("userid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" 
		if oRs("isdeleted") then response.write "["
		response.write oRs("firstname") & " " & oRs("lastname")
		if oRs("isdeleted") then response.write "]"
		response.write "</option>"
		oRs.MoveNext
	Loop

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' void showDisplayOrder sSelectName, iMaxRows, iDisplayOrder, iRowCount 
'--------------------------------------------------------------------------------------------------
Sub showDisplayOrder( ByVal sSelectName, ByVal iMaxRows, ByVal iDisplayOrder, ByVal iRowCount ) 
	Dim x

	response.write vbcrlf & "<select name=""" & sSelectName & iRowCount & """ id=""" & sSelectName & iRowCount & """>"
	For x = 1 To iMaxRows
		response.write vbcrlf & "<option value=""" & x & """"
		If CLng(x) = CLng(iDisplayOrder) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & x & "</option>"
	Next 
	response.write vbcrlf & "</select>"

End Sub 


'--------------------------------------------------------------------------------------------------
' integer GetMaxFeeRows( iPermitTypeid )
'--------------------------------------------------------------------------------------------------
Function GetMaxFeeRows( ByVal iPermitTypeid )
	Dim sSql, oRs

	sSql = "SELECT COUNT(permitfeetypeid) AS hits FROM egov_permittypes_to_permitfeetypes "
	sSql = sSql & " WHERE permittypeid = " & iPermitTypeid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then
			GetMaxFeeRows = CLng(oRs("hits"))
		Else
			GetMaxFeeRows = 0
		End If 
	Else
		GetMaxFeeRows = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' integer GetMaxInspectionRows( iPermitTypeid ) 
'--------------------------------------------------------------------------------------------------
Function GetMaxInspectionRows( ByVal iPermitTypeid ) 
	Dim sSql, oRs

	sSql = "SELECT COUNT(permitinspectiontypeid) AS hits FROM egov_permittypes_to_permitinspectiontypes "
	sSql = sSql & " WHERE permittypeid = " & iPermitTypeid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then
			GetMaxInspectionRows = CLng(oRs("hits"))
		Else
			GetMaxInspectionRows = 0
		End If 
	Else
		GetMaxInspectionRows = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' integer GetMaxReviewRows( iPermitTypeid )
'--------------------------------------------------------------------------------------------------
Function GetMaxReviewRows( ByVal iPermitTypeid )
	Dim sSql, oRs

	sSql = "SELECT COUNT(permitreviewtypeid) AS hits FROM egov_permittypes_to_permitreviewtypes "
	sSql = sSql & " WHERE permittypeid = " & iPermitTypeid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then
			GetMaxReviewRows = CLng(oRs("hits"))
		Else
			GetMaxReviewRows = 0
		End If 
	Else
		GetMaxReviewRows = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowFeeTypePicks( iPermitFeeTypeId, iRowCount )
'--------------------------------------------------------------------------------------------------
Sub ShowFeeTypePicks( ByVal iPermitFeeTypeId, ByVal iRowCount )
	Dim sSql, oRs

	sSql = "SELECT F.permitfeetypeid, F.permitfee, F.isupfrontfee, F.isreinspectionfee, M.permitfeemethod "
	sSql = sSql & " FROM egov_permitfeetypes F, egov_permitfeemethods M "
	sSql = sSql & " WHERE F.permitfeemethodid = M.permitfeemethodid AND F.orgid = " & SESSION("orgid") & " ORDER BY F.permitfee"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If not oRs.EOF Then
		response.write vbcrlf & "<select id=""permitfeetypeid" & iRowCount & """ name=""permitfeetypeid" & iRowCount & """ class=""permitfeetypeDD"">"
		response.write vbcrfl & "<option value=""0"">Select a Fee Type</option>"
		Do While NOT oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("permitfeetypeid") & """"  
			If CLng(iPermitFeeTypeId) = CLng(oRs("permitfeetypeid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("permitfee")
'			If oRs("isbuildingpermitfee") Then
'				response.write " (Bldg Prmt)"
'			End If 
			response.write " (" & Left(oRs("permitfeemethod"),23) & ")"
			response.write "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowInspectionTypePicks( iPermitInspectionTypeId, iRowCount )
'--------------------------------------------------------------------------------------------------
Sub ShowInspectionTypePicks( ByVal iPermitInspectionTypeId, ByVal iRowCount )
	Dim sSql, oRs

	sSql = "SELECT permitinspectiontypeid, permitinspectiontype "
	sSql = sSql & " FROM egov_permitinspectiontypes WHERE orgid = " & SESSION("orgid") & " ORDER BY permitinspectiontype"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If not oRs.EOF Then
		response.write vbcrlf & "<select id=""permitinspectiontypeid" & iRowCount & """ name=""permitinspectiontypeid" & iRowCount & """ class=""permitinspectiontypeDD"">"
		response.write vbcrfl & "<option value=""0"">Select an Inspection Type</option>"
		Do While NOT oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("permitinspectiontypeid") & """"  
			If CLng(iPermitInspectionTypeId) = CLng(oRs("permitinspectiontypeid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("permitinspectiontype")
'			If oRs("isbuildingpermittype") Then
'				response.write " (Bldg Prmt)"
'			End If 
			response.write "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If

	oRs.Close
	Set oRs = Nothing

End Sub  


'--------------------------------------------------------------------------------------------------
' Sub ShowInspectorPicks( iInspectorUserId, iRowCount )
'--------------------------------------------------------------------------------------------------
Sub ShowInspectorPicks( ByVal iInspectorUserId, ByVal iRowCount )
	Dim sSql, oRs

	sSql = "SELECT userid, firstname, lastname,isdeleted FROM users WHERE (isdeleted = 0 OR userid = '" & iInspectorUserId & "') AND ispermitinspector = 1 AND orgid = " & SESSION("orgid")
	sSql = sSql & " ORDER BY isdeleted,lastname, firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1
	
	response.write vbcrlf & "<select id=""permitinspectorid" & iRowCount & """ name=""permitinspectorid" & iRowCount & """>"
	response.write vbcrlf & "<option value=""0"" "
	If CLng(iInspectorUserId) = CLng(0) Then
		response.write " selected=""selected"" "
	End If 
	response.write ">Unassigned</option>"

	Do While NOT oRs.EOF 
		response.write vbcrlf & "<option value=""" & oRs("userid") & """"  
		If CLng(iInspectorUserId) = CLng(oRs("userid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" 
		if oRs("isdeleted") then response.write "[" 
		response.write oRs("firstname") & " " & oRs("lastname")
		if oRs("isdeleted") then response.write "]" 
		response.write "</option>"
		oRs.MoveNext
	Loop

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowReviewerPicks( iReviewerUserId, iRowCount )
'--------------------------------------------------------------------------------------------------
Sub ShowReviewerPicks( ByVal iReviewerUserId, ByVal iRowCount )
	Dim sSql, oRs

	sSql = "SELECT userid, firstname, lastname,isdeleted FROM users WHERE ispermitreviewer = 1 AND (isdeleted = 0 OR userid = '" & iReviewerUserId & "') AND orgid = " & SESSION("orgid")
	sSql = sSql & " ORDER BY isdeleted,lastname, firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	response.write vbcrlf & "<select id=""permitreviewerid" & iRowCount & """ name=""permitreviewerid" & iRowCount & """>"
	response.write vbcrlf & "<option value=""0"">Unassigned</option>"

	Do While NOT oRs.EOF 
		response.write vbcrlf & "<option value=""" & oRs("userid") & """"  
		If CLng(iReviewerUserId) = CLng(oRs("userid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" 
		if oRs("isdeleted") then response.write "[" 
		response.write oRs("firstname") & " " & oRs("lastname")
		if oRs("isdeleted") then response.write "]" 
		response.write "</option>"
		oRs.MoveNext
	Loop
		
	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowReviewTypePicks( iPermitReviewTypeId, iRowCount )
'--------------------------------------------------------------------------------------------------
Sub ShowReviewTypePicks( ByVal iPermitReviewTypeId, ByVal iRowCount )
	Dim sSql, oRs

	sSQL = "SELECT permitreviewtypeid, permitreviewtype "
	sSql = sSql & " FROM egov_permitreviewtypes WHERE orgid = " & SESSION("orgid") & " ORDER BY permitreviewtype"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If not oRs.EOF Then
		response.write vbcrlf & "<select id=""permitreviewtypeid" & iRowCount & """ name=""permitreviewtypeid" & iRowCount & """ class=""permitreviewtypeDD"">"
		response.write vbcrfl & "<option value=""0"">Select a Review Type</option>"
		Do While NOT oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("permitreviewtypeid") & """"  
			If CLng(iPermitReviewTypeId) = CLng(oRs("permitreviewtypeid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("permitreviewtype")
'			If oRs("isbuildingpermittype") Then
'				response.write " (Bldg Prmt)"
'			End If 
			response.write "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If

	oRs.Close
	Set oRs = Nothing

End Sub  


'--------------------------------------------------------------------------------------------------
' void ShowPermitNumberPrefixes sPermitNumberPrefix 
'--------------------------------------------------------------------------------------------------
Sub ShowPermitNumberPrefixes( ByVal sPermitNumberPrefix )
	Dim sSql, oRs

	sSql = "SELECT permitnumberprefix, permitnumberprefixtype, isnone FROM egov_permitnumberprefixes WHERE orgid = " & SESSION("orgid")
	sSql = sSql & " ORDER BY displayorder, permitnumberprefix"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If not oRs.EOF Then
		Do While NOT oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("permitnumberprefix") & """"  
			If sPermitNumberPrefix = oRs("permitnumberprefix") Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" 
			If Not oRs("isnone") Then 
				response.write oRs("permitnumberprefix") & " &ndash; "
			End If 
			If Not IsNull(oRs("permitnumberprefixtype")) Then
				response.write oRs("permitnumberprefixtype")
			End If 
			response.write "</option>"
			oRs.MoveNext
		Loop
	Else
		response.write "<option value="">No Prefixes Available</option>"
	End If

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' integer ShowInspectionAlertTypeTable( iPermitTypeid )
'--------------------------------------------------------------------------------------------------
Function ShowInspectionAlertTypeTable( ByVal iPermitTypeid )
	Dim oRs, sSql, iRowCount, sRowClass

	iRowCount = -1

	sSql = "SELECT permitalerttypeid, notifyuserid "
	sSql = sSql & " FROM egov_permittypes_to_permitalerttypes "
	sSql = sSql & " WHERE isforinspections = 1 AND permittypeid = " & iPermitTypeid
	sSql = sSql & " ORDER BY permitalertid"

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
			response.write "<input type=""checkbox"" id=""removeinspectionalert" & iRowCount & """ name=""removeinspectionalert" & iRowCount & """ /></td>"
			response.write "<td align=""center"">"
			ShowInspectionAlertTypePicks oRs("permitalerttypeid"), iRowCount
			response.write "</td>"
			response.write "<td align=""center"">"
			ShowNotifyPicks oRs("notifyuserid"), iRowCount, "ispermitinspector", "notifyinspectoruserid"
			response.write "</td>"
			response.write "</tr>"
			oRs.MoveNext 
		Loop 
	Else
		' put in a starter row.
		iRowCount = 0
		response.write vbcrlf & "<tr><td class=""firstcell"">"
		response.write "<input type=""checkbox"" id=""removeinspectionalert" & iRowCount & """ name=""removeinspectionalert" & iRowCount & """ /></td>"
		response.write "<td align=""center"">"
		ShowInspectionAlertTypePicks 0, iRowCount
		response.write "</td>"
		response.write "<td align=""center"">"
		ShowNotifyPicks 0, iRowCount, "ispermitinspector", "notifyinspectoruserid"
		response.write "</td>"
		response.write "</tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowInspectionAlertTypeTable = iRowCount

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowInspectionAlertTypePicks( iPermitAlertTypeId, iRowCount )
'--------------------------------------------------------------------------------------------------
Sub ShowInspectionAlertTypePicks( ByVal iPermitAlertTypeId, ByVal iRowCount )
	Dim sSql, oRs

	sSQL = "SELECT permitalerttypeid, permitalert FROM egov_permitalerttypes "
	sSql = sSql & " WHERE isforbuildingpermits = 1 AND isforinspections = 1 AND orgid = " & SESSION("orgid")
	sSql = sSql & " ORDER BY permitalert"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then
		response.write vbcrlf & "<select id=""permitinspectionalerttypeid" & iRowCount & """ name=""permitinspectionalerttypeid" & iRowCount & """>"
		response.write vbcrfl & "<option value=""0"">Select an Inspection Alert Type</option>"
		Do While NOT oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("permitalerttypeid") & """ "  
			If CLng(iPermitAlertTypeId) = CLng(oRs("permitalerttypeid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("permitalert")
			response.write "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If

	oRs.Close
	Set oRs = Nothing

End Sub  


'--------------------------------------------------------------------------------------------------
' void ShowlicenseTypes( iPermitTypeId )
'--------------------------------------------------------------------------------------------------
Sub ShowlicenseTypes( ByVal iPermitTypeId, ByVal sIsRequired )
	Dim sSql, oRs, sName

	If sIsRequired = "0" Then
		sName = "show"
	Else
		sName = ""
	End If 

	sSql = "SELECT licensetypeid, licensetype FROM egov_permitlicensetypes WHERE orgid = " & SESSION("orgid")
	sSql = sSql & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	Do While Not oRs.EOF
		response.write "<input type=""checkbox"" name=""" & sName & "licensetypeid"" value=""" & oRs("licensetypeid") & """"
		If PermitTypeRequiresLicenseType( iPermitTypeId, oRs("licensetypeid"), sIsRequired ) Then
			response.write " checked=""checked"" "
		End If 
		response.write " />&nbsp;" & oRs("licensetype") & "&nbsp;&nbsp;"
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' boolean PermitTypeRequiresLicenseType( iPermitTypeId, iLicenseTypeId )
'--------------------------------------------------------------------------------------------------
Function PermitTypeRequiresLicenseType( ByVal iPermitTypeId, ByVal iLicenseTypeId, ByVal sIsRequired )
	Dim sSql, oRs

	sSql = "SELECT COUNT(licensetypeid) AS hits FROM egov_permittypes_to_permitlicensetypes "
	sSql = sSql & " WHERE permittypeid = " & iPermitTypeId & " AND licensetypeid = " & iLicenseTypeId
	sSql = sSql & " AND isrequired = " & sIsRequired

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then
			PermitTypeRequiresLicenseType = True 
		Else
			PermitTypeRequiresLicenseType = False 
		End If 
	Else 
		PermitTypeRequiresLicenseType = False 
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowPermitDocuments( iDocumentId )
'--------------------------------------------------------------------------------------------------
Sub ShowPermitDocuments( ByVal iDocumentId )
	Dim sSql, oRs

	sSql = "SELECT documentid, document FROM egov_permitdocuments "
	sSql = sSql & " WHERE orgid = " & SESSION("orgid")
	sSql = sSql & " ORDER BY document"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select id=""documentid"" name=""documentid"">"
	response.write vbcrfl & "<option value=""0"">Use the setup defined below</option>"
	
	Do While NOT oRs.EOF 
		response.write vbcrlf & "<option value=""" & oRs("documentid") & """ "  
		If CLng(iDocumentId) = CLng(oRs("documentid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("document")
		response.write "</option>"
		oRs.MoveNext
	Loop

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' integer ShowDocumentTable( iPermitTypeid )
'--------------------------------------------------------------------------------------------------
Function ShowDocumentTable( ByVal iPermitTypeid )
	Dim oRs, sSql, iRowCount, sRowClass

	iRowCount = -1

	sSql = "SELECT permitdocumentid, documentid, documentlabel "
	sSql = sSql & " FROM egov_permittypes_to_permitdocuments"
	sSql = sSql & " WHERE permittypeid = " & iPermitTypeid & " ORDER BY documentlabel"

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
			response.write vbcrlf & "<tr" & sRowClass & ">"
			response.write "<td class=""firstcell"">"
			response.write "<input type=""checkbox"" id=""removedocument" & iRowCount & """ name=""removedocument" & iRowCount & """ />"
			response.write "<input type=""hidden"" id=""permitdocumentid" & iRowCount & """ name=""permitdocumentid" & iRowCount & """ value=""" & oRs("permitdocumentid") & """ />"
			response.write "</td>"
			response.write "<td>"
			response.write "<input type=""text"" size=""50"" maxlength=""50"" name=""documentlabel" & iRowCount & """ id=""documentlabel" & iRowCount & """ value=""" & oRs("documentlabel") & """ />"
			response.write "</td>"
			response.write "<td>"
			ShowPermitTypeDocumentPicks oRs("documentid"), iRowCount
			response.write "</td>"
			response.write "</tr>"
			oRs.MoveNext 
		Loop
	Else
		' put in a starter row.
		iRowCount = 0
		response.write vbcrlf & "<tr>"
		response.write "<td class=""firstcell"">"
		response.write "<input type=""checkbox"" id=""removedocument" & iRowCount & """ name=""removedocument" & iRowCount & """ />"
		response.write "<input type=""hidden"" name=""permitdocumentid" & iRowCount & """ value=""0"" />"
		response.write "</td>"
		response.write "<td>"
		response.write "<input type=""text"" size=""50"" maxlength=""50"" name=""documentlabel" & iRowCount & """ id=""documentlabel" & iRowCount & """ value="""" />"
		response.write "</td>"
		response.write "<td>"
		ShowPermitTypeDocumentPicks 0, iRowCount
		response.write "</td>"
		response.write "</tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowDocumentTable = iRowCount

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowPermitDocuments( iDocumentId, iRowCount )
'--------------------------------------------------------------------------------------------------
Sub ShowPermitTypeDocumentPicks( ByVal iDocumentId, ByVal iRowCount )
	Dim sSql, oRs

	sSQL = "SELECT documentid, document FROM egov_permitdocuments "
	sSql = sSql & " WHERE orgid = " & SESSION("orgid")
	sSql = sSql & " ORDER BY document"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select id=""documentid" & iRowCount & """ name=""documentid" & iRowCount & """>"
	response.write vbcrfl & "<option value=""0"">Select a Document</option>"
	
	Do While NOT oRs.EOF 
		response.write vbcrlf & "<option value=""" & oRs("documentid") & """ "  
		If CLng(iDocumentId) = CLng(oRs("documentid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("document")
		response.write "</option>"
		oRs.MoveNext
	Loop

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowDetailFields iPermitTypeid 
'--------------------------------------------------------------------------------------------------
Sub ShowDetailFields( ByVal iPermitTypeid )
	Dim sSql, oRs, iRowCount

	iRowCount = -1

	sSql = "SELECT detailfieldid, detailfieldlabel FROM egov_permitdetailfields "
	sSql = sSql & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While NOT oRs.EOF 
		iRowCount = iRowCount + 1
		If iRowCount Mod 2 = 0 Then 
			sRowClass = ""
		Else
			sRowClass = " class=""altrow"" "
		End If 
		response.write vbcrlf & "<tr" & sRowClass & ">"
		response.write vbcrlf & "<td><input type=""checkbox"" name=""detailfieldid"" value=""" & oRs("detailfieldid") & """"
		If DetailFieldIsSelected( iPermitTypeid, oRs("detailfieldid") ) Then
			response.write " checked=""checked"""
		End If
		response.write " /></td><td>" & oRs("detailfieldlabel") & "</td></tr>"
		oRs.MoveNext
	Loop

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' boolean DetailFieldIsSelected( iPermitTypeid, iDetailFieldId )
'--------------------------------------------------------------------------------------------------
Function DetailFieldIsSelected( ByVal iPermitTypeid, ByVal iDetailFieldId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(detailfieldid) AS hits FROM egov_permittypes_to_permitdetailfields "
	sSql = sSql & " WHERE permittypeid = " & iPermitTypeId & " AND detailfieldid = " & iDetailFieldId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then
			DetailFieldIsSelected = True 
		Else
			DetailFieldIsSelected = False 
		End If 
	Else 
		DetailFieldIsSelected = False 
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowPermitCategoryPicks iPermitCategoryId 
'--------------------------------------------------------------------------------------------------
Sub ShowPermitCategoryPicks( ByVal iPermitCategoryId )
	Dim sSql, oRs

	sSql = "SELECT permitcategoryid, permitcategory FROM egov_permitcategories "
	sSql = sSql & " WHERE orgid = " & SESSION("orgid")
	sSql = sSql & " ORDER BY permitcategory"
	'Response.Write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select id=""permitcategoryid"" name=""permitcategoryid"">"
	
	Do While NOT oRs.EOF 
		response.write vbcrlf & "<option value=""" & oRs("permitcategoryid") & """"  
		If CLng(iPermitCategoryId) = CLng(oRs("permitcategoryid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("permitcategory")
		response.write "</option>"
		oRs.MoveNext
	Loop

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' integer ShowCustomFieldTypeTable( iPermitTypeid, bOrgHasPermitTypeReport )
'--------------------------------------------------------------------------------------------------
Function ShowCustomFieldTypeTable( ByVal iPermitTypeid, ByVal bOrgHasPermitTypeReport )
	Dim oRs, sSql, iRowCount, sRowClass, iMaxRows

	iRowCount = -1
	iMaxRows = GetMaxCustomFieldTypeRows( iPermitTypeid ) 

	sSql = "SELECT customfieldtypeid, customfieldorder, includeonreport "
	sSql = sSql & " FROM egov_permittypes_to_permitcustomfieldtypes "
	sSql = sSql & " WHERE permittypeid = " & iPermitTypeid
	sSql = sSql & " ORDER BY customfieldorder"

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
			response.write "<input type=""checkbox"" id=""removecustomfield" & iRowCount & """ name=""removecustomfield" & iRowCount & """ /></td><td>"
			
			ShowCustomFieldTypePicks oRs("customfieldtypeid"), iRowCount
			
			response.write "</td>"

			If bOrgHasPermitTypeReport Then 
				response.write "<td align=""center"">"
				response.write "<input type=""checkbox"" id=""includeonreport" & iRowCount & """ name=""includeonreport" & iRowCount & """"
				If oRs("includeonreport") Then
					response.write " checked=""checked"""
				End If 
				response.write " />"
				response.write "</td>"
			End If 
			
			response.write "<td align=""center"">"
			If Not bOrgHasPermitTypeReport Then 
				response.write "<input type=""hidden"" name=""includeonreport" & iRowCount & """ value=""no"" />"
			End If 
			showDisplayOrder "customfieldorder", iMaxRows, oRs("customfieldorder"), iRowCount
			response.write "</td>"

			response.write "</tr>"

			oRs.MoveNext 
		Loop 
	Else
		' put in a starter row.
		iRowCount = 0
		response.write vbcrlf & "<tr><td class=""firstcell"">"
		response.write "<input type=""checkbox"" id=""removecustomfield" & iRowCount & """ name=""removecustomfield" & iRowCount & """ /></td><td>"
		ShowCustomFieldTypePicks 0, iRowCount
		response.write "</td>"

		If bOrgHasPermitTypeReport Then 
			response.write "<td align=""center"">"
			response.write "<input type=""checkbox"" id=""includeonreport" & iRowCount & """ name=""includeonreport" & iRowCount & """ />"
			response.write "</td>"
		End If 

		response.write "<td align=""center"">"
		showDisplayOrder "customfieldorder", 1, 1, iRowCount
		response.write "</td>"
		response.write "</tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowCustomFieldTypeTable = iRowCount

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowCustomFieldTypePicks iCustomFieldTypeId, iRowCount
'--------------------------------------------------------------------------------------------------
Sub ShowCustomFieldTypePicks( ByVal iCustomFieldTypeId, ByVal iRowCount )
	Dim sSql, oRs

	sSQL = "SELECT customfieldtypeid, fieldname "
	sSql = sSql & " FROM egov_permitcustomfieldtypes WHERE orgid = " & SESSION("orgid")
	sSql = sSql & " AND isactive = 1 ORDER BY fieldname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If not oRs.EOF Then
		response.write vbcrlf & "<select id=""customfieldtypeid" & iRowCount & """ name=""customfieldtypeid" & iRowCount & """ class=""permitcustomfieldtypeDD"">"
		response.write vbcrfl & "<option value=""0"">Select a Custom Field</option>"
		Do While NOT oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("customfieldtypeid") & """"  
			If CLng(iCustomFieldTypeId) = CLng(oRs("customfieldtypeid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("fieldname")
			response.write "</option>"
			oRs.MoveNext
		Loop

		response.write vbcrlf & "</select>"

	End If

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' integer GetMaxCustomFieldTypeRows( iPermitTypeid )
'--------------------------------------------------------------------------------------------------
Function GetMaxCustomFieldTypeRows( ByVal iPermitTypeid )
	Dim sSql, oRs

	sSql = "SELECT COUNT(customfieldtypeid) AS hits FROM egov_permittypes_to_permitcustomfieldtypes "
	sSql = sSql & " WHERE permittypeid = " & iPermitTypeid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then
			GetMaxCustomFieldTypeRows = CLng(oRs("hits"))
		Else
			GetMaxCustomFieldTypeRows = 0
		End If 
	Else
		GetMaxCustomFieldTypeRows = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowLocationRequirements iPermitLocationRequirementId
'--------------------------------------------------------------------------------------------------
Sub ShowLocationRequirements( ByVal iPermitLocationRequirementId )
	Dim sSql, oRs

	sSQL = "SELECT permitlocationrequirementid, permitlocationrequirement "
	sSql = sSql & "FROM egov_permitlocationrequirements "
	sSql = sSql & "ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If not oRs.EOF Then
		response.write vbcrlf & "<select id=""permitlocationrequirementid"" name=""permitlocationrequirementid"">"

		Do While NOT oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("permitlocationrequirementid") & """"  
			If CLng(iPermitLocationRequirementId) = CLng(oRs("permitlocationrequirementid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("permitlocationrequirement")
			response.write "</option>"
			oRs.MoveNext
		Loop

		response.write vbcrlf & "</select>"

	End If

	oRs.Close
	Set oRs = Nothing

End Sub 



%>
