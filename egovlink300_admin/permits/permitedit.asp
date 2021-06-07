<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitedit.asp
' AUTHOR: Steve Loar
' CREATED: 03/06/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Edits permits
'
' MODIFICATION HISTORY
' 1.0   03/06/2008	Steve Loar - INITIAL VERSION
' 1.1	01/18/2010	Steve Loar - Added notification of applicant of inspection scheduling
' 1.2	08/17/2010	Steve Loar - Changes to allow changing the permit type.
' 1.3	10/27/2010	Steve Loar - Changes to allow any type of permits
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, sPermitStatus, iNextStatus, sButtonText, sPermitNo, sApplied, sExpires, sProposedUse
Dim sExistingUse, iWorkClassId, iConstructionTypeId, iOccupancyTypeId, sDescriptionOfWork, sLegalDescription
Dim iUseTypeId, sReleased, sIssued, sCompleted, sListedOwner, iMaxContractors, iPermitAddressId
Dim sJobValue, sTotalSqFt, sFinishedSqFt, sUnFinishedSqFt, sOtherSqFt, sLinearFt, sCuFt, sExaminationHours, sFeeTotal, iMaxFees
Dim iActiveTabId, sInvoicedTotal, sPaidTotal, sDueTotal, sNonInvoicedTotal, iPermitStatusId, bCanPrintPermit
Dim bCanChangeExpirationDate, bCanSaveChanges, bStatusCanChange, bWaiveAllFees, sWaivedTotal, iMaxPriorContacts
Dim iMaxReviews, sApproved, iInvoicePicks, iMaxInspections, iMaxAttachments, bHasExpirationDate, sPermitnotes
Dim bIsOnHold, bIsCompleted, sPriorStatus, bCanPlaceHolds, bIsVoided, bIsExpired, sAlertMsg, sUpFrontFeeTotal
Dim sCounty, sParcelid, iPlansByContactId, sResidentialUnits, sApprovedAs, sOccupants, sTempCONotes, sCONotes
Dim bHasTempCO, bHasCO, dTempCOIssuedDate, dCOIssuedDate, sTempCOButtonText, sCOButtonText, sTempCOAction
Dim sCOAction, bCanIssueTempCO, bCanIssueCO, sPrimaryContact, iDocumentid, iPromptOnNoJobValue
Dim iUseClassId, iWorkScopeId, sStructureLength, sStructureWidth, sStructureHeight, sZoning, sPlanNumber
Dim sDemolishExistingStructure, sLandFillName, sLandFillCity, sLandFillPhone, bAlertApplicant
Dim bUserCanChangeCriticalDates, bUserIsRootAdmin, iPermitLocationRequirementId, sPermitLocation
Dim bNeedsAddress, bNeedsLocation, iMaxCustomPermitFields, bPermitIsInBuildingPermitCategory
Dim sPermitIsInBuildingPermitCategory, bOrgHasVolume, bHasReqLicenses, bSomeFeesSetToZero, bAllFeesPaidOrWaived
Dim bAllReviewsComplete

sLevel = "../" ' Override of value from common.asp

PageDisplayCheck "edit permits", sLevel	' In common.asp

iPermitId = CLng(request("permitid"))

bPermitIsInBuildingPermitCategory = PermitIsInBuildingPermitCategory( iPermitId )

If bPermitIsInBuildingPermitCategory Then
	sPermitIsInBuildingPermitCategory = "yes"
Else
	sPermitIsInBuildingPermitCategory = "no"
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

iNextStatusOrder = 0
sPermitNo = ""
sApplied = ""
sExpires = ""
sProposedUse = ""
sExistingUse = ""
iWorkClassId = 0
iUseClassId = 0
iWorkScopeId = 0
iConstructionTypeId = 0
iOccupancyTypeId = 0
sDescriptionOfWork = ""
sLegalDescription = ""
iUseTypeId = 0 
sReleased = ""
sApproved = ""
sIssued = ""
sCompleted = ""
sListedOwner = ""
iMaxContractors = 0
iMaxPriorContacts = 0
iPermitAddressId = 0
sJobValue = FormatNumber(0.00,2,,,0)
sTotalSqFt = FormatNumber(0.00,2,,,0)
sFinishedSqFt = FormatNumber(0.00,2,,,0)
sUnFinishedSqFt = FormatNumber(0.00,2,,,0)
sOtherSqFt = FormatNumber(0.00,2,,,0)
'sLinearFt = FormatNumber(0.00,2,,,0)
'sCuFt = FormatNumber(0.00,2,,,0)
sExaminationHours = FormatNumber(0.00,2,,,0)
sFeeTotal = FormatNumber(0.00,2)
sUpFrontFeeTotal = FormatNumber(0.00,2)
iMaxFees = 0
iMaxReviews = 0
iMaxInspections = 0
iMaxAttachments = 0
iPermitStatusId = 1
bCanPrintPermit = False 
bCanChangeExpirationDate = False 
bHasExpirationDate = True 
bWaiveAllFees = False 
iInvoicePicks = CLng(0)
sPermitnotes = ""
bIsCompleted = False 
sPriorStatus = ""
bCanPlaceHolds = False 
sAlertMsg = ""
iPlansByContactId = CLng(0)
sResidentialUnits = 0
sApprovedAs = ""
sOccupants = ""
sTempCONotes = ""
sCONotes = ""
dTempCOIssuedDate = ""
dCOIssuedDate = ""
sPrimaryContact = ""
sStructureLength = ""
sStructureWidth = ""
sStructureHeight = ""
sZoning = ""
sPlanNumber = ""
sDemolishExistingStructure = ""
sLandFillName = ""
sLandFillCity = ""
sLandFillPhone = ""
bAlertApplicant = False 
iPermitLocationRequirementId = 1
sPermitLocation = ""
bNeedsAddress = False 
bNeedsLocation = False 
iMaxCustomPermitFields = 0
bOrgHasVolume = False 


bIsOnHold = GetPermitIsOnHold( iPermitId )		' in permitcommonfunctions.asp  
bIsVoided = GetPermitIsVoided( iPermitId )		' in permitcommonfunctions.asp
bIsExpired = GetPermitIsExpired( iPermitId )	' in permitcommonfunctions.asp
bHasTempCO = GetPermitPermitTypeFlag( iPermitid, "hastempco" )	' in permitcommonfunctions.asp
bHasCO = GetPermitPermitTypeFlag( iPermitid, "hasco" )	' in permitcommonfunctions.asp
iDocumentid = GetPermitPermitTypeValue( iPermitid, "documentid" )	' in permitcommonfunctions.asp
bOrgHasVolume = OrgHasFeature("volume total")

If bIsOnHold Then 
	sPermitStatus = "On Hold"
	bCanSaveChanges = False 
	sButtonText = ""
	bStatusCanChange = False 
	bCanIssueTempCO = False 
	bCanIssueCO = False 
Else 
	If bIsVoided Then 
		sPermitStatus = "Voided"
		bCanSaveChanges = False 
		sButtonText = ""
		bStatusCanChange = False
		bCanIssueTempCO = False 
		bCanIssueCO = False 
	Else 
		sPermitStatus = GetPermitStatus( iPermitId, iNextStatus, iPermitStatusId, bCanPrintPermit, bCanChangeExpirationDate, bHasExpirationDate, bIsCompleted )
		If bIsExpired Then 
			sPermitStatus = "Expired"
		End If 
		bCanSaveChanges = StatusAllowsSaveChanges( iPermitStatusId ) 	' in permitcommonfunctions.asp
		sButtonText = GetButtonText( iNextStatus )
		bStatusCanChange = CheckIfStatusCanChange( iPermitId, iPermitStatusId ) 	' in permitcommonfunctions.asp
	End If 
End If 


GetPermitDetails iPermitId

sInvoicedTotal = GetInvoicedTotal( iPermitId ) 	' in permitcommonfunctions.asp

sPaidTotal = GetPaidTotal( iPermitId ) 	' in permitcommonfunctions.asp

sWaivedTotal = GetWaivedTotal( iPermitId ) 	' in permitcommonfunctions.asp

sDueTotal = FormatNumber(CDbl(sInvoicedTotal) - ( CDbl(sPaidTotal) + CDbl(sWaivedTotal) ),2,,,0)

sNonInvoicedTotal = FormatNumber(CDbl(sFeeTotal) - CDbl(sInvoicedTotal),2,,,0)

If bIsCompleted Then 
	sPriorStatus = GetPriorPermitStatus( iPermitStatusId )   	' in permitcommonfunctions.asp
End If 

If OrgHasFeature("prompt on no job value") Then
	iPromptOnNoJobValue = 1 
Else
	iPromptOnNoJobValue = 0 
End If 

GetLocationRequirements iPermitLocationRequirementId, bNeedsAddress, bNeedsLocation


' Check these user rights here as they are used in several places
bUserIsRootAdmin = UserIsRootAdmin(session("UserID"))
bUserCanChangeCriticalDates = UserHasPermission( Session("UserId"), "can change critical dates" )

%>


<html>
<head>
	<title>E-Gov Administration Console</title>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" />

	<link rel="stylesheet" type="text/css" href="../yui/build/tabview/assets/skins/sam/tabview.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<script type="text/javascript" src="../yui/yahoo-dom-event.js"></script>  
	<script type="text/javascript" src="../yui/element-min.js"></script>  
	<script type="text/javascript" src="../yui/tabview-min.js"></script>

	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="JavaScript" src="../scripts/formatnumber.js"></script>
	<script language="JavaScript" src="../scripts/removespaces.js"></script>
	<script language="JavaScript" src="../scripts/removecommas.js"></script>
	<script language="JavaScript" src="../scripts/textareamaxlength.js"></script>
	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>
	<script language="JavaScript" src="../scripts/tooltip.js"></script> 
	<script language="JavaScript" src="../scripts/layers.js"></script>
	<script language="javascript" src="../scripts/isvaliddate.js"></script>
	<script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
	<script language="JavaScript" src="../scriptaculous/src/scriptaculous.js"></script>
  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
	
	<script language="Javascript">
	<!--
  		$( function() {
    		$( ".datepicker" ).datepicker({
      		changeMonth: true,
      		showOn: "both",
      		buttonText: "<i class=\"fa fa-calendar\"></i>",
      		changeYear: true
    		});
  		} );
		var tabView;
		var winHandle;
		var w = (screen.width - 640)/2;
		var h = (screen.height - 480)/2;

		(function() {
			tabView = new YAHOO.widget.TabView('demo');
			tabView.set('activeIndex', <%=iActiveTabId%>);

		})();

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

		function ViewCO( iPermitId, sCOType, sAction, sQuestionText )
		{
			if (sAction == 'issue')
			{
				if (! confirm("Issue" + sQuestionText + " Certificate of Occupancy?"))
				{
					return;
				}
			}
			//popup the Temp Co or CO doc
			//winHandle = eval('window.open("viewtempcoandco.asp?permitid=' + iPermitId + '&cotype=' + sCOType + '", "_contact", "width=900,height=700,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,resizable=1,left=' + w + ',top=' + h + '")');
			//winHandle.focus();
			showModal('viewtempcoandco.asp?permitid=' + iPermitId + '&cotype=' + sCOType, sQuestionText + ' Certificate of Occupancy', 80, 90);
		}

		function ViewAttachment( iAttachmentId )
		{
			location.href = "permitattachmentview.asp?permitattachmentid=" + iAttachmentId;
		}

		function CheckConstRate()
		{
			//alert(document.frmPermit.constructiontypeid.value);
			doAjax('checkconstructionrate.asp', 'constructiontypeid=' + document.frmPermit.constructiontypeid.value + '&occupancytypeid=' + document.frmPermit.occupancytypeid.value, 'RateDisplay', 'get', '0');
		}

		function RateDisplay( sResults )
		{
			alert( 'The Rate is ' + sResults );
		}

		function VoidInvoice( iInvoiceId )
		{
			if (confirm("Void invoice #" + iInvoiceId + "?"))
			{
				//alert("Voiding");
				doAjax('voidinvoice.asp', 'invoiceid=' + iInvoiceId, 'RefreshPageAfterVoid', 'get', '0');
			}
		}

		function RefreshPageAfterVoid( sResults )
		{
			//alert(sResults);
			setTimeout(function() {location.href = "permitedit.asp?permitid=<%=iPermitId%>&v2a=" + sResults + "&activetab=" + tabView.get("activeIndex");}, 200);
		}

		function SelectContact( sType )
		{
			//winHandle = eval('window.open("contactpicker.asp?permitid=<%=iPermitId%>&stype=' + sType + '", "_contact", "width=900,height=600,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('contactpicker.asp?permitid=<%=iPermitId%>&stype=' + sType, 'Edit Contact', 45, 30);
		}

		function SelectFee( )
		{
			//winHandle = eval('window.open("feepicker.asp?permitid=<%=iPermitId%>", "_contact", "width=800,height=300,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('feepicker.asp?permitid=<%=iPermitId%>', 'Add A Fee', 30, 30);
		}

		function SelectReviews( )
		{
			//winHandle = eval('window.open("reviewpicker.asp?permitid=<%=iPermitId%>", "_contact", "width=800,height=300,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('reviewpicker.asp?permitid=<%=iPermitId%>', 'Permit Review', 30, 30);
		}

		function SelectInspections( )
		{
			//winHandle = eval('window.open("inspectionpicker.asp?permitid=<%=iPermitId%>", "_contact", "width=800,height=300,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('inspectionpicker.asp?permitid=<%=iPermitId%>', 'Permit Inspection', 30, 30);
		}

		function AddAttachments( )
		{
			//winHandle = eval('window.open("permitattachment.asp?permitid=<%=iPermitId%>", "_contact", "width=800,height=350,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('permitattachment.asp?permitid=<%=iPermitId%>', 'Add An Attachment', 40, 40);
		}

		function EditContact( sContactId, sType )
		{
			myRand = parseInt(Math.random() * 99999999 );
			//winHandle = eval('window.open("permitcontactedit.asp?permitcontactid=' + sContactId +'&permitid=<%=iPermitId%>&permitstatusid=<%=iPermitStatusId%>&type=' + sType + '&rand=' + myRand + '", "_contact", "width=900,height=600,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('permitcontactedit.asp?permitcontactid=' + sContactId +'&permitid=<%=iPermitId%>&permitstatusid=<%=iPermitStatusId%>&type=' + sType + '&rand=' + myRand, 'Edit Contact', 55, 90);
		}

		function EditApplicant( sContactType, sContactId )
		{
			myRand = parseInt(Math.random() * 99999999 );
			if (sContactType == 'U')
			{
				//winHandle = eval('window.open("permitapplicantedit.asp?userid=' + sContactId + '&updatetitle=1&detailid=applicantdetails&permitid=<%=iPermitId%>&permitstatusid=<%=iPermitStatusId%>&rand=' + myRand + '", "_contact", "width=800,height=800,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
				showModal('permitapplicantedit.asp?userid=' + sContactId + '&updatetitle=1&detailid=applicantdetails&permitid=<%=iPermitId%>&permitstatusid=<%=iPermitStatusId%>&rand=' + myRand, 'Edit Permit Applicant', 50, 90);
			}
			else
			{
				//winHandle = eval('window.open("permitcontactedit.asp?permitcontactid=' + sContactId +'&permitid=<%=iPermitId%>&permitstatusid=<%=iPermitStatusId%>&type=applicant&rand=' + myRand + '", "_contact", "width=900,height=600,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
				showModal('permitcontactedit.asp?permitcontactid=' + sContactId +'&permitid=<%=iPermitId%>&permitstatusid=<%=iPermitStatusId%>&type=applicant&rand=' + myRand, 'Edit Permit Contact', 50, 90);
			}
		}

		function EditPrimaryContact( sContactId )
		{
			myRand = parseInt(Math.random() * 99999999 );
			//winHandle = eval('window.open("permitapplicantedit.asp?userid=' + sContactId + '&updatetitle=1&detailid=primarycontactdetails&permitid=<%=iPermitId%>&permitstatusid=<%=iPermitStatusId%>&rand=' + myRand + '", "_contact", "width=800,height=800,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function SelectPrimaryContact( )
		{
			//winHandle = eval('window.open("primarycontactpicker.asp?permitid=<%=iPermitId%>", "_contact", "width=800,height=800,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function EditPermitAddress( iPermitAddressId, iCanSave )
		{
			var title = "Edit";
			if (iCanSave != 1)
			{
				title = "View";
			}
			myRand = parseInt(Math.random() * 99999999 );
			//winHandle = eval('window.open("permitaddressedit.asp?permitaddressid=' + iPermitAddressId + '&cansave=' + iCanSave + '&permitid=<%=iPermitId%>&permitstatusid=<%=iPermitStatusId%>&rand=' + myRand + '", "_contact", "width=900,height=800,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('permitaddressedit.asp?permitaddressid=' + iPermitAddressId + '&cansave=' + iCanSave + '&permitid=<%=iPermitId%>&permitstatusid=<%=iPermitStatusId%>&rand=' + myRand, title + ' Permit Address', 50, 90);
		}

		function EditManualPrice( iPermitFeeId )
		{
			//winHandle = eval('window.open("manualfeeedit.asp?permitfeeid=' + iPermitFeeId + '", "_contact", "width=400,height=300,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('manualfeeedit.asp?permitfeeid=' + iPermitFeeId, 'Edit Manual Fee', 30, 30);
		}

		function EditPermitNumber( )
		{
			//winHandle = eval('window.open("permitnumberedit.asp?permitid=<%=iPermitId%>", "_contact", "width=400,height=300,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('permitnumberedit.asp?permitid=<%=iPermitId%>', 'Edit Permit Number', 20, 30);
		}

		function ViewResidentialUnitFee( iPermitFeeId )
		{
			//winHandle = eval('window.open("viewresidentialunitfee.asp?permitfeeid=' + iPermitFeeId + '", "_contact", "width=650,height=320,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('viewresidentialunitfee.asp?permitfeeid=' + iPermitFeeId, 'Residential Unit Fee Formula', 30, 30);
		}

		function ViewValuationFee( iPermitFeeId )
		{
			//winHandle = eval('window.open("viewvaluationfee.asp?permitfeeid=' + iPermitFeeId + '", "_contact", "width=650,height=320,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('viewvaluationfee.asp?permitfeeid=' + iPermitFeeId, 'Valuation Fee Formula', 30, 35);
		}

		function ViewPercentageFee( iPermitFeeId )
		{
			//winHandle = eval('window.open("viewpercentagefee.asp?permitfeeid=' + iPermitFeeId + '", "_contact", "width=650,height=320,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('viewpercentagefee.asp?permitfeeid=' + iPermitFeeId, 'Percentage Fee Formula', 30, 30);
		}

		function ViewSqFootageFee( iPermitFeeId )
		{
			//winHandle = eval('window.open("viewsqfootagefee.asp?permitfeeid=' + iPermitFeeId + '", "_contact", "width=650,height=320,location=0toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('viewsqfootagefee.asp?permitfeeid=' + iPermitFeeId, 'Fee Formula', 30, 30);
		}

		function ViewCuFtFee( iPermitFeeId )
		{
			//winHandle = eval('window.open("viewcuftfee.asp?permitfeeid=' + iPermitFeeId + '", "_contact", "width=650,height=320,location=0toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('viewsqftfee.asp?permitfeeid=' + iPermitFeeId, 'Fee Formula', 30, 30);
		}

		function EditHourlyRateFee( iPermitFeeId )
		{
			//winHandle = eval('window.open("hourlyratefeeedit.asp?permitfeeid=' + iPermitFeeId + '", "_contact", "width=650,height=320,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('hourlyratefeeedit.asp?permitfeeid=' + iPermitFeeId, 'Hourly Rate Fee', 30, 40);
		}

		function ViewConstructionTypeFee( iPermitFeeId )
		{
			//winHandle = eval('window.open("viewconstructiontypefee.asp?permitfeeid=' + iPermitFeeId + '", "_contact", "width=650,height=300,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('viewconstructiontypefee.asp?permitfeeid=' + iPermitFeeId, 'Construction Type Fee', 30, 30);
		}

		function EditFixtureFee( iPermitFeeId )
		{
			//winHandle = eval('window.open("fixturefeeedit.asp?permitfeeid=' + iPermitFeeId + '", "_contact", "width=900,height=600,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('fixturefeeedit.asp?permitfeeid=' + iPermitFeeId, 'Fixture Fees', 35, 70);
		}

		function PayInvoices()
		{
			//winHandle = eval('window.open("paypermitinvoices.asp?permitid=<%=iPermitId%>", "_contact", "width=900,height=700,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,resizable=1,left=' + w + ',top=' + h + '")');
			showModal('paypermitinvoices.asp?permitid=<%=iPermitId%>', '', 70, 70);
		}

		function CreateInvoice( iPromptOnNoJobValue )
		{
			var bOpenWin = false;

			if ( iPromptOnNoJobValue == 1)
			{
				if ($("#jobvalue").val() == '0.00')
				{
					if (confirm('The Job Value on this permit is set to $0.00, do you still wish to create this invoice?'))
					{
						bOpenWin = true;
					}
					else
					{
						// Invoices tab
						tabView.set('activeIndex',2);
						$("#jobvalue").focus();
						// Already set bOpenWin to false
					}
				}
				else
				{
					bOpenWin = true;
				}
			}
			else
			{
				bOpenWin = true;
			}

			if ( bOpenWin )
			{
				$("#newinvoicebtn").prop('disabled',true);
				//winHandle = eval('window.open("invoicecreate.asp?permitid=<%=iPermitId%>", "_contact", "width=900,height=600,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
				showModal('invoicecreate.asp?permitid=<%=iPermitId%>', 'New Invoice', 35, 70);
			}
		}

		function ViewInvoice( iInvoiceid )
		{
			//winHandle = eval('window.open("viewinvoice.asp?invoiceid=' + iInvoiceid + '", "_contact", "width=900,height=700,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,resizable=1,left=' + w + ',top=' + h + '")');
			showModal('viewinvoice.asp?invoiceid=' + iInvoiceid, 'View Invoice', 70, 90);
		}

		function ViewReview( iPermitReviewId )
		{
			//winHandle = eval('window.open("permitreviewedit.asp?permitreviewid=' + iPermitReviewId +'", "_contact", "width=790,height=700,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,resizable=1,left=' + w + ',top=' + h + '")');
			showModal('permitreviewedit.asp?permitreviewid=' + iPermitReviewId, 'Permit Review', 50, 80);
		}

		function ViewInspection( iPermitInspectionId )
		{
			//winHandle = eval('window.open("permitinspectionedit.asp?permitinspectionid=' + iPermitInspectionId +'", "_contact", "width=800,height=800,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('permitinspectionedit.asp?permitinspectionid=' + iPermitInspectionId, 'Permit Inspection', 50, 80);
		}
		function ViewIR( iPermitIRId )
		{
			//winHandle = eval('window.open("permitinspectionedit.asp?permitinspectionid=' + iPermitInspectionId +'", "_contact", "width=800,height=800,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('inspectionreport.asp?permitinspectionreportid=' + iPermitIRId, 'Permit Inspection Report', 50, 80);
		}
		function EmailIR( iPermitIRId )
		{
			//winHandle = eval('window.open("permitinspectionedit.asp?permitinspectionid=' + iPermitInspectionId +'", "_contact", "width=800,height=800,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('emailinspectionreport.asp?permitinspectionreportid=' + iPermitIRId, 'Email Permit Inspection Report', 30, 50);
		}
		function PrintIR( iPermitIRId )
		{
			//winHandle = eval('window.open("permitinspectionedit.asp?permitinspectionid=' + iPermitInspectionId +'", "_contact", "width=800,height=800,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('inspectionreport.asp?print=true&permitinspectionreportid=' + iPermitIRId, 'Print Permit Inspection Report', 50, 80);
		}
		function CopyIR( iPermitIRId )
		{
			//winHandle = eval('window.open("permitinspectionedit.asp?permitinspectionid=' + iPermitInspectionId +'", "_contact", "width=800,height=800,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('inspectionreport.asp?copy=true&permitinspectionreportid=' + iPermitIRId, 'Permit Inspection Report', 50, 80);
		}
		function NewIR( )
		{
			//winHandle = eval('window.open("inspectionpicker.asp?permitid=<%=iPermitId%>", "_contact", "width=800,height=300,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('inspectionreport.asp?permitid=<%=iPermitId%>', 'Permit Inspection Report', 50, 80);
		}

		function ViewInvoiceSummary( )
		{
			var iContactId = document.frmPermit.viewinvoicecontactid.options[document.frmPermit.viewinvoicecontactid.selectedIndex].value;
			//winHandle = eval('window.open("viewinvoicesummary.asp?permitid=<%=iPermitId%>&permitcontactid=' + iContactId + '", "_contact", "width=900,height=700,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,resizable=1,left=' + w + ',top=' + h + '")');
			showModal('viewinvoicesummary.asp?permitid=<%=iPermitId%>&permitcontactid=' + iContactId, 'View Invoice Summary', 70, 90);
		}

		function ViewPermit( iDocumentid )
		{
			if (iDocumentid > 0)
			{
				//winHandle = eval('window.open("viewpermitpdf.asp?permitid=<%=iPermitId%>", "_contact", "width=900,height=700,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,resizable=1,left=' + w + ',top=' + h + '")');
				showModal('viewpermitpdf.asp?permitid=<%=iPermitId%>&permitdoc=1&permitdocumentid=' + iDocumentid, 'Print Permit', 50, 80);
			}
			else
			{
				//winHandle = eval('window.open("viewpermit.asp?permitid=<%=iPermitId%>", "_contact", "width=900,height=700,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,resizable=1,left=' + w + ',top=' + h + '")');
				showModal('viewpermit.asp?permitid=<%=iPermitId%>', 'Print Permit', 50, 80);
			}
		}

		function ViewPermitXMLDocument( iDocumentid )
		{
			if (iDocumentid > 0)
			{
				//winHandle = eval('window.open("viewpermitXMLPDF.asp?permitid=<%=iPermitId%>&permitdoc=1&permitdocumentid=' + iDocumentid + '", "_contact", "width=900,height=700,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,resizable=1,left=' + w + ',top=' + h + '")');
				showModal('viewpermitXMLPDF.asp?permitid=<%=iPermitId%>&permitdoc=1&permitdocumentid=' + iDocumentid, 'Print Permit', 50, 80);
			}
			else
			{
				//winHandle = eval('window.open("viewpermit.asp?permitid=<%=iPermitId%>", "_contact", "width=900,height=700,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,resizable=1,left=' + w + ',top=' + h + '")');
				showModal('viewpermit.asp?permitid=<%=iPermitId%>', 'Print Permit', 50, 80);
			}
		}

		function ViewPermitDocument( )
		{
			var iPermitDocumentId;
			iPermitDocumentId = document.frmPermit.permitdocument.options[document.frmPermit.permitdocument.selectedIndex].value;
			//winHandle = eval('window.open("viewpermitdocument.asp?permitid=<%=iPermitId%>&permitdocumentid=' + iPermitDocumentId + '", "_contact", "width=900,height=700,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,resizable=1,left=' + w + ',top=' + h + '")');
		}

		function ViewPermitWord( )
		{
			var iPermitDocumentId;
			iPermitDocumentId = document.frmPermit.permitdocument.options[document.frmPermit.permitdocument.selectedIndex].value;
			//winHandle = eval('window.open("viewpermitwordtest.asp?permitid=<%=iPermitId%>", "_contact", "width=900,height=700,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,resizable=1,left=' + w + ',top=' + h + '")');
		}

		function ViewPermitXML( )
		{
			var iPermitDocumentId;
			iPermitDocumentId = document.frmPermit.permitdocument.options[document.frmPermit.permitdocument.selectedIndex].value;
			//winHandle = eval('window.open("viewpermitXMLPDF.asp?permitid=<%=iPermitId%>&permitdoc=0&permitdocumentid=' + iPermitDocumentId + '", "_contact", "width=900,height=700,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,resizable=1,left=' + w + ',top=' + h + '")');
			showModal('viewpermitXMLPDF.asp?permitid=<%=iPermitId%>&permitdoc=0&permitdocumentid=' + iPermitDocumentId, 'Print Permit', 50, 80);
		}

		function ViewDetails()
		{
			//winHandle = eval('window.open("viewpermitdetails.asp?permitid=<%=iPermitId%>", "_contact", "width=900,height=700,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,resizable=1,left=' + w + ',top=' + h + '")');
			showModal('viewpermitdetails.asp?permitid=<%=iPermitId%>', 'Permit Details', 50, 80);
		}

		function EditExpirationDate( )
		{
			//winHandle = eval('window.open("expirationdateedit.asp?permitid=<%=iPermitId%>", "_contact", "width=330,height=200,toolbar=0,statusbar=0,scrollbars=0,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function EditCriticalDate( iCriticalDateType )
		{
			//winHandle = eval('window.open("criticaldateedit.asp?permitid=<%=iPermitId%>&criticaldatetype=' + iCriticalDateType + '", "_contact", "width=330,height=200,toolbar=0,statusbar=0,scrollbars=0,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('criticaldateedit.asp?permitid=<%=iPermitId%>&criticaldatetype=' + iCriticalDateType, '', 20, 35);
		}

		function EditInvoiceDate( iInvoiceid, iFieldId ) 
		{
			//winHandle = eval('window.open("invoicedateedit.asp?invoiceid=' + iInvoiceid + '&updatefield=invoicedate' + iFieldId + '", "_contact", "width=330,height=200,toolbar=0,statusbar=0,scrollbars=0,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('invoicedateedit.asp?invoiceid=' + iInvoiceid + '&updatefield=invoicedate', 'Invoice Date Edit', 20, 30);
		}

		function EditPaidDate( iPaymentid, iFieldId, iInvoiceid )
		{
			//winHandle = eval('window.open("paiddateedit.asp?permitid=<%=iPermitId%>&paymentid=' + iPaymentid + '&updatefield=paymentdate' + iFieldId + '&invoiceid=' + iInvoiceid + '", "_contact", "width=350,height=200,toolbar=0,statusbar=0,scrollbars=0,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('paiddateedit.asp?permitid=<%=iPermitId%>&paymentid=' + iPaymentid + '&updatefield=paymentdate', 'Paid Date Edit', 20, 30);
		}

		function SetAlert()
		{
			//winHandle = eval('window.open("permitalertedit.asp?permitid=<%=iPermitId%>", "_contact", "width=650,height=320,location=0,toolbar=0,statusbar=0,scrollbars=0,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('permitalertedit.asp?permitid=<%=iPermitId%>', 'Set Alert', 30, 30);
		}

		function ChangePermitType()
		{
			showModal('permittypechange.asp?permitid=<%=iPermitId%>', 'Change Permit Type', 20, 30);
		}

		function ShowNone()
		{
			return false;
		}

		function DeletePermit( iPermitId )
		{
			if (confirm("Are you sure you want to delete this permit?"))
			{
				$("deletepermitbutton").prop('disabled',true);
				doAjax('deletepermit.asp', 'permitid=' + iPermitId, 'BackToList', 'get', '0');
			}
		}

		function BackToList( sReturn )
		{
			location.href = 'permitlist.asp';
		}

		function RemoveFees()
		{
			if (confirm("Remove the selected fees?"))
			{
				//alert('Removing the fees.');
				var iRow = 1;
				var tbl = document.getElementById("feelist");
				// Check the feelist rows for any selected for removal
				for (var t = 0; t <= parseInt(document.frmPermit.maxfees.value); t++)
				{
					// See if a row exists for this one
					if ($("#removefee" + t).length)
					{
						// If it is marked for removal, remove it
						if ($("#removefee" + t).is(':checked'))
						{
							//alert(document.getElementById("permitfeeid" + t).value);
							doAjax('removepermitfee.asp', 'permitfeeid=' + $("#permitfeeid" + t).val(), '', 'get', '0');
							//alert(iRow);
							tbl.deleteRow(iRow);
							//iRow--;
						}
						else
						{
							iRow++;
						}
					}
				}
			}
		}

		function RemoveReviews()
		{
			if (confirm("Remove the selected reviews?"))
			{
				var iRow = 1;
				var tbl = document.getElementById("reviewlist");
				// Check the reviewlist rows for any selected for removal
				for (var t = 0; t <= parseInt(document.frmPermit.maxreviews.value); t++)
				{
					// See if a row exists for this one
					if (document.getElementById("removereview" + t))
					{
						// If it is marked for removal, remove it
						if (document.getElementById("removereview" + t).checked == true)
						{
							//alert(document.getElementById("permitreviewid" + t).value);
							doAjax('removepermitreview.asp', 'permitreviewid=' + document.getElementById("permitreviewid" + t).value, '', 'get', '0');
							//alert(iRow);
							tbl.deleteRow(iRow);
							//iRow--;
						}
						else
						{
							iRow++;
						}
					}
				}
			}
		}

		function RemoveInspections()
		{
			if (confirm("Remove the selected inspections?"))
			{
				var iRow = 1;
				var tbl = document.getElementById("inspectionlist");
				// Check the inspectionlist rows for any selected for removal
				for (var t = 0; t <= parseInt($("#maxinspections").val()); t++)
				{
					// See if a row exists for this one
					if ($("#removeinspection" + t).length)
					{
						// If it is marked for removal, remove it
						if ($("#removeinspection" + t).is(':checked'))
						{
							//alert(document.getElementById("permitreviewid" + t).value);
							doAjax('removepermitinspection.asp', 'permitinspectionid=' + $("#permitinspectionid" + t).val(), '', 'get', '0');
							tbl.deleteRow(iRow);
						}
						else
						{
							iRow++;
						}
					}
				}
			}
		}
		function RemoveIRs()
		{
			if (confirm("Remove the selected inspection reports?"))
			{
				var iRow = 1;
				var tbl = document.getElementById("inspectionreportlist");
				// Check the inspectionlist rows for any selected for removal
				for (var t = 0; t <= parseInt($("#maxirs").val()); t++)
				{
					// See if a row exists for this one
					if ($("#removeIR" + t).length)
					{
						// If it is marked for removal, remove it
						if ($("#removeIR" + t).is(':checked'))
						{
							//alert(document.getElementById("permitreviewid" + t).value);
							doAjax('removepermitinspectionreport.asp', 'permitinspectionreportid=' + $("#permitinspectionreportid" + t).val(), '', 'get', '0');
							tbl.deleteRow(iRow);
						}
						else
						{
							iRow++;
						}
					}
				}
			}
		}

		function RemoveContractorRows()
		{
			var iRow = 0;
			var tbl = document.getElementById("contractorlist");
			// Check the Contractor rows for any selected for removal
			for (var t = 0; t <= parseInt(document.frmPermit.maxcontractors.value); t++)
			{
				// See if a row exists for this one
				if (document.getElementById("removepermitcontactid" + t))
				{
					// If it is marked for removal, remove it
					if (document.getElementById("removepermitcontactid" + t).checked == true)
					{
						// Fire off an Ajax Job to remove them
						//alert(document.getElementById("permitcontactid" + t).value);
						doAjax('removepermitcontact.asp', 'permitcontactid=' + document.getElementById("permitcontactid" + t).value, '', 'get', '0');
						tbl.deleteRow(iRow);
						iRow--;
					}
					else
					{
						iRow++;
					}
				}
			}
		}

		function RemoveAttachments()
		{
			if (confirm("Remove the selected attachments?"))
			{
				var iRow = 1;
				var tbl = document.getElementById("attachmentlist");
				// Check the attachment rows for any selected for removal
				for (var t = 0; t <= parseInt(document.frmPermit.maxattachments.value); t++)
				{
					// See if a row exists for this one
					if ($("#removeattachment" + t).length)
					{
						// If it is marked for removal, remove it
						if ($("#removeattachment" + t).is(':checked') == true)
						{
							// Fire off an Ajax Job to remove them
							doAjax('permitattachmentremove.asp', 'permitattachmentid=' + $("#permitattachmentid" + t).val(), '', 'get', '0');
							tbl.deleteRow(iRow);
							//iRow--;
						}
						else
						{
							iRow++;
						}
					}
				}
			}
		}

		function validate()
		{
			//alert(tabView.get("activeIndex"));
			$("#activetab").val(tabView.get("activeIndex"));

			// Check that a locaton is entered if that field exists
			if ($("#permitlocation").length)
			{
				if ($("#permitlocation").val() == "")
				{
					alert("A location is required for this permit.\nPlease enter one and try saving again.");
					$("#permitlocation").focus();
					return;
				}
			}

			if ($("#permitisinbuildingpermitcategory").val() == 'yes')
			{
				// Validate that the Use Type has been selected
				if (document.frmPermit.usetypeid.options[document.frmPermit.usetypeid.selectedIndex].value == 0 )
				{
					tabView.set('activeIndex',0);
					alert("Use Type is required for various reports.\nPlease select one and try saving again.");
					$("usetypeid").focus();
					return;
				}
			}

			// Validate the residential units
			if ($("#residentialunits").length)
			{
				if ($("#residentialunits").val() != '')
				{
					// Remove any extra spaces
					$("#residentialunits").val(removeSpaces($("#residentialunits").val()));
					//Remove commas that would cause problems in validation
					$("#residentialunits").val(removeCommas($("#residentialunits").val()));

					rege = /^\d*$/;
					Ok = rege.test($("#residentialunits").val());
					if ( ! Ok )
					{
						tabView.set('activeIndex',0);
						alert("The value of 'Residential Units' must be a positive integer.\nPlease correct this and try saving again.");
						$("#residentialunits").focus();
						return;
					}
				}
			}

			// Validate the Occupants
			if ($("#occupants").length)
			{
				if ($("#occupants").val() != '')
				{
					// Remove any extra spaces
					$("#occupants").val(removeSpaces($("#occupants").val()));
					//Remove commas that would cause problems in validation
					$("#occupants").val(removeCommas($("#occupants").val()));

					rege = /^\d*$/;
					Ok = rege.test($("#occupants").val());
					if ( ! Ok )
					{
						tabView.set('activeIndex',0);
						alert("The value of 'Occupants' must be a positive integer, or blank.\nPlease correct this and try saving again.");
						$("occupants").focus();
						return;
					}
				}
			}

			// validate the custom fields on the details tab (0)
			if (parseInt($("#maxcustompermitfields").val()) > parseInt(0))
			{
				for (var t = 1; t <= parseInt($("#maxcustompermitfields").val()); t++)
				{
					if ($("#customfield" + t).val() != '')
					{
						// need to validate the values in the date, money and integer field types
						if ($("#fieldtypebehavior" + t).val() == 'date')
						{
							if (! isValidDate($("#customfield" + t).val()))
							{
								tabView.set('activeIndex',0);
								//alert("The value entered into this date field must be a valid date in the format of MM/DD/YYYY.\nPlease enter it again.");
								$("#customfield" + t).focus();
								inlineMsg($("#customfield" + t).attr('id'),'<strong>Invalid Value: </strong>This field must be a valid date in the format of MM/DD/YYYY.',8,$("#customfield" + t).attr('id'));
								return false;
							}
						}

						if ($("#fieldtypebehavior" + t).val() == 'integer')
						{
							
							var rege = /^-{0,1}\d*$/
							var Ok = rege.exec($("#customfield" + t).val());
							if ( ! Ok )
							{
								tabView.set('activeIndex',0);
								//alert("The value entered into this field must be a valid integer.\nPlease enter it again.");
								$("#customfield" + t).focus();
								inlineMsg($("#customfield" + t).attr('id'),'<strong>Invalid Value: </strong>This field must be a valid integer.',8,$("#customfield" + t).attr('id'));
								return false;
							}
						}

						if ($("#fieldtypebehavior" + t).val() == 'money')
						{
							// force it to have one decimal point and 2 numbers after it
							var rege = /^\d*\.{1}\d{2}$/
							var Ok = rege.exec($("#customfield" + t).val());
							if ( ! Ok )
							{
								tabView.set('activeIndex',0);
								//alert("The value entered into this field must be a valid number in currency format.\nPlease enter it again.");
								$("#customfield" + t).focus();
								inlineMsg($("#customfield" + t).attr('id'),'<strong>Invalid Value: </strong>This field must be a valid number in currency format.',8,$("#customfield" + t).attr('id'));
								return false;
							}
						}
					}
				}
			}

			// Validate the Job Value
			if ($("#jobvalue") != '')
			{
				// Remove any extra spaces
				$("#jobvalue").val(removeSpaces($("#jobvalue").val()));
				//Remove commas that would cause problems in validation
				$("#jobvalue").val(removeCommas($("#jobvalue").val()));

				rege = /^\d*\.?\d{0,2}$/;
				Ok = rege.test($("#jobvalue").val());
				if ( ! Ok )
				{
					tabView.set('activeIndex',2);
					alert("The 'Job Value' must be in currency format.\nPlease correct this and try saving again.");
					$("#jobvalue").focus();
					return;
				}
				else
				{
					$("#jobvalue").val(format_number(Number($("#jobvalue").val()),2));
				}
			}

			// Validate the Finished Sq Ft
			if (document.getElementById("finishedsqft").value != '')
			{
				// Remove any extra spaces
				document.getElementById("finishedsqft").value = removeSpaces(document.getElementById("finishedsqft").value);
				//Remove commas that would cause problems in validation
				document.getElementById("finishedsqft").value = removeCommas(document.getElementById("finishedsqft").value);

				rege = /^\d*\.?\d{0,2}$/;
				Ok = rege.test(document.getElementById("finishedsqft").value);
				if ( ! Ok )
				{
					tabView.set('activeIndex',2);
					alert("The 'Finished Sq Ft' must be numeric with up to two decimal places.\nPlease correct this and try saving again.");
					document.getElementById("finishedsqft").focus();
					return;
				}
				else
				{
					document.getElementById("finishedsqft").value = format_number(Number(document.getElementById("finishedsqft").value),2);
				}
			}

			// Validate the UnFinished Sq Ft
			if (document.getElementById("unfinishedsqft").value != '')
			{
				// Remove any extra spaces
				document.getElementById("unfinishedsqft").value = removeSpaces(document.getElementById("unfinishedsqft").value);
				//Remove commas that would cause problems in validation
				document.getElementById("unfinishedsqft").value = removeCommas(document.getElementById("unfinishedsqft").value);

				rege = /^\d*\.?\d{0,2}$/;
				Ok = rege.test(document.getElementById("unfinishedsqft").value);
				if ( ! Ok )
				{
					tabView.set('activeIndex',2);
					alert("The 'Unfinished Sq Ft' must be numeric with up to two decimal places.\nPlease correct this and try saving again.");
					document.getElementById("unfinishedsqft").focus();
					return;
				}
				else
				{
					document.getElementById("unfinishedsqft").value = format_number(Number(document.getElementById("unfinishedsqft").value),2);
				}
			}

			// Validate the Other Sq Ft
			if (document.getElementById("othersqft").value != '')
			{
				// Remove any extra spaces
				document.getElementById("othersqft").value = removeSpaces(document.getElementById("othersqft").value);
				//Remove commas that would cause problems in validation
				document.getElementById("othersqft").value = removeCommas(document.getElementById("othersqft").value);

				rege = /^\d*\.?\d{0,2}$/;
				Ok = rege.test(document.getElementById("othersqft").value);
				if ( ! Ok )
				{
					tabView.set('activeIndex',2);
					<% if bOrgHasVolume then %>
						alert("The 'Volume' must be numeric with up to two decimal places.\nPlease correct this and try saving again.");
					<% else %>
						alert("The 'Other Sq Ft' must be numeric with up to two decimal places.\nPlease correct this and try saving again.");
					<% end if %>
						document.getElementById("othersqft").focus();
					return;
				}
				else
				{
					document.getElementById("othersqft").value = format_number(Number(document.getElementById("othersqft").value),2);
				}
			}

			// Validate the Linear Ft
//			if (document.getElementById("linearft").value != '')
//			{
				// Remove any extra spaces
//				document.getElementById("linearft").value = removeSpaces(document.getElementById("linearft").value);
				//Remove commas that would cause problems in validation
//				document.getElementById("linearft").value = removeCommas(document.getElementById("linearft").value);

//				rege = /^\d*\.?\d{0,2}$/;
//				Ok = rege.test(document.getElementById("linearft").value);
//				if ( ! Ok )
//				{
//					tabView.set('activeIndex',2);
//					alert("The 'Linear Ft' must be numeric with up to two decimal places.\nPlease correct this and try saving again.");
//					document.getElementById("linearft").focus();
//					return;
//				}
//				else
//				{
//					document.getElementById("linearft").value = format_number(Number(document.getElementById("linearft").value),2);
//				}
//			}

			// Validate the Cubic Ft
//			if (document.getElementById("cuft").value != '')
//			{
				// Remove any extra spaces
//				document.getElementById("cuft").value = removeSpaces(document.getElementById("cuft").value);
//				//Remove commas that would cause problems in validation
//				document.getElementById("cuft").value = removeCommas(document.getElementById("cuft").value);

//				rege = /^\d*\.?\d{0,2}$/;
//				Ok = rege.test(document.getElementById("cuft").value);
//				if ( ! Ok )
//				{
//					tabView.set('activeIndex',2);
//					alert("The 'Cubic Ft' must be numeric with up to two decimal places.\nPlease correct this and try saving again.");
//					document.getElementById("cuft").focus();
//					return;
//				}
//				else
//				{
//					document.getElementById("cuft").value = format_number(Number(document.getElementById("cuft").value),2);
//				}
//			}

			// Validate the Total Sq Ft
//			if (document.getElementById("totalsqft").value != '')
//			{
//				// Remove any extra spaces
//				document.getElementById("totalsqft").value = removeSpaces(document.getElementById("totalsqft").value);
//				//Remove commas that would cause problems in validation
//				document.getElementById("totalsqft").value = removeCommas(document.getElementById("totalsqft").value);
//
//				rege = /^\d*\.?\d{0,2}$/;
//				Ok = rege.test(document.getElementById("totalsqft").value);
//				if ( ! Ok )
//				{
//					tabView.set('activeIndex',2);
//					alert("The 'Total Sq Ft' must be numeric with up to two decimal places.\nPlease correct this and try saving again.");
//					document.getElementById("totalsqft").focus();
//					return;
//				}
//				else
//				{
//					document.getElementById("totalsqft").value = format_number(Number(document.getElementById("totalsqft").value),2);
//				}
//			}


			// Validate the Hours
			if (document.getElementById("examinationhours").value != '')
			{
				// Remove any extra spaces
				document.getElementById("examinationhours").value = removeSpaces(document.getElementById("examinationhours").value);
				//Remove commas that would cause problems in validation
				document.getElementById("examinationhours").value = removeCommas(document.getElementById("examinationhours").value);

				rege = /^\d*\.?\d{0,2}$/;
				Ok = rege.test(document.getElementById("examinationhours").value);
				if ( ! Ok )
				{
					tabView.set('activeIndex',2);
					alert("The 'Hours' must be numeric with up to two decimal places.\nPlease correct this and try saving again.");
					document.getElementById("examinationhours").focus();
					return;
				}
				else
				{
					document.getElementById("examinationhours").value = format_number(Number(document.getElementById("examinationhours").value),2);
				}
			}

			// Post the page
			//alert( "Validated");
			document.frmPermit.submit();
		}

		function SetUpPage()
		{
			setMaxLength();
		}

		function ChangeStatus(newstatus)
		{
			if (confirm("Are you sure that you want to change the status of this permit???"))
			{
				//$("#changestatusbutton").prop('disabled',true);
				$(".wfbtnlock").prop('disabled',true);
				// Fire off job to change the status
				doAjax('statuschange.asp', 'permitid=' + document.getElementById("permitid").value + '&newstatus=' + newstatus, 'StatusReturn', 'get', '0');
			}
		}

		function StatusReturn( sReturn )
		{
			//alert(sReturn);
			//if (sReturn == 'UPDATED')
			//{
				location.href = "permitedit.asp?permitid=<%=iPermitId%>&activetab=" + tabView.get("activeIndex");
			//}
		}

		function CorrectInvoiceJobValue( iPermitId )
		{
			doAjax('correctinvoicejobvalue.asp', 'permitid=' + iPermitId, 'CorrectionReturn', 'get', '0');
		}

		function CorrectionReturn( sReturn )
		{
			alert(sReturn + ': Job Values are in sync, now.');
		}

		function ChangeHold( iChange )
		{
			var sMessage = "Are you sure that you want to ";
			if (iChange == 1)
			{
				sMessage += "place this permit on hold?";
			}
			else
			{
				sMessage += "release the hold on this permit?"
			}

			if (confirm(sMessage))
			{
				// check that internal notes are entered for placing on hold.
				if (iChange == 1 && document.getElementById("internalcomment").value == "")
				{
					tabView.set('activeIndex',7);
					alert('Please provide some internal notes about why this permit is being placed on hold.\nThen click the "Place Hold" button again.');
					document.getElementById("internalcomment").focus();
					return;
				}
				var sParameter = 'permitid=' + encodeURIComponent(document.getElementById("permitid").value);
				sParameter += '&isonhold=' + iChange;
				sParameter += '&internalcomment=' + encodeURIComponent(document.getElementById("internalcomment").value);
				sParameter += '&externalcomment=' + encodeURIComponent(document.getElementById("externalcomment").value);

				// Fire off job to change the hold flag
				doAjax('changeholdflag.asp', sParameter , 'StatusReturn', 'post', '0');
			}
		}

		function VoidPermit( iChange )
		{
			var sMessage = "Are you sure that you want to ";
			if (iChange == 1)
			{
				sMessage += "void this permit?";
			}
			else
			{
				sMessage += "remove the void on this permit?"
			}

			if (confirm(sMessage))
			{
				// check that internal notes are entered for placing on hold.
				if (iChange == 1 && document.getElementById("internalcomment").value == "")
				{
					tabView.set('activeIndex',7);
					alert('Please provide some internal notes about why this permit is being voided.\nThen click the "Void Permit" button again.');
					$("internalcomment").focus();
					return;
				}
				var sParameter = 'permitid=' + encodeURIComponent(document.getElementById("permitid").value);
				sParameter += '&isvoided=' + iChange;
				sParameter += '&internalcomment=' + encodeURIComponent(document.getElementById("internalcomment").value);
				sParameter += '&externalcomment=' + encodeURIComponent(document.getElementById("externalcomment").value);

				// Fire off job to change the void flag
				doAjax('changevoidflag.asp', sParameter , 'StatusReturn', 'post', '0');
			}
		}

		function UnComplete( sPriorStatus )
		{
			if (confirm( 'Are you sure you wish to change this permit back to ' + sPriorStatus + ' status?' ))
			{
				if ($("#internalcomment").val() == "")
				{
					tabView.set('activeIndex',7);
					alert('Please provide some internal notes about why this permit is being moved back to ' + sPriorStatus + ' status.\nThen click the "Move Permit" button again.');
					$("internalcomment").focus();
					return;
				}
				var sParameter = 'permitid=' + encodeURIComponent($("#permitid").val());
				sParameter += '&internalcomment=' + encodeURIComponent($("#internalcomment").val());
				sParameter += '&externalcomment=' + encodeURIComponent($("#externalcomment").val());

				// Fire off job to change the uncomplete the permit
				doAjax('uncomplete.asp', sParameter , 'StatusReturn', 'post', '0');
			}
		}

		var myMouseOver = new Object(); 
		myMouseOver.eventHandler = function(event) { 
			event.srcElement.style.backgroundColor = '#93bee1';
			event.srcElement.style.cursor='pointer';
		} 

		var myMouseOut = new Object(); 
		myMouseOut.eventHandler = function(event) { 
			event.srcElement.style.backgroundColor = '';
			event.srcElement.style.cursor='';
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

	<style>

		textarea
		{
			width:100%;
		}
		input[size="100"]
		{
			width:100%;
		}
	</style>
</head>

<body class="yui-skin-sam" onload="SetUpPage();">

	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="permitalert"><%=sAlertMsg%></div>
		<div id="centercontent" class="col3">

			<form name="frmPermit" action="permitupdate.asp" method="post">
				<input type="hidden" name="permitid" id="permitid" value="<%=iPermitId%>" />
				<input type="hidden" name="activetab" id="activetab" value="<%=iActiveTabId%>" />
				<input type="hidden" name="feetotal" id="feetotal" value="<%=sFeeTotal%>" />
				<input type="hidden" name="invoicedtotal" id="invoicedtotal" value="<%=sInvoicedTotal%>" />
				<input type="hidden" name="paidtotal" id="paidtotal" value="<%=sPaidTotal%>" />
				<input type="hidden" name="statuschg" id="statuschg" value= "" />
				<input type="hidden" name="permitisinbuildingpermitcategory" id="permitisinbuildingpermitcategory" value="<%=sPermitIsInBuildingPermitCategory%>" />

				<!--TWF:= <%=request.cookies("user")("userid")%> -->
				<input type="button" class="button ui-button ui-widget ui-corner-all" value="<< Back to Permit List" onclick="javascript:window.location='permitlist.asp';" />
				<br>
				<br>
				<div style="float:left;width:40%;">
					Permit #: <span class="keyinfo"><span id="permitnumberdisplay"><%=sPermitNo%></span></span>
<%					
					CanChangeNumber = UserHasPermission( Session("UserId"), "permit number override" )
					href = "javascript:EditPermitNumber( );"
					tooltipclass=""
					tooltip = ""
					If not CanChangeNumber or LCase(sPermitNo) = "none" or not bCanSaveChanges or bIsOnHold or bIsVoided then
						tooltipclass="tooltip"
						tooltip = "<span class=""tooltiptext"">The permit # cannot change because<br />"
						if not CanChangeNumber then tooltip = tooltip & "You don't have permission.<br />"
						if not bCanSaveChanges then tooltip = tooltip & "You cannot save changes.<br />"
						if LCase(sPermitNo) = "none"  then tooltip = tooltip & "The Permit Number isn't set.<br />"
						if bIsOnHold then tooltip = tooltip & "Permit is on hold.<br />"
						if bIsVoided then tooltip = tooltip & "Permit is voided.<br />"
						tooltip = tooltip & "</span>"
						href="javascript: void(0)"
					End If	%>
					<a class="<%=tooltipclass%>" href="<%=href%>"><i class="fa fa-pencil"></i><%=tooltip%></a>
					<br />
					<br />
					Permit Type: <span class="keyinfo"><span id="permittype"><%=GetPermitTypeDesc( iPermitId, true ) %></span></span>
<%					
					href="javascript:ChangePermitType();"
					tooltipclass=""
					tooltip = ""
					CanChangeType = UserHasPermission( session("UserID"), "can change permit types" )
					StatusAllowsDeletes = PermitStatusAllowsDeletes( iPermitId )
					If not bUserIsRootAdmin and ( not CanChangeType or not StatusAllowsDeletes  ) Then
						tooltipclass="tooltip"
						tooltip = "<span class=""tooltiptext"">The permit type cannot change because<br />"
						if not CanChangeType then tooltip = tooltip & "You don't have permission.<br />"
						if not StatusAllowsDeletes then tooltip = tooltip & "The permit status.<br />"
						tooltip = tooltip & "</span>"
						href="javascript: void(0)"
					End If	%>
					<a class="<%=tooltipclass%>" href="<%=href%>"><i class="fa fa-pencil"></i><%=tooltip%></a>
					<br />
					<br />
					Permit Status: <span class="keyinfo"><%=sPermitStatus%></span> &nbsp; &nbsp;
<%					If sButtonText <> "" And bStatusCanChange Then %>					
						<!--input type="button" id="changestatusbutton" class="button ui-button ui-widget ui-corner-all" value="<%=sButtonText%>" onclick="ChangeStatus();" /> &nbsp; &nbsp; -->
<%					End If %>


				</div>

				<div style="float:left;width:60%;">
<%				If bNeedsAddress Or bNeedsLocation Then			%>
					<div id="jobinfo" style="height:100%">
						<table cellpadding="0" border="0" cellspacing="0" id="LDtable">
<%						If bNeedsAddress Then	%>
							<tr>
								<%
									strStreetNumber = ""
									strStreetName = ""
								%>
								<td nowrap="nowrap" class="LDtagcell">Job Site Address:</td><td><span class="keyinfo"><span id="jobaddress"><%=GetPermitLocation( iPermitId, sLegalDescription, sListedOwner, iPermitAddressId, sCounty, sParcelid, False )%></span></span>
<%								If bCanSaveChanges Then		%>
									<a href="javascript:EditPermitAddress('<%=iPermitAddressId%>',1);"><i class="fa fa-pencil"></i></a>
<%								Else 	%>
									 &nbsp; &nbsp; <input type="button" class="button ui-button ui-widget ui-corner-all" value="View" onclick="EditPermitAddress('<%=iPermitAddressId%>',0)" />
<%								End If	%>
								<%

 								lcl_orghasfeature_issuelocation                    = orghasfeature("issue location")
								%>
       								<%if lcl_orghasfeature_issuelocation then%>
									&nbsp; &nbsp; <input type="button" class="button ui-button ui-widget ui-corner-all" value="Check For Code Violations" onClick="showModal('../action_line/action_line_list.asp?selectIssueStreetNumber=<%=strStreetNumber%>&selectIssueStreet=<%=strStreetName%>&fromdate=1/1/2000&todate=<%=date()%>','Code Violations',90,90);" />
								<%end if %>
								</td>
							</tr>
							<tr>
								<td nowrap="nowrap" class="LDtagcell">Listed Owner:</td><td><span class="keyinfo1"><span id="listedowner"><%=sListedOwner%></span></span></td>
							</tr>
							<tr>
								<td nowrap="nowrap" class="LDtagcell">Legal Description:</td><td><span class="keyinfo1"><span id="legaldescription"><%=sLegalDescription%></span></span></td>
							</tr>
<%						End If		

						If bNeedsLocation Then	%>
							<tr>
								<td class="label" nowrap="nowrap" valign="top">Location:</td>
								<td><textarea id="permitlocation" name="permitlocation" rows="5" cols="100" maxlength="1000"><%=sPermitLocation%></textarea></td>
							</tr>
<%						End If		

						' in case it is decided to show these in the future
						If Not bNeedsAddress And Not bNeedsLocation Then		%>
							<tr>
								<td class="label" nowrap="nowrap" valign="top">&nbsp</td>
								<td><strong>This permit does not require an address or a location.</strong></td>
							</tr>
<%						End If 
%>
						</table>
					</div>
				</p>
<%				End If			%>
				</div>
				<div style="clear:both;"></div>

				<p id="criticaldates">
						<table cellpadding="2" cellspacing="0" border="0" class="tableadmin">
							<tr>
								<th>Applied<br />For Permit</th>
								<th>Permit<br />Released</th>
								<th>Permit<br />Approved</th>
								<th>Permit<br />Issued</th>
								<th>Permit<br />Completed</th>
								<th>Application<br />Expires</th>
							</tr>
							<tr>
								<td align="center">
									<span class="detaildata" id="applieddate"><%=sApplied%></span>
<%
									href = "javascript:EditCriticalDate( '1' );"
									tooltipclass=""
									tooltip = ""
									If (not bUserIsRootAdmin and not bUserCanChangeCriticalDates) or bIsCompleted Then
										tooltipclass="tooltip"
										tooltip = "<span class=""tooltiptext"">This date cannot be changed because<br />"
										if not bUserCanChangeCriticalDates then tooltip = tooltip & "You don't have permission.<br />"
										if bIsCompleted then tooltip = tooltip & "The permit is complete.<br />"
										tooltip = tooltip & "</span>"
										href="javascript: void(0)"
									End If		%>
									<a class="<%=tooltipclass%>" href="<%=href%>"><i class="fa fa-pencil"></i><%=tooltip%></a>
								</td>
								<td align="center">
									<span class="detaildata" id="releaseddate"><%=sReleased%></span>
<%
									href = "javascript:EditCriticalDate( '2' );"
									tooltipclass=""
									tooltip = ""
									If (not bUserIsRootAdmin and not bUserCanChangeCriticalDates) or bIsCompleted Then
										tooltipclass="tooltip"
										tooltip = "<span class=""tooltiptext"">This date cannot be changed because<br />"
										if not bUserCanChangeCriticalDates then tooltip = tooltip & "You don't have permission.<br />"
										if bIsCompleted then tooltip = tooltip & "The permit is complete.<br />"
										tooltip = tooltip & "</span>"
										href="javascript: void(0)"
									End If		%>
									<%if sReleased <> "" then%><a class="<%=tooltipclass%>" href="<%=href%>"><i class="fa fa-pencil"></i><%=tooltip%></a><%end if%>
								</td>
								<td align="center">
									<span class="detaildata" id="approveddate"><%=sApproved%></span>
<%
									href = "javascript:EditCriticalDate( '3' );"
									tooltipclass=""
									tooltip = ""
									If (not bUserIsRootAdmin and not bUserCanChangeCriticalDates) or bIsCompleted Then
										tooltipclass="tooltip"
										tooltip = "<span class=""tooltiptext"">This date cannot be changed because<br />"
										if not bUserCanChangeCriticalDates then tooltip = tooltip & "You don't have permission.<br />"
										if bIsCompleted then tooltip = tooltip & "The permit is complete.<br />"
										tooltip = tooltip & "</span>"
										href="javascript: void(0)"
									End If		%>
									<%if sApproved <> "" then%><a class="<%=tooltipclass%>" href="<%=href%>"><i class="fa fa-pencil"></i><%=tooltip%></a><%end if%>
								</td>
								<td align="center">
									<span class="detaildata" id="issueddate"><%=sIssued%></span>
<%
									href = "javascript:EditCriticalDate( '4' );"
									tooltipclass=""
									tooltip = ""
									If (not bUserIsRootAdmin and not bUserCanChangeCriticalDates) or bIsCompleted Then
										tooltipclass="tooltip"
										tooltip = "<span class=""tooltiptext"">This date cannot be changed because<br />"
										if not bUserCanChangeCriticalDates then tooltip = tooltip & "You don't have permission.<br />"
										if bIsCompleted then tooltip = tooltip & "The permit is complete.<br />"
										tooltip = tooltip & "</span>"
										href="javascript: void(0)"
									End If		%>
									<%if sIssued <> "" then%><a class="<%=tooltipclass%>" href="<%=href%>"><i class="fa fa-pencil"></i><%=tooltip%></a><%end if%>
								</td>
								<td align="center"><span class="detaildata"><%=sCompleted%></span></td>
								<td align="center">
			<%						If bHasExpirationDate Then 		%>
										<span class="detaildata" id="expirationdate"><%=sExpires%></span>
<%
										href = "javascript:EditCriticalDate( '5' );"
										tooltipclass=""
										tooltip = ""
										If not bCanChangeExpirationDate Then
											tooltipclass="tooltip"
											tooltip = "<span class=""tooltiptext"">This date cannot be changed because of the permit status.</span>"
											href="javascript: void(0)"
										End If		%>
										<%if sExpires <> "" then%><a class="<%=tooltipclass%>" href="<%=href%>"><i class="fa fa-pencil"></i><%=tooltip%></a><%end if%>
			<%						Else	%>
										&nbsp;
			<%						End If	%>
								</td>

							</tr>
							</table>
				</p>
				<p>
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
				<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" id="savebutton" onclick="validate();" >Save Changes<%=tooltip%></button> &nbsp; &nbsp; 
				</p>

				<div id="demo" class="yui-navset">
					<ul class="yui-nav">
						<li><a href="#tab1"><em>Details</em></a></li>
						<li><a href="#tab2"><em>Contacts</em></a></li>
						<li><a href="#tab3"><em>Reviews</em></a></li>
						<li><a href="#tab4"><em>Fees</em></a></li>
						<li><a href="#tab5"><em>Invoices</em></a></li>
						<li><a href="#tab6"><em>Inspections</em></a></li>
						<li><a href="#tab7"><em>Attachments</em></a></li>
						<li><a href="#tab8"><em>Notes</em></a></li>
						<li><a href="#tab9"><em>Documents</em></a></li>
						<% if session("orgid") = "139" or session("orgid") = "181" then %>
						<!--li><a href="#tab10"><em>Inspect. Reports</em></a></li-->
						<%end if%>
					</ul>            
					<div class="yui-content">
						<div id="tab1"> <!-- Details -->

								<table cellpadding="2" cellspacing="0" border="0" id="permitdetail">
<%								If bPermitIsInBuildingPermitCategory Then	%>
									<tr><td class="label" nowrap="nowrap">Use Type:</td><td colspan="3"><% ShowUseTypes iUseTypeId %></td></tr>
<%								Else	%>
									<input type="hidden" id="usetypeid" name="usetypeid" value="0" />
<%								End If 

								If PermitHasDetail( iPermitid, "useclassid" ) Then	%>
									<tr><td class="label" nowrap="nowrap">Use Class:</td><td colspan="3"><% ShowUseClasses iUseClassId %></td></tr>
<%								End If	
								If PermitHasDetail( iPermitid, "descriptionofwork" ) Then	%>
									<tr><td class="label" nowrap="nowrap">Description of Work:</td><td colspan="3"><span class="detaildata"><input type="text" name="descriptionofwork" value="<%=sDescriptionOfWork%>" size="100" maxlength="200" /></span></td></tr>
<%								End If	
								If PermitHasDetail( iPermitid, "workclass" ) Then	%>
									<tr><td class="label" nowrap="nowrap">Work Class:</td><td colspan="3"><% ShowWorkClass iWorkClassId %></td></tr>
<%								End If	
								If PermitHasDetail( iPermitid, "workscope" ) Then	%>
									<tr><td class="label" nowrap="nowrap">Work Scope:</td><td colspan="3"><% ShowWorkScopes iWorkScopeId %></td></tr>
<%								End If	
								If PermitHasDetail( iPermitid, "constructiontype" ) Then	%>
									<tr><td class="label" nowrap="nowrap">Type of Construction:</td><td><% ShowConstructionTypes iConstructionTypeId %> &nbsp; &nbsp; <input type="button" class="button ui-button ui-widget ui-corner-all" value="Check Rate" onclick="CheckConstRate();" /></td></tr>
									<tr><td class="label" nowrap="nowrap">Occupancy Type:</td><td nowrap="nowrap"><% ShowOccupancyTypes iOccupancyTypeId %></td></tr>
<%								End If	
								If PermitHasDetail( iPermitid, "existinguse" ) Then	%>
									<tr><td class="label" nowrap="nowrap">Existing Use:</td><td colspan="3"><span class="detaildata"><input type="text" name="existinguse" value="<%=sExistingUse%>" size="100" maxlength="150" /></span></td></tr>
<%								End If	
								If PermitHasDetail( iPermitid, "proposeduse" ) Then	%>
									<tr><td class="label" nowrap="nowrap">Proposed Use:</td><td colspan="3"><span class="detaildata"><input type="text" name="proposeduse" value="<%=sProposedUse%>" size="100" maxlength="150" /></span></td></tr>
<%								End If	
								If PermitHasDetail( iPermitid, "approvedas" ) Then	%>
									<tr><td class="label" nowrap="nowrap">Approved As:</td><td colspan="3"><span class="detaildata"><input type="text" name="approvedas" value="<%=sApprovedAs%>" size="100" maxlength="150" /></span></td></tr>
<%								End If	
								If PermitHasDetail( iPermitid, "residentalunits" ) Then	%>
								<tr><td class="label" nowrap="nowrap">New Residential Units:</td><td colspan="3"><span class="detaildata"><input type="text" id="residentialunits" name="residentialunits" value="<%=sResidentialUnits%>" size="10" maxlength="10" /></span></td></tr>
<%								End If	
								If PermitHasDetail( iPermitid, "occupants" ) Then	%>
									<tr><td class="label" nowrap="nowrap">Occupants:</td><td colspan="3"><span class="detaildata"><input type="text" id="occupants" name="occupants" value="<%=sOccupants%>" size="10" maxlength="10" /></span></td></tr>
<%								End If	

								' Fields For Lansing, IL
								If PermitHasDetail( iPermitid, "structuredimensions" ) Then	%>
									<tr><td class="label" nowrap="nowrap">Structure Dimensions:</td><td colspan="3">Length:<span class="detaildata"><input type="text" id="structurelength" name="structurelength" value="<%=sStructureLength%>" size="10" maxlength="10" /></span> &nbsp; Width:<span class="detaildata"><input type="text" id="structurewidth" name="structurewidth" value="<%=sStructureWidth%>" size="10" maxlength="10" /></span> &nbsp; Height:<span class="detaildata"><input type="text" id="structureheight" name="structureheight" value="<%=sStructureHeight%>" size="10" maxlength="10" /></span></td>
<%								End If	
								If PermitHasDetail( iPermitid, "zoning" ) Then	%>
									<tr><td class="label" nowrap="nowrap">Zoning:</td><td colspan="3"><span class="detaildata"><input type="text" id="zoning" name="zoning" value="<%=sZoning%>" size="25" maxlength="25" /></span></td></tr>
<%								End If		
								If PermitHasDetail( iPermitid, "planno" ) Then	%>
									<tr><td class="label" nowrap="nowrap">Plan #:</td><td colspan="3"><span class="detaildata"><input type="text" id="plannumber" name="plannumber" value="<%=sPlanNumber%>" size="25" maxlength="25" /></span></td></tr>
<%								End If		
								If PermitHasDetail( iPermitid, "demolishexisting" ) Then	%>
									<tr><td class="label" nowrap="nowrap">&nbsp;</td><td colspan="3" align="left"><span class="detaildata"><input type="checkbox" id="demolishexistingstructure" name="demolishexistingstructure" <%=sDemolishExistingStructure%> /></span> Demolish Existing Structure</td></tr>
<%								End If		
								If PermitHasDetail( iPermitid, "landfill" ) Then	%>
									<tr><td class="label" nowrap="nowrap">Landfill Name:</td><td colspan="3"><span class="detaildata"><input type="text" id="landfillname" name="landfillname" value="<%=sLandFillName%>" size="30" maxlength="30" /></span></td></tr>
									<tr><td class="label" nowrap="nowrap">Landfill City:</td><td colspan="3"><span class="detaildata"><input type="text" id="landfillcity" name="landfillcity" value="<%=sLandFillCity%>" size="30" maxlength="30" /></span></td></tr>
									<tr><td class="label" nowrap="nowrap">Landfill Phone:</td><td colspan="3"><span class="detaildata"><input type="text" id="landfillphone" name="landfillphone" value="<%=sLandFillPhone%>" size="30" maxlength="30" /></span></td></tr>
<%								End If	
								' End of Lansing IL fields

								If PermitHasDetail( iPermitid, "permitnotes" ) Then	%>
									<tr><td class="label" nowrap="nowrap" valign="top">Permit Notes:</td><td colspan="3"><textarea name="permitnotes" rows="5" cols="100" maxlength="1000"><%=sPermitnotes%></textarea></td></tr>
<%								End If	
								If PermitHasDetail( iPermitid, "tempconotes" ) Then	
									If bHasTempCO Then		%>
										<tr><td class="label" nowrap="nowrap" valign="top" colspan="4">Temporary CO Stipulations, Conditions, Variances:</td></tr>
										<tr><td class="label" nowrap="nowrap" valign="top">&nbsp;</td><td colspan="3"><textarea name="tempconotes" rows="5" cols="100" maxlength="1000"><%=sTempCONotes%></textarea></td></tr>
<%									End If			%>
<%								End If	
								If PermitHasDetail( iPermitid, "conotes" ) Then	
									If bHasCO Then		%>
										<tr><td class="label" nowrap="nowrap" valign="top" colspan="4">Cert of Occupancy Stipulations, Conditions, Variances:</td></tr>
										<tr><td class="label" nowrap="nowrap" valign="top">&nbsp;</td><td colspan="3"><textarea name="conotes" rows="5" cols="100" maxlength="1000"><%=sCONotes%></textarea></td></tr>
<%									End If			
								End If		

								' Show custom fields here
								iMaxCustomPermitFields = ShowCustomPermitFields( iPermitid )
%>
								</table>
								<input type="hidden" id="maxcustompermitfields" name="maxcustompermitfields" value="<%=iMaxCustomPermitFields%>" />

						</div>
						<div id="tab2"> <!-- Contacts -->
						<br />
							<%
							If PermitHasLicenseRequirement( iPermitId ) Then
								response.write "<p id=""requiredlicenses"">This permit requires the following licenses: "
								ShowRequiredPermitLicenses iPermitId, "1"
								response.write "</p>"
							End If		
%>
							<table cellpadding="2" cellspacing="0" border="0" id="permitcontact">
								<tr><td class="contactlabel" nowrap="nowrap" valign="top">Applicant:</td><td nowrap="nowrap" valign="top"><% ShowPermitApplicant iPermitId %></td></tr>
								<tr><td class="contactlabel" nowrap="nowrap" valign="top">Primary Contact:</td><td nowrap="nowrap" valign="top"><input type="text" id="primarycontact" name="primarycontact" value="<%=sPrimaryContact %>" size="100" maxlength="150" /></td></tr>
								<tr><td class="contactlabel" nowrap="nowrap" valign="top">Billing Contact:</td><td nowrap="nowrap" valign="top"><% GetPermitContact iPermitId, "isbillingcontact", False %></td></tr>
								<tr><td class="contactlabel" nowrap="nowrap" valign="top">Primary Contractor:</td><td nowrap="nowrap" valign="top"><% GetPermitContact iPermitId, "isprimarycontractor", True %></td></tr>
								<tr><td class="contactlabel" nowrap="nowrap" valign="top">Architect/Engineer:</td><td nowrap="nowrap" valign="top"><% GetPermitContact iPermitId, "isarchitect", True %></td></tr>
								<tr><td class="contactlabel" nowrap="nowrap" valign="top">Plans By:</td><td nowrap="nowrap" valign="top"><% ShowPermitPlansBy iPermitId, iPlansByContactId %></td></tr>
								<tr><td class="contactlabel" nowrap="nowrap" valign="top">Other Contractors:</td>
									<td nowrap="nowrap" valign="top">
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
										<button id="addcontractor" <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" onclick="SelectContact( 'iscontractor' );" >Add Contractor<%=tooltip%></button> &nbsp; 
										<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" onClick="RemoveContractorRows()"  >Remove Selected<%=tooltip%></button>
										<br /><br />
										<table cellpadding="0" cellspacing="0" border="0" id="contractorlist">
<%											iMaxContractors = ShowContractorList( iPermitId )		%>									
										</table>
										<input type="hidden" id="maxcontractors" name="maxcontractors" value="<%=iMaxContractors%>" />
									</td>
								</tr>
								<tr>
									<td class="contactlabel" nowrap="nowrap" valign="top">Prior Contacts:</td>
									<td nowrap="nowrap" valign="top">
										<table cellpadding="0" cellspacing="0" border="0" id="priorcontactlist">
<%											iMaxPriorContacts = ShowPriorContactList( iPermitId )		%>		
										</table>
										<input type="hidden" id="maxpriorcontacts" name="maxpriorcontacts" value="<%=iMaxPriorContacts%>" />
									</td>
								</tr>
							</table>
						</div>

						<div id="tab3"> <!-- Reviews -->
							<p class="tabpage">
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
							&nbsp; <button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" onclick="SelectReviews( );" >Add A Review<%=tooltip%></button> &nbsp;&nbsp; 
							<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" onclick="RemoveReviews();" >Remove Selected Reviews<%=tooltip%></button> &nbsp;&nbsp; 
							<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" onClick="EditPermitReviewTypes();" >Edit Permit Review Types</button> &nbsp;&nbsp;
							<%if GetPermitStatusBlockReview( iPermitId ) then response.write "<span class=""red"">Permit must be released before Reviews can be completed.</span>"	' in permitcommonfunctions.asp%>
							</p>
							<table cellpadding="2" cellspacing="0" border="0" class="feetable" id="reviewlist">
								<tr><th>Remove</th><th>Review</th><th>Status</th><th>Date</th><th>Reviewer</th></tr>
<%								iMaxReviews = ShowReviewList( iPermitId )		%>									
							</table>
							<input type="hidden" id="maxreviews" name="maxreviews" value="<%=iMaxReviews%>" />
						</div>

						<div id="tab4"> <!-- Fees -->
							<p class="tabpage">
								<table cellpadding="2" cellspacing="0" border="0" class="feetable">
									<tr><th>Finished Sq Ft</th><th>+</th><th>Unfinished Sq Ft</th><th>=</th><th>Total Sq Ft</th>
									<% If bOrgHasVolume Then %>
										<th>Volume</th>
									<% Else %>
										<th>Other Sq Ft</th>
									<% end if %>
									<th>Job Value</th><th>Examination<br />Hours</th><th>Fee Total</th></tr>

									<tr>
										<td align="center">
											<input type="text" name="finishedsqft" id="finishedsqft" value="<%=sFinishedSqFt%>" size="10" maxlength="10" />
										</td>
										<td align="center">+</td>
										<td align="center">
											<input type="text" name="unfinishedsqft" id="unfinishedsqft" value="<%=sUnFinishedSqFt%>" size="10" maxlength="10" />
										</td>
										<td align="center">=</td>
										<td align="center">
											<%=sTotalSqFt%>
										</td>
										<td align="center">
											<input type="text" name="othersqft" id="othersqft" value="<%=sOtherSqFt%>" size="10" maxlength="10" />
										</td>
										<td align="center">
											<input type="text" name="jobvalue" id="jobvalue" value="<%=sJobValue%>" size="10" maxlength="10" />
										</td>
										<td align="center">
											<input type="text" name="examinationhours" id="examinationhours" value="<%=sExaminationHours%>" size="10" maxlength="10" />
										</td>
										<td align="center">
											<span id="feetabfeetotal"><%=sFeeTotal%></span>
										</td>
									</tr>
								</table>
							</p>
							
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
							<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" onclick="SelectFee( );" >Add A Fee<%=tooltip%></button> &nbsp;&nbsp; 
							<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" onclick="RemoveFees();" >Remove Selected Fees<%=tooltip%></button> &nbsp;&nbsp; 
							<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" onClick="EditFeeTypes()" >Edit Fee Types<%=tooltip%></button> &nbsp; &nbsp;
							<input <%=disabled%> type="checkbox" name="waiveallfees" 
							<%	If bWaiveAllFees Then %>
								checked="checked" 
							<%	End If %>
							 /> <strong>Waive All Fees</strong>

							<p class="tabpage">
								<table cellpadding="2" cellspacing="0" border="0" class="feetable" id="feelist">
									<tr><th>Remove</th><th>Category</th><th>Description</th><th>Method</th>
<%									If OrgHasFeature( "up front fees" ) Then		%>
										<th>Up Front<br />Amount</th>
<%									End If							%>
									<th>Fee Amount</th></tr>
<%									iMaxFees = ShowFeeList( iPermitId )		%>									
								</table>
								<input type="hidden" id="maxfees" name="maxfees" value="<%=iMaxFees%>" />
							</p>
						</div>
						<div id="tab5"> <!-- Invoices -->
							<table cellpadding="2" cellspacing="0" border="0" class="feetable" id="balancelist">
							<caption>Balance</caption>
								<tr>
<%									If OrgHasFeature( "up front fees" ) Then		%>
										<th>Up Front<br />Fees</th>
<%									End If							%>
									<th>Total Fees</th><th>Non-Invoiced Fees</th><th>Invoiced Fees</th><th>Total Waived</th><th>Total Paid</th><th>Total Due</th></tr>
								<tr>
<%									If OrgHasFeature( "up front fees" ) Then		%>
										<td align="center"><%=sUpFrontFeeTotal%></td>
<%									End If							%>
									<td align="center"><span id="invoicetabfeetotal"><%=sFeeTotal%></span></td>
									<td align="center"><span id="invoicetabnoninvoicedtotal"><%=sNonInvoicedTotal%></span></td>
									<td align="center"><span id="invoicetabinvoicedtotal"><%=sInvoicedTotal%></span></td>
									<td align="center"><span id="invoicetabwaivedtotal"><%=sWaivedTotal%></span></td>
									<td align="center"><span id="invoicetabpaidtotal"><%=sPaidTotal%></span></td>
									<td align="center"><span id="invoicetabduetotal"><%=sDueTotal%></span></td>
								</tr>
							</table>
							</td></tr>
							<tr><td style="padding-left:50px;"><br /> &nbsp;
							<%
							onclick = "CreateInvoice(" & iPromptOnNoJobValue & ");"
							tooltipclass=""
							tooltip = ""
							disabled = ""
							If not StatusAllowsNewInvoices( iPermitStatusId ) or CDbl(sNonInvoicedTotal) <= CDbl(0.00) or bIsCompleted Then		' in permitcommonfunctions.asp  
								tooltipclass="tooltip"
								disabled = " disabled "
								tooltip = "<span class=""tooltiptext"">You cannot add an invoice because:<br />"
								if not StatusAllowsNewInvoices( iPermitStatusId ) then tooltip = tooltip & "The permit status.<br />"
								if CDbl(sNonInvoicedTotal) <= CDbl(0.00) then tooltip = tooltip & "The non invoiced fees aren't<br />greater than zero.<br />"
								if bIsCompleted then tooltip = tooltip & "The permit is complete.<br />"
								tooltip = tooltip & "</span>"
								onclick="void(0);"
							end if
							%>
							<button type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" <%=disabled%> id="newinvoicebtn" onclick="<%=onclick%>" >New Invoice<%=tooltip%></button> &nbsp;&nbsp; 
							<%
							onclick = "PayInvoices();"
							tooltipclass=""
							tooltip = ""
							disabled = ""
							If CDbl(sInvoicedTotal) <= CDbl(0.00) or CDbl(sDueTotal) <= CDbl(0.00) or bIsCompleted Then 
								tooltipclass="tooltip"
								disabled = " disabled "
								tooltip = "<span class=""tooltiptext"">You cannot pay the invoice(s) because:<br />"
								if CDbl(sInvoicedTotal) <= CDbl(0.00) then tooltip = tooltip & "The invoiced fees aren't greater than zero.<br />"
								if CDbl(sDueTotal) <= CDbl(0.00) then tooltip = tooltip & "The total due isn't greater than zero.<br />"
								if bIsCompleted then tooltip = tooltip & "The permit is complete.<br />"
								tooltip = tooltip & "</span>"
								onclick="void(0);"
							end if
							%>
							<button type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" <%=disabled%> onclick="<%=onclick%>" >Pay Invoices<%=tooltip%></button> &nbsp;&nbsp; 

							<%
							onclick = "ViewInvoiceSummary();"
							tooltipclass=""
							tooltip = ""
							disabled = ""
							If CDbl(sInvoicedTotal) <= CDbl(0.00) Then 
								tooltipclass="tooltip"
								disabled = " disabled "
								tooltip = "<span class=""tooltiptext"">You cannot view the invoice summary because:<br />"
								if CDbl(sInvoicedTotal) <= CDbl(0.00) then tooltip = tooltip & "The invoiced fees aren't greater than zero.<br />"
								tooltip = tooltip & "</span>"
								onclick="void(0);"
							end if
							%>
							<% ShowInvoiceContacts iPermitId %>
							<button type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" <%=disabled%> onclick="<%=onclick%>" >View Invoice Summary<%=tooltip%></button>
							<%
							onclick = "CorrectInvoiceJobValue( " & iPermitId & " );"
							tooltipclass=""
							tooltip = ""
							disabled = ""
							If CDbl(sInvoicedTotal) <= CDbl(0.00) or bIsCompleted or bIsOnHold or clng(iPromptOnNoJobValue) <> clng(1) or not PermitJobValueNotEqualToInvices( iPermitId, sJobValue ) Then 
								tooltipclass="tooltip"
								disabled = " disabled "
								tooltip = "<span class=""tooltiptext"">You cannot update this because:<br />"
								if CDbl(sInvoicedTotal) <= CDbl(0.00) then tooltip = tooltip & "The invoiced fees aren't greater than zero.<br />"
								if bIsCompleted then tooltip = tooltip & "The permit is complete.<br />"
								if bIsOnHold then tooltip = tooltip & "The permit is on hold.<br />"
								if clng(iPromptOnNoJobValue) <> clng(1) then tooltip = tooltip & "The invoices have a job value.<br />"
								if not PermitJobValueNotEqualToInvices( iPermitId, sJobValue ) then tooltip = tooltip & "The job value is equal to the invoices.<br />"
								tooltip = tooltip & "</span>"
								onclick="void(0);"
							end if
							%>
							<button type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" <%=disabled%> onclick="<%=onclick%>" >Update Invoice Job Value<%=tooltip%></button>

								<br /><br />
							</td></tr>
							<tr><td>
							<table cellpadding="2" cellspacing="0" border="0" class="feetable" id="invoicelist">
								<caption>Invoices</caption>
								<tr><th>Invoice #</th><th>Invoice<br />Date</th><th>Billed To</th><th>Status</th><th>Invoice Total</th><th>Paid<br />Date</th><th>Amount<br />Paid/Waived</th><th>Void</th></tr>
								<% ShowInvoices iPermitId %>
							</table>
						</div>

						<div id="tab6"> <!-- Inspections -->
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
							<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" onclick="SelectInspections( );">Add An Inspection<%=tooltip%></button> &nbsp;&nbsp; 
							<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" onclick="RemoveInspections();">Remove Selected Inspections<%=tooltip%></button> &nbsp;&nbsp; 
							<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" onClick="EditPermitInspectionTypes();" >Edit Permit Inspection Types</button> &nbsp;&nbsp;
							<input type="checkbox" name="alertapplicantofinspections" id = "alertapplicantofinspections" 
								<%	If bAlertApplicant Then 
										response.write " checked=""checked"" "
									End If		%>
							/> Email the applicant when an inspection is scheduled
<%							If Not GetPermitIsIssued( iPermitId ) Then %>
								<br /><span class="red">The permit must be issued before inspections can be scheduled.</span>
<%							End If		%>
							<br />
							<br />
							<table cellpadding="2" cellspacing="0" border="0" class="feetable" id="inspectionlist">
								<tr><th>Remove</th><th>Inspection</th><th>Reinspection</th><th>Status</th><th>Scheduled<br />Date</th><th>Inspected<br />Date</th><th>Inspector</th></tr>
<%								iMaxInspections = ShowInspectionList( iPermitId )		%>									
							</table>
							<input type="hidden" id="maxinspections" name="maxinspections" value="<%=iMaxInspections%>" />
						</div>
						<div id="tab7"> <!-- Attachments -->
							<p class="tabpage">
<%							If bCanSaveChanges Then		%>
								&nbsp; <input type="button" class="button ui-button ui-widget ui-corner-all" value="Add An Attachment" onclick="AddAttachments( );" /> &nbsp;&nbsp; 
<%								If UserHasPermission( Session("UserId"), "remove permit attachments" ) Then	%>
									<input type="button" class="button ui-button ui-widget ui-corner-all" value="Remove Selected Attachments" onclick="RemoveAttachments();" /> &nbsp;&nbsp; 
<%								End If	%>
<%							End If %>
							</p>
							<p>
								<table cellpadding="2" cellspacing="0" border="0" class="feetable" id="attachmentlist">
									<tr>
<%										If UserHasPermission( Session("UserId"), "remove permit attachments" ) Then	%>
											<th>Remove</th>
<%										End If	%>
									<th>File Name</th><th>Description</th><th>Date Added</th><th>Added By</th></tr>
<%									iMaxAttachments = ShowAttachmentList( iPermitId )		%>		
								</table>

								<input type="hidden" id="maxattachments" name="maxattachments" value="<%=iMaxAttachments%>" />
							</p>
						</div>
						<div id="tab8"> <!-- Notes -->
							<div id="newnotes_expand" onMouseOver="this.style.cursor='pointer';" onClick="toggleDisplayShow( 'newnotes' );">
								<strong><span id="newnotesimg">&ndash;</span> <u>New Notes:</u></strong>
							</div>
							<div id="newnotes" style="padding:5px;margin-top:5px;border:solid 1px #000000;">
							<table>
								<tr><td><strong>Internal Notes:</strong><br />
										<textarea id="internalcomment" name="internalcomment" rows="5" cols="80" maxlength="1000"></textarea>
									</td>
								</tr>
								<tr><td><strong>Public Notes:</strong><br />
										<textarea id="externalcomment" name="externalcomment" rows="5" cols="80" maxlength="1000"></textarea>
									</td>
								</tr>
							</table>
							</div>
							<div id="priornotes_expand" onMouseOver="this.style.cursor='pointer';" onClick="toggleDisplayShow( 'priornotes' );">
								<strong><span id="priornotesimg">&ndash;</span> <u>Prior Notes:</u></strong>
							</div>
							<div id="priornotes">
<%								ShowPermitNotes iPermitId		%>
							</div>
						</div>
						
						<div id="tab9"> <!-- Documents -->
							<p class="tabpage"><br />
<%
								ListPermitDocumentButtons iPermitId
%>
							</p>
						</div>
						<% if session("orgid") = "139" or session("orgid") = "181" then %>
						<!-- Inspection Reports -->
						<!--div id="tab10"> 
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
							<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" onclick="NewIR( );">Add An Inspection Report<%=tooltip%></button> &nbsp;&nbsp; 
							<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" onclick="RemoveIRs();">Remove Selected Inspection Reports<%=tooltip%></button>
							<br />
							<br />
							<style>
								table#inspectionreportlist tbody tr:hover,
								table#inspectionreportlist tbody tr:hover td
								{
									cursor:pointer;
									background-color:#93bee1;
								}

							</style>
							<table cellpadding="2" cellspacing="0" border="0" class="feetable" id="inspectionreportlist">
								<thead>
								<tr><th>Remove</th><th>Inspection Report</th><th></th><th></th><th></th></tr>
								</thead>
								<tbody>
<%								iMaxIRs = ShowInspectionReportList( iPermitId )		%>									
								</tbody>
							</table>
							<input type="hidden" id="maxirs" name="maxirs" value="<%=iMaxIRs%>" />
						</div-->
						<% end if%>

					</div>
				</div>
				<p>
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
				<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" id="savebutton" onclick="validate();" >Save Changes<%=tooltip%></button> &nbsp; &nbsp; 
				</p>
			</form>
			<!--END: EDIT FORM-->

		</div>
		<div class="col1">
     			<div class="dropdown right">
  				<button class="ui-button ui-widget ui-corner-all dd-green"><i class="fa fa-bars" aria-hidden="true"></i> Tools</button>
  				<div class="dropdown-content">
<%					' Delete Permit Button   -   
					CanDeletePermit = UserHasPermission( session("UserID"), "can delete permits" )
					StatusCanDelete = PermitStatusAllowsDeletes( iPermitId )
					tooltipclass=""
					tooltip = ""
					href = "javascript:DeletePermit(" & iPermitId & ");"
					If not bUserIsRootAdmin and (not CanDeletePermit or not StatusCanDelete ) Then
						tooltipclass="tooltip"
						tooltip = "<span class=""tooltiptext"">This cannot be deleted because<br />"
						if not CanDeletePermit then tooltip = tooltip & "You don't have permission.<br />"
						if not StatusCanDelete then tooltip = tooltip & "The permit status.<br />"
						tooltip = tooltip & "</span>"
						href="javascript: void(0)"
					end if %>
					<a class="<%=tooltipclass%>" href="<%=href%>" id="deletepermitbutton">Delete Permit<%=tooltip%></a>
<%					
					tooltipclass=""
					tooltip = ""
					If OrgHasFeature("xml pdf display") Then
						href = "javascript:ViewPermitXMLDocument(" & iDocumentid & ");"
					else
						href = "javascript:ViewPermit(" & iDocumentid & ");"
					end if
					If not bCanPrintPermit Then		
						tooltipclass="tooltip"
						tooltip = "<span class=""tooltiptext"">This cannot be printed because "
						tooltip = tooltip & "the current status Issued or Completed.</span>"
						href="javascript: void(0)"
					End If %>
					<a class="<%=tooltipclass%>" href="<%=href%>">Print Permit Document<%=tooltip%></a>

					<%
					href = "javascript:SetAlert();"
					tooltipclass=""
					tooltip = ""
					if not bCanSaveChanges then
						tooltipclass="tooltip"
						tooltip = "<span class=""tooltiptext"">You cannot set an alert because you cannot save changes.</span>"
						href="javascript: void(0)"
					end if
					%>
					<a class="<%=tooltipclass%>" href="<%=href%>">Set Alert<%=tooltip%></a>
					<%
					href = "javascript:ViewCO( " & iPermitId & ", 'temp', " & sTempCOAction & ", ' Temporary' );"
					tooltipclass=""
					tooltip = ""
					If not bCanIssueTempCO Then
						tooltipclass="tooltip"
						tooltip = "<span class=""tooltiptext"">You cannot Issue a Temporary CO<br />because the status doesn't allow it.</span>"
						href="javascript: void(0)"
					end if
					%>
					<!--a class="<%=tooltipclass%>" href="<%=href%>">Issue Temporary Certificate of Occupancy<%=tooltip%></a-->
					<%
					href = "javascript:ViewCO( " & iPermitId & ", '', " & sTempCOAction & ", '' );"
					tooltipclass=""
					tooltip = ""
					If not bCanIssueCO Then
						tooltipclass="tooltip"
						tooltip = "<span class=""tooltiptext"">You cannot Issue a CO<br />because the status doesn't allow it.</span>"
						href="javascript: void(0)"
					end if
					%>
					<!--a class="<%=tooltipclass%>" href="<%=href%>">Issue Certificate of Occupancy<%=tooltip%></a-->
					<%
					btntext = "Place"
					fxVal = 1
					If bIsOnHold Then	
						btntext = "Remove"
						fxVal = 0
					end if 
					href = "javascript:ChangeHold( " & fxVal & " );"
					tooltipclass=""
					tooltip = ""
					If not UserHasPermission( Session("UserId"), "can hold permits" ) or bIsCompleted Then 
						tooltipclass="tooltip"
						tooltip = "<span class=""tooltiptext"">You cannot change the hold status because<br />"
						if not UserHasPermission( Session("UserId"), "can hold permits" )then tooltip = tooltip & "You don't have permissions.<br />"
						if bIsCompleted then tooltip = tooltip & "The permit is complete.</span>"
						href="javascript: void(0)"
					end if
					%>
					<a class="<%=tooltipclass%>" href="<%=href%>"><%=btntext%> Hold<%=tooltip%></a>
					<%
					fxVal = 1
					btntext = "Void Permit"
					If bIsVoided Then
						fxVal = 0
						btntext = "Remove Void"
					end if
					href = "javascript:VoidPermit( " & fxVal & " );"
					tooltipclass=""
					tooltip = ""
					If not UserHasPermission( Session("UserId"), "can void permits" ) or bIsCompleted or bIsOnHold Then
						tooltipclass="tooltip"
						tooltip = "<span class=""tooltiptext"">You cannot change the void status because<br />"
						if not UserHasPermission( Session("UserId"), "can void permits" )then tooltip = tooltip & "You don't have permissions.<br />"
						if bIsOnHold then tooltip = tooltip & "The permit is on hold.<br />"
						if bIsCompleted then tooltip = tooltip & "The permit is complete.</span>"
						href="javascript: void(0)"
					end if
					%>
					<a class="<%=tooltipclass%>" href="<%=href%>"><%=btntext%><%=tooltip%></a>
					<a href="javascript:ViewDetails();">Print Permit Details</a>



				</div>
			</div>
			<br />
			<br />
			<!--WORKFLOW-->
			<fieldset>
				<legend>Workflow</legend>
				<%			
				no = "<img src=""../images/x.png"" class=""wficon"">"
				yes = "<img src=""../images/check.png"" class=""wficon"">"
				disabled = ""
				tooltipclass=""
				tooltip = ""
				If sButtonText = "" or sButtonText <> "Release" or not bStatusCanChange Then 
					disabled = " disabled "
					tooltipclass="tooltip"
					tooltip = "<span class=""tooltiptext"">This button is disabled because<br />the status isn't ""Applied"".</span>"
				End If %>
				<%=yes%>Permit is Created<br />
				<button <%=disabled%> type="button" id="changestatusbutton" class="button ui-button ui-widget ui-corner-all wfbtnlock <%=tooltipclass%>" onclick="ChangeStatus('Released');" />Release Permit<%=tooltip%></button><br />
				<br />
				<br />
				<%			
				disabled = ""
				tooltipclass=""
				tooltip = ""
				If sButtonText = "" or sButtonText <> "Approve" or not bStatusCanChange Then 
					disabled = " disabled "
					tooltipclass="tooltip"
					tooltip = "<span class=""tooltiptext"">This button is disabled because<br />the status isn't ""Released"" or<br />the reviews aren't complete.</span>"
				End If %>
				<!--<%if sPermitStatus = "Released" then response.write yes else response.write no end if%>Permit is in Released status<br /-->
				<%if bAllReviewsComplete then response.write yes else response.write no end if%>Reviews are completed<br />
				<button <%=disabled%> type="button" id="changestatusbutton" class="button ui-button ui-widget ui-corner-all wfbtnlock <%=tooltipclass%>" onclick="ChangeStatus('Approved');">Approve Permit<%=tooltip%></button><br />
				<br />
				<br />
				<%			
				disabled = ""
				tooltipclass=""
				tooltip = ""
				If sButtonText = "" or sButtonText <> "Issue Permit" or not bStatusCanChange Then 
					disabled = " disabled "
					tooltipclass="tooltip"
					tooltip = "<span class=""tooltiptext"">This button is disabled because<br />the status isn't ""Approved"" or<br />the fees aren't complete.</span>"
				End If %>
				<!--<%if sPermitStatus = "Approved" then response.write yes else response.write no end if%>Permit is in Approved status<br /-->
				<%if not bSomeFeesSetToZero then response.write yes else response.write no end if%>All Fee calculations are complete<br />
				<%if bAllFeesPaidOrWaived and not bSomeFeesSetToZero then response.write yes else response.write no end if%>All Fees are paid or waived<br />
				<%If PermitHasLicenseRequirement( iPermitId ) Then%>
				<%if bHasReqLicenses then response.write yes else response.write no end if%>Required Licenses are present<br />
				<% end if %>
				<button <%=disabled%> type="button" id="changestatusbutton" class="button ui-button ui-widget ui-corner-all wfbtnlock <%=tooltipclass%>" onclick="ChangeStatus('Issued');">Issue Permit<%=tooltip%></button><br />
				<br />
				<br />
				<%			
				disabled = ""
				tooltipclass=""
				tooltip = ""
				'bAllInspectionsPassed = AllInspectionsPassed( iPermitId )
				bAllInspectionsPassed = PermitHasNoPendingInspections( iPermitId, 0)
				If sButtonText = "" or sButtonText <> "Complete Permit" or not bAllInspectionsPassed Then 
					disabled = " disabled "
					tooltipclass="tooltip"
					tooltip = "<span id=""completett"" class=""tooltiptext"">This button is disabled because<br />the status isn't ""Issued"" or<br />the inspections haven't passed.</span>"
				End If %>
				<!--<%if sPermitStatus = "Issued" then response.write yes else response.write no end if%>Permit is in Issued status<br /-->
				<%if bAllInspectionsPassed then response.write replace(yes, "class", "id=""apipimg"" class") else response.write replace(no, "class", "id=""apipimg"" class") end if%>All Permit Inspections passed<br />
				<button <%=disabled%> type="button" id="changestatusbutton" name="completepermitbtn" class="button ui-button ui-widget ui-corner-all wfbtnlock <%=tooltipclass%>" onclick="ChangeStatus('Completed');">Complete Permit<%=tooltip%></button><br />
				<br />
				<br />
				<style>
					#moveback .tooltiptext {margin-left:-80px;}
				</style>
<%			
				disabled = ""
				tooltipclass=""
				tooltip = ""
				If not bIsCompleted Then
					disabled = " disabled "
					tooltipclass="tooltip"
					tooltip = "<span class=""tooltiptext"">This button is disabled because<br />the status isn't ""Completed"".</span>"
				End If		%>
				<!--<%if sPermitStatus = "Completed" then response.write yes else response.write no end if%>Permit is in Completed status<br />-->
				<button <%=disabled%> id="moveback" type="button" class="button ui-button ui-widget ui-corner-all wfbtnlock <%=tooltipclass%>" onclick="UnComplete( '<%=sPriorStatus%>' );">Move Permit Back to Issued Status<%=tooltip%></button> &nbsp; &nbsp; 
				<br />
				<br />
			</fieldset>
		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  
	<!--#Include file="modal.asp"-->  
	<iframe id="makeics" src="about:blank" style="display:none"></iframe>
	<script>
	<%
	v2a = request.querystring("v2a")
	arrV2A = split(v2a,"|")
	if v2a <> "" then
		if instr(v2a,"NewInvoice") > 0 then response.write "ViewInvoice(" & arrV2A(1) & ");"
		if instr(v2a,"PaidInvoice") > 0 then 
			response.write "$('#viewinvoicecontactid').val('" & arrV2A(1) & "');"
			response.write "ViewInvoiceSummary();"
		end if
		if instr(v2a,"InspectionChange") > 0 then 
			response.write "ViewInspection(" & arrV2A(1) & ");"

			if arrV2A(2) = "yes" then response.write "$('#makeics').attr('src','makeics.asp?permitinspectionid=" & arrV2A(1) & "');"
		end if
	end if
	%>
	$("#invoicelist td a").click(function(e) {
   		// Do something
   		e.stopPropagation();
	});
	</script>

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

'-------------------------------------------------------------------------------------------------
' void GetPermitDetails iPermitId 
'-------------------------------------------------------------------------------------------------
Sub GetPermitDetails( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT permitnumberprefix, permitnumberyear, ISNULL(permitnumber,0) AS permitnumber, applieddate, waiveallfees, "
	sSql = sSql & " expirationdate, ISNULL(proposeduse, '') AS proposeduse, ISNULL(existinguse, '') AS existinguse, alertmsg, "
	sSql = sSql & " ISNULL(workclassid, 0) AS workclassid, ISNULL(workscopeid, 0) AS workscopeid, ISNULL(constructiontypeid,0) AS constructiontypeid, ISNULL(permitnotes,'') AS permitnotes, "
	sSql = sSql & " ISNULL(occupancytypeid, 0) AS occupancytypeid, ISNULL(descriptionofwork,'') AS descriptionofwork, "
	sSql = sSql & " ISNULL(usetypeid,0) AS usetypeid, ISNULL(useclassid,0) AS useclassid, releaseddate, approveddate, issueddate, completeddate, ISNULL(jobvalue,0.00) AS jobvalue, "
	sSql = sSql & " ISNULL(totalsqft,0.00) AS totalsqft, ISNULL(finishedsqft,0.00) AS finishedsqft, ISNULL(unfinishedsqft,0.00) AS unfinishedsqft, "
	sSql = sSql & " ISNULL(othersqft,0.00) AS othersqft, ISNULL(examinationhours,0.00) AS examinationhours, ISNULL(feetotal,0.00) AS feetotal, "
	sSql = sSql & " ISNULL(plansbycontactid,0) AS plansbycontactid, ISNULL(residentialunits,0) AS residentialunits, ISNULL(approvedas,'') AS approvedas, "
	sSql = sSql & " occupants, ISNULL(tempconotes,'') AS tempconotes, ISNULL(conotes,'') AS conotes, ISNULL(primarycontact,'') AS primarycontact, "
	sSql = sSql & " ISNULL(structurelength,'') AS structurelength, ISNULL(structurewidth,'') AS structurewidth, ISNULL(structureheight,'') AS structureheight, "
	sSql = sSql & " ISNULL(zoning,'') AS zoning, ISNULL(plannumber,'') AS plannumber, demolishexistingstructure, alertapplicantofinspections, "
	sSql = sSql & " ISNULL(landfillname,'') AS landfillname, ISNULL(landfillcity,'') AS landfillcity, ISNULL(landfillphone,'') AS landfillphone, "
	sSql = sSql & " ISNULL(permitlocation,'') AS permitlocation, permitlocationrequirementid "
	sSql = sSql & " FROM egov_permits WHERE permitid = " & iPermitId 
	'response.write "<!--" & sSql & "-->"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CLng(oRs("permitnumber")) > CLng(0) Then 
			'sPermitNo = oRs("permitnumberyear") & oRs("permitnumberprefix") & oRs("permitnumber")
			sPermitNo = GetPermitNumber( iPermitId )
		Else
			sPermitNo = "None"
		End If 
		sApplied = FormatDateTime(oRs("applieddate"),2)
		If IsNull(oRs("expirationdate")) Then 
			sExpires = ""
		Else 
			sExpires = FormatDateTime(oRs("expirationdate"), 2)
		End If 
		iPermitLocationRequirementId = oRs("permitlocationrequirementid")
		sPermitLocation = Replace(oRs("permitlocation"), Chr(34), "&quot;" )
		sProposedUse = Replace(oRs("proposeduse"), Chr(34), "&quot;" )
		sExistingUse = Replace(oRs("existinguse"), Chr(34), "&quot;" )
		iWorkClassId= oRs("workclassid")
		iWorkScopeId= oRs("workscopeid")
		iConstructionTypeId = oRs("constructiontypeid")
		iOccupancyTypeId = oRs("occupancytypeid")
		sDescriptionOfWork= Replace(oRs("descriptionofwork"), Chr(34), "&quot;" )
		iUseTypeId = oRs("usetypeid")
		iUseClassId = oRs("useclassid")
		If Not IsNull(oRs("releaseddate")) Then 
			sReleased = FormatDateTime(oRs("releaseddate"), 2)
		End If 
		If Not IsNull(oRs("approveddate")) Then 
			sApproved = FormatDateTime(oRs("approveddate"), 2)
		End If 
		If Not IsNull(oRs("issueddate")) Then 
			sIssued = FormatDateTime(oRs("issueddate"), 2)
		End If 
		If Not IsNull(oRs("completeddate")) Then 
			sCompleted = FormatDateTime(oRs("completeddate"), 2)
		End If 
		sJobValue = FormatNumber(oRs("jobvalue"),2,,,0)
		sTotalSqFt = FormatNumber(oRs("totalsqft"),2,,,0)
		sFinishedSqFt = FormatNumber(oRs("finishedsqft"),2,,,0)
		sUnFinishedSqFt = FormatNumber(oRs("unfinishedsqft"),2,,,0)
		sOtherSqFt = FormatNumber(oRs("othersqft"),2,,,0)
		sExaminationHours = FormatNumber(oRs("examinationhours"),2,,,0)
		sFeeTotal = FormatNumber(oRs("feetotal"),2,,,0)
		bWaiveAllFees = oRs("waiveallfees")
		sPermitnotes = oRs("permitnotes")
		sAlertMsg = oRs("alertmsg")
		iPlansByContactId = CLng(oRs("plansbycontactid"))
		sResidentialUnits = oRs("residentialunits")
		sApprovedAs = Replace(oRs("approvedas"), Chr(34), "&quot;" )
		If Not IsNull(oRs("occupants")) Then 
			sOccupants = CLng(oRs("occupants"))
		End If 
		sTempCONotes = oRs("tempconotes")
		sCONotes = oRs("conotes")
		sPrimaryContact = Replace(oRs("primarycontact"), Chr(34), "&quot;" )
		sStructureLength = oRs("structurelength")
		sStructureWidth = oRs("structurewidth")
		sStructureHeight = oRs("structureheight")
		sZoning = oRs("zoning")
		sPlanNumber = oRs("plannumber")
		If oRs("demolishexistingstructure") Then 
			sDemolishExistingStructure = " checked=""checked"" "
		Else
			sDemolishExistingStructure = ""
		End If 
		sLandFillName = Replace(oRs("landfillname"), Chr(34), "&quot;" )
		sLandFillCity = Replace(oRs("landfillcity"), Chr(34), "&quot;" )
		sLandFillPhone = Replace(oRs("landfillphone"), Chr(34), "&quot;" )
		If oRs("alertapplicantofinspections") Then 
			bAlertApplicant = True 
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub  


'--------------------------------------------------------------------------------------------------
' string GetPermitStatus( iPermitId, iNextStatus, iPermitStatusId, bCanPrintPermit, bCanChangeExpirationDate, bHasExpirationDate, bIsCompleted )
'--------------------------------------------------------------------------------------------------
Function GetPermitStatus( ByVal iPermitId, ByRef iNextStatus, ByRef iPermitStatusId, ByRef bCanPrintPermit, ByRef bCanChangeExpirationDate, ByRef bHasExpirationDate, ByRef bIsCompleted )
	Dim sSql, oRs

	sSql = "SELECT P.permitstatusid, S.permitstatus, S.nextpermitstatusid, S.canprintpermit, S.canchangeexpirationdate, "
	sSql = sSql & " S.hasexpirationdate, S.iscompletedstatus "
	sSql = sSql & " FROM egov_permits P, egov_permitstatuses S "
	sSql = sSql & " WHERE P.permitstatusid = S.permitstatusid AND P.permitid = " & iPermitId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitStatus = oRs("permitstatus")
		iNextStatus = CLng(oRs("nextpermitstatusid"))
		iPermitStatusId = oRs("permitstatusid")
		bCanPrintPermit = oRs("canprintpermit")
		bCanChangeExpirationDate = oRs("canchangeexpirationdate")
		bHasExpirationDate = oRs("hasexpirationdate")
		bIsCompleted = oRs("iscompletedstatus")
	Else
		GetPermitStatus = ""
		iNextStatus = 0
		iPermitStatusId = 1
		bCanPrintPermit = False 
		bCanChangeExpirationDate = False 
		bHasExpirationDate = False 
		bIsCompleted = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetButtonText( iNextStatus )
'--------------------------------------------------------------------------------------------------
Function GetButtonText( ByVal iNextStatus )
	Dim sSql, oRs

	sSql = "SELECT buttontext FROM egov_permitstatuses "
	sSql = sSql & " WHERE orgid = " & session("orgid") & " AND permitstatusid = " & iNextStatus 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetButtonText = oRs("buttontext")
	Else 
		GetButtonText = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowWorkClass iWorkClassId 
'--------------------------------------------------------------------------------------------------
Sub ShowWorkClass( ByVal iWorkClassId )
	Dim sSql, oRs

	sSql = "SELECT workclassid, workclass FROM egov_permitworkclasses "
	sSql = sSql & " WHERE orgid = " & session("orgid") & " ORDER BY workclass" 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""workclassid"">"
		response.write vbcrlf & "<option value=""0"">Select...</option>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("workclassid") & """"
			If CLng(iWorkClassId) = CLng(oRs("workclassid")) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("workclass") & "</option>"
			oRs.MoveNext 
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' void ShowWorkScopes iWorkScopeId 
'--------------------------------------------------------------------------------------------------
Sub ShowWorkScopes( ByVal iWorkScopeId )
	Dim sSql, oRs

	sSql = "SELECT workscopeid, workscope FROM egov_permitworkscope "
	sSql = sSql & " WHERE orgid = " & session("orgid") & " ORDER BY workscope" 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select id=""workscopeid"" name=""workscopeid"">"
	response.write vbcrlf & "<option value=""0"">Select a Work Scope...</option>"

	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("workscopeid") & """"
		If CLng(iWorkScopeId) = CLng(oRs("workscopeid")) Then 
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("workscope") & "</option>"
		oRs.MoveNext 
	Loop 
	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowConstructionTypes iConstructionTypeId 
'--------------------------------------------------------------------------------------------------
Sub ShowConstructionTypes( ByVal iConstructionTypeId )
	Dim sSql, oRs

	sSql = "SELECT constructiontypeid, constructiontype FROM egov_constructiontypes "
	sSql = sSql & " WHERE orgid = " & session("orgid") & " ORDER BY displayorder" 
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select name=""constructiontypeid"">"
	response.write vbcrlf & "<option value=""0"">Select...</option>"

	If Not oRs.EOF Then
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("constructiontypeid") & """"
			If CLng(iConstructionTypeId) = CLng(oRs("constructiontypeid")) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("constructiontype") & "</option>"
			oRs.MoveNext 
		Loop 
	End If 

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowOccupancyTypes iOccupancyTypeId 
'--------------------------------------------------------------------------------------------------
Sub ShowOccupancyTypes( ByVal iOccupancyTypeId )
	Dim sSql, oRs

	sSql = "SELECT occupancytypeid, ISNULL(usegroupcode,'') AS usegroupcode, ISNULL(occupancytype, '') AS occupancytype " 
	sSql = sSql & " FROM egov_occupancytypes WHERE orgid = " & session("orgid") & " ORDER BY usegroupcode, occupancytype" 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select name=""occupancytypeid"">"
	response.write vbcrlf & "<option value=""0"">Select...</option>"

	If Not oRs.EOF Then
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("occupancytypeid") & """"
			If CLng(iOccupancyTypeId) = CLng(oRs("occupancytypeid")) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" 
			If oRs("usegroupcode") <> "" Then 
				response.write oRs("usegroupcode") & " "
			End If 
			response.write oRs("occupancytype") & "</option>"
			oRs.MoveNext 
		Loop 
	End If 

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' void ShowPermitApplicant iPermitId 
'--------------------------------------------------------------------------------------------------
Sub ShowPermitApplicant( ByVal iPermitId )
	Dim sSql, oRs, sToolTip

	sToolTip = ""

	sSql = "SELECT permitcontactid, userid, ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, "
	sSql = sSql & " ISNULL(company,'') AS company, ISNULL(address,'') AS address, ISNULL(city,'') AS city, "
	sSql = sSql & " ISNULL(state,'') AS state, ISNULL(zip,'') AS zip, ISNULL(phone,'') AS phone, contacttype " 
	sSql = sSql & " FROM egov_permitcontacts WHERE isapplicant = 1 AND ispriorcontact = 0 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write "<table cellpadding=""0"" cellspacing=""0"" border=""0""><tr><td valign=""top"">"
		response.write " &nbsp; <span id=""applicantdetails"" "
		If oRs("firstname") <> "" Then 
			sToolTip = "<strong>" & oRs("firstname") & " " & oRs("lastname") & "</strong><br />"
		End If 
		If Not IsNull(oRs("company")) And oRs("company") <> "" Then 
			If sToolTip = "" Then 
				sToolTip = "<strong>" & oRs("company") & "</strong><br />"
			Else 
				sToolTip = sToolTip & oRs("company") & "<br />"
			End If 
		End If 
		If Not IsNull(oRs("address")) And oRs("address") <> "" Then 
			sToolTip = sToolTip &  oRs("address") & "<br />"
		End If 
		If Not IsNull(oRs("city")) And oRs("city") <> "" Then
			sToolTip = sToolTip &  oRs("city") & ", " & oRs("state") & " " & oRs("zip") & "<br />" 
		End If 
		If Not IsNull(oRs("phone")) And oRs("phone") <> "" Then
			sToolTip = sToolTip &  FormatPhoneNumber( oRs("phone") )
		End If 
		sToolTip = Replace(sToolTip,"'","\'")
		'sToolTip = sToolTip &  "</p>"
		response.write " onMouseover=""ddrivetip('" & sToolTip & "', 350)""; onMouseout=""hideddrivetip()""; "
		response.write " >"

		If oRs("company") <> "" Then
			response.write oRs("company")
			If oRs("firstname") <> "" Then 
				response.write " &mdash; "
			End If 
		End If 
		If oRs("firstname") <> "" Then 
			response.write oRs("lastname") & ", " & oRs("firstname")
		End If 

'		If oRs("firstname") <> "" Then
'			response.write oRs("firstname") & " " & oRs("lastname")
'		End If 
'		If oRs("company") <> "" Then
'			'response.write " ( " & oRs("company") & " ) "
'			If oRs("firstname") <> "" Then 
'				response.write " ( " & oRs("company") & " ) "
'			Else
'				response.write oRs("company")
'			End If 
'		End If 

		response.write "&nbsp;<a href=""javascript:EditApplicant('" & oRs("contacttype") & "', '"
		If oRs("contacttype") = "U" Then 
			response.write oRs("userid")
		Else
			response.write oRs("permitcontactid")
		End If 
		response.write "')""><i class=""fa fa-pencil""></i>"
		response.write "</a></span>"

		' If they are a contractor then show the latest license
		If ( CLng(oRs("permitcontactid")) > CLng(0) ) Then
			response.write "</td><td>"
			ShowLatestContactLicense iPermitId, oRs("permitcontactid") 
		End If 
		response.write "</td></tr></table>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowPrimaryContact iPermitId 
'--------------------------------------------------------------------------------------------------
Sub ShowPrimaryContact( ByVal iPermitId )
	Dim sSql, oRs, sToolTip

	sSql = "SELECT permitcontactid, userid, ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, "
	sSql = sSql & " ISNULL(company,'') AS company, ISNULL(address,'') AS address, ISNULL(city,'') AS city, "
	sSql = sSql & " ISNULL(state,'') AS state, ISNULL(zip,'') AS zip, ISNULL(phone,'') AS phone " 
	sSql = sSql & " FROM egov_permitcontacts WHERE isprimarycontact = 1 AND ispriorcontact = 0 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write " &nbsp; <span id=""primarycontactdetails""><a href=""javascript:EditPrimaryContact('" & oRs("userid") & "')"" "
		'title='<p>"
		sToolTip = "<strong>" & oRs("firstname") & " " & oRs("lastname") & "</strong><br />"
		If Not IsNull(oRs("company")) And oRs("company") <> "" Then 
			sToolTip = sToolTip & "" & oRs("company") & "<br />"
		End If 
		If Not IsNull(oRs("address")) And oRs("address") <> "" Then 
			sToolTip = sToolTip & oRs("address") & "<br />"
		End If 
		If Not IsNull(oRs("city")) And oRs("city") <> "" Then
			sToolTip = sToolTip & oRs("city") & ", " & oRs("state") & " " & oRs("zip") & "<br />" 
		End If 
		If Not IsNull(oRs("phone")) And oRs("phone") <> "" Then
			sToolTip = sToolTip & FormatPhoneNumber( oRs("phone") )
		End If 
		'response.write "</p>'"
		sToolTip = Replace(sToolTip,"'","\'")
		response.write " onMouseover=""ddrivetip('" & sToolTip & "', 350)""; onMouseout=""hideddrivetip()""; "
		response.write ">"
		
		If oRs("company") <> "" Then
			response.write oRs("company")
			If oRs("firstname") <> "" Then 
				response.write " &mdash; "
			End If 
		End If 
		If oRs("firstname") <> "" Then 
			response.write oRs("lastname") & ", " & oRs("firstname")
		End If 

'		response.write oRs("firstname") & " " & oRs("lastname")
'		If oRs("company") <> "" Then
'			response.write " ( " & oRs("company") & " ) "
'		End If 

		response.write "</span></a>"
		response.write "<input type=""hidden"" name=""isprimarycontactpermitcontactid"" value=""" & oRs("permitcontactid") & """ />"
		response.write "<input type=""hidden"" name=""isprimarycontactoriginaluserid"" value=""" & oRs("userid") & """ />"
		response.write "<input type=""hidden"" id=""isprimarycontactuserid"" name=""isprimarycontactuserid"" value=""" & oRs("userid") & """ />"
	Else
		response.write " &nbsp; <span id=""primarycontactdetails"">None Selected</span>"
		response.write "<input type=""hidden"" name=""isprimarycontactpermitcontactid"" value=""0"" />"
		response.write "<input type=""hidden"" name=""isprimarycontactoriginaluserid"" value=""0"" />"
		response.write "<input type=""hidden"" id=""isprimarycontactuserid"" name=""isprimarycontactuserid"" value=""0"" />"
	End If 
	tooltipclass=""
	tooltip = ""
	disabled = ""
	If not bCanSaveChanges Then
		tooltipclass="tooltip"
		disabled = " disabled "
		tooltip = "<span class=""tooltiptext"">You don't have permission to save changes.</span>"
	end if
	response.write " &nbsp; <button " & disabled & " type=""button"" class=""button ui-button ui-widget ui-corner-all " & tooltipclass & """ onclick=""SelectPrimaryContact( );"" >Select" & tooltip & "</button>"

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void GetPermitContact iPermitId, sContactType, bShowLicense 
'--------------------------------------------------------------------------------------------------
Sub GetPermitContact( ByVal iPermitId, ByVal sContactType, ByVal bShowLicense )
	Dim sSql, oRs, sToolTip

	sToolTip = ""
	sSql = " SELECT permitcontactid, permitcontacttypeid, ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, "
	sSql = sSql & " ISNULL(company,'') AS company, ISNULL(address,'') AS address, ISNULL(city,'') AS city, "
	sSql = sSql & " ISNULL(state,'') AS state, ISNULL(zip,'') AS zip, ISNULL(phone,'') AS phone " 
	sSql = sSql & " FROM egov_permitcontacts WHERE " & sContactType & " = 1 AND ispriorcontact = 0 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write "<table cellpadding=""0"" cellspacing=""0"" border=""0""><tr><td valign=""top"">"

	If Not oRs.EOF Then
		response.write " &nbsp; <span id=""" & sContactType & "display""><span id=""" & sContactType & "details"" "
		'title=""<p>"
		If oRs("firstname") <> "" Then 
			sToolTip = "<strong>" & oRs("firstname") & " " & oRs("lastname") & "</strong><br />"
		End If 
		If oRs("company") <> "" Then 
			If sToolTip = "" Then 
				sToolTip = "<strong>" & oRs("company") & "</strong><br />" 
			Else 
				sToolTip = sToolTip & oRs("company") & "<br />" 
			End If 
		End If 
		If Trim(oRs("address")) <> "" Then 
			sToolTip = sToolTip & oRs("address") & "<br />" 
		End If 
		If Trim(oRs("city")) <> "" Then 
			sToolTip = sToolTip & oRs("city") & ", " & oRs("state") & " " & oRs("zip") & "<br />"
		End If 
		If Not IsNull(oRs("phone")) And Trim(oRs("phone")) <> "" Then 
			sToolTip = sToolTip & FormatPhoneNumber( oRs("phone") ) 
		End If 
		sToolTip = Replace(sToolTip,"'","\'")
		'response.write "</p>"
		response.write " onMouseover=""ddrivetip('" & sToolTip & "', 350)""; onMouseout=""hideddrivetip()""; "
		response.write ">"   

		If oRs("company") <> "" Then
			response.write oRs("company")
			If oRs("firstname") <> "" Then 
				response.write " &mdash; "
			End If 
		End If 
		If oRs("firstname") <> "" Then 
			response.write oRs("lastname") & ", " & oRs("firstname")
		End If 

'		If oRs("firstname") <> "" Then 
'			response.write oRs("firstname") & " " & oRs("lastname")
'		End If 
'		If oRs("company") <> "" Then
'			If oRs("firstname") <> "" Then 
'				response.write " ( " & oRs("company") & " ) "
'			Else
'				response.write oRs("company")
'			End If 
'		End If 
		response.write " <a href=""javascript:EditContact('" & oRs("permitcontactid") & "', '" & sContactType & "');"">"
		response.write "<i class=""fa fa-pencil""></i></a></span>"
		response.write "</span>"
		response.write "<input type=""hidden"" name=""" & sContactType & "permitcontacttypeid"" id=""" & sContactType & "permitcontacttypeid"" value=""" & oRs("permitcontacttypeid") & """ />"
		response.write "<input type=""hidden"" name=""" & sContactType & "originalpermitcontacttypeid"" value=""" & oRs("permitcontacttypeid") & """ />"
		response.write "<input type=""hidden"" name=""" & sContactType & "permitcontactid"" value=""" & oRs("permitcontactid") & """ />"
		
	Else
		response.write " &nbsp; <span id=""" & sContactType & "display"">None Selected </span>"
		response.write "<input type=""hidden"" name=""" & sContactType & "permitcontacttypeid"" id=""" & sContactType & "permitcontacttypeid"" value=""0"" />"
		response.write "<input type=""hidden"" name=""" & sContactType & "originalpermitcontacttypeid"" value=""0"" />"
		response.write "<input type=""hidden"" name=""" & sContactType & "permitcontactid"" value=""0"" />"
	End If 

	tooltipclass=""
	tooltip = ""
	disabled = ""
	If not bCanSaveChanges Then
		tooltipclass="tooltip"
		disabled = " disabled "
		tooltip = "<span class=""tooltiptext"">You don't have permission to save changes.</span>"
	end if
	response.write " &nbsp; <button " & disabled & " type=""button"" class=""button ui-button ui-widget ui-corner-all " & tooltipclass & """ onclick=""SelectContact( '" & sContactType & "' );"" >Select New Contact" & tooltip & "</button>"

	If bShowLicense Then 
		If Not oRs.EOF Then 
			If ( CLng(oRs("permitcontactid")) > CLng(0) ) Then
				response.write "</td><td>"
				ShowLatestContactLicense iPermitId, oRs("permitcontactid") 
			End If 
		End If 
	End If 
	 
	response.write "</td></tr></table>"

	oRs.Close
	Set oRs = Nothing 

End Sub  


'--------------------------------------------------------------------------------------------------
' integer ShowContractorList( iPermitId )
'--------------------------------------------------------------------------------------------------
Function ShowContractorList( ByVal iPermitId )
	Dim sSql, oRs, iRecCount, sToolTip, sContractorType

	'<input type="checkbox" name="permitcontactid0" value="123" /> <a href="">Link of contractor to edit them</a><br />
	iRecCount = -1
	sSql = " SELECT permitcontactid, permitcontacttypeid, ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, "
	sSql = sSql & " ISNULL(company,'') AS company, ISNULL(address,'') AS address, ISNULL(city,'') AS city, "
	sSql = sSql & " ISNULL(state,'') AS state, ISNULL(zip,'') AS zip, ISNULL(phone,'') AS phone, " 
	sSql = sSql & " ISNULL(company,'') + ISNULL(lastname,'') + ISNULL(firstname,'') AS sortname, ISNULL(contractortypeid,0) AS contractortypeid "
	sSql = sSql & " FROM egov_permitcontacts WHERE iscontractor = 1 AND ispriorcontact = 0 AND permitid = " & iPermitId
	sSql = sSql & " ORDER BY 11"

	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		sToolTip = ""
		iRecCount = iRecCount + 1
'		If CLng(iRecCount) > CLng(1) Then 
'			response.write "<br />"
'		End If 

		' class=""contactpick""
		response.write vbcrlf & "<tr><td valign=""top"" nowrap=""nowrap"">"

		' table for each contractor
		response.write vbcrlf & "<table cellpadding=""0"" cellspacing=""0"" border=""0"" class=""contractorlisting"">"
		response.write vbcrlf & "<tr><td valign=""top"" nowrap=""nowrap"">"

		' Remove checkbox
		response.write vbcrlf & "<input type=""checkbox"" id=""removepermitcontactid" & iRecCount & """ name=""removepermitcontactid" & iRecCount & """ />&nbsp;" '</td>"
		'response.write "<td valign=""middle""><br />"
		response.write vbcrlf & "<input type=""hidden"" name=""contractor" & iRecCount & """ id = ""contractor" & iRecCount & """ value=""" & oRs("permitcontacttypeid") & """ />"
		response.write vbcrlf & "<input type=""hidden"" name=""permitcontactid" & iRecCount & """ id = ""permitcontactid" & iRecCount & """ value=""" & oRs("permitcontactid") & """ />"
		
		' Contractor name with tool tip
		response.write "&nbsp;<span id=""contractor" & oRs("permitcontactid") & """><a href=""javascript:EditContact('" & oRs("permitcontactid") & "', 'iscontractor');"" "
		'title=""<p>"
		If oRs("firstname") <> "" Then 
			sToolTip = "<strong>" & oRs("firstname") & " " & oRs("lastname") & "</strong><br />"
		End If 
		If oRs("company") <> "" Then 
			If sToolTip = "" Then 
				sToolTip = "<strong>" & oRs("company") & "</strong><br />" 
			Else 
				sToolTip = sToolTip & oRs("company") & "<br />" 
			End If  
		End If 
		If Trim(oRs("address")) <> "" Then 
			sToolTip = sToolTip & oRs("address") & "<br />" 
		End If 
		If Trim(oRs("city")) <> "" Then 
			sToolTip = sToolTip & oRs("city") & ", " & oRs("state") & " " & oRs("zip") & "<br />"
		End If 
		If Not IsNull(oRs("phone")) And Trim(oRs("phone")) <> "" Then 
			sToolTip = sToolTip & FormatPhoneNumber( oRs("phone") ) 
		End If
		response.write " onMouseover=""ddrivetip('" & sToolTip & "', 350)""; onMouseout=""hideddrivetip()""; "
		response.write ">"  
		
		If oRs("company") <> "" Then
			response.write oRs("company")
			If oRs("firstname") <> "" Then 
				response.write " &mdash; "
			End If 
		End If 
		If oRs("firstname") <> "" Then 
			response.write oRs("lastname") & ", " & oRs("firstname")
		End If 
		response.write "</a></span>"

		' Contractor type
		sContractorType = GetContractorType( oRs("contractortypeid") ) 
		If sContractorType <> "" Then 
			response.write " &ndash; (<strong>" & sContractorType & "</strong>)"
		End If 
		response.write "</td>"
		
		' Licenses
		response.write "<td nowrap=""nowrap"">"
		' Licenses are only shown if some are required
		ShowLatestContactLicense iPermitId, oRs("permitcontactid") 
		response.write "</td></tr></table>"

		response.write vbcrlf & "</td></tr>"
		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 

	ShowContractorList = iRecCount

End Function 


'--------------------------------------------------------------------------------------------------
' integer ShowPriorContactList( iPermitId )
'--------------------------------------------------------------------------------------------------
Function ShowPriorContactList( ByVal iPermitId )
	Dim sSql, oRs, iRecCount

	iRecCount = 0

	sSql = "SELECT permitcontactid, permitcontacttypeid, ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, "
	sSql = sSql & " ISNULL(company,'') AS company, ISNULL(address,'') AS address, ISNULL(city,'') AS city, "
	sSql = sSql & " ISNULL(state,'') AS state, ISNULL(zip,'') AS zip, ISNULL(phone,'') AS phone, " 
	sSql = sSql & " ISNULL(lastname,'') + ISNULL(firstname,'') + ISNULL(company,'') AS sortname, "
	sSql = sSql & " isbillingcontact, isprimarycontact, isprimarycontractor, isarchitect, iscontractor "
	sSql = sSql & " FROM egov_permitcontacts WHERE ispriorcontact = 1 AND permitid = " & iPermitId
	sSql = sSql & " ORDER BY 11"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		sToolTip = ""
		iRecCount = iRecCount + 1
		response.write vbcrlf & "<tr>"
		response.write "<td>"
		If oRs("company") <> "" Then
			response.write oRs("company")
			If oRs("firstname") <> "" Then 
				response.write " &mdash; "
			End If 
		End If 
		If oRs("firstname") <> "" Then 
			response.write oRs("lastname") & ", " & oRs("firstname")
		End If 
'		If oRs("firstname") <> "" Then 
'			response.write oRs("firstname") & " " & oRs("lastname")
'		End If 
'		If oRs("company") <> "" Then
'			If oRs("firstname") <> "" Then 
'				response.write " (" & oRs("company") & ") "
'			Else
'				response.write oRs("company")
'			End If 
'		End If 
		If oRs("isbillingcontact") Then
			response.write " &ndash; Prior Billing Contact"
		Else 
			If oRs("isprimarycontact") Then
				response.write " &ndash; Prior Primary Contact"
			Else 
				If oRs("isprimarycontractor") Then
					response.write " &ndash; Prior Primary Contractor"
				Else 
					If oRs("isarchitect") Then
						response.write " &ndash; Prior Architect"
					Else 
						If oRs("iscontractor") Then
							response.write " &ndash; Prior Contractor"
						End If
					End If
				End If
			End If
		End If 
		response.write "</td></tr>"
		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 
	
	ShowPriorContactList = iRecCount

End Function 


'--------------------------------------------------------------------------------------------------
' integer ShowFeeList( iPermitId )
'--------------------------------------------------------------------------------------------------
Function ShowFeeList( ByVal iPermitId )
	Dim sSql, oRs, iRecCount, sClass, sClick

	iRecCount = 0

	sSql = " SELECT F.permitfeeid, F.isrequired, F.includefee, f.isfixturetypefee, F.permitfeeprefix, F.permitfee, F.isresidentialunittypefee, "
	sSql = sSql & " F.isvaluationtypefee, F.isconstructiontypefee, F.feeamount, ISNULL(F.paymentid,0) AS paymentid, M.permitfeemethod, "
	sSql = sSql & " M.isflatfee, M.ismanual, M.isfixture, F.isupfrontfee, M.ishourly, M.istotalsqft, M.isfinishedsqft, M.iscuft, "
	sSql = sSql & " M.isunfinishedsqft, M.isothersqft, F.ispercentagetypefee, ISNULL(F.upfrontamount,0.00) AS upfrontamount "
	sSql = sSql & " FROM egov_permitfees F, egov_permitfeemethods M "
	sSql = sSql & " WHERE F.permitfeemethodid = M.permitfeemethodid AND F.permitid = " & iPermitId
	sSql = sSql & " ORDER BY F.displayorder, F.permitfeeid"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		sClick = ""
		If oRs("ismanual") Then
			sClick = " title=""click to edit"" onclick=""EditManualPrice('" & oRs("permitfeeid") & "');"" "
		End If 
		If oRs("isresidentialunittypefee") Then
			sClick = " title=""click to view"" onclick=""ViewResidentialUnitFee('" & oRs("permitfeeid") & "');"" "
		End If
		If oRs("isvaluationtypefee") Then
			sClick = " title=""click to view"" onclick=""ViewValuationFee('" & oRs("permitfeeid") & "');"" "
		End If 
		If oRs("ispercentagetypefee") Then
			sClick = " title=""click to view"" onclick=""ViewPercentageFee('" & oRs("permitfeeid") & "');"" "
		End If 
		If oRs("isconstructiontypefee") Then
			sClick = " title=""click to view"" onclick=""ViewConstructionTypeFee('" & oRs("permitfeeid") & "');"" "
		End If 
		If oRs("isfixturetypefee") Then
			sClick = " title=""click to edit"" onclick=""EditFixtureFee('" & oRs("permitfeeid") & "');"" "
		End If 
		If oRs("ishourly") Then
			sClick = " title=""click to view"" onclick=""EditHourlyRateFee('" & oRs("permitfeeid") & "');"" "
		End If 
		If oRs("istotalsqft") Or oRs("isfinishedsqft") Or oRs("isunfinishedsqft") Or oRs("isothersqft") Then
			sClick = " title=""click to view"" onclick=""ViewSqFootageFee('" & oRs("permitfeeid") & "');"" "
		End If 
		If oRs("iscuft") Then
			sClick = " title=""click to view"" onclick=""ViewCuFtFee('" & oRs("permitfeeid") & "');"" "
		End If 
		iRecCount = iRecCount + 1
		
		response.write vbcrlf & "<tr"   ' id=""" & iRecCount & """" & sClass 
		If iRecCount Mod 2 = 0 Then
			response.write " class=""altrow"" "
		End If 
		response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"

		response.write "<td align=""center"">"  ' Remove Fee Cell
		response.write "<input type=""hidden"" id=""permitfeeid" & iRecCount & """ name=""permitfeeid" & iRecCount & """ value=""" & oRs("permitfeeid") & """ />"
		response.write "<input type=""checkbox"" name=""removefee" & iRecCount & """ id=""removefee" & iRecCount & """"
		If oRs("isrequired") Then
			response.write " disabled=""disabled"" "
		End If 
		response.write " />"
		response.write "</td>"

		response.write "<td nowrap=""nowrap"" align=""center""" & sClick & ">"  ' Category cell
		response.write oRs("permitfeeprefix")
		response.write "</td>"

		response.write "<td" & sClick & ">"  ' Description cell
			response.write oRs("permitfee")
		response.write "</td>"

		response.write "<td align=""center""" & sClick & ">"  ' Method cell
			response.write oRs("permitfeemethod")
		response.write "</td>"

		' Up Front Amount
		If OrgHasFeature( "up front fees" ) Then
			response.write "<td nowrap=""nowrap"" align=""center""" & sClick & ">"  ' Up Front cell
			response.write FormatNumber(oRs("upfrontamount"),2,,,0)
			sUpFrontFeeTotal = FormatNumber((CDbl(sUpFrontFeeTotal) + CDbl(oRs("upfrontamount"))),2,,,0)
			response.write "</td>"
		End If 
		
		response.write "<td align=""right""" & sClick & "><span id=""fee" & oRs("permitfeeid") & """>"  ' Fee amount cell
		response.write FormatNumber(oRs("feeamount"),2,,,0)
		response.write "</span>&nbsp;</td>"
		response.write "</tr>"
		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 

	ShowFeeList = iRecCount

End Function 


'--------------------------------------------------------------------------------------------------
' integer ShowAttachmentList( iPermitId )
'--------------------------------------------------------------------------------------------------
Function ShowAttachmentList( ByVal iPermitId )
	Dim sSql, oRs, iRecCount

	iRecCount = 0

	sSql = "SELECT permitattachmentid, attachmentname, ISNULL(description,'') AS description, attachmentpath, "
	sSql = sSql & " ISNULL(adminuserid,0) AS adminuserid, dateadded, fileextension "
	sSql = sSql & " FROM egov_permitattachments WHERE permitid = " & iPermitId
	sSql = sSql & " ORDER BY 1 DESC"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			iRecCount = iRecCount + 1
			response.write vbcrlf & "<tr"
			If iRecCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write ">"
			'response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"

			' Remove Attachment Cell
			If UserHasPermission( Session("UserId"), "remove permit attachments" ) Then
				response.write "<td align=""center"">"  
				response.write "<input type=""hidden"" id=""permitattachmentid" & iRecCount & """ name=""permitattachmentid" & iRecCount & """ value=""" & oRs("permitattachmentid") & """ />"
				response.write "<input type=""checkbox"" name=""removeattachment" & iRecCount & """ id=""removeattachment" & iRecCount & """ />"
				response.write "</td>"
			End If 

'			response.write "<td align=""center"" title=""Click to View"" onclick=""ViewAttachment(" & oRs("permitattachmentid") & ");"">" & oRs("dateadded") & "</td>"
'			response.write "<td align=""center"" title=""Click to View"" onclick=""ViewAttachment(" & oRs("permitattachmentid") & ");"">" & GetAdminName( oRs("adminuserid") ) & "</td>"
'			response.write "<td align=""center"" title=""Click to View"" onclick=""ViewAttachment(" & oRs("permitattachmentid") & ");"">" & oRs("attachmentname") & "</td>"
'			response.write "<td align=""center"" title=""Click to View"" onclick=""ViewAttachment(" & oRs("permitattachmentid") & ");"">" & oRs("description") & "</td>"

			If oRs("attachmentpath") = "..\permitattachments" Then 
				sLink = "<a class=""permitattachments"" href=""" & oRs("attachmentpath") & "/" & oRs("permitattachmentid") & "." & oRs("fileextension") & """ target=""_blank"">"
			Else
				sLink = "<a class=""permitattachments"" href=""" & oRs("attachmentpath") & "/" & oRs("permitattachmentid") & "_" & Replace(Server.URLEncode(oRs("attachmentname")),"+","%20") & """ target=""_blank"">"
			End If
			response.write "<td align=""left"" title=""Click to View"">" & sLink & oRs("attachmentname") & "</a></td>"
			response.write "<td align=""center"">" & oRs("description") & "</td>"
			response.write "<td align=""center"">" & DateValue(oRs("dateadded")) & "</td>"
			response.write "<td align=""center"">" & GetAdminName( oRs("adminuserid") ) & "</td>"
			
			response.write "</tr>"
			oRs.MoveNext
		Loop
	End If 
	
	oRs.Close
	Set oRs = Nothing 

	ShowAttachmentList = iRecCount

End Function 


'--------------------------------------------------------------------------------------------------
' integer ShowReviewList( iPermitId )
'--------------------------------------------------------------------------------------------------
Function ShowReviewList( ByVal iPermitId )
	Dim sSql, oRs, iRecCount

	iRecCount = 0

	sSql = "SELECT R.permitreviewid, R.permitreviewtype, R.isrequired, R.isincluded, S.reviewstatus, "
	sSql = sSql & " ISNULL(R.revieweruserid,0) AS revieweruserid, R.reviewed "
	sSql = sSql & " FROM egov_permitreviews R, egov_reviewstatuses S "
	sSql = sSql & " WHERE R.reviewstatusid = S.reviewstatusid AND R.permitid = " & iPermitId
	sSql = sSql & " ORDER BY R.revieworder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			iRecCount = iRecCount + 1
			response.write vbcrlf & "<tr"
			If iRecCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"

			response.write "<td align=""center"">"  ' Remove Review Cell
			response.write "<input type=""hidden"" id=""permitreviewid" & iRecCount & """ name=""permitreviewid" & iRecCount & """ value=""" & oRs("permitreviewid") & """ />"
			response.write "<input type=""checkbox"" name=""removereview" & iRecCount & """ id=""removereview" & iRecCount & """"
			If oRs("isrequired") Then
				response.write " disabled=""disabled"" "
			End If 
			response.write " />"
			response.write "</td>"

			' Review type
			response.write "<td title=""Click to View"" onclick=""ViewReview(" & oRs("permitreviewid") & ");"">" & oRs("permitreviewtype") & "</td>"

			' Status
			response.write "<td align=""center"" title=""Click to View"" onclick=""ViewReview(" & oRs("permitreviewid") & ");"">" & oRs("reviewstatus") & "</td>"

			' Reviewed
			response.write "<td align=""center"" title=""Click to View"" onclick=""ViewReview(" & oRs("permitreviewid") & ");"">"
			If IsNull(oRs("reviewed")) Then
				response.write "&nbsp;"
			Else
				response.write FormatDateTime(oRs("reviewed"),2)
			End If 
			response.write "</td>"

			' Reviewer
			response.write "<td align=""center"" title=""Click to View"" onclick=""ViewReview(" & oRs("permitreviewid") & ");"">"
			If CLng(oRs("revieweruserid")) > CLng(0) Then 
				response.write GetPermitReviewerName( CLng(oRs("revieweruserid")) )
			Else
				response.write "Unassigned"
			End If 
			response.write "</td>"

			response.write "</tr>"
			oRs.MoveNext
		Loop
	End If 
	
	oRs.Close
	Set oRs = Nothing 

	ShowReviewList = iRecCount

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowInvoices iPermitId 
'--------------------------------------------------------------------------------------------------
Sub ShowInvoices( ByVal iPermitId )
	Dim sSql, oRs, iRecCount

	sSql = "SELECT I.invoiceid, I.invoicedate, I.totalamount, ISNULL(I.paymentid,0) AS paymentid, I.permitcontactid, "
	sSql = sSql & " S.invoicestatus, I.allfeeswaived, S.isvoid FROM egov_permitinvoices I, egov_invoicestatuses S "
	sSql = sSql & " WHERE I.invoicestatusid = S.invoicestatusid AND I.permitid = " & iPermitId & " ORDER BY invoiceid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			iRecCount = iRecCount + 1
			response.write vbcrlf & "<tr"
			If iRecCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
			
			' Invoice Number
			response.write "<td align=""center"" title=""Click to View Invoice"" onclick=""ViewInvoice(" & oRs("invoiceid") & ");"">" & oRs("invoiceid") & "</td>"
			
			' Invoice Date
			response.write "<td align=""center"" title=""Click to View Invoice"" onclick=""ViewInvoice(" & oRs("invoiceid") & ");"" >"
			response.write "<span id=""invoicedate" & iRecCount & """>" &  DateValue(oRs("invoicedate")) & "</span>"
					href="javascript:EditInvoiceDate( '" & oRs("invoiceid") & "', '" & iRecCount & "' );"
					tooltipclass=""
					tooltip = ""
					CanChangeType = UserHasPermission( session("UserID"), "can change permit types" )
					StatusAllowsDeletes = PermitStatusAllowsDeletes( iPermitId )
					If (not bUserIsRootAdmin and not bUserCanChangeCriticalDates) or bIsCompleted or bIsOnHold or oRs("isvoid") Then 
						tooltipclass="tooltip"
						tooltip = "<span class=""tooltiptext"">The invoice date cannot change because<br />"
						if not bUserCanChangeCriticalDates then tooltip = tooltip & "You don't have permission.<br />"
						if bIsCompleted then tooltip = tooltip & "The permit is complete.<br />"
						if bIsOnHold then tooltip = tooltip & "The permit is on hold.<br />"
						if oRs("isvoid") then tooltip = tooltip & "The invoice is void.<br />"
						tooltip = tooltip & "</span>"
						href="javascript: void(0)"
					End If	%>
					<a class="<%=tooltipclass%>" href="<%=href%>"><i class="fa fa-pencil"></i><%=tooltip%></a>
					<%
			response.write "</td>"
			
			' Billed To
			response.write "<td align=""center"" title=""Click to View Invoice"" onclick=""ViewInvoice(" & oRs("invoiceid") & ");"">"
			response.write GetInvoiceContact( oRs("permitcontactid") )
			response.write "</td>"

			' Invoid Status
			response.write "<td align=""center"" title=""Click to View Invoice"" onclick=""ViewInvoice(" & oRs("invoiceid") & ");"">" & oRs("invoicestatus") & "</td>"
			' Inoice Total
			response.write "<td align=""center"" title=""Click to View Invoice"" onclick=""ViewInvoice(" & oRs("invoiceid") & ");"">" & FormatNumber(oRs("totalamount"),2,,,0) & "</td>"

			' Paid Date
			response.write "<td align=""center"" "
			If CLng(oRs("paymentid")) > CLng(0) And Not oRs("isvoid") Then 
				response.write "title=""Click to View Invoice"" onclick=""ViewInvoice(" & oRs("invoiceid") & ");"""
				response.write ">"

				' Get the paid date of the payment id
				response.write "<span id=""paymentdate" & iRecCount & """>" &  GetInvoicePaymentDate( oRs("paymentid") ) & "</span>"
					href="javascript:EditPaidDate( '" & oRs("paymentid") & "', '" & iRecCount & "', '" & oRs("invoiceid") & "' );"
					tooltipclass=""
					tooltip = ""
					CanChangeType = UserHasPermission( session("UserID"), "can change permit types" )
					StatusAllowsDeletes = PermitStatusAllowsDeletes( iPermitId )
					If (not bUserIsRootAdmin and not bUserCanChangeCriticalDates) or bIsCompleted or bIsOnHold Then 
						tooltipclass="tooltip"
						tooltip = "<span class=""tooltiptext"">The invoice date cannot change because<br />"
						if not bUserCanChangeCriticalDates then tooltip = tooltip & "You don't have permission.<br />"
						if bIsCompleted then tooltip = tooltip & "The permit is complete.<br />"
						if bIsOnHold then tooltip = tooltip & "The permit is on hold.<br />"
						tooltip = tooltip & "</span>"
						href="javascript: void(0)"
					End If	%>
					<a class="<%=tooltipclass%>" href="<%=href%>"><i class="fa fa-pencil"></i><%=tooltip%></a>
					<%
			Else
				response.write "title=""Click to View Invoice"" onclick=""ViewInvoice(" & oRs("invoiceid") & ");"">"
				response.write "&nbsp;"
			End If 
			response.write "</td>"
			
			' Amount paid or waived
			response.write "<td align=""center"" title=""Click to View Invoice"" onclick=""ViewInvoice(" & oRs("invoiceid") & ");"">"
			If oRs("allfeeswaived") Then 
				response.write FormatNumber(oRs("totalamount"),2,,,0)
			Else
				If Not oRs("isvoid") Then
					response.write GetInvoicePaymentTotal( CLng(oRs("invoiceid")) )   ' in permitcommonfunctions.asp
				Else
					response.write "&nbsp;"
				End If 
			End If 
			response.write "</td>"
			
			' Void Button
			response.write "<td align=""center"""
			If Not oRs("isvoid") Then
				response.write ">"
				If Not bIsCompleted And Not bIsOnHold Then 
					' can only void an invoice until the permit is completed
				End If 
				tooltipclass=""
				tooltip = ""
				disabled = ""
				If bIsCompleted or bIsOnHold Then 
					' can only void an invoice until the permit is completed
					tooltipclass="tooltip"
					tooltip = "<span class=""tooltiptext"">The invoice date cannot change because<br />"
					tooltip = tooltip & "The permit is complete or on hold.<br /></span>"
					disabled = " disabled "
				End If
				response.write "<button " & disabled & " type=""button"" class=""button ui-button ui-widget ui-corner-all " & tooltipclass & """ onclick=""VoidInvoice(" & oRs("invoiceid") & ");"">Void" & tooltip & "</button>"
			Else	
				response.write " title=""Click to View Invoice"" onclick=""ViewInvoice(" & oRs("invoiceid") & ");"">&nbsp;"
			End If 
			response.write "</td>"

			response.write "</tr>"
			oRs.MoveNext
		Loop
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowInvoiceContacts iPermitId 
'--------------------------------------------------------------------------------------------------
Sub ShowInvoiceContacts( ByVal iPermitId )
	Dim sSql, oRs, bName

	sSql = "SELECT DISTINCT C.permitcontactid, ISNULL(C.company,'') AS company, ISNULL(C.firstname,'') AS firstname, ISNULL(C.lastname,'') AS lastname, "
	sSql = sSql & " ISNULL(C.company,'') AS company, ISNULL(C.lastname,'') + ISNULL(C.firstname,'') + ISNULL(C.company,'') AS sortname "
	sSql = sSql & " FROM egov_permitcontacts C, egov_permitinvoices I WHERE I.permitcontactid = C.permitcontactid AND I.permitid = " & iPermitId
	sSql = sSql & " ORDER BY 5"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select id=""viewinvoicecontactid"" name=""viewinvoicecontactid"">"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("permitcontactid") & """>" 
			If oRs("firstname") <> "" Then
				response.write oRs("firstname") & " " & oRs("lastname")
				bName = True 
			Else
				bName = False 
			End If 
			If oRs("company") <> "" Then
				If bName Then 
					response.write " ("
				End If 
				response.write oRs("company")
				If bName Then
					response.write ")"
				End If 
			End If 
 			response.write "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowPermitNotes iPermitId 
'--------------------------------------------------------------------------------------------------
Sub ShowPermitNotes( ByVal iPermitId )
	Dim sSql, oRs, iRowCount

	iRowCount = 0

	sSql = "SELECT entrydate, ISNULL(internalcomment,'') AS internalcomment, ISNULL(externalcomment,'') AS externalcomment, S.permitstatus, ISNULL(L.adminuserid,0) AS adminuserid, "
	sSql = sSql & " ISNULL(activitycomment,'') AS activitycomment "
	sSql = sSql & " FROM egov_permitlog L, egov_permitstatuses S "
	sSql = sSql & " WHERE isactivityentry = 1 AND S.permitstatusid = L.permitstatusid AND permitid = " & iPermitId
	sSql = sSql & " ORDER BY permitlogid DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<table id=""permitpriornotes"" cellpadding=""3"" cellspacing=""0"" border=""0"">"
		Do While Not oRs.EOF 
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr"
			If iRowCount Mod 2 = 1 Then
				response.write " class=""altrow"" "
			End If 
			response.write "><td>"
			If CLng(oRs("adminuserid")) > CLng(0) then
				response.write GetAdminName( CLng(oRs("adminuserid")) ) ' In common.asp
			Else
				response.write "System Generated"
			End If 
			response.write " &ndash; " & oRs("permitstatus") & " &ndash; " & oRs("entrydate") & "<br />"
			If oRs("activitycomment") <> "" Then 
				response.write oRs("activitycomment") & "<br />"
			End If 
			If oRs("internalcomment") <> "" Then 
				response.write " &nbsp; <strong>Internal Note:</strong> " & oRs("internalcomment") & "<br />"
			End If 
			If oRs("externalcomment") <> "" Then 
				response.write " &nbsp; <strong>Public Note:</strong> " & oRs("externalcomment")
			End If 
			response.write "</td></tr>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</table>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' integer ShowInspectionList( iPermitId )
'--------------------------------------------------------------------------------------------------
Function ShowInspectionList( ByVal iPermitId )
	Dim sSql, oRs, iRecCount, bAllowInspections, sAttributes

	iRecCount = 0

	'bAllowInspections = PermitStatusAllowsInspections( iPermitId )   ' in permitcommonfunctions.asp

	sSql = "SELECT I.permitinspectionid, I.permitinspectiontype, I.inspectiondescription, I.isrequired, S.inspectionstatus, "
	sSql = sSql & " I.inspecteddate, I.scheduleddate, I.isreinspection, ISNULL(I.inspectoruserid,0) AS inspectoruserid, isfinal "
	sSql = sSql & " FROM egov_permitinspections I, egov_inspectionstatuses S "
	sSql = sSql & " WHERE I.inspectionstatusid = S.inspectionstatusid AND I.permitid = " & iPermitId
	sSql = sSql & " ORDER BY I.inspectionorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			iRecCount = iRecCount + 1
			response.write vbcrlf & "<tr id=""inspectionrow" & oRs("permitinspectionid") & """ "
			If iRecCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
			
			' Remove Inspection Cell
			response.write "<td align=""center"">"  
			response.write "<input type=""hidden"" id=""permitinspectionid" & iRecCount & """ name=""permitinspectionid" & iRecCount & """ value=""" & oRs("permitinspectionid") & """ />"
			response.write "<input type=""checkbox"" name=""removeinspection" & iRecCount & """ id=""removeinspection" & iRecCount & """"
			If oRs("isrequired") Then
				response.write " disabled=""disabled"" "
			End If 
			response.write " />"
			response.write "</td>"
			
			sAttributes = " onclick=""ViewInspection(" & oRs("permitinspectionid") & ");"""
			if (not oRs("isfinal") and GetPermitIsIssued( iPermitId )) or (oRs("isfinal") and AllOtherInspectionsAreDone( iPermitId, oRs("Permitinspectionid") ) And GetPermitIsIssued( iPermitId ) And PermitFeesArePaid( iPermitId )) Then
				sAttributes = sAttributes & " title=""Click to View"""
			else
				sAttributes = sAttributes & " title="""
				if oRs("isfinal") and not AllOtherInspectionsAreDone( iPermitId, oRs("permitinspectionid") ) then
					sAttributes = sAttributes & " All other inspections are not done."
				end if
				if not GetPermitIsIssued( iPermitId ) then
					sAttributes = sAttributes & " Permit is not issued."
				end if
				if oRs("isfinal") and not PermitFeesArePaid( iPermitId ) then
					sAttributes = sAttributes & " Permit fees are not paid."
				end if
				sAttributes = sAttributes & """"
			end if

			' Inspection
			response.write "<td" & sAttributes & ">" & oRs("permitinspectiontype") & " &mdash; " & oRs("inspectiondescription") & "</td>"

			' Reinspection
			response.write "<td align=""center""" & sAttributes & ">"
			If oRs("isreinspection") Then
				response.write "Reinspection"
			Else
				response.write "&nbsp;"
			End If 

			response.write "</td>"

			' Status
			response.write "<td id=""InStatus" & oRs("permitinspectionid") & """ align=""center""" & sAttributes & ">" & oRs("inspectionstatus") & "</td>"

			' Scheduled Date
			response.write "<td id=""InSchedDate" & oRs("permitinspectionid") & """ align=""center""" & sAttributes & ">" 
			If IsNull(oRs("scheduleddate")) Then
				response.write "&nbsp;"
			Else 
				response.write FormatDateTime(oRs("scheduleddate"),2) 
			End If 
			response.write "</td>"

			' Inspected Date
			response.write "<td id=""InInspectDate" & oRs("permitinspectionid") & """ align=""center""" & sAttributes & ">" 
			If IsNull(oRs("inspecteddate")) Then
				response.write "&nbsp;"
			Else 
				response.write FormatDateTime(oRs("inspecteddate"),2) 
			End If 
			response.write "</td>"

			' Inspector
			response.write "<td id=""InInspector" & oRs("permitinspectionid") & """ align=""center""" & sAttributes & ">"

			If CLng(oRs("inspectoruserid")) > CLng(0) Then 
				response.write GetAdminName( CLng(oRs("inspectoruserid")) )
			Else
				response.write "Unassigned"
			End If 

			response.write "</td>"

			response.write "</tr>"
			oRs.MoveNext
		Loop
	End If 
	
	oRs.Close
	Set oRs = Nothing 

	ShowInspectionList = iRecCount

End Function 

'--------------------------------------------------------------------------------------------------
' integer ShowInspectionReportList( iPermitId )
'--------------------------------------------------------------------------------------------------
Function ShowInspectionReportList( ByVal iPermitId )
	Dim sSql, oRs, iRecCount, bAllowInspections, sAttributes

	iRecCount = 0


	sSql = "SELECT permitinspectionreportid,datecreated "
	sSql = sSql & " FROM egov_permitinspectionreports  "
	sSql = sSql & " WHERE permitid = " & iPermitId
	sSql = sSql & " ORDER BY DateCreated"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			iRecCount = iRecCount + 1
			response.write vbcrlf & "<tr id=""IRrow" & oRs("permitinspectionreportid") & """ "
			If iRecCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
			
			' Remove Inspection Cell
			response.write "<td align=""center"">"  
			response.write "<input type=""hidden"" id=""permitinspectionreportid" & iRecCount & """ name=""permitinspectionreportid" & iRecCount & """ value=""" & oRs("permitinspectionreportid") & """ />"
			response.write "<input type=""checkbox"" name=""removeIR" & iRecCount & """ id=""removeIR" & iRecCount & """/>"
			response.write "</td>"
			
			sAttributes = " onclick=""ViewIR(" & oRs("permitinspectionreportid") & ");"""
			' Inspection
			response.write "<td" & sAttributes & ">" & oRs("datecreated") & "</td>"
			if bCanSaveChanges then
				response.write "<td><a href=""javascript:CopyIR(" & oRs("PermitInspectionReportID") & ")"">Create Copy</a></td>"
			else
				response.write "<td></td>"
			end if
			response.write "<td><a href=""javascript:PrintIR(" & oRs("permitinspectionreportid") & ")"">Print</a></td>"
			response.write "<td><a href=""javascript:EmailIR(" & oRs("permitinspectionreportid") & ")"">Email</a></td>"


			response.write "</tr>"
			oRs.MoveNext
		Loop
	End If 
	
	oRs.Close
	Set oRs = Nothing 

	ShowInspectionReportList = iRecCount

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowPermitPlansBy iPermitId, iPlansByContactId 
'--------------------------------------------------------------------------------------------------
Sub ShowPermitPlansBy( ByVal iPermitId, ByVal iPlansByContactId )
	Dim sSql, oRs

	sSql = "SELECT permitcontactid, ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, "
	sSql = sSql & " ISNULL(company,'') AS company, ISNULL(company,'') + ISNULL(lastname,'') + ISNULL(firstname,'') AS sortname, "
	sSql = sSql & " isapplicant, isbillingcontact, isprimarycontact, isprimarycontractor, isarchitect, iscontractor, ISNULL(contractortypeid,0) AS contractortypeid "
	sSql = sSql & " FROM egov_permitcontacts WHERE ispriorcontact = 0 AND permitid = " & iPermitId
	sSql = sSql & " ORDER BY 5"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select name=""plansbycontactid"">"
	response.write vbcrlf & "<option value=""0"">Select a contact...</option>"
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("permitcontactid") & """"
		If CLng(iPlansByContactId) = CLng(oRs("permitcontactid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">"

		If oRs("company") <> "" Then
			response.write oRs("company")
			If oRs("firstname") <> "" Then 
				response.write " &mdash; "
			End If 
		End If 
		If oRs("firstname") <> "" Then 
			response.write oRs("lastname") & ", " & oRs("firstname")
		End If 

'		If oRs("firstname") <> "" Then
'			response.write oRs("firstname") & " " & oRs("lastname")
'			bName = True 
'		Else
'			bName = False 
'		End If 
'		If oRs("company") <> "" Then
'			If bName Then 
'				response.write " ("
'			End If 
'			response.write oRs("company")
'			If bName Then
'				response.write ")"
'			End If 
'		End If 
		If oRs("isapplicant") Then
			response.write " &ndash; Applicant"
		Else 
			If oRs("isbillingcontact") Then
				response.write " &ndash; Billing Contact"
			Else 
				If oRs("isprimarycontact") Then
					response.write " &ndash; Primary Contact"
				Else 
					If oRs("isprimarycontractor") Then
						response.write " &ndash; Primary Contractor"
					Else 
						If oRs("isarchitect") Then
							response.write " &ndash; Architect/Engineer"
						Else 
							If oRs("iscontractor") Then
								sContractorType = GetContractorType( oRs("contractortypeid") )
								If sContractorType <> "" Then 
									response.write " &ndash; " & sContractorType
								Else 
									response.write " &ndash; Contractor"
								End If 
							End If
						End If
					End If
				End If
			End If 
		End If 
		response.write "</option>"
		oRs.MoveNext 
	Loop
	response.write vbcrlf & "</select>"
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowRequiredPermitLicenses iPermitId 
'--------------------------------------------------------------------------------------------------
Sub ShowRequiredPermitLicenses( ByVal iPermitId, ByVal sIsRequired )
	Dim sSql, oRs, iCount

	iCount = 0
	sSql = "SELECT licensetype FROM egov_permits_to_permitlicensetypes WHERE permitid = " & iPermitId
	sSql = sSql & " AND isrequired = " & sIsRequired & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		If iCount > 0 Then
			response.write ","
		End If 
		iCount = iCount + 1
		response.write "&nbsp;&nbsp;" & oRs("licensetype")
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' boolean PermitJobValueNotEqualToInvices( iPermitId, sJobValue )
'--------------------------------------------------------------------------------------------------
Function PermitJobValueNotEqualToInvices( ByVal iPermitId, ByVal sJobValue )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(SUM(netjobvalue),0.00) AS netjobvalue FROM egov_permitinvoices "
	sSql = sSql & " WHERE isvoided = 0 AND permitid = " & iPermitId 
	response.write "<!--" & sSql & "-->"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CDbl(sJobValue) <> CDbl(oRs("netjobvalue")) Then
			PermitJobValueNotEqualToInvices = True 
		Else
			PermitJobValueNotEqualToInvices = False 
		End If 
	Else 
		' No invoices so we do not want to try to update them
		PermitJobValueNotEqualToInvices = False  
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void ListPermitDocumentButtons iPermitId 
'--------------------------------------------------------------------------------------------------
Sub ListPermitDocumentButtons( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT permitdocumentid, documentlabel "
	sSql = sSql & " FROM egov_permittypes_to_permitdocuments D, egov_permits P "
	sSql = sSql & " WHERE D.permittypeid = P.permittypeid AND P.permitid = " & iPermitId
	sSql = sSql & " ORDER BY documentlabel"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select name=""permitdocument"" id=""permitdocument"">"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("permitdocumentid") & """ >" & oRs("documentlabel") & "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select> &nbsp; &nbsp; "
		If OrgHasFeature("xml pdf display") Then
			' Newer Acrobat 9 PDF
			response.write "<input type=""button"" class=""button ui-button ui-widget ui-corner-all"" value=""Print Selected Document"" onclick=""ViewPermitXML();"" />"
		Else 
			' Old Acrobat 6 PDF
			response.write "<input type=""button"" class=""button ui-button ui-widget ui-corner-all"" value=""Print Selected Document"" onclick=""ViewPermitDocument();"" />"
		End If 

		
	Else
		response.write vbcrlf & "There are no documents associated with this type of permit."
	End If 
	
	oRs.CLose
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' integer ShowCustomPermitFields( iPermitid )
'--------------------------------------------------------------------------------------------------
Function ShowCustomPermitFields( ByVal iPermitid )
	Dim sSql, oRs, iCount, sDateValue, sMoneyValue, sIntValue, aChoices, x, sChoice, bCheckFirstRadio
	Dim sSelectedValue, aPicks, bHasChecks

	iCount = clng(0)

	sSql = "SELECT P.customfieldid, F.fieldtypebehavior, P.prompt, P.valuelist, P.fieldsize, "
	sSql = sSql & "ISNULL(P.simpletextvalue,'') AS simpletextvalue, ISNULL(P.largetextvalue,'') AS largetextvalue, "
	sSql = sSql & "P.datevalue, moneyvalue, P.intvalue "
	sSql = sSql & "FROM egov_permitcustomfields P, egov_permitfieldtypes F "
	sSql = sSql & "WHERE P.fieldtypeid = F.fieldtypeid AND P.permitid = " & iPermitid
	sSql = sSql & "ORDER BY P.displayorder"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<tr><td class=""label"" nowrap=""nowrap"" valign=""top"" colspan=""4"">&nbsp;</td></tr>"
		Do While Not oRs.EOF
			iCount = iCount + 1
			response.write vbcrlf & "<tr><td class=""label"" nowrap=""nowrap"" valign=""top"">" & oRs("prompt") & "</td><td colspan=""3"">"
			response.write "<input type=""hidden"" id=""customfieldid" & iCount & """ name=""customfieldid" & iCount & """ value=""" & oRs("customfieldid") & """ />"
			response.write "<input type=""hidden"" id=""fieldtypebehavior" & iCount & """ name=""fieldtypebehavior" & iCount & """ value=""" & oRs("fieldtypebehavior") & """ />"
			
			Select Case oRs("fieldtypebehavior")
				
				Case "radio"
					sSelectedValue = oRs("simpletextvalue")
					If sSelectedValue = "" Then
						bCheckFirstRadio = True 
					Else 
						bCheckFirstRadio = False 
					End If 
					'response.write bCheckFirstRadio & "<br />"
					aChoices = Split(oRs("valuelist"),Chr(10))
			
					for x = 0 to UBound(aChoices)
						sChoice = Replace(aChoices(x), Chr(13), "")
						response.write vbcrlf & "<input type=""radio"" id=""customfield" & iCount & """ name=""customfield" & iCount & """ value=""" & sChoice & """"
						If (sSelectedValue = sChoice) Or bCheckFirstRadio Then
							response.write " checked=""checked"""
							bCheckFirstRadio = False
						End If 
						response.write " /> " & sChoice & "<br />"
					Next
					
				Case "select"
					sSelectedValue = oRs("simpletextvalue")
					response.write vbcrlf & "<select id=""customfield" & iCount & """ name=""customfield" & iCount & """>"
					aChoices = Split(oRs("valuelist"),Chr(10))
			
					for x = 0 to UBound(aChoices)
						sChoice = Replace(aChoices(x), Chr(13), "")
						response.write vbcrlf & "<option value=""" & sChoice & """"
						If sSelectedValue = sChoice Then
							response.write " selected=""selected"""
						End If 
						response.write " /> " & sChoice & "</option>"
					Next
					response.write vbcrlf & "</select>"
				
				Case "checkbox"
					sSelectedValue = Replace(oRs("simpletextvalue"), Chr(13), "")
					If oRs("simpletextvalue") <> "" Then 
						aPicks = Split(sSelectedValue,Chr(10))
						bHasChecks = True 
					Else
						bHasChecks = False 
					End If 
					aChoices = Split(oRs("valuelist"),Chr(10))
			
					For x = 0 To UBound(aChoices)
						sChoice = Replace(aChoices(x), Chr(13), "")
						response.write vbcrlf & "<input type=""checkbox"" id=""customfield" & iCount & """ name=""customfield" & iCount & """ value=""" & sChoice & """"
						If bHasChecks Then 
							For y = 0 To UBound(aPicks)
								If aPicks(y) = sChoice Then
									response.write " checked=""checked"""
									Exit For 
								End If 
							Next 
						End If 
						response.write " /> " & sChoice & "<br />"
					Next

				Case "textbox"
					response.write "<input type=""text"" id=""customfield" & iCount & """ name=""customfield" & iCount & """ value=""" & Replace(oRs("simpletextvalue"), Chr(34), "&quot;" ) & """ size=""" & oRs("fieldsize") & """ maxlength=""" & oRs("fieldsize") & """ />"

				Case "textarea"
					response.write "<textarea class=""customfields"" id=""customfield" & iCount & """ name=""customfield" & iCount & """ maxlength=""" & oRs("fieldsize") & """ rows=""100"" cols=""100"">" & oRs("largetextvalue") & "</textarea>"

				Case "date"
					If IsNull(oRs("datevalue")) Then
						sDateValue = ""
					Else 
						sDateValue = DateValue(oRs("datevalue"))
					End If 
					response.write "<input type=""text"" class=""datepicker"" id=""customfield" & iCount & """ name=""customfield" & iCount & """ value=""" & sDateValue & """  size=""10"" maxlength=""10"" />"
					' put a date picker
					'response.write "&nbsp;<a href=""javascript:void doCalendar('customfield" & iCount & "');""><img src=""../images/calendar.gif"" border=""0"" /></a>"
				
				Case "money"
					If IsNull(oRs("moneyvalue")) Then
						sMoneyValue = ""
					Else 
						sMoneyValue = FormatNumber(oRs("moneyvalue"),2,,,0)
					End If 
					response.write "<input type=""text"" id=""customfield" & iCount & """ name=""customfield" & iCount & """ value=""" & sMoneyValue & """ size=""" & oRs("fieldsize") & """ maxlength=""" & oRs("fieldsize") & """ />"
				
				Case "integer"
					If IsNull(oRs("intvalue")) Then
						sIntValue = ""
					Else 
						sIntValue = oRs("intvalue")
					End If 
					response.write "<input type=""text"" id=""customfield" & iCount & """ name=""customfield" & iCount & """ value=""" & sIntValue & """ size=""" & oRs("fieldsize") & """ maxlength=""" & oRs("fieldsize") & """ />"
			
			End Select 
			response.write "</tr>"
			response.write vbcrlf & "<tr><td class=""label"" nowrap=""nowrap"" valign=""top"">&nbsp;</td><td colspan=""3"">"

			response.write vbcrlf & "</td></tr>"

			oRs.MoveNext
		Loop 
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowCustomPermitFields = iCount

End Function 



%>
