<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../classes/class_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: poolpass_type_report.asp
' AUTHOR: Steve Loar
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the report of passes by type
'
' MODIFICATION HISTORY
' 1.0 05/09/06	Steve Loar - INITIAL VERSION
' 1.1	10/05/06	Steve Loar - Header and nav changed
' 1.2 04/02/09 David Boyer - Added Period Description if org has "Membership Alternate Layout" feature on
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 dim iYear, iAdminCnt, iOnlineCnt, iMembershipId, iPeriodId
 dim sShowMembershipPicks

 sLevel = "../"  'Override of value from common.asp

 if not userhaspermission(session("userid"),"membership list") then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

 if request("iyear") <> "" then
	   iYear = request("iyear")
 else
	   iYear = Year(Now())
 end if

 iAdminCnt  = 0
 iOnlineCnt = 0

 if request("membershipid") = "" then
   	iMembershipId = GetFirstMembershipId()
 else
 	  iMembershipId = CLng(request("membershipid"))
 end if

 if request("periodid") = "" then
   	iPeriodId = CLng(0)
 else
 	  iPeriodId = CLng(request("periodid"))
 end if

'Get all of the membership picks
 sShowMembershipPicks = ShowMembershipPicks(session("orgid"), _
                                            iMembershipId)

 sShowPeriodPicks = ShowPeriodPicks(session("orgid"), _
                                    iPeriodId)

 sShowYearChoices = ShowYearChoices(session("orgid"), _
                                    iYear)
%>
<html>
<head>
	 <title>E-Gov Administration Console {Membership Types Report}</title>

	 <link rel="stylesheet" type="text/css" href="../global.css" />
	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" type="text/css" href="../classes/classes.css" />
	 <link rel="stylesheet" type="text/css" href="poolpass.css" />
	 <link rel="stylesheet" type="text/css" media="print" href="receiptprint.css" />

<style type="text/css">
  .fieldset_poolpass {
     border: 1pt solid #808080;
       border-radius:         5px;
       -webkit-border-radius: 5px;
       -moz-border-radius:    5px;
  }
</style>

<script language="javascript">
<!--
	window.onload = function()
	{
	  //factory.printing.header = "Printed on &d"
	  factory.printing.footer       = "&bPrinted on &d - Page:&p/&P";
	  factory.printing.portrait     = true;
	  factory.printing.leftMargin   = 0.5;
	  factory.printing.topMargin    = 0.5;
	  factory.printing.rightMargin  = 0.5;
	  factory.printing.bottomMargin = 0.5;
	 
	  // enable control buttons
	  var templateSupported = factory.printing.IsTemplateSupported();
	  var controls = idControls.all.tags("input");
	  for ( i = 0; i < controls.length; i++ ) 
	  {
		controls[i].disabled = false;
		if ( templateSupported && controls[i].className == "ie55" )
		  controls[i].style.display = "inline";
	  }
	}

	function YearPick()
	{
		//alert('year = ' + document.YearForm.iyear.value);
		document.YearForm.submit();
	}

//-->
</script>

</head>
<body>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<div id=""idControls"" class=""noprint"">" & vbcrlf
  response.write "	 <input type=""button"" disabled=""disabled"" value=""Print the page"" onclick=""factory.printing.Print(true)"" />&nbsp;&nbsp;" & vbcrlf
  response.write "	 <input class=""ie55"" disabled=""disabled"" type=""button"" value=""Print Preview..."" onclick=""factory.printing.Preview()"" />" & vbcrlf
  response.write "</div>" & vbcrlf

  response.write "<object id=""factory"" viewastext  style=""display:none""" & vbcrlf
  response.write "  classid=""clsid:1663ed61-23eb-11d2-b92f-008048fdd814""" & vbcrlf
  response.write "  codebase=""../includes/smsx.cab#Version=6,3,434,12"">" & vbcrlf
  response.write "</object>" & vbcrlf

  response.write "<div id=""content"">" & vbcrlf
  response.write "	<div id=""centercontent"">" & vbcrlf

  response.write "		<h3>" & session("sOrgName") & " Membership Counts</h3><br /><br />" & vbcrlf
  response.write "  <form name=""YearForm"" method=""post"" action=""poolpass_type_report.asp"">" & vbcrlf
  response.write "  			<fieldset id=""search"" class=""fieldset_poolpass"">" & vbcrlf
  response.write "  				<legend><strong>View Member Counts For</strong></legend><br />" & vbcrlf
  response.write "  				<table cellspacing=""0"" cellpadding=""3"" border=""0"" style=""width:260px;"">" & vbcrlf
  response.write "  					<tr>" & vbcrlf
  response.write "           <td><strong>Membership Type:</strong></td>" & vbcrlf
  response.write "           <td>" & sShowMembershipPicks & "</td>" & vbcrlf
  response.write "       </tr>" & vbcrlf
  response.write "  					<tr>" & vbcrlf
  response.write "           <td><strong>Membership Period:</strong></td>" & vbcrlf
  response.write "           <td>" & sShowPeriodPicks & "</td>" & vbcrlf
  response.write "       </tr>" & vbcrlf
  response.write "  					<tr>" & vbcrlf
  response.write "           <td><strong>Purchase Year:</strong></td>" & vbcrlf
  response.write "           <td>" & sShowYearChoices & "</td>" & vbcrlf
  response.write "       </tr>" & vbcrlf
  response.write "  				</table>" & vbcrlf
  response.write "  			</fieldset>" & vbcrlf
  response.write "  		</form>" & vbcrlf
  response.write "  		<p>" & vbcrlf

  GetPoolPassCountByYear iYear, _
                         iMembershipId, _
                         iPeriodId

  response.write "  		</p>" & vbcrlf
  response.write "  	</div>" & vbcrlf
  response.write "  </div>" & vbcrlf
%>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>
<%
'------------------------------------------------------------------------------
function ShowMembershipPicks(iOrgID, _
                             iMembershipId)
	 dim sSql, oMember, lcl_return, lcl_selected

  lcl_return = ""

	 sSQL= "SELECT membershipid, "
  sSQL = sSQL & " membershipdesc "
  sSQL = sSQL & " FROM egov_memberships "
  sSQL = sSQL & " WHERE orgid = " & iOrgID
	
 	set oMember = Server.CreateObject("ADODB.Recordset")
	 oMember.Open sSQL, Application("DSN"), 3, 1
	
	 if not oMember.eof then
   		lcl_return = "<select name=""membershipid"" onchange=""javascript:YearPick();"">" & vbcrlf
	 	  lcl_return = lcl_return & "<option value=""0"">All</option>" & vbcrlf

   		do while not oMember.eof
        lcl_selected = ""

     			if CLng(oMember("membershipid")) = CLng(iMembershipId) then
       				lcl_selected = " selected=""selected"""
    		 	end if

		     	lcl_return = lcl_return & "<option value=""" & oMember("membershipid") & """" & lcl_selected & ">" & oMember("membershipdesc") & "</option>" & vbcrlf

     			oMember.MoveNext
		   loop

  		 lcl_return = lcl_return & "</select>" & vbcrlf

  end if
	
 	oMember.close
 	set oMember = Nothing

  ShowMembershipPicks = lcl_return

end function

'------------------------------------------------------------------------------
function ShowPeriodPicks(iOrgID, _
                         iPeriodId)

	 dim sSql, oPeriod, lcl_return, lcl_selected

	 sSQL = "SELECT periodid, "
  sSQL = sSQL & " period_desc "
  sSQL = sSQL & " FROM egov_membership_periods "
  sSQL = sSQL & " WHERE orgid = " & iOrgID
	
	 set oPeriod = Server.CreateObject("ADODB.Recordset")
	 oPeriod.Open sSQL, Application("DSN"), 3, 1
	
	 if not oPeriod.eof then
		   lcl_return = "<select name=""periodid"" onchange=""javascript:YearPick();"">" & vbcrlf
   		lcl_return = lcl_return & "<option value=""0"">All</option>" & vbcrlf

   		do while not oPeriod.eof
        lcl_selected = ""

     			if CLng(oPeriod("periodid")) = CLng(iPeriodId) then
       				lcl_selected = " selected=""selected"""
     			end if

			     lcl_return = lcl_return & "<option value=""" & oPeriod("periodid") & """" & lcl_selected & ">" & oPeriod("period_desc") & "</option>" & vbcrlf

     			oPeriod.MoveNext
   		loop

   		lcl_return = lcl_return & "</select>" & vbcrlf
  end if
	
 	oPeriod.close
	 set oPeriod = nothing

  ShowPeriodPicks = lcl_return

end function

'------------------------------------------------------------------------------
function ShowYearChoices(iOrgID, _
                         iDefaultYear)
	 dim sSql, oYears, lcl_return, lcl_selected

	 sSQL = "SELECT distinct year(paymentdate) as year "
  sSQL = sSQL & " FROM egov_poolpasspurchases "
  sSQL = sSQL & " WHERE orgid = " & iOrgID
  sSQL = sSQL & " ORDER BY 1 desc"

	 set oYears = Server.CreateObject("ADODB.Recordset")
	 oYears.Open sSQL, Application("DSN"), 0, 1

	 lcl_return = "<select name=""iyear"" id=""iyear"" onchange=""YearPick();"">" & vbcrlf

	 do while not oYears.eof
    lcl_selected = ""

  		if clng(iDefaultYear) = clng(oYears("year")) then
		    	lcl_selected = " selected=""selected"""
  		end if

		  lcl_return = lcl_return & "<option value=""" & oYears("year") & """" & lcl_selected & ">" & oYears("year") & "</option>" & vbcrlf

  		oYears.MoveNext
	 loop

	 lcl_return = lcl_return & "</select>"

	 oYears.close
	 set oYears = nothing

  ShowYearChoices = lcl_return

end function

'------------------------------------------------------------------------------
Function GetFirstMembershipId()
	Dim sSql, oMember

	sSQL= "Select MIN(membershipid) as membershipid FROM egov_memberships WHERE orgid = " & session("orgid") 
	
	Set oMember = Server.CreateObject("ADODB.Recordset")
	oMember.Open sSQL, Application("DSN"), 3, 1
	
	If IsNull(oMember("membershipid")) Then
		GetFirstMembershipId = 0
	Else
		GetFirstMembershipId = oMember("membershipid")
	End If 
	
	oMember.close
	Set oMember = Nothing
End Function 

'------------------------------------------------------------------------------
sub GetPoolPassCountByYear( iYear, iMembershipId, iPeriodId )
	Dim  oCmd, oCounts, sOldResident, iTotal, sSql, sWhere

	sOldResident = "s"
	iTotal       = 0
	sWhere       = ""

 if CLng(iMembershipId) > CLng(0) then
  		sWhere = " AND P.membershipid = " & iMembershipId
 end if
	
	if CLng(iPeriodId) > CLng(0) then
  		sWhere = sWhere & " AND P.periodid = " & iPeriodId
 end if

	sSQL = "SELECT T.description AS restype, R.description AS ratetype, "

 if CLng(iPeriodId) = CLng(0) then
    sSQL = sSQL & " (select mp.period_desc "
    sSQL = sSQL &  " from egov_membership_periods mp "
    sSQL = sSQL &  " where mp.orgid = " & session("orgid")
    sSQL = sSQL &  " and mp.periodid = p.periodid) AS period_desc, "
 end if

 sSQL = sSQL & " COUNT(P.poolpassid) AS passes "
	sSQL = sSQL & " FROM egov_poolpassresidenttypes T, egov_poolpassrates R, egov_poolpasspurchases P "
	sSQL = sSQL & " WHERE T.orgid = " & session("orgid")
 sSQL = sSQL & " AND T.orgid = R.orgid "
	sSQL = sSQL & " AND T.orgid = P.orgid "
	sSQL = sSQL & " AND T.resident_type = R.residenttype "
	sSQL = sSQL & " AND P.rateid = R.rateid "
	sSQL = sSQL & " AND YEAR(P.paymentdate) = " & iYear 
 sSQL = sSQL & " AND P.paymentresult <> 'Pending' "
	sSQL = sSQL & " AND P.paymentresult <> 'Declined' "
	sSQL = sSQL & sWhere

 if CLng(iPeriodId) = CLng(0) then
    lcl_groupby_periodid = ", P.periodid"
 else
    lcl_groupby_periodid = ""
 end if

	sSQL = sSQL & " GROUP BY T.description, R.description" & lcl_groupby_periodid & ", T.displayorder, R.displayorder "
	sSQL = sSQL & " ORDER BY T.displayorder, R.displayorder"
	
	set oCounts = Server.CreateObject("ADODB.Recordset")
	oCounts.Open sSQL, Application("DSN"), 3, 1

	do while not oCounts.eof
  		if sOldResident <> oCounts("restype") then
    			sOldResident = oCounts("restype")

    			response.write "<h3>&nbsp;</h3>" & vbcrlf
    			response.write "<h3>" & oCounts("restype") & "</h3>" & vbcrlf

   	end if

    if CLng(iPeriodId) = CLng(0) then
       lcl_display_period_desc = " (" & oCounts("period_desc") & ")"
    else
       lcl_display_period_desc = ""
    end if

  		response.write "<div class=""ratetype"">" & oCounts("ratetype") & lcl_display_period_desc & "<div class=""typecount"">" & oCounts("passes") & "</div></div>" & vbcrlf

  		iTotal = iTotal + clng(oCounts("passes"))
  		oCounts.movenext
 loop

	response.write "<h3>&nbsp;</h3>" & vbcrlf
	response.write "<h3>Total</h3>" & vbcrlf
	response.write "<div class=""ratetype"">Combined<div class=""typecount"">" & iTotal & "</div></div>" & vbcrlf

	oCounts.close
	set oCounts = nothing 
	
end sub
%>
