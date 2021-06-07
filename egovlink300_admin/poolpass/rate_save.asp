<%
Dim x, iRateId, iPublicCanPurchase

lcl_resident_type   = request("sResidentType")
lcl_membership_id   = request("iMembershipId")
lcl_membership_type = request("sMembershipType")
lcl_orgid           = request("orgid")

'Add Rate ---------------------------------------------------------------------
 if request("total_rates") = "" then
    lcl_rateid       = request("rateid_0")
    lcl_description  = request("description_0")
    lcl_amount       = request("amount_0")
    lcl_maxsignups   = request("maxsignups_0")
    lcl_message      = left(request("message_0"),500)
    lcl_displayorder = request("displayorder_0")
    lcl_periodid     = request("iperiodid_0")

    lcl_PublicCanPurchase        = 0
    lcl_isrenewable              = 0
    lcl_renewalstartdate         = ""
    lcl_renewalTimeAfterExpire   = 0
    lcl_cardid                   = "NULL"
    iIsPunchcard                 = 0
    iPunchCardLimit              = 0
    lcl_nonresident_prestartdate = ""
    lcl_nonresident_preenddate   = ""
    lcl_isEnabled                = 1

    lcl_attendancetypeid = request("attendancetypeid_0")

    if request("publiccanpurchase_0") = "on" then
       lcl_PublicCanPurchase = 1
    end if

    if request("isrenewable_0") <> "" then
       lcl_isrenewable = request("isrenewable_0")
    end if

    if request("renewalstartdate_0") <> "" then
       lcl_renewalstartdate = CDate(request("renewalstartdate_0"))
    end if

    if request("renewalTimeAfterExpire_0") <> "" then
       lcl_renewalTimeAfterExpire = request("renewalTimeAfterExpire_0")
    end if

    if request("cardid_0") <> "" then
       lcl_cardid = request("cardid_0")
    end if

    if request("isPunchcard_0") = "on" then
       iIsPunchcard = 1
    end if

    if request("punchcard_limit_0") <> "" then
       iPunchCardLimit = request("punchcard_limit_0")
    end if

    if request("nonresident_prestartdate_0") <> "" then
       lcl_nonresident_prestartdate = CDate(request("nonresident_prestartdate_0"))
    end if

    if request("nonresident_preenddate_0") <> "" then
       lcl_nonresident_preenddate = CDate(request("nonresident_preenddate_0"))
    end if

   'Update each record
    iRateId = SaveRate(lcl_orgid, _
                       lcl_resident_type, _
                       lcl_rateid, _
                       lcl_description, _
                       lcl_amount, _
                       lcl_maxsignups, _
                       lcl_message, _
                       lcl_displayorder, _
                       lcl_membership_id, _
                       lcl_periodid, _
                       lcl_PublicCanPurchase, _
                       lcl_attendancetypeid, _
                       lcl_isrenewable, _
                       lcl_renewalstartdate, _
                       lcl_renewalTimeAfterExpire, _
                       lcl_cardid, _
                       iIsPunchcard, _
                       iPunchCardLimit, _
                       lcl_nonresident_prestartdate, _
                       lcl_nonresident_preenddate, _
                       lcl_isEnabled)

    if request("relation_0").count > 0 then
       for x = 1 to request("relation_0").count
         		AddPreselects iRateId, request("relation_0")(x)
       next
    end if

    lcl_success = "&success=SN"

'Save Rates -------------------------------------------------------------------
 else
   
    for e = 1 to request("total_rates")
        lcl_rateid       = request("rateid_"       & e)
        lcl_description  = request("description_"  & e)
        lcl_amount       = request("amount_"       & e)
        lcl_maxsignups   = request("maxsignups_"   & e)
        lcl_message      = left(request("message_" & e),500)
        lcl_displayorder = request("displayorder_" & e)
        lcl_periodid     = request("iperiodid_"    & e)

        lcl_PublicCanPurchase        = 0 
        lcl_isrenewable              = 0
        lcl_renewalstartdate         = ""
        lcl_renewalTimeAfterExpire   = 0
        lcl_cardid                   = "NULL"
        iIsPunchcard                 = 0
        iPunchCardLimit              = 0
        lcl_nonresident_prestartdate = ""
        lcl_nonresident_preenddate   = ""
        lcl_isEnabled                = 1

        lcl_attendancetypeid = request("attendancetypeid_"&e)

        if request("publiccanpurchase_" & e) = "on" then
           lcl_PublicCanPurchase = 1
        end if

        if request("disableRate_" & e) = "Y" then
           lcl_isEnabled = 0
        end if

        if request("isrenewable_" & e) <> "" then
           lcl_isrenewable = request("isrenewable_" & e)
        end if

        if request("renewalstartdate_" & e) <> "" then
           lcl_renewalstartdate = CDate(request("renewalstartdate_" & e))
        end if

        if request("renewalTimeAfterExpire_" & e) <> "" then
           lcl_renewalTimeAfterExpire = request("renewalTimeAfterExpire_" & e)
        end if

        if request("cardid_" & e) <> "" then
           lcl_cardid = request("cardid_" & e)
        end if

        if request("isPunchcard_" & e) = "on" then
           iIsPunchcard = 1
        end if

        if request("punchcard_limit_"&e) <> "" then
           iPunchCardLimit = request("punchcard_limit_" & e)
        end if

        if request("nonresident_prestartdate_" & e) <> "" then
           lcl_nonresident_prestartdate = CDate(request("nonresident_prestartdate_" & e))
        end if

        if request("nonresident_preenddate_" & e) <> "" then
           lcl_nonresident_preenddate = CDate(request("nonresident_preenddate_" & e))
        end if

       'Update each record
        iRateId = SaveRate(lcl_orgid, _
                           lcl_resident_type, _
                           lcl_rateid, _
                           lcl_description, _
                           lcl_amount, _
                           lcl_maxsignups, _
                           lcl_message, _
                           lcl_displayorder, _
                           lcl_membership_id, _
                           lcl_periodid, _
                           lcl_PublicCanPurchase, _
                           lcl_attendancetypeid, _
                           lcl_isrenewable, _
                           lcl_renewalstartdate, _
                           lcl_renewalTimeAfterExpire, _
                           lcl_cardid, _
                           iIsPunchcard, _
                           iPunchCardLimit, _
                           lcl_nonresident_prestartdate, _
                           lcl_nonresident_preenddate, _
                           lcl_isEnabled)

        'iRateId = SaveRate(request("sResidentType"), request("rateid"), request("description"), request("amount"), request("maxsignups"), _
                           'request("message"), request("displayorder"), request("iMembershipId"), request("iPeriodId"), iPublicCanPurchase, _
                           'request("attendancetypeid") )

       'Add any Preselected Family Member types
        if request("relation_"&e).count > 0 then
           for x = 1 to request("relation_"&e).count
             		AddPreselects iRateId, request("relation_"&e)(x)
           next
        end if
    next

    lcl_success = "&success=SU"

end if

'REDIRECT TO Pool Pass Rates PAGE
 response.redirect("poolpass_rates.asp?sResidentType=" & lcl_resident_type & "&iMembershipId=" & lcl_membership_id & "&sMembershipType=" & lcl_membership_type & lcl_success)

'------------------------------------------------------------------------------
Function SaveRate(sOrgID, _
                  sResidentType, _
                  iRateId, _
                  sDescription, _
                  nAmount, _
                  iMaxsignups, _
                  sMessage, _
                  iDisplayOrder, _
                  iMembershipId, _
                  iPeriodId, _
                  iPublicCanPurchase, _
                  iAttendanceTypeId, _
                  p_isRenewable, _
                  p_RenewalStartDate, _
                  p_renewalTimeAfterExpire, _
                  p_cardid, _
                  sIsPunchcard, _
                  sPunchCardLimit, _
                  sNonResidentPreStartDate, _
                  sNonResidentPreEndDate, _
                  iIsEnabled)

	Dim sSql, oCmd, oInsert

	if iRateId = "0" then
	  'Insert new records
 		 sSQL = "SET NOCOUNT ON;"
  		sSQL = sSQL & "INSERT INTO egov_poolpassrates ("
		  sSQL = sSQL & "orgid, "
    sSQL = sSQL & "residenttype, "
    sSQL = sSQL & "description, "
    sSQL = sSQL & "amount, "
    sSQL = sSQL & "displayorder, "
    sSQL = sSQL & "maxsignups, "
    sSQL = sSQL & "message, "
    sSQL = sSQL & "membershipid, "
    sSQL = sSQL & "periodid, "
  		sSQL = sSQL & "publiccanpurchase, "
    sSQL = sSQL & "attendancetypeid, "
    sSQL = sSQL & "isRenewable, "
    sSQL = sSQL & "renewalstartdate, "
    sSQL = sSQL & "renewalTimeAfterExpire, "
    sSQL = sSQL & "cardid, "
    sSQL = sSQL & "isPunchcard, "
    sSQL = sSQL & "punchcard_limit, "
    sSQL = sSQL & "isEnabled "
', nonresident_prestartdate, nonresident_preenddate
    sSQL = sSQL & ") VALUES (" 
  		sSQL = sSQL &       sOrgID                     & ", "
		  sSQL = sSQL & "'" & dbsafe(sResidentType)      & "', "
  		sSQL = sSQL & "'" & dbsafe(sDescription)       & "', "
		  sSQL = sSQL &       nAmount                    & ", "
  		sSQL = sSQL &       iDisplayOrder              & ", "
		  sSQL = sSQL &       iMaxsignups                & ", "
   	sSQL = sSQL & "'" & dbsafe(sMessage)           & "', "
  		sSQL = sSQL &       iMembershipId              & ", "
		  sSQL = sSQL &       iPeriodId                  & ", "
  		sSQL = sSQL &       iPublicCanPurchase         & ", "
		  sSQL = sSQL &       iAttendanceTypeId          & ", "
    sSQL = sSQL &       p_isRenewable              & ", "
    sSQL = sSQL & "'" & dbsafe(p_RenewalStartDate) & "', "
    sSQL = sSQL &       p_renewalTimeAfterExpire   & ", "
    sSQL = sSQL &       p_cardid                   & ", "
    sSQL = sSQL &       sIsPunchcard               & ", "
    sSQL = sSQL &       sPunchCardLimit            & ", "
    sSQL = sSQL &       iIsEnabled
'    sSQL = sSQL & "'" & sNonResidentPreStartDate & "', "
'    sSQL = sSQL & "'" & sNonResidentPreEndDate   & "' "
  		sSQL = sSQL & ");"
		  sSQL = sSQL & "SELECT @@IDENTITY AS ROWID;"
		
   	set oInsert = Server.CreateObject("ADODB.Recordset")
  	 oInsert.Open sSQL, Application("DSN") , 3, 1

  		SaveRate = oInsert("ROWID")

  		oInsert.close
		  set oInsert = nothing

	else
 		'Update existing records
  		sSQL = "UPDATE egov_poolpassrates SET "
		  sSQL = sSQL & "description = '"              & dbsafe(sDescription)     & "', "
  		sSQL = sSQL & "amount = "                    & nAmount                  & ", "
		  sSQL = sSQL & "maxsignups = "                & iMaxsignups              & ", "
  		sSQL = sSQL & "message = '"                  & dbsafe(sMessage)         & "', "
		  sSQL = sSQL & "periodid = "                  & iPeriodId                & ", "
  		sSQL = sSQL & "membershipid = "              & iMembershipId            & ", "
		  sSQL = sSQL & "publiccanpurchase = "         & iPublicCanPurchase       & ", "
  		sSQL = sSQL & "attendancetypeid = "          & iAttendanceTypeId        & ", "
    sSQL = sSQL & "isRenewable = "               & p_isRenewable            & ", "
    sSQL = sSQL & "renewalstartdate = '"         & p_RenewalStartDate       & "', "
    sSQL = sSQL & "renewalTimeAfterExpire = "    & p_renewalTimeAfterExpire & ", "
    sSQL = sSQL & "cardid = "                    & p_cardid                 & ", "
    sSQL = sSQL & "isPunchcard = "               & sIsPunchcard             & ", "
    sSQL = sSQL & "punchcard_limit = "           & sPunchCardLimit          & ", "
    sSQL = sSQL & "isEnabled = "                 & iIsEnabled
'    sSQL = sSQL & "nonresident_prestartdate = '" & sNonResidentPreStartDate & "', "
'    sSQL = sSQL & "nonresident_preenddate = '"   & sNonResidentPreEndDate   & "' "
  		sSQL = sSQL & " WHERE rateid = " & iRateId & ""

   	set oUpdateRate = Server.CreateObject("ADODB.Recordset")
  	 oUpdateRate.Open sSQL, Application("DSN") , 3, 1

  	'Delete from the preselected family member table so we can re-add them in the next step
    sSQL = "DELETE FROM egov_poolpasspreselected WHERE rateid = " & iRateid

   	set oDeletePreSet = Server.CreateObject("ADODB.Recordset")
  	 oDeletePreSet.Open sSQL, Application("DSN") , 3, 1

    set oUpdateRate   = nothing
    set oDeletePreSet = nothing

	  	SaveRate = iRateid
	end if

end function

'------------------------------------------------------------------------------
' SUB AddPreselects( iRateid, sRelation )
' AUTHOR: Steve Loar
' CREATED: 02/09/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'------------------------------------------------------------------------------

Sub AddPreselects( iRateid, sRelation )
	Dim sSql

	sSQL = "INSERT INTO egov_poolpasspreselected (rateid, relation) VALUES (" & iRateid & ", '" & sRelation & "')"

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing

End Sub 

'------------------------------------------------------------------------------
function dbsafe(p_value)
  lcl_return = ""

  if p_value <> "" then
     lcl_return = replace(p_value,"'","''")
  end if

  dbsafe = lcl_return

end function

'------------------------------------------------------------------------------
function dtb_debug(p_value)

  sSQLi = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"
 	set oDTB = Server.CreateObject("ADODB.Recordset")
	 oDTB.Open sSQLi, Application("DSN") , 3, 1

  set oDTB = nothing

end function
%>