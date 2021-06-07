<!-- #include file="../includes/common.asp" //-->
<%
  lcl_orgid            = 0
  lcl_userid           = 0
  lcl_memberid         = 0
  lcl_memberBarcodeID  = 0
  lcl_statusid         = 0
  lcl_isStatusActive   = true
  lcl_barcode          = ""
  lcl_comments         = ""
  lcl_removeIsChecked  = "N"
  lcl_lastmodifieddate = "'" & ConvertDateTimetoTimeZone() & "'"
  lcl_createddate      = "'" & ConvertDateTimetoTimeZone() & "'"
  lcl_canAddUpdate     = true

  if request("orgid") <> "" then
     if isnumeric(request("orgid")) then
        lcl_orgid = request("orgid")
     end if
  end if

  if request("userid") <> "" then
     if isnumeric(request("userid")) then
        lcl_userid = request("userid")
     end if
  end if

  if request("memberid") <> "" then
     if isnumeric(request("memberid")) then
        lcl_memberid = request("memberid")
     end if
  end if

  if request("memberBarcodeID") <> "" then
    if isnumeric(request("memberBarcodeID")) then
       lcl_memberBarcodeID = request("memberBarcodeID")
    end if
  end if

  if request("statusid") <> "" then
     if isnumeric(request("statusid")) then
        lcl_statusid = request("statusid")
     end if
  end if

  if request("barcode") <> "" then
     lcl_barcode = request("barcode")
  end if

  if request("comments") <> "" then
     lcl_comments = request("comments")
  end if

  if request("removeIsChecked") = "Y" then
     lcl_removeIsChecked = "Y"
  end if

  if lcl_removeIsChecked = "Y" then
     sSQL = "DELETE FROM egov_poolpassmembers_to_barcodes WHERE memberbarcodeid = " & lcl_memberBarcodeID

     RunSQLStatement sSQL

     lcl_success = "SD"
  else
    '1. Check to see if the current status is "active"
    '2. If "no" then check for a duplicate barcode.
    '3. If "yes" then check for at least one different "active" barcode for the member.
     lcl_isStatusActive = isStatusActive(lcl_orgid, _
                                         lcl_statusid)

     lcl_memberHasActiveBarcode = checkMemberHasDifferentActiveBarcode(lcl_orgid, _
                                                                       lcl_memberid, _
                                                                       lcl_memberBarcodeID, _
                                                                       lcl_barcode)

     lcl_memberHasDuplicateBarcode = checkMemberForDuplicateBarcode(lcl_barcode, _
                                                                    lcl_orgid, _
                                                                    lcl_memberBarcodeID)

     if lcl_isStatusActive then
        if not lcl_memberHasActiveBarcode then
           if lcl_memberHasDuplicateBarcode then
              lcl_success      = "DuplicateBarcodeExists"
              lcl_canAddUpdate = false
           end if
        else
           lcl_success      = "ActiveBarcodeForMemberExists"
           lcl_canAddUpdate = false
        end if
     else
        if lcl_memberHasDuplicateBarcode then
           lcl_success      = "DuplicateBarcodeExists"
           lcl_canAddUpdate = false
        end if
     end if
     
     if lcl_canAddUpdate then
        if lcl_barcode <> "" then
           lcl_barcode = dbsafe(lcl_barcode)
        end if

        if lcl_comments <> "" then
           lcl_comments = dbsafe(lcl_comments)
        end if

        lcl_barcode  = "'" & lcl_barcode  & "'"
        lcl_comments = "'" & lcl_comments & "'"

        if lcl_memberBarcodeID > 0 then
           sSQL = "UPDATE egov_poolpassmembers_to_barcodes SET "
           sSQL = sSQL & "barcode = "          & lcl_barcode  & ", "
           sSQL = sSQL & "barcode_statusid = " & lcl_statusid & ", "
           sSQL = sSQL & "barcode_comments = " & lcl_comments & ", "
           sSQL = sSQL & "lastmodifiedbyid = " & lcl_userid   & ", "
           sSQL = sSQL & "lastmodifieddate = " & lcl_lastmodifieddate
           sSQL = sSQL & " WHERE memberbarcodeid = " & lcl_memberBarcodeID

           RunSQLStatement sSQL

           lcl_success = "Saved"
        else
           sSQL = "INSERT INTO egov_poolpassmembers_to_barcodes ("
           sSQL = sSQL & "orgid, "
           sSQL = sSQL & "memberid, "
           sSQL = sSQL & "barcode, "
           sSQL = sSQL & "barcode_statusid, "
           sSQL = sSQL & "barcode_comments, "
           sSQL = sSQL & "createdbyid, "
           sSQL = sSQL & "createddate "
           sSQL = sSQL & ") VALUES ("
           sSQL = sSQL & lcl_orgid    & ", "
           sSQL = sSQL & lcl_memberid & ", "
           sSQL = sSQL & lcl_barcode  & ", "
           sSQL = sSQL & lcl_statusid & ", "
           sSQL = sSQL & lcl_comments & ", "
           sSQL = sSQL & lcl_userid   & ", "
           sSQL = sSQL & lcl_createddate
           sSQL = sSQL & ") "

           lcl_memberBarcodeID = RunInsertStatement(sSQL)

           lcl_success = "Added"
        end if
     end if
  end if

  response.write lcl_success

'------------------------------------------------------------------------------
function checkMemberHasDifferentActiveBarcode(iOrgID, _
                                              iMemberID, _
                                              iMemberBarcodeID, _
                                              iBarcode)

  dim lcl_return, sSQL, sBarcode, sTotalBarcodes

  lcl_return = false

  sBarcode       = ""
  sTotalBarcodes = 0

  if iBarcode <> "" then
     sBarcode = dbsafe(iBarcode)
  end if

  sBarcode = "'" & sBarcode & "'"

 'Check to see if there is a different barcode that is in "active" status for the memberid
  sSQL = "SELECT count(mtb.memberbarcodeid) as totalBarcodes "
  sSQL = sSQL & " FROM egov_poolpassmembers_to_barcodes mtb "
  sSQL = sSQL &      " INNER JOIN egov_poolpassmembers_barcode_statuses bs ON bs.statusid = mtb.barcode_statusid "
  sSQL = sSQL & " WHERE bs.isEnabled = 1 "
  sSQL = sSQL & " AND bs.isActiveStatus = 1 "
  sSQL = sSQL & " AND bs.orgid = mtb.orgid "
  sSQL = sSQL & " AND mtb.orgid = "            & iOrgID
  sSQL = sSQL & " AND mtb.memberid = "         & iMemberID
  sSQL = sSQL & " AND mtb.memberbarcodeid <> " & iMemberBarcodeID
  sSQL = sSQL & " AND mtb.barcode <> "         & sBarcode

  set oCheckDifferentBarcodeExists = Server.CreateObject("ADODB.Recordset")
  oCheckDifferentBarcodeExists.Open sSQL, Application("DSN"), 3, 1

  if not oCheckDifferentBarcodeExists.eof then
     sTotalBarcodes = oCheckDifferentBarcodeExists("totalBarcodes")
  end if

  oCheckDifferentBarcodeExists.close
  set oCheckDifferentBarcodeExists = nothing

  if sTotalBarcodes > 0 then
     lcl_return = true
  end if

  checkMemberHasDifferentActiveBarcode = lcl_return

end function

'------------------------------------------------------------------------------
function checkMemberForDuplicateBarcode(iBarcode, _
                                        iOrgID, _
                                        iMemberBarcodeID)

  dim lcl_return, sSQL, sMemberBarcodeID

  lcl_return       = false
  sMemberBarcodeID = 0
  sTotalBarcodes   = 0

  if iMemberBarcodeID <> "" then
     if isnumeric(iMemberBarcodeID) then
        sMemberBarcodeID = clng(iMemberBarcodeID)
     end if
  end if

  if iBarcode <> "" then
     sBarcode = dbsafe(iBarcode)
  end if

  sBarcode = "'" & sBarcode & "'"

 'Check to see if the barcode value will be duplicate of another member/barcode assignment that is 'Enabled'
  sSQL = "SELECT count(memberbarcodeid) as totalBarcodes "
  sSQL = sSQL & " FROM egov_poolpassmembers_to_barcodes "
  sSQL = sSQL & " WHERE orgid = " & iOrgID
  sSQL = sSQL & " AND barcode = " & sBarcode
  sSQL = sSQL & " AND barcode_statusid IN (select statusid "
  sSQL = sSQL &                          " from egov_poolpassmembers_barcode_statuses "
  sSQL = sSQL &                          " where isEnabled = 1 "
  sSQL = sSQL &                          " and orgid = " & iOrgID & ") "

  if sMemberBarcodeID > 0 then
     sSQL = sSQL & " AND memberBarcodeID <> " & sMemberBarcodeID
  end if

  set oCheckForDuplicateBarcodes = Server.CreateObject("ADODB.Recordset")
  oCheckForDuplicateBarcodes.Open sSQL, Application("DSN"), 3, 1

  if not oCheckForDuplicateBarcodes.eof then
     sTotalBarcodes = oCheckForDuplicateBarcodes("totalBarcodes")
  end if

  oCheckForDuplicateBarcodes.close
  set oCheckForDuplicateBarcodes = nothing

  if sTotalBarcodes > 0 then
     lcl_return = true
  end if

  checkMemberForDuplicateBarcode = lcl_return

end function

'------------------------------------------------------------------------------
function isStatusActive(iOrgID, _
                        iStatusID)

  dim lcl_return, sSQL

  lcl_return = false

  sSQL = "SELECT isActiveStatus "
  sSQL = sSQL & " FROM egov_poolpassmembers_barcode_statuses "
  sSQL = sSQL & " WHERE orgid = " & iOrgID
  sSQL = sSQL & " AND statusid = " & iStatusID
  sSQL = sSQL & " AND isEnabled = 1 "

  set oIsStatusActive = Server.CreateObject("ADODB.Recordset")
  oIsStatusActive.Open sSQL, Application("DSN"), 3, 1

  if not oIsStatusActive.eof then
     if oIsStatusActive("isActiveStatus") then
        lcl_return = true
     end if
  end if

  oIsStatusActive.close
  set oIsStatusActive = nothing

  isStatusActive = lcl_return

end function
%>