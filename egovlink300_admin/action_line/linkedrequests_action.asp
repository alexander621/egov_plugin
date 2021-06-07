<!-- #include file="../includes/common.asp" //-->
<%
 lcl_action = ""
 lcl_isAjax = "N"
 lcl_rowid  = 1
 lcl_orgid  = 0

'Variables for delete
 lcl_currentRequestID   = 0
 lcl_requestToBeRemoved = 0

'Variables for adding/editing
 lcl_linkedrequestid = 0
 lcl_description     = ""

'Variables for adding
 lcl_parent_trackingnumber = 0
 lcl_parent_requestid      = 0
 lcl_linked_trackingnumber = 0
 lcl_linked_requestid      = 0

 if request("action") <> "" then
    lcl_action = ucase(request("action"))
 end if

 if request("isAjax") <> "" then
    lcl_isAjax = ucase(request("isAjax"))
 end if

 if request("rowid") <> "" then
    if isnumeric(request("rowid")) then
       'lcl_rowid = clng(request("rowid"))
       lcl_rowid = request("rowid")
    end if
 end if

 if lcl_action = "D" then
    if request("currentrequestid") <> "" then
       if isnumeric(request("currentrequestid")) then
          lcl_currentrequestid = clng(request("currentrequestid"))
       end if
    end if

    if request("requestToBeRemoved") <> "" then
       if isnumeric(request("requestToBeRemoved")) then
          lcl_requestToBeRemoved = clng(request("requestToBeRemoved"))
       end if
    end if

    deleteLinkedRequests lcl_rowid, lcl_currentrequestid, lcl_requestToBeRemoved

 else
    if request("description") <> "" then
       lcl_description = request("description")
    end if

    if lcl_action = "A" then
       if request("orgid") <> "" then
          if isnumeric(request("orgid")) then
             lcl_orgid = clng(request("orgid"))
          end if
       end if

       if request("parent_trackingnumber") <> "" then
          if isnumeric(request("parent_trackingnumber")) then
             lcl_parent_trackingnumber = request("parent_trackingnumber")
          end if
       end if

       if request("parent_requestid") <> "" then
          if isnumeric(request("parent_requestid")) then
             lcl_parent_requestid = request("parent_requestid")
          end if
       end if

       if request("linked_trackingnumber") <> "" then
          if isnumeric(request("linked_trackingnumber")) then
             lcl_linked_trackingnumber        = request("linked_trackingnumber")
             lcl_linked_trackingnumber_length = len(lcl_linked_trackingnumber)
             lcl_linked_requestid             = MID(lcl_linked_trackingnumber,1,lcl_linked_trackingnumber_length - 4)
          end if
       end if

       maintainLinkedRequest lcl_rowid, lcl_orgid, lcl_linkedrequestid, _
                             lcl_parent_trackingnumber, lcl_parent_requestid, _
                             lcl_linked_trackingnumber, lcl_linked_requestid, _
                             lcl_description
    else
       if request("linkedrequestid") <> "" then
          if isnumeric(request("linkedrequestid")) then
             lcl_linkedrequestid = clng(request("linkedrequestid"))
          end if
       end if

       maintainLinkedRequest lcl_rowid, lcl_orgid, lcl_linkedrequestid, _
                             lcl_parent_trackingnumber, lcl_parent_requestid, _
                             lcl_linked_trackingnumber, lcl_linked_requestid, _
                             lcl_description
    end if
 end if

'------------------------------------------------------------------------------
sub maintainLinkedRequest(iRowID, iOrgID, iLinkedRequestID, iParent_trackingNumber, iParent_requestID, _
                          iLinked_trackingNumber, iLinked_requestID, iDescription)

  dim sLinkedRequestID, sParent_trackingNumber, sParent_requestID
  dim sLinked_trackingNumber, sLinked_requestID, sDescription
  dim sSQL, lcl_success, sRowID, sOrgID

  sLinkedRequestID       = 0
  sParent_trackingNumber = 0
  sParent_requestid      = 0
  sLinked_trackingNumber = 0
  sLinked_requestid      = 0
  sDescription           = ""
  lcl_success            = ""
  sRowID                 = "0"
  sOrgID                 = 0

  if iLinkedRequestID <> "" then
     sLinkedRequestID = clng(iLinkedRequestID)
  end if

  if iRowID <> "" then
     if isnumeric(iRowID) then
        sRowID = clng(iRowID)
     end if
  end if

  if iOrgID <> "" then
     if isnumeric(iOrgID) then
        sOrgID = clng(iOrgID)
     end if
  end if

  if iDescription <> "" then
     sDescription = dbsafe(iDescription)
     sDescription = "'" & sDescription & "'"
  else
     sDescription = "NULL"
  end if

 'Determine if we are to ADD or UPDATE.
  if sLinkedRequestID > 0 then
     sSQL = "UPDATE egov_actionline_linkedrequests SET "
     sSQL = sSQL & " description = " & sDescription
     sSQL = sSQL & " WHERE linkedrequestid = " & sLinkedRequestID

    	set oSaveLRequests = Server.CreateObject("ADODB.Recordset")
   	 oSaveLRequests.Open sSQL, Application("DSN"), 3, 1

     set oSaveLRequests = nothing

     lcl_success = "S_" & sRowID

     response.write lcl_success

  else
     if iParent_trackingNumber <> "" then
        if isnumeric(iParent_trackingNumber) then
           sParent_trackingNumber = iParent_trackingNumber
        end if
     end if

     if iParent_requestID <> "" then
        if isnumeric(iParent_requestID) then
           sParent_requestID = iParent_requestID
        end if
     end if

     if iLinked_trackingNumber <> "" then
        if isnumeric(iLinked_trackingNumber) then
           sLinked_trackingNumber = iLinked_trackingNumber
        end if
     end if

     if iLinked_requestID <> "" then
        if isnumeric(iLinked_requestID) then
           sLinked_requestID = iLinked_requestID
        end if
     end if

    'Determine if the linked_trackingnumber is already "linked" to the current request.
     lcl_exists = checkLinkRequestComboExists(sOrgID, sParent_requestID, sLinked_requestID)

     if not lcl_exists then
        sSQL = "INSERT INTO egov_actionline_linkedrequests ("
        sSQL = sSQL & "orgid, "
        sSQL = sSQL & "parent_trackingnumber, "
        sSQL = sSQL & "parent_requestid, "
        sSQL = sSQL & "linked_trackingnumber, "
        sSQL = sSQL & "linked_requestid, "
        sSQL = sSQL & "description "
        sSQL = sSQL & ") VALUES ("
        sSQL = sSQL & sOrgID                 & ", "
        sSQL = sSQL & sParent_trackingNumber & ", "
        sSQL = sSQL & sParent_requestID      & ", "
        sSQL = sSQL & sLinked_trackingNumber & ", "
        sSQL = sSQL & sLinked_requestID      & ", "
        sSQL = sSQL & sDescription
        sSQL = sSQL & ") "

        sLinkedRequestID = RunInsertStatement(sSQL)

        response.write sRowID & "," & sLinkedRequestID
     else
        response.write "DUPLICATE"
     end if

     'lcl_success = sRowID & "-" & sLinkedRequestID

  end if

end sub

'------------------------------------------------------------------------------
function checkLinkRequestComboExists(iOrgID, iParentRequestID, iLinkedRequestID)
  dim lcl_return, sSQL, sOrgID, sParent_requestID, sLinked_requestID

  lcl_return        = false
  sParent_requestid = 0
  sLinked_requestid = 0
  sOrgID            = 0

  if iOrgID <> "" then
     if isnumeric(iOrgID) then
        sOrgID = clng(iOrgID)
     end if
  end if

  if iParentRequestID <> "" then
     if isnumeric(iParentRequestID) then
        sParent_requestID = iParentRequestID
     end if
  end if

  if iLinkedRequestID <> "" then
     if isnumeric(iLinkedRequestID) then
        sLinked_requestID = iLinkedRequestID
     end if
  end if

  sSQL = "SELECT DISTINCT 'Y' as lcl_exists "
  sSQL = sSQL & " FROM egov_actionline_linkedrequests "
  sSQL = sSQL & " WHERE (parent_requestid = " & sParent_requestID & " AND linked_requestid = " & sLinked_requestID & ") "
  sSQL = sSQL &    " OR (parent_requestid = " & sLinked_requestID & " AND linked_requestid = " & sParent_requestID & ") "
  sSQL = sSQL & " AND orgid = " & sOrgID

  set oLRequestsExists = Server.CreateObject("ADODB.Recordset")
  oLRequestsExists.Open sSQL, Application("DSN"), 3, 1

  if not oLRequestsExists.eof then
     lcl_return = true
  end if

  oLRequestsExists.close
  set oLRequestsExists = nothing  

  checkLinkRequestComboExists = lcl_return

end function

'------------------------------------------------------------------------------
sub deleteLinkedRequests(iRowID, iCurrentRequestID, iRequestToBeRemovedID)
  dim sCurrentRequestID, sRequestToBeRemovedID, sSQL, lcl_success, sRowID

  sCurrentRequestID     = 0
  sRequestToBeRemovedID = 0
  lcl_success           = ""
  sRowID                = 0

  if iCurrentRequestID <> "" then
     sCurrentRequestID = clng(iCurrentRequestID)
  end if

  if iRequestToBeRemovedID <> "" then
     sRequestToBeRemovedID = clng(iRequestToBeRemovedID)
  end if

  if iRowID <> "" then
     sRowID = clng(iRowID)
  end if

  if sCurrentRequestID > 0 AND sRequestToBeRemovedID > 0 then
     sSQL = "DELETE FROM egov_actionline_linkedrequests "
     sSQL = sSQL & " WHERE (parent_requestid = " & sCurrentRequestID     & " AND linked_requestid = " & sRequestToBeRemovedID & ") "
     sSQL = sSQL &    " OR (parent_requestid = " & sRequestToBeRemovedID & " AND linked_requestid = " & sCurrentRequestID     & ") "

    	set oLRequests = Server.CreateObject("ADODB.Recordset")
   	 oLRequests.Open sSQL, Application("DSN"), 3, 1

     set oLRequests = nothing

     lcl_success = "D"
  end if

  response.write lcl_success

end sub

'------------------------------------------------------------------------------
sub dtb_debug(p_value)

  if p_value <> "" then
     sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"
    	set oDTB = Server.CreateObject("ADODB.Recordset")
   	 oDTB.Open sSQL, Application("DSN"), 3, 1
  end if

end sub
%>