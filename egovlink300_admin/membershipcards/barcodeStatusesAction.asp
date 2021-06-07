<!-- #include file="../includes/common.asp" //-->
<%
  lcl_member_id      = request("memberid")
  lcl_action         = request("action")
  lcl_poolpassid     = request("poolpassid")
  lcl_orgid          = request("orgid")
  lcl_totalStatuses  = 0
  lcl_isActiveStatus = 0
  lcl_isEnabled      = 0

  if request("totalStatuses") <> "" then
     lcl_totalStatuses = clng(request("totalStatuses"))
  end if

  if lcl_totalStatuses > 0 AND lcl_orgid > 0 then
     lcl_statusid   = clng(0)
     lcl_statusname = ""

     for i = 1 to lcl_totalStatuses
        lcl_removeStatus = ""

        if request("statusid" & i) <> "" then
           lcl_statusid = clng(request("statusid" & i))
        end if

        if request("removeStatus" & i) <> "" then
           lcl_removeStatus = request("removeStatus" & i)
        end if

        if lcl_removeStatus = "Y" then
           sSQL = "DELETE FROM egov_poolpassmembers_barcode_statuses WHERE statusid = " & lcl_statusid

           RunSQLStatement sSQL
        else
           lcl_statusname     = request("statusname" & i)
           lcl_isActiveStatus = 0
           lcl_isEnabled      = 0

           if lcl_statusname <> "" then
              lcl_statusname = dbsafe(lcl_statusname)
           end if

           lcl_statusname = "'" & lcl_statusname & "'"

           if request("isActiveStatus") <> "" then
              if clng(cstr(request("isActiveStatus"))) = clng(cstr(lcl_statusid)) then
                 lcl_isActiveStatus = 1
              end if
           end if

           if request("isEnabled" & i) <> "" then
              if request("isEnabled" & i) = "Y" then
                 lcl_isEnabled = 1
              end if
           end if

           if clng(lcl_statusid) = clng(0) then
              sSQL = "INSERT INTO egov_poolpassmembers_barcode_statuses ("
              sSQL = sSQL & " statusname, "
              sSQL = sSQL & " isActiveStatus, "
              sSQL = sSQL & " isEnabled, "
              sSQL = sSQL & " orgid "
              sSQL = sSQL & " ) VALUES ( "
              sSQL = sSQL & lcl_statusname     & ", "
              sSQL = sSQL & lcl_isActiveStatus & ", "
              sSQL = sSQL & lcl_isEnabled      & ", "
              sSQL = sSQL & lcl_orgid
              sSQL = sSQL & ") "

              lcl_statusid = RunInsertStatement(sSQL)
           else
              sSQL = "UPDATE egov_poolpassmembers_barcode_statuses SET "
              sSQL = sSQL & " statusname = "     & lcl_statusname       & ", "
              sSQL = sSQL & " isActiveStatus = " & lcl_isActiveStatus & ", "
              sSQL = sSQL & " isEnabled = "      & lcl_isEnabled
              sSQL = sSQL & " WHERE orgid = " & lcl_orgid
              sSQL = sSQL & " AND statusid = " & lcl_statusid

              RunSQLStatement sSQL
           end if
        end if
     next
  end if

  lcl_return_url = "barcodeStatusesMaint.asp"
  lcl_return_url = lcl_return_url & "?memberid=" & lcl_member_id
  lcl_return_url = lcl_return_url & "&action=" & lcl_action
  lcl_return_url = lcl_return_url & "&poolpassid=" & lcl_poolpassid
  lcl_return_url = lcl_return_url & "&success=SU"

  response.redirect lcl_return_url
%>
