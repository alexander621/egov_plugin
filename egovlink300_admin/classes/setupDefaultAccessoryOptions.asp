<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
 lcl_orgid         = 0
 lcl_classid       = 0
 lcl_isAjaxRoutine = "Y"

 if request("orgid") <> "" then
    lcl_orgid = clng(request("orgid"))
 end if

 if request("classid") <> "" then
    lcl_classid = clng(request("classid"))
 end if

 if request("isAjaxRoutine") <> "" then
    if not containsApostrophe(request("isAjaxRoutine")) then
       lcl_isAjaxRoutine = request("isAjaxRoutine")
    end if
 end if
 
 if lcl_orgid > clng(0) AND lcl_classid > clng(0) then
    setupDefaultAccessoryOptions lcl_orgid, _
                                 lcl_classid, _
                                 lcl_isAjaxRoutine
 else
    if lcl_isAjaxRoutine = "Y" then
       response.write "error"
    end if
 end if

'------------------------------------------------------------------------------
sub setupDefaultAccessoryOptions(p_orgid, p_classid, lcl_isAjaxRoutine)
  dim sOrgID, sClassID, sIsAjaxRoutine, lcl_total_accessories, sLineCount

  sOrgID                = 0
  sClassID              = 0
  sIsAjaxRoutine        = "Y"
  lcl_total_accessories = 0
  sLineCount            = 0

  if p_orgid <> "" then
     sOrgID = clng(p_orgid)
  end if

  if p_classid <> "" then
     sClassID = clng(p_classid)
  end if

  if p_isAjaxRoutine <> "" then
     if not containsApostrophe(p_isAjaxRoutine) then
        sIsAjaxRoutine = p_isAjaxRoutine
     end if
  end if

 'Check to see if this accessory type for this class already has values assigned
  lcl_total_accessories = getTotalAccessories(sOrgID)

 'If no accessories exist for the org then insert the default accessories
  if lcl_total_accessories < 1 then
     sSQL = "SELECT accessorytype,  "
     sSQL = sSQL & " accessoryname, "
     sSQL = sSQL & " accessoryvalue, "
     sSQL = sSQL & " displayorder "
     sSQL = sSQL & " FROM egov_class_teamroster_accessories "
     sSQL = sSQL & " WHERE orgid = 0 "
     sSQL = sSQL & " ORDER BY accessorytype, displayorder "

     set oDefaultAccessories = Server.CreateObject("ADODB.Recordset")
     oDefaultAccessories.Open sSQL, Application("DSN"), 3, 1

     if not oDefaultAccessories.eof then
        do while not oDefaultAccessories.eof
           sLineCount = sLineCount + 1

           lcl_default_accessorytype  = "NULL"
           lcl_default_accessoryname  = "NULL"
           lcl_default_accessoryvalue = "NULL"

           if trim(oDefaultAccessories("accessorytype")) <> "" then
              lcl_default_accessorytype = oDefaultAccessories("accessorytype")
              lcl_default_accessorytype = dbsafe(lcl_default_accessorytype)
              lcl_default_accessorytype = "'" & lcl_default_accessorytype & "'"
           end if

           if trim(oDefaultAccessories("accessoryname")) <> "" then
              lcl_default_accessoryname = oDefaultAccessories("accessoryname")
              lcl_default_accessoryname = dbsafe(lcl_default_accessoryname)
              lcl_default_accessoryname = "'" & lcl_default_accessoryname & "'"
           end if

           if trim(oDefaultAccessories("accessoryvalue")) <> "" then
              lcl_default_accessoryvalue = oDefaultAccessories("accessoryvalue")
              lcl_default_accessoryvalue = dbsafe(lcl_default_accessoryvalue)
              lcl_default_accessoryvalue = "'" & lcl_default_accessoryvalue & "'"
           end if

           sSQLi = "INSERT INTO egov_class_teamroster_accessories ("
           sSQLi = sSQLi & "orgid, "
           sSQLi = sSQLi & "accessorytype, "
           sSQLi = sSQLi & "accessoryname, "
           sSQLi = sSQLi & "accessoryvalue, "
           sSQLi = sSQLi & "displayorder "
           sSQLi = sSQLi & ") VALUES ("
           sSQLi = sSQLi & sOrgID                     & ", "
           sSQLi = sSQLi & lcl_default_accessorytype  & ", "
           sSQLi = sSQLi & lcl_default_accessoryname  & ", "
           sSQLi = sSQLi & lcl_default_accessoryvalue & ", "
           sSQLi = sSQLi & oDefaultAccessories("displayorder")
           sSQLi = sSQLi & ") "

           lcl_accessoryid = RunInsertStatement(sSQLi)

           oDefaultAccessories.movenext
        loop
     end if

     oDefaultAccessories.close
     set oDefaultAccessories = nothing
  end if

  if lcl_isAjaxRoutine = "Y" then
     response.write "success"
  end if

end sub

'------------------------------------------------------------------------------
function getTotalAccessories(p_orgid)
  dim lcl_return, lcl_orgid

  lcl_return = 0
  lcl_orgid  = 0

  if p_orgid <> "" then
     lcl_orgid = clng(p_orgid)
  end if

  if lcl_orgid > 0 then
     sSQL = "SELECT COUNT(accessoryid) as total_accessories "
     sSQL = sSQL & " FROM egov_class_teamroster_accessories "
     sSQL = sSQL & " WHERE orgid = " & lcl_orgid

     set oGetTotalAccessories = Server.CreateObject("ADODB.Recordset")
     oGetTotalAccessories.Open sSQL, Application("DSN"), 3, 1

     if not oGetTotalAccessories.eof then
        lcl_return = oGetTotalAccessories("total_accessories")
     end if

     oGetTotalAccessories.close
     set oGetTotalAccessories = nothing
  end if

  getTotalAccessories = lcl_return

end function

'------------------------------------------------------------------------------
sub dtb_debug(p_value)
  sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"

  set oDTB = Server.CreateObject("ADODB.Recordset")
  oDTB.Open sSQL, Application("DSN"), 3, 1
end sub
%>