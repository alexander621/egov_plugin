<!-- #include file="../includes/common.asp" //-->
<%
 lcl_orgid           = 0
 lcl_mappoint_typeid = 0
 lcl_sectionid       = 0
 lcl_sectionlocation = "L"
 lcl_sectionorder    = 1

 if request("mappoint_typeid") <> "" then
    if isnumeric(request("mappoint_typeid")) then
       lcl_mappoint_typeid = request("mappoint_typeid")
    end if
 end if

 if request("orgid") <> "" then
    if isnumeric(request("orgid")) then
       lcl_orgid = request("orgid")
    end if
 end if

 if request("sectionid") <> "" then
    if isnumeric(request("sectionid")) then
       lcl_sectionid = request("sectionid")
    end if
 end if

 if request("sectionlocation") <> "" then
    lcl_sectionlocation = ucase(request("sectionlocation"))
 end if

 if request("sectionorder") <> "" then
    if isnumeric(request("sectionorder")) then
       lcl_sectionorder = request("sectionorder")
    end if
 end if

 if request("isAjax") <> "" then
    lcl_isAjax = UCASE(request("isAjax"))
 else
    lcl_isAjax = "N"
 end if

 if lcl_mappoint_typeid > 0 then
    updateMPTSection lcl_orgid, lcl_mappoint_typeid, lcl_sectionid, _
                     lcl_sectionlocation, lcl_sectionorder, lcl_isAjax
 else
    if lcl_isAjax = "Y" then
       response.write "Failed to update section order - Error in AJAX Routine"
    else
       response.write "mappoints_types_maint.asp?mappoint_typeid=" & lcl_mappoint_typeid & "&success=AJAX_ERROR"
    end if
 end if

'------------------------------------------------------------------------------
sub updateMPTSection(iOrgID, iMapPointTypeID, iSectionID, iSectionLocation, iSectionOrder, iIsAjax)

  if iMapPointTypeID <> "" AND iSectionID <> "" then
     if iSectionLocation <> "" then
        sSectionLocation = ucase(iSectionLocation)
        sSectionlocation = "'" & dbsafe(sSectionLocation) & "'"
     else
        sSectionLocation = "'L'"
     end if

     if iSectionOrder <> "" then
        sSectionOrder = iSectionOrder
     else
        sSectionOrder = "1"
     end if

    '1. need to check to see if the section exists for the mappoint type
    '2. if "no" then insert the record
    '3. if "yes" then update the record

     if iSectionID > 0 then
        sSQL = "UPDATE egov_mappoints_types_sections SET "
        sSQL = sSQL & "sectionlocation = " & sSectionLocation & ", "
        sSQL = sSQL & "sectionorder = "    & sSectionOrder
        sSQL = sSQL & " WHERE mp_sectionid = " & iSectionID

       	set oUpdateMPTSection = Server.CreateObject("ADODB.Recordset")
      	 oUpdateMPTSection.Open sSQL, Application("DSN"), 3, 1

        set oUpdateMPTSection = nothing

        lcl_success   = "SU"
        lcl_isAjaxmsg = "Sucessfully Updated"
     else
        lcl_success   = "RSS_ERROR"
        lcl_isAjaxMsg = "ERROR: Failed to send to RSS..."
     end if

  else
     lcl_success   = "RSS_ERROR"
     lcl_isAjaxMsg = "ERROR: Failed to send to RSS..."
  end if

  if iIsAjax = "Y" then
     response.write lcl_isAjaxMsg
  else
     response.redirect "list_faq.asp?faqtype=" & iFAQType & "&success=" & lcl_success
  end if

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