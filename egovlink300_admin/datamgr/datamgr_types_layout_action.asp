<!-- #include file="../includes/common.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
  lcl_orgid     = request("orgid")
  lcl_dm_typeid = request("dm_typeid")
  lcl_layoutid  = request("layoutid")
  lcl_feature   = request("f")

  lcl_redirect_url = "datamgr_types_layout_maint.asp"
  lcl_redirect_url = lcl_redirect_url & "?dm_typeid=" & lcl_dm_typeid
  lcl_redirect_url = lcl_redirect_url & "&layoutid="  & lcl_layoutid
  lcl_redirect_url = lcl_redirect_url & "&f="         & lcl_feature
  lcl_redirect_url = lcl_redirect_url & "&success=SU"

 'Determine which action the user has selected.
 '  options: CHANGE_LAYOUT, SAVE_LAYOUT
  if request("user_action") <> "" then
     lcl_user_action = UCASE(request("user_action"))
  end if

  if lcl_user_action = "CHANGE_LAYOUT" then
     updateDMTypeLayout lcl_dm_typeid, lcl_layoutid
  else

     lcl_total_items = request("totalitems")
     i = 0

     if lcl_total_items > 0 then
        for i = 1 to lcl_total_items

            lcl_dm_sectionid    = request("dm_sectionid_"    & i)
            lcl_sectionid       = request("sectionid_"       & i)
            lcl_sectionlocation = request("sectionlocation_" & i)
            lcl_sectionorder    = request("sectionorder_"    & i)
            lcl_sectionactive   = request("sectionactive_"   & i)

            'dtb_debug("dm_sectionid: [" & lcl_dm_sectionid & "] - sectionid: [" & lcl_sectionid & "] - location: [" & lcl_sectionlocation & "] - order: [" & lcl_sectionorder & "] - active: [" & lcl_sectionactive & "]")
            maintainDMTSection lcl_dm_typeid, lcl_orgid, lcl_dm_sectionid, lcl_sectionid, _
                               lcl_sectionlocation, lcl_sectionorder, lcl_sectionactive

        next
     end if
  end if

  response.redirect lcl_redirect_url
%>