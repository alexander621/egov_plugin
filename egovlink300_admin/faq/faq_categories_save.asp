<!--#include file="../includes/common.asp" //-->
<!--#include file="faq_global_functions.asp" //-->
<%
'Add FAQ Category -------------------------------------------------------------
 if CLng(request("total_faq_categories")) = CLng(0) then
    lcl_orgid           = request("orgid0")
    lcl_faqcategoryid   = request("FAQCategoryID0")
    lcl_displayorder    = request("displayorder0")
    lcl_faqcategoryname = request("FAQCategoryName0")

    if request("internalonly0") = "on" then
       lcl_internalonly = 1
    else
       lcl_internalonly = 0
    end if

    if request("faqtype0") <> "" then
       lcl_faqtype = UCASE(request("faqtype0"))
    else
       lcl_faqtype = "FAQ"
    end if

    saveCategory lcl_orgid, lcl_faqtype, lcl_faqcategoryid, lcl_faqcategoryname, lcl_displayorder, lcl_internalonly

    lcl_success = "SA"

'EDIT FAQ Category ------------------------------------------------------------
 else
    for e = 1 to request("total_faq_categories")
        lcl_orgid           = request("orgid"&e)
        lcl_faqcategoryid   = request("FAQCategoryID"&e)
        lcl_displayorder    = request("displayorder"&e)
        lcl_faqcategoryname = request("FAQCategoryName"&e)

        if request("internalonly"&e) = "on" then
           lcl_internalonly = 1
        else
           lcl_internalonly = 0
        end if

        if request("faqtype"&e) <> "" then
           lcl_faqtype = UCASE(request("faqtype"&e))
        else
           lcl_faqtype = "FAQ"
        end if

        saveCategory lcl_orgid, lcl_faqtype, lcl_faqcategoryid, lcl_faqcategoryname, lcl_displayorder, lcl_internalonly
    next

    lcl_success = "SU"

 end if

'Reorder the Categories
 reorderFAQCategories session("orgid"), lcl_faqtype

 response.redirect "faq_categories.asp?faqtype=" & lcl_faqtype & "&success=" & lcl_success

'------------------------------------------------------------------------------
sub saveCategory(iOrgID, iFAQType, iFAQCategoryID, iFAQCategoryName, iDisplayOrder, iInternalOnly)

'Insert new record
	if iFAQCategoryID = "0" then
 		 sSQL = "INSERT INTO faq_categories (orgid, FAQCategoryName, displayorder, internalonly, faqtype) VALUES ("
    sSQL = sSQL &       iOrgID                   & ", "
    sSQL = sSQL & "'" & dbsafe(iFAQCategoryName) & "', "
    sSQL = sSQL &       iDisplayOrder            & ", "
    sSQL = sSQL &       iInternalOnly            & ", "
    sSQL = sSQL & "'" & dbsafe(iFAQType)         & "'"
  		sSQL = sSQL & ")"

 else  'Update record
    sSQL = "UPDATE faq_categories SET "
    sSQL = sSQL & "FAQCategoryName = '" & dbsafe(iFAQCategoryName) & "', "
    sSQL = sSQL & "displayorder = "     & iDisplayOrder            & ", "
    sSQL = sSQL & "internalonly = "     & iInternalOnly            & ", "
    sSQL = sSQL & "faqtype = '"         & dbsafe(iFAQType)         & "'"
    sSQL = sSQL & " WHERE faqcategoryid = " & iFAQCategoryID

 end if

	set oSaveCat = Server.CreateObject("ADODB.Recordset")
 oSaveCat.Open sSQL, Application("DSN") , 3, 1

 set oSaveCat = nothing

end sub
%>