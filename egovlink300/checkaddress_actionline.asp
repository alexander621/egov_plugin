<!-- #include file="includes/common.asp" //-->
<!-- #include file="include_top_functions.asp" //-->
<!-- #include file="class/classOrganization.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: checkaddress_actionline.asp
' AUTHOR: Steve Loar
' CREATED: 08/28/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This checks that the passed address is in the loaded address list, is called via AJAX
'
' MODIFICATION HISTORY
' 1.0 08/28/2007	 Steve Loar - INITIAL VERSION
' 1.1	02/05/2008	 Steve Loar - Changed handling of street number to handle none provided
' 1.2 04/10/2008  David Boyer - Modified adderss format
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 dim oActionOrg

 set oActionOrg = New classOrganization

 lcl_streetnumber = ""
 lcl_stname       = ""
 lcl_addresstype  = ""
 lcl_returntype   = "CHECK"

 if request("stnumber") <> "" then
    if IsNumeric(request("stnumber")) then
     		lcl_streetnumber = CLng(request("stnumber"))
  	 else
 		    lcl_streetnumber = dbsafe(request("stnumber"))
  	end if
 else
   	lcl_streetnumber = ""
 end if

 lcl_stname = request("stname")

 if request("addresstype") <> "" then
    if not containsapostrophe(request("addresstype")) then
       lcl_addresstype = request("addresstype")
       lcl_addresstype = ucase(lcl_addresstype)
    end if
 end if

'Determine if we are "CHECKING" to see if the address exists
'or "DISPLAYING" a list of valid addresses
 if request("returntype") <> "" then
    if not containsapostrophe(request("returntype")) then
       lcl_returntype = request("returntype")
       lcl_returntype = ucase(lcl_returntype)
    end if
 end if

'Determine which action to take.
 if lcl_returntype = "DISPLAY_OPTIONS" then
    buildAddressOptions iorgid, _
                 lcl_stname, _
                 lcl_addresstype
 else 'CHECK
    checkAddress iorgid, _
                 lcl_streetnumber, _
                 lcl_stname, _
                 lcl_addresstype
 end if

'------------------------------------------------------------------------------
sub checkAddress(iOrgID, iStreetNumber, iStreetName, iAddressType)

 sAddressType  = ""
 sStreetNumber = ""
 sStreetName   = ""

 if iAddressType <> "" then
    sAddressType = iAddressType
 end if

 if iStreetNumber <> "" then
    sStreetNumber = dbsafe(iStreetNumber)
 end if

 if iStreetName <> "" then
    sStreetName = dbsafe(iStreetName)
 end if

 if sStreetNumber <> "" AND sStreetName <> "" then
    sSQL = "SELECT COUNT(residentaddressid) AS hits "
    sSQL = sSQL & " FROM egov_residentaddresses "
    sSQL = sSQL & " WHERE orgid = " & iOrgID

    if sAddressType = "LARGE" then
       sSQL = sSQL & " AND residentstreetnumber = '" & sStreetNumber & "' "
       sSQL = sSQL & " AND (residentstreetname = '" & sStreetName & "' "
       sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix = '" & sStreetName & "' "
       sSQL = sSQL & " OR residentstreetname + ' ' + streetdirection = '" & sStreetName & "' "
       sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & sStreetName & "' "
       sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname = '" & sStreetName & "' "
       sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix = '" & sStreetName & "' "
       sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetdirection = '" & sStreetName & "' "
       sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & sStreetName & "'"
       sSQL = sSQL & ")"
    else
       sSQL = sSQL & " AND residentaddressid = " & sStreetName
    end if

    set oRs = Server.CreateObject("ADODB.Recordset")
    oRs.Open sSQL, Application("DSN"), 3, 1

    if not oRs.eof then
  	    if CLng(oRs("hits")) > CLng(0) then
	  	      sResults = "FOUND CHECK"
       else
		        sResults = "NOT FOUND"
      	end if
    else
	      sResults = "NOT FOUND"
    end if

    oRs.close
    set oRs = nothing
 else
    sResults = "NOT FOUND"
 end if

 response.write sResults

end sub

'--------------------------------------------------------------------------------------------------
sub buildAddressOptions( iOrgID, iStreetName , iAddressType)
	dim sSql, oAddress, sOption

 lcl_display_options = ""
 sStreetName         = ""
 sAddressType        = ""

 if iStreetName <> "" then
    sStreetName = dbsafe(iStreetName)
 end if

 if iAddressType <> "" then
    sAddressType = UCASE(iAddressType)
 end if

	sSQL = "SELECT DISTINCT residentstreetnumber, residentstreetname, CAST(residentstreetnumber AS INT) AS ordernumb, "
	sSQL = sSQL & " ISNULL(residentstreetprefix,'') AS residentstreetprefix, ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection "
	sSQL = sSQL & " FROM egov_residentaddresses "
 sSQL = sSQL & " WHERE orgid = " & iOrgID

 if sAddressType = "LARGE" then
    sSQL = sSQL & " AND (residentstreetname = '" & sStreetName & "' "
    sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix = '" & sStreetName & "' "
    sSQL = sSQL & " OR residentstreetname + ' ' + streetdirection = '" & sStreetName & "' "
    sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & sStreetName & "' "
    sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname = '" & sStreetName & "' "
    sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix = '" & sStreetName & "' "
    sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetdirection = '" & sStreetName & "' "
    sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & sStreetName & "'"
    sSQL = sSQL & " ) "
 else
    sSQL = sSQL & " AND residentaddressid = " & sStreetName
 end if

 sSQL = sSQL & " AND excludefromactionline = 0 "
	sSQL = sSQL & " ORDER BY 2, 5, 6, 4, 3, 1 "

	set oAddress = Server.CreateObject("ADODB.Recordset")
	oAddress.Open sSQL, Application("DSN"), 0, 1
	
	if NOT oAddress.eof then
    lcl_display_options = lcl_display_options & "<strong>Valid Address Choices </strong><br />" & vbcrlf
		  lcl_display_options = lcl_display_options & "<select id=""stnumber"" name=""stnumber"" size=""10"">" & vbcrlf

    do while NOT oAddress.eof

      'Build the street name
       sOption = buildStreetAddress(oAddress("residentstreetnumber"), oAddress("residentstreetprefix"), oAddress("residentstreetname"), oAddress("streetsuffix"), oAddress("streetdirection"))

    			lcl_display_options = lcl_display_options & "<option value=""" & oAddress("residentstreetnumber") & """ >" & sOption & "</option>" & vbcrlf

    			oAddress.MoveNext
    loop

  		lcl_display_options = lcl_display_options & "</select>" & vbcrlf

	end if

	oAddress.close
	set oAddress = nothing

 response.write lcl_display_options

end sub

'--------------------------------------------------------------------------------------------------
function DBsafe( strDB )

If Not VarType( strDB ) = vbString Then 
 		DBsafe = strDB
Else 
  	DBsafe = Replace( strDB, "'", "''" )
End If 

end function
%>