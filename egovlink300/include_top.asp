<%
 if bCustomButtonsOn then
    sImgDir = "/custom/images/" & sorgVirtualSiteName & "/"
 else
    sImgDir = "/img/"
 end if

 iStartSecs = timer
 sUserName  = ""

'BEGIN: Javascript for Midland Menu -------------------------------------------
 if iorgid = 29 then
%>
	<!-- JavaScript for the Midland Menu -->
	<script type="text/javascript" src="http://www.midland-mi.org/scripts/mm_menu.js"></script>
	<script type="text/javascript" src="http://www.midland-mi.org/scripts/MM_menu_links.js"></script>
<%
 end if
'END: Javascript for Midland Menu ---------------------------------------------
%>
<!--meta name="viewport" content="width=device-width, initial-scale=1" /-->
<script>
	window.addEventListener('load', function () {
		if (navigator.userAgent.indexOf('Safari') != -1 && navigator.userAgent.indexOf('Chrome') == -1 && navigator.userAgent.indexOf('CriOS') == -1) {
				if ( window.location !== window.parent.location ) {	  
     					document.getElementsByTagName("body")[0].innerHTML =  "<center><input type=\"button\" class=\"reserveformbutton\" style=\"width:auto;text-align:center;\" name=\"continue\" id=\"continueButton\" value=\"Safari Users must click here to continue\" onclick=\"window.open(window.location, '_blank');\" /></center>";
				}
		}
	});
</script>
<link href="https://fonts.googleapis.com/css?family=Open+Sans" rel="stylesheet">
<script type="text/javascript">
<!--
	function SetFocus()
	{
		// Steve Loar 2/17/2006 - To set focus on login form
		var formnames = document.getElementsByName("frmLogin");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmLogin["email"]);
			if(bexists)
			{
				document.frmLogin.email.focus();
			}
		}
		// This will set focus on the family_members.asp page
		var formfamily = document.getElementsByName("addFamily");
		if (formfamily.length == 1)
		{
			var bexists = eval(document.addFamily["firstname"]);
			if(bexists)
			{
				document.addFamily.firstname.focus();
			}
		}
	}

	function HideThings()
	{
		// Steve Loar 2/21/2006 - To hide form selects that block the dropdown menu

		// events/calendar.asp
		var formnames = document.getElementsByName("frmDate");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmDate["month"]);
			if(bexists)
			{
				document.frmDate.month.style.visibility="hidden";
			}
			bexists = eval(document.frmDate["year"]);
			if(bexists)
			{
				document.frmDate.year.style.visibility="hidden";
			}
		}
		// recreation/facility_availability.asp
		formnames = document.getElementsByName("frmcal");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmcal["selfacility"]);
			if(bexists)
			{
				document.frmcal.selfacility.style.visibility="hidden";
			}
		}
	}

	function UnhideThings()
	{
		// Steve Loar 2/21/2006 - To unhide form selects that block the dropdown menu

		// events/calendar.asp
		var formnames = document.getElementsByName("frmDate");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmDate["month"]);
			if(bexists)
			{
				document.frmDate.month.style.visibility="visible";
			}
			bexists = eval(document.frmDate["year"]);
			if(bexists)
			{
				document.frmDate.year.style.visibility="visible";
			}
		}
		// recreation/facility_availability.asp
		formnames = document.getElementsByName("frmcal");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmcal["selfacility"]);
			if(bexists)
			{
				document.frmcal.selfacility.style.visibility="visible";
			}
		}
	}

	function iframecheck()
	{
 		if (window.top!=window.self)
		{
 			document.body.classList.add("iframeformat") // In a Frame or IFrame
 			//var element = document.getElementById("egovhead");
 			//element.classList.add("iframeformat");
		}
	}



 //-->
 </script>
 <style>

	.respHeader
	{
		max-height:145px;
		height:auto;
	}
</style>

<%
'BEGIN: Setup the BODY tag ----------------------------------------------------
 if iorgid = 29 then
    response.write "<body topmargin=""0"" leftmargin=""0"" onLoad=""javascript:MM_preloadImages(linkFiles);SetFocus();"">" & vbcrlf
    response.write "<script language=""JavaScript1.2"" type=""text/javascript"">mmLoadMenus();</script>" & vbcrlf
    response.write "<div style=""margin-left:150px;"">" & vbcrlf
 else
    lcl_onload = "SetFocus();" & lcl_onload

    'response.write "<body topmargin=""0"" leftmargin=""0"" onload=""javascript:SetFocus();"" class=""yui-skin-sam"">" & vbcrlf
    response.write "<body topmargin=""0"" leftmargin=""0"" onload=""" & lcl_onload & """ onunload=""" & lcl_onunload & """ id=""egovhead"" class=""yui-skin-sam""><script>iframecheck();</script>" & vbcrlf
 end if
'END: Setup the BODY tag ------------------------------------------------------
     	if request.servervariables("HTTPS") <> "on" then
	   sNavBaseURL = sEgovWebsiteURL
	else
	   sNavBaseURL = replace(sEgovWebsiteURL,"http://www.egovlink.com","https://secure.egovlink.com")
	end if
 response.write "<div id=""iframenav"" style=""display:none;"">"
 response.write "<div class=""iframenavlink iframenavbutton""><a href=""" & sNavBaseURL & "/rd_classes/class_categories.aspx"">Classes and Events</a></div>"
 response.write "<div class=""iframenavlink iframenavbutton""><a href=""" & sNavBaseURL & "/rentals/rentalcategories.asp"">Rentals</a></div>"
 response.write "<div class=""iframenavlink iframenavbutton""><a href=""" & sNavBaseURL & "/user_login.asp"">Login</a></div>"
 response.write "<div class=""searchMenuDiv"">    "
	 response.write "<div class=""searchBoxText iframenavbutton"" onClick=""expandiframeSearchBox()""><span>Search</span></div>    "
	 response.write "<div class=""searchBox"">      "
		 response.write "<div id=""iframeclassesSearchBox"" class=""classesSearchBox"" onmouseleave=""expandiframeSearchBox()"">"
			 response.write "<input type=""text"" id=""iframetxtsearchphrase"" name=""txtsearchphrase"" class=""txtsearchphrase"" value="""" size=""40"" />        "
			 response.write "<input type=""button"" name=""searchButton"" class=""searchButton"" value=""Find"" onClick=""iframeSearch()"" />      "
		 response.write "</div>    "
	 response.write "</div>  "
 response.write "</div>"
 response.write "</div>"
 response.write "<div id=""footerbug"" style=""display:none;""><a href=""http://www.egovlink.com"" target=""_top"">Powered By EGovLink</a></div>"
if iorgid = 37 or iorgid = 60 then 
	response.write "<script> function clearMsg(id) { if(document.getElementById('msg'+id)) { document.getElementById('msg'+id).style.display = ""none""; } } </script>"
	'response.write "<script type=""text/javascript"" src=""/eclink/rd_scripts/jquery-1.7.2.min.js""></script>"
	response.write "<script type=""text/javascript"" src=""/eclink/rd_scripts/egov_navigation_asp.js""></script>"
end if

'Call check for Google Analytic Code ------------------------------------------
 fnInserGoogleAnalytics( sGoogleAnalyticAccnt )

 response.write "<table cellspacing=""0"" cellpadding=""0"" border=""0"" bordercolor=""#008000"">" & vbcrlf

'BEGIN: Top Graphic -----------------------------------------------------------
	select case iorgid

  'BEGIN: Header row for Carrboro NC, orgid = 43 ------------------------------
	 case 43

					if request.servervariables("HTTPS") <> "on" then
					   sImgBaseURL = sEgovWebsiteURL
					else
					   sImgBaseURL = replace(sEgovWebsiteURL,"http://www.egovlink.com","https://secure.egovlink.com")
					end if

     response.write "  <tr class=""topbanner"">" & vbcrlf
     response.write "      <td height=""" & iHeaderSize & """ width=""100%"" valign=""top"" background=""" & sTopGraphicRighURL & """ bgcolor=""243D53"">" & vbcrlf
     response.write "					     <table width=""100%"" bgcolor=""243D53"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbcrlf
     response.write "						      <tr>" & vbcrlf
     response.write "							         <td>" & vbcrlf
     response.write "								            <table width=""729"" border=""0"" cellspacing=""0"" cellpadding=""0"" bgcolor=""243D53"">" & vbcrlf
     response.write "								              <tr>" & vbcrlf
     response.write "									                 <td><img src=""" & sEgovWebsiteURL & "/custom/images/carrboro/hdr-1.jpg"" width=""280"" height=""52"" /></td>" & vbcrlf
     response.write "									                 <td><img src=""" & sEgovWebsiteURL & "/custom/images/carrboro/hdr-2-srvcs.jpg"" width=""367""  height=""52"" /></td>" & vbcrlf
     response.write "	                							  <td><a href="""  & sHomeWebsiteURL & """><img src=""" & sEgovWebsiteURL & "/custom/images/carrboro/hdr-3a.jpg"" alt=""Go to Home Page"" width=""82"" height=""52"" border=""0"" /></a></td>" & vbcrlf
     response.write "								              </tr>" & vbcrlf
     response.write "								            </table>" & vbcrlf
     response.write "							         </td>" & vbcrlf
     response.write "						      </tr>" & vbcrlf
     response.write "						      <tr>" & vbcrlf
     response.write "							         <td bgcolor=""243D53"">" & vbcrlf
     response.write "								            <table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""729"" bgcolor=""243D53"">" & vbcrlf
     response.write "								              <tr>" & vbcrlf
     response.write "									                 <td><a href=""" & sHomeWebsiteURL & "/gov.htm"" title=""""><img src="""   & sEgovWebsiteURL & "/custom/images/carrboro/hdr-btn-1.jpg"" border=""0"" width=""130"" height=""25"" alt="" /></a></td>" & vbcrlf
     response.write "									                 <td><a href=""" & sHomeWebsiteURL & "/dept.htm"" title=""""><img src="""  & sEgovWebsiteURL & "/custom/images/carrboro/hdr-btn-2.jpg"" border=""0"" width=""95""  height=""25"" alt="" /></a></td>" & vbcrlf
     response.write "					               				  <td><a href=""" & sHomeWebsiteURL & "/srvcs.htm"" title=""""><img src=""" & sEgovWebsiteURL & "/custom/images/carrboro/hdr-btn-3.jpg"" border=""0"" width=""110""  height=""25"" alt="" /></a></td>" & vbcrlf
     response.write "						               			  <td><a href=""" & sHomeWebsiteURL & "/comm.htm"" title=""""><img src="""  & sEgovWebsiteURL & "/custom/images/carrboro/hdr-btn-4.jpg"" border=""0"" width=""80""  height=""25"" alt="" /></a></td>" & vbcrlf
     response.write "						               			  <td><a href=""" & sHomeWebsiteURL & "/msg.htm"" title=""""><img src="""   & sEgovWebsiteURL & "/custom/images/carrboro/hdr-btn-5.jpg"" border=""0"" width=""110""  height=""25"" alt="" /></a></td>" & vbcrlf
     response.write "						               			  <td><a href=""" & sHomeWebsiteURL & "/docs.htm"" title=""""><img src="""  & sEgovWebsiteURL & "/custom/images/carrboro/hdr-btn-6.jpg"" border=""0"" width=""122""  height=""25"" alt="" /></a></td>" & vbcrlf
     response.write "							               		  <td><a href=""" & sHomeWebsiteURL & """ title=""Go to Home Page""><img src=""" & sEgovWebsiteURL & "/custom/images/carrboro/hdr-3b.jpg"" border=""0"" width=""82""  height=""25"" alt=""Go to Home Page"" /></a></td>" & vbcrlf
     response.write "					             		  </tr>" & vbcrlf
     response.write "					             		</table>" & vbcrlf
     response.write "			         				</td>" & vbcrlf
     response.write "						      </tr>" & vbcrlf
     response.write "					     </table>" & vbcrlf
     response.write "				  </td>" & vbcrlf
     response.write "				  <td width=""1"" height=""" & iHeaderSize & """ background=""" & sTopGraphicRighURL & """>" & vbcrlf
     response.write "					     <img src=""" & sImgBaseURL & "/img/clearshim.gif"" border=""0"" width=""1"" height=""" & iHeaderSize & """ />" & vbcrlf
     response.write "				  </td>" & vbcrlf
     response.write "			  </tr>" & vbcrlf

 'BEGIN: Header row for Midland, MI, orgid = 29 -------------------------------
	 case 29

     response.write "<tr>" & vbcrlf
     response.write "				<td>" & vbcrlf
     response.write "					   <script language=""JavaScript1.2"" type=""text/javascript"">MM_outputHeader();</script>" & vbcrlf
     response.write "				</td>" & vbcrlf
     response.write "</tr>" & vbcrlf

 'BEGIN: Header Graphics for most organizations -------------------------------
		case else

     if request.servervariables("HTTPS") <> "on" then
					   sImgBaseURL = sEgovWebsiteURL
					else
					   'sImgBaseURL = replace(sEgovWebsiteURL,"http://www.egovlink.com","https://secure.egovlink.com")
					   sImgBaseURL = replace(sEgovWebsiteURL,"http:","https:")
					end if

     if sTopGraphicLeftURL <> "" then
   			  response.write "<tr class=""topbanner"">" & vbcrlf
	  			  response.write "				<td class=""respHeader"" height=""" & iHeaderSize & """ width=""100%"" valign=""top"" background=""" & sTopGraphicRighURL & """>" & vbcrlf
	       response.write "        <a href=""" & sHomeWebsiteURL & """><img name=""City Logo"" src=""" & sTopGraphicLeftURL & """ border=""0"" alt=""Click here to return to the E-Government Services start page"" /></a>" & vbcrlf
   	    response.write "    </td>" & vbcrlf
	       response.write "    <td class=""respHeader"" width=""1"" height=""" & iHeaderSize & """ background=""" & sTopGraphicRighURL & """>" & vbcrlf
   	    response.write "        <img src=""" & sImgBaseURL & "/img/clearshim.gif"" class=""respHeader"" border=""0"" width=""1"" height=""" & iHeaderSize & """ />" & vbcrlf
	       response.write "    </td>" & vbcrlf
	       response.write "</tr>" & vbcrlf
     end if

 end select
'END: Top Graphic -------------------------------------------------------------

 if blnMenuOn then
    if blnCustomMenu then
    		'This is the Dropdown menu picks
    		'response.write ReadFile(server.mappath( "/" & sorgVirtualSiteName & "/custom_html/custom_menu.asp"))
     		Dim oOrg
     		Set oOrg = New classOrganization

       response.write "<tr class=""topbanner"">" & vbcrlf
       response.write "    <td>" & vbcrlf

      'BEGIN: Menu ------------------------------------------------------------
       response.write "        <div id=""listmenu"">" & vbcrlf
       response.write "        <ul>" & vbcrlf
       response.write "          <li onmouseover=""javascript:HideThings();"" onmouseout=""javascript:UnhideThings();""><a href=""#"">Main Menu</a>" & vbcrlf
       response.write "          <ul>" & vbcrlf

      'City Home (maintained in Org Properites)
       if oOrg.checkMenuOptionEnabled("CITY") then
          lcl_label_city = oOrg.getMenuOptionLabel("CITY")

          'response.write "<li><a href=""" & oOrg.GetOrgURL() & """>" & oOrg.GetOrgDisplayName("homewebsitetag") & "</a></li>" & vbcrlf
          response.write "<li><a href=""" & oOrg.GetOrgURL() & """>" & lcl_label_city & "</a></li>" & vbcrlf
       end if

      'E-Gov Home (maintained in Org Properties)
       if oOrg.checkMenuOptionEnabled("EGOV") then
          lcl_label_egov = oOrg.getMenuOptionLabel("EGOV")

          response.write "<li><a href=""" & oOrg.GetEgovURL() & """>" & lcl_label_egov & "</a></li>" & vbcrlf
       end if

       oOrg.ShowPublicDropDownMenu

       response.write "          </ul>" & vbcrlf
       response.write "				      </li>" & vbcrlf
       response.write "        </ul>" & vbcrlf
       response.write "      </div>" & vbcrlf
      'END: Menu --------------------------------------------------------------

       response.write "    </td>" & vbcrlf
       response.write "</tr>" & vbcrlf

      'BEGIN: Fade Lines ------------------------------------------------------
       response.write "<tr class=""fadeline""><td height=""1"" colspan=""2""></td></tr>" & vbcrlf
      'END: Fade Lines --------------------------------------------------------

       response.write "<tr class=""hdsp""><td>&nbsp;</td></tr>" & vbcrlf

       set oOrg = nothing 

    else  'This is the traditional button based navagation	
      'BEGIN: Button Row ------------------------------------------------------
       response.write "  <tr>" & vbcrlf
       response.write "      <td background=""" & sEgovWebsiteURL & sImgDir & "button_finish.gif"">" & vbcrlf
       response.write "          <table cellspacing=""0"" cellpadding=""0"" height=""24"" border=""0"" bordercolor=""#ff0000"">" & vbcrlf
       response.write "            <tr>" & vbcrlf

      'Main City Website
       response.write "                <td><a href=""" & sHomeWebsiteURL & """><img src=""" & sEgovWebsiteURL & sImgDir & "button_city.gif"" border=""0"" /></a></td>" & vbcrlf
       response.write "                <td><img src=""" & sEgovWebsiteURL & sImgDir & "button_line.gif"" border=""0"" /></td>" & vbcrlf

      'E-Gov Link City Home
       response.write "                <td><a href=""" & sEgovWebsiteURL & "/""><img src=""" & sEgovWebsiteURL & sImgDir & "button_egov.gif"" border=""0"" /></a></td>" & vbcrlf
       response.write "                <td><img src=""" & sEgovWebsiteURL & sImgDir & "button_line.gif"" border=""0"" /></td>" & vbcrlf

      'Action Line TAB
       if blnOrgAction then
          response.write "                <td><a href=""" & sEgovWebsiteURL & "/action.asp""><img src=""" & sEgovWebsiteURL & sImgDir & "button_action.gif"" border=""0"" /></a></td>" & vbcrlf
          response.write "                <td><img src=""" & sEgovWebsiteURL & sImgDir & "button_line.gif"" border=""0"" /></td>" & vbcrlf
       end if

      'Calendar TAB
       if blnOrgCalendar then
          response.write "                <td><a href=""" & sEgovWebsiteURL & "/events/calendar.asp""><img src=""" & sEgovWebsiteURL & sImgDir & "button_calendar.gif"" border=""0"" /></a></td>" & vbcrlf
          response.write "                <td><img src=""" & sEgovWebsiteURL & sImgDir & "button_line.gif"" border=""0"" /></td>" & vbcrlf
       end if

      'Document TAB
       if blnOrgDocument then
          response.write "                <td><a href=""" & sEgovWebsiteURL & "/docs/menu/home.asp""><img src=""" & sEgovWebsiteURL & sImgDir & "button_docs.gif"" border=""0"" /></a></td>" & vbcrlf
          response.write "                <td><img src=""" & sEgovWebsiteURL & sImgDir & "button_line.gif"" border=""0"" /></td>" & vbcrlf
       end if

      'Payment TAB
       if blnOrgPayment then
          response.write "                <td><a href=""" & sEgovWebsiteURL & "/payment.asp""><img src=""" & sEgovWebsiteURL & sImgDir & "button_permits.gif"" border=""0"" /></a></td>" & vbcrlf
          response.write "                <td><img src=""" & sEgovWebsiteURL & sImgDir & "button_line.gif"" border=""0"" /></td>" & vbcrlf
       end if

      'FAQ TAB
       if blnOrgFaq then
          response.write "                <td><a href=""" & sEgovWebsiteURL & "/faq.asp""><img src=""" & sEgovWebsiteURL & sImgDir & "button_faq.gif"" border=""0"" /></a></td>" & vbcrlf
       end if

       response.write "            </tr>" & vbcrlf
       response.write "          </table>" & vbcrlf
       response.write "      </td>" & vbcrlf
       response.write "      <td width=""1"" height=""24"" background=""" & sEgovWebsiteURL & sImgDir & "button_finish.gif""><img src=""" & sEgovWebsiteURL & sImgDir & "clearshim.gif"" border=""0"" width=""1"" height=""24"" /></td>" & vbcrlf
       response.write "  </tr>" & vbcrlf
      'END: Button Row --------------------------------------------------------

      'BEGIN: Button Row Shadow -----------------------------------------------
       response.write "  <tr>" & vbcrlf
       response.write "      <td background=""" & sEgovWebsiteURL & sImgDir & "horiz_shadow_14px.gif""><img src=""" & sEgovWebsiteURL & sImgDir & "horiz_shadow_14px.gif"" border=""0"" height=""14"" /></td>" & vbcrlf
       response.write "      <td width=""1"" height=""14"" background=""" & sEgovWebsiteURL & sImgDir & "horiz_shadow_14px.gif""><img src=""" & sEgovWebsiteURL & sImgDir & "clearshim.gif"" border=""0"" width=""1"" height=""14"" /></td>" & vbcrlf
       response.write "  </tr>" & vbcrlf
      'END: Button Row Shadow -------------------------------------------------
    end if
 else
    'Menu is turned off
 end if

'BEGIN: Main Body Content -----------------------------------------------------
 response.write "  <tr>" & vbrlf
 response.write "      <td valign=""top"" class=""indent20"">" & vbrlf
%>

<!--#include file="include_top_functions.asp"-->

<!--#include file="class/classOrganization.asp"-->
