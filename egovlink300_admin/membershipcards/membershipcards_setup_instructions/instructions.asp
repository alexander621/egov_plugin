<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: instructions.asp
' AUTHOR:   David Boyer
' CREATED:  04/30/08
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Allows users to access the required files and documentation needed to setup equipment for membership card use
'
' MODIFICATION HISTORY
' 1.0  09/08/2011  David Boyer - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("memberships") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../../"  'Override of value from common.asp

 if NOT UserHasPermission( session("userid"), "memberships_setup" ) then
    response.redirect sLevel & "permissiondenied.asp"
 end if
%>
<html>
<head>
  <title>E-GovLink {Membership Equipment Setup}</title>
  
	<link rel="stylesheet" type="text/css" href="../../global.css">
	<link rel="stylesheet" type="text/css" href="../../menu/menu_scripts/menu.css" />
 <script language="javascript" src="../../scripts/modules.js"></script>

<style type="text/css">
  #windowsXP,
  #windows7 {
     font-size:   12pt;
     font-weight: bold;
  }
</style>

</head>
<body>
<% ShowHeader sLevel %>
<!--#Include file="../../menu/menu.asp"-->

<div id="content">
    <div id="centercontent">

      <p>
        <fieldset class="fieldset">
          <legend id="windowsXP">Windows XP</legend>
          <span style="color:#ff0000">Download the following files in order to set up the membership equipment properly.</span><br />

          <p>- [<a href="membershipcard_setup_instructions.doc">Membership Card Setup Instructions</a>]</p>
          <p>- [<a href="VSTwain.zip">VSTwain.zip</a>]&nbsp;(contains: <span style="color:#ff0000">vstwain37-setup.exe, VSTwain.dll, piccleanup.bat</span>)</p>
        </fieldset>
      </p>

      <p>
        <fieldset class="fieldset">
          <legend id="windows7">Windows 7</legend>
<!--          <p>- [<a href="membershipcard_setup_instructions.doc">Membership Card Setup Instructions</a>]</p> -->
          <p>1. Pull in Hercules webcam and allow drivers to install.</p>
          <p>2. Download the VintaSoft 'exe' file ---> [<a href="VSTwain52-setup.exe">VSTwain52-setup.exe</a>]</p>
          <p>3. Close browser and re-open it and attempt to take picture.  If you receive an error message like "Cannot save file" then
                in your browser click:<br />
                <ul>
                  <li>Tools -> Internet Options -> Security (tab)</li>
                  <li>Click on "Trusted Sites" where it asks you to select a "zone".</li>
                  <li>Click on the "Sites" button.</li>
                  <li>"http://www.egovlink.com" should be the default website entered, if not enter it.</li>
                  <li>Once the URL has been entered, make sure the "Require server verification (https:) for all sites in this zone" checkbox is UNCHECKED.</li>
                  <li>Click the "Add" button.</li>
                  <li>Click the "Close" button.</li>
                  <li>Click the "OK" button</li>
                  <li>Close the browser and re-open the E-Gov site (admin) and attempt to retake the picture.</li>
                </ul>
          </p>
        </fieldset>
      </p>
    </div>
</div>

<!--#Include file="../../admin_footer.asp"-->  

</body>
</html>
