<!-- #include file="SelectedLanguage.asp" //-->
<!-- #include file="lang.asp" //-->
<%
' THIS FILE CONTAINS CONTAIN TO DISPLAY THE SPECIFIED TABS
Const custLoginRequired	= false				'This determines whether guest user automatically goes past login page
Const custGraphic		= "custom/images/" 	'This is the directory for the custom logo


' DEFAULT TAB CONFIGURATION
custTabVisible	= "YY.YY....YY....."


' DISPLAY OPTIONAL REGISTRATION TAB
If session("orgregistration") Then
	custTabVisible = LEFT(custTabVisible,12) & "Y" & RIGHT(custTabVisible,3)
End If


' DISPLAY OPTIONAL INTERNAL REQUEST TAB
If session("OrgInternalEntry") Then
	custTabVisible = LEFT(custTabVisible,11) & "Y" & RIGHT(custTabVisible,4)
End If


' DISPLAY OPTIONAL RECREATION TAB
'If session("ORGID")=26 Or session("ORGID")=37  Then
If OrgHasFeature( "recreation" ) Then 
	custTabVisible = LEFT(custTabVisible,13) & "Y" & RIGHT(custTabVisible,2)
End If


%>
