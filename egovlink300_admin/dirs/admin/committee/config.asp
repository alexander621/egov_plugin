<%
sConnForum =Application("DSN")
dim conTABLEFIELDS()
dim l_fields()
	conTABLENAME="groups"
TableDescription=array( _ 
"GroupID int PRIMARY KEY","OrgID int not null","groupname [varchar](50)", _
"GroupDescription [varchar](100)","createtime datetime")  
fields_description=fields_description_committee
l_register=array(0,0,1,1,0)
l_display=array(1,0,1,1,0)
l_modify=array(0,0,1,1,0)
const cPageSize =20
'+++++++++++++++++++++++++++++++++++++++++++++++++
%>


<%
redim conTABLEFIELDS(UBound(TableDescription))
redim l_fields(UBound(TableDescription))
For I=0 to UBound(TableDescription)
temp=trim(TableDescription(I))
'response.write "I="&I&temp&space(13)
splittemp = Split(temp)
temp=splittemp(0)
temp=replace(temp,"]","")
temp=replace(temp,"[","")
'response.write temp+"<br>"
'---- conTABLEFIELDS name  ----
conTABLEFIELDS(I)=temp
'----------------------------------
if instr(TableDescription(I)," text") then ' the field is text area
	l_fields(I)=250 '1000 is very large number, just show it is text
elseif instr(TableDescription(I)," int") then
     l_fields(I)=4
elseif instr(TableDescription(I)," datetime") then
     l_fields(I)=20
elseif instr(TableDescription(I)," bit") then
     l_fields(I)=5
else
'---- calculate conTABLEFIELDS ----
	endx=instr(TableDescription(I),")")
	lefts=left(TableDescription(I),endx)
	startx=instr(lefts,"(")
	rights=right(lefts,len(lefts)-startx)
	rights=replace(rights,"(","")
	rights=replace(rights,")","")
'---- calculate fields    ----
'if isempty(trim(rights)) or isnull(rights) then rights="10"
	if isnumeric(rights) then l_fields(I)=clng(rights) else l_fields(I)=0
end if
'------------------------------
Next

	function IIf(bCheck, sTrue, sFalse)
		if bCheck then IIf = sTrue Else IIf = sFalse
	end function
%>

