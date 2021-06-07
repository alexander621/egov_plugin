<!--#include file="config.asp"-->
<!--#include file="db.asp"-->
<%
	Class DBActionClass
		private cDB			'Database Class
        private tablefields 'table field columns
		public intColumns  'number of elements in array columns
        public tablename	'after done, I think it should be removed to config.asp 
		public currentpage
		public pagesize
		public TotalPages 
		public totalrecords
		public numstartid
		public numendid
		public strSQL
		private aSorted()
		private irec
		public sub Class_Initialize()
		    set cDB = new DBClass
			cDB.Open sConnForum
 	    tablename=conTABLENAME
		tablefields=conTABLEFIELDS
		intColumns=UBound(conTABLEFIELDS)
		end sub
'---------------------------------------------------		
		public sub Class_Terminate()
			set cDB = Nothing
		end sub
'-------------------------------------------------
public function AddNew
strSQL = "select * from " & tablename & " where "&conTABLEFIELDS(0)&"=0"
set rsform=cDB.GetRS(strSQL)
rsform.AddNew
With request
for i =0 to intColumns
'response.write rsform.fields(i).name&"<br>"
EachName=rsform.fields(i).name
'response.write "<br>i="&i&" fld="&EachName&" value="&.form(EachName)
if l_register(i) then
'response.write "***"
rsform(EachName)=.form(EachName)
end if
next
end with
rsform.Update  'No need to do here, since register is same for the administrator
destroy rsform
end function
'-----------------------------------------------------			
public function saveupdate(id)
strSQL = "select * from " & tablename & " where "&conTABLEFIELDS(IDCol)&"="&id
set rsform=cDB.GetRS(strSQL)
With request
for i =0 to intColumns
'response.write rsform.fields(i).name&"<br>"
EachName=rsform.fields(i).name
'response.write "<br>i="&i&" fld="&EachName&" value="&.form(EachName)
if l_modify(i) then
'response.write "***"
rsform(EachName)=.form(EachName)
end if
next
end with
rsform.update
destroy rsform
end function

'-----------------------------------------------------			
public function getrecordarray
'strSQL has already been set in calling program before calling getrecordarray
set rs=cDB.GetRS(strSQL)
totalrecords = rs.RecordCount
if totalrecords=0 then
response.write "*** 0 ****"
response.end
end if
iFields = rs.Fields.Count
TotalPages = (totalrecords \ pagesize) + 1  '\means integer/integer
			if TotalPages < 1 then TotalPages = 1

if isNumeric(CurrentPage) then
				if CurrentPage < 1 then CurrentPage = 1
				if CurrentPage > TotalPages then CurrentPage = TotalPages
			else
				CurrentPage = 1
			end if
numstartid	= (CurrentPage-1) * PageSize
'response.write "numstartid="&numstartid
numendid	= IIf(numstartid + PageSize < totalrecords, numstartid+pagesize- 1, totalrecords - 1)
'response.write  "numendid:" &numendid&"<br>hello"
getrecordarray=rs.GetRows
destroy rs
end function
'-----------------------------------------------------		
public function deleterecord(id)
' strSQL = "delete from " & tablename &" where "&conTABLEFIELDS(j)&"="&id
strSQL = "delete from " & DeleteTablename &" where "&conTABLEFIELDS(j)&"="&id
cDB.Execute(strSQL)
 end function
end class
'-----------------------------------------------------	
function getid
fileurl=server.mappath("counter.txt")
Set fs = CreateObject("Scripting.FileSystemObject") 
Set a = fs.OpenTextFile(fileurl,1, True) 
counter= Clng(a.ReadLine) 
counter = counter + 1 
a.close 
Set a = fs.CreateTextFile(fileurl,True) 
a.WriteLine(counter) 
a.Close 
Set a=nothing 
Set fs=nothing 
getid=counter
end function

function IIf(bCheck, sTrue, sFalse)
		if bCheck then IIf = sTrue Else IIf = sFalse
	end function
%>