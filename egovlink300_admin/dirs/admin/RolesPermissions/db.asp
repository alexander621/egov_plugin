<%
	Class DBClass
		private connOfTemp
		
		public sub Class_Initialize()
			set connOfTemp = Server.CreateObject("ADODB.Connection")
		end sub
		
		public sub Class_Terminate()
			Destroy connOfTemp
		end sub
		
		public sub Open(sConn)
			connOfTemp.Open sConn
		end sub

		public sub Execute(strsqr)
			connOfTemp.Execute strsqr
		end sub
		'Returns a recordset object for a given an sql statement
		public function GetRS(sSQL)
			dim rsOfTemp			
			'Create a temporary recordset object
			set rsOfTemp = Server.CreateObject("ADODB.Recordset")
			set rsOfTemp.ActiveConnection = connOfTemp
			rsOfTemp.CursorLocation = 3 'adUseClient
			rsOfTemp.CursorType = 3 'adOpenStatic
			rsOfTemp.Open sSQL,,, 2 'adCmdTable

			'Return a copy of recordset
			set GetRS = rsOfTemp

			'Remove the reference from the memory but keep the recordset open
			set rsOfTemp = nothing
		end function
		
	End Class

	sub Destroy(obj)
		'A generic object desctruction function
		on error resume next
		select case lcase(typename(obj))
			case "recordset", "connection"
				obj.close
			case "variant()"	'array
				erase obj
		end select
		set obj = nothing
	end sub	
%>