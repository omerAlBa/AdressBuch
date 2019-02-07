
<%
	DIM PickUpXML
	Dim Body
 
Set v_Page = Server.CreateObject("ADODB.Connection")
			v_Page.Open "Provider=SQLOLEDB; Server=10.150.2.5; uid=sa; pwd=sl34150; database=extract" 

			sSql = "Select Name, Vorname, Geschlecht, APID From AdressenAP"

			set adoRs = v_Page.Execute(sSql)

			'GetRows Retrieves multiple records of a Recordset object into an array.
			PageData = adoRs.GetRows()
			
			For v = 0 to ubound(PageData, 2)
			Next

			anzahl = 10

			seite = Request("Pointer1")
			start = (seite * anzahl) - anzahl
			start = seite * anzahl
			ende  = start + anzahl -1 
			

			For i = start to ende

				if i=0 then
					Body = "[" & vbcrlf
				end if
				if PageData(2,i)=0 then
					Body = Body & "{" & """Geschlecht""" & ":" & """Frau""" & ","
				end if
				if PageData(2,i)=255 then
					Body = Body & "{" & """Geschlecht""" & ":" & """Herr""" & ","
				end if
				Body = Body & """Name""" & ":" & """" & PageData(0,i) & """" &"," 
				Body = Body & """Vorname""" & ":" & """" & PageData(1,i) & """"

				if i< ende then
					body = Body & "}" & "," & vbcrlf
				else
					Body = Body & "}" 
				end if
				if i = ende then
					Body = Body & vbcrlf & "]"
				end if
			

				
			Next
			Response.contenttype="text/javascript"
			Response.write(body)

			adoRs.Close
			v_Page.Close
	
			set adoRs = Nothing
			Set v_Page = Nothing
		%>
		