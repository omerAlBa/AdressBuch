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
			
			if ende>v then
				ende=v-1
			end if

			For i = start to ende

				Body= Body & (vbTab & "<adress>" & vbcrlf)
				Body=Body & (vbTab & vbTab & "<Geschlecht>" & PageData(2,i) & "</Geschlecht>" & vbcrlf)
				Body=Body & (vbTab & vbTab & "<Name>" & PageData(0,i) & "</Name>"  & vbcrlf) 
				Body=Body & (vbTab & vbTab & "<Vorname>" & PageData (1,i) & "</Vorname>" & vbcrlf)
				Body=Body &	(vbTab & vbTab & "<APID>" & PageData (3,i) & "</APID>" & vbcrlf)
				Body=Body & (vbTab & "</adress>" & vbcrlf)

			
			Next
			' Content-Disposition:attachment;filename:yourfile.ext

			Response.contenttype="application/octetstream"
			Response.AddHeader "content-disposition", "attachment; filename=addressBook.xml;"
			Response.Write "<?xml version=""1.0"" encoding=""UTF-8""?>"
			PickUpXML = vbcrlf & "<addressBook>"  & vbcrlf &  Body & "</addressBook>"
			Response.Write PickUpXML
			adoRs.Close
			v_Page.Close
	
			set adoRs = Nothing
			Set v_Page = Nothing
		%>