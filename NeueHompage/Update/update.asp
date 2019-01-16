<html>
<body>


	<%SET o_cnn = Server.CreateObject("ADODB.Connection")
		  o_cnn.open "Provider=SQLOLEDB; Server=10.150.2.5; uid=sa; pwd=sl34150; database=extract"

		  sql = "UPDATE AdressenAP " &_
			"SET Name='" & Request("Name") & "', Vorname='" & Request("Vorname") & 	"'" &_
			"WHERE APID='" & Request("APID") & "'"

			o_cnn.execute sql
			o_cnn.close

			Response.Redirect "../ProjectNr.3/homepage.asp"
	%>

</body>
</html>