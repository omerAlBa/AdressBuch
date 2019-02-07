<html>
<head>
	<title>Changer</title>
	<style type="text/css">
#Block
		{
			background: lightblue;
			width: 300px;
			height: auto;
			padding-left: 20px;
			color: white;
			position: fixed;
			left: 500px;
			top: 200px;
			padding-bottom: 10px;
			border-top-left-radius: 45px;
			text-align: left;
			font-size: x-large;

		}
		#ChangePicture
		{
			width: 50px;
			height: 50px;
			position: fixed;
			left: 800px;
			top: 150px;
		}
	</style>
</head>
<body>

	<div id="Block">

	<%SET o_cnn = Server.CreateObject("ADODB.Connection")
		  o_cnn.open "Provider=SQLOLEDB; Server=10.150.2.5; uid=sa; pwd=sl34150; database=extract"

		  SET rs=Server.CreateObject("ADODB.recordset")
		  rs.Open "SELECT Name, Vorname, Geschlecht, APID FROM AdressenAP WHERE APID='" & Request("APID") & "'",o_cnn

		  do until rs.EOF
		  	Response.Write(rs.Fields("Name").Value & " " & rs.Fields("Vorname").value & " ")
		  	Response.Write("APID" & ":" & " " & rs.Fields("APID").value)


		  	rs.MoveNext
		  Loop
		  	o_cnn.close
		  	%>
		  	


<form method="get" action="update.asp">
		Gesclecht: <select class="VT" name="Gesclecht">
						<option class="VT" value="0">Herr</option>
						<option class="VT" value="255">Frau</option>
					</select><br> 
		Name:	<input id="NameTEXT" value=<% 
			SET o_cnn = Server.CreateObject("ADODB.Connection")
			  o_cnn.open "Provider=SQLOLEDB; Server=10.150.2.5; uid=sa; pwd=sl34150; database=extract"

			  SET rs=Server.CreateObject("ADODB.recordset")
			  rs.Open "SELECT Name FROM AdressenAP WHERE APID='" & Request("APID") & "'",o_cnn

		Response.Write(rs.Fields("Name").value) 
		%> 
		name="APID"><br>

		Vorname:<input id="VornameTEXT" value=
		<%
		SET o_cnn = Server.CreateObject("ADODB.Connection")
			  o_cnn.open "Provider=SQLOLEDB; Server=10.150.2.5; uid=sa; pwd=sl34150; database=extract"

			  SET rs=Server.CreateObject("ADODB.recordset")
			  rs.Open "SELECT Vorname FROM AdressenAP WHERE APID='" & Request("APID") & "'",o_cnn

		Response.Write(rs.Fields("Vorname").value) 
		%>

		 name="Vorname"><br>

		 <input type="hidden" id="ID" value="<%
		  	SET o_cnn = Server.CreateObject("ADODB.Connection")
		  o_cnn.open "Provider=SQLOLEDB; Server=10.150.2.5; uid=sa; pwd=sl34150; database=extract"
		  SET rs=Server.CreateObject("ADODB.recordset")
		  rs.Open "SELECT Name, Vorname, APID FROM AdressenAP WHERE APID='" & Request("APID") & "'",o_cnn
		  	Response.Write(rs.Fields("APID").value)%>" name="APID">

		<input type="submit" value="Change!" id="Sender">
	</form>
	<image id="ChangePicture" src="https://image.flaticon.com/icons/svg/126/126794.svg"></image>
</body>
<script src="https://code.jquery.com/jquery-latest.js"></script>
</html>