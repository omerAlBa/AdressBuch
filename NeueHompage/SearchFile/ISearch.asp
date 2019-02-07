<!DOCTYPE html>
<html>
<link rel="stylesheet" href="font-awesome/css/font-awesome.min.css">
	<style type="text/css">
	#Block
		{
			background: lightblue;
			width: 260px;
			max-height: 350px;
			padding-left: 20px;
			color: white;
			position: fixed;
			left: 500px;
			top: 200px;
			padding-bottom: 10px;
			padding-top: 10px;
			border-top-left-radius: 45px;
			text-align: left;
			font-size: x-large;
			overflow: scroll;
		}
		#searchPicture
		{
			width: 80%;
			height: 60%;
		}
		.verlinkung
		{
			margin-left: 8.5px;
			width: 20px;
		}
		a
		{
			position: fixed;
			width: 20px;	
		}
		#HomeBTN
		{
			width: 70px;
			height: 70px;
		}

	</style>
<body>
	<div id="Block">
	<%
		Dim Zaehler
		Zaehler=0

		SET o_cnn = Server.CreateObject("ADODB.Connection")
		  o_cnn.open "Provider=SQLOLEDB; Server=10.150.2.5; uid=sa; pwd=sl34150; database=extract"

		  SET rs=Server.CreateObject("ADODB.recordset")
		  rs.Open "SELECT Name, Vorname, APID FROM AdressenAP WHERE Name LIKE '%" & Request("name") & "%'", o_cnn

		  do until rs.EOF
		  	Response.Write(rs.Fields("Name").Value & " " & rs.Fields("Vorname").value)
		  		Response.Write("<a class=""verlinkung""href=""../Update/Changer.asp?APID=" & rs.Fields("APID") & """><image class=""verlinkung""src=""https://image.flaticon.com/icons/svg/126/126794.svg""></a>" & "<br>")
		  Zaehler=Zaehler +1

		  	rs.MoveNext
		  	
		  Loop
		  Response.Write("Hits:" & " " & Zaehler)

		  if Zaehler<1 then
		  	Response.Write  "<br>" & "No Result!"
		  end if
		  	o_cnn.close
		%>
	</div>
	<a href="../index.asp"><image id="HomeBTN" src="https://image.flaticon.com/icons/svg/1371/1371153.svg"></a>
	<image id="searchPicture" src="https://image.flaticon.com/icons/svg/148/148834.svg"></image>
	
</body>
</html>
