<!DOCTYPE html>
<html>
<head>
	<link rel="stylesheet" href="font-awesome/css/font-awesome.min.css">
	<style type="text/css">
		#Block
		{
			background: lightblue;
			width: 350px;
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
		#AdressBlock
		{
			position: fixed;
			left:509px;
			top: 175px;
			color: blue;
			font-size: x-large;
		}
		.AddUserPic
		{
			width: 20px;
			height: 30px;
			position-left: 700px;

		}
		.verlinkung
		{
			margin-left: 8.5px;
		}
		#neuerEintrag
		{
			position: fixed;
			left: 510px;
			top: 578px;
			width: 50px;
		}
		#neuerEintrag1
		{
			position: fixed;
			left: 570px;
			top: 600px;
		}
		#SucheDIV
		{
			position: fixed;
			left: 670px;
			top:170px;

		}
		#SeitenAnzahl
		{
			position: fixed;

		}
		#AnzeigeDerSeite
		{
			position: fixed;
			left:570px; 
			top:605px;
		}
		.recht
		{
			position: fixed;
			left:520px;
			top:650px;
			width: 30px;
			height: 30px;
		}
		.links
		{
			position: fixed;
			left:710px;
			top:650px;
			width: 30px;
			height: 30px;

		}
		#SeitenZahlEingabe
		{
			width: 350px;
			height: auto;
		}
		.AnzeigeDerSeite
		{
			position: fixed;
			left:566px;
			top:655px;
		}
		#SeitenAngeber
		{
			position: fixed;
			top: 650px;
			left: 595px;
		}
		#untereKasten
		{
			position: fixed;
			left:500px; 
			top:640px;
			background: lightblue;
			width: 370px;
			height: 60px;
			border-bottom-left-radius: 45px;
			text-align: left;
			font-size: x-large; 
		}
		#alleSeitenDerDB
		{
			color: red;
		}


		
	</style>
	<title>Page number</title>
</head>
<body>
	<script type="text/javascript">
		var Pointer;
	</script>
	<div id="Block">
	<%
	Request("Pointer1")
	
		Set v_Page = Server.CreateObject("ADODB.Connection")

		v_Page.Open "Provider=SQLOLEDB; Server=10.150.2.5; uid=sa; pwd=sl34150; database=extract" 


		sSql = "Select Name, Vorname, Geschlecht, APID From AdressenAP"

		set adoRs = v_Page.Execute(sSql)
		
		'GetRows Retrieves multiple records of a Recordset object into an array.
		PageData = adoRs.GetRows()
		anzahl = 10

		seite = Request("Pointer1")
		start = (seite * anzahl) - anzahl
		start = seite * anzahl
		ende  = start + anzahl -1 

		'For i = 0 to ubound(PageData, 2)
		For i = start to ende
			if PageData(2,i)=0 then
				Response.Write "Herr" & " "
			end if
			if PageData(2,i)=255 then
				Response.Write "Frau" & " "
			end if
			Response.Write Herr
			Response.Write  PageData(0,i) & " "
			Response.Write  PageData(1,i)
			'Add Links'
			Response.Write("<a class=""verlinkung"" href=""DeleteFile/DeleteUser.asp?APID=" & PageData(3,i) & """><image class=""AddUserPic""src=""https://image.flaticon.com/icons/svg/126/126831.svg""></a>")

			Response.Write("<a class=""verlinkung""href=""InfoFile/info.asp?APID=" & PageData(3,i) & """><image class=""AddUserPic""src=""https://image.flaticon.com/icons/svg/1076/1076337.svg""></a>") 

			Response.Write("<a class=""verlinkung""href=""Update/Changer.asp?APID=" & PageData(3,i) & """><image class=""AddUserPic""src=""https://image.flaticon.com/icons/svg/126/126794.svg""></a>")
		  	
		  	Response.Write("<br>")

		Next
		adoRs.Close
		v_Page.Close
		%></div>
		<div id="untereKasten"><div id="SeitenAngeber" class="AnzeigeDerSeite"><%

		Response.Write ("<br>" & "<br>" & Request("Pointer1" ))
		Response.Write (" " & "von" & " "& ende)
		%><div id="alleSeitenDerDB"><%Response.Write v%></div><%

		set adoRs = Nothing
		Set v_Page = Nothing
	%></div><br>
	<!----Seiten Zahleingabe--->
	<div id="untereKasten">
	<div id="SeitenZahlEingabe" class="AnzeigeDerSeite">
	<a class="recht" href="index.asp?Pointer1=<%Response.Write(Request("Pointer1"))-1%>"><image class="recht" src="https://image.flaticon.com/icons/svg/1286/1286897.svg"></a>
		<script src="https://code.jquery.com/jquery-latest.js"></script>
		<form method="get" action="index.asp">
		<input type="text" value="<%
		Response.Write(Request("Pointer1"))
		%>" name="Pointer1">
		</form>
	<a class="links" href="index.asp?Pointer1=<%Response.Write(Request("Pointer1"))+1%>"><image class="links" src="https://image.flaticon.com/icons/svg/1286/1286867.svg"></a>
</div></div>
	<!-----------------------------Close Paging/Homepage-------------------------------------------->

	<div id="neuerEintrag">
		<a href="AddUserFile/index.html" id="ADDBTN"><image src="https://image.flaticon.com/icons/svg/104/104779.svg"></a></image></div><div id="neuerEintrag1">Neuer Eintrag?
	</div>

	<div id="SucheDIV">
		<form method="get" action="SearchFile/ISearch.asp">		
			<input type="text" name="name"> <input type="submit" value="Search">
		</form>
	</div>
	</div>
	<!---Function------>
	<script>
		if (<%Response.Write(start)%>==0){
			$(".recht").hide(); 
		}
		
	</script>
</body>
</html>
