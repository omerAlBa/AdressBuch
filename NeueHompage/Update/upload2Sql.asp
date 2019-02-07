<!--#INCLUDE file="upload.ase"-->
<!--#INCLUDE file="io.ase"-->
	<%
		Dim Uploader, count
		Set Uploader = New Upload
		Uploader.Upload
		' Check if any files were uploaded
If Uploader.Files(0).isMissing Then
    Response.Write "File(s) not uploaded."
Else
    v_ContentType = Uploader.Files(0).ContentType
    If v_ContentType = "text/xml" Then

     	Response.BinaryWrite Uploader.Files(0).data
      Response.Write "<br>"

      Set o_cnn = Server.CreateObject("ADODB.Connection")
          o_cnn.open "Provider=SQLOLEDB; Server=10.150.2.5; uid=sa; pwd=sl34150; database=extract"

      Set objXMLDoc = Server.CreateObject("MSXML2.DOMDocument.3.0")    
          objXMLDoc.async = False    
          objXMLDoc.load Uploader.Files(0).data

      Dim xmlProduct

      For Each xmlProduct In objXMLDoc.documentElement.selectNodes("adress")
          Geschlecht = xmlProduct.selectSingleNode("Geschlecht").text
          Geschlecht = Replace(Geschlecht,"'","''")
          Response.Write (Geschlecht) & " "
          Name = xmlProduct.selectSingleNode("Name").text
          Name = Replace(Name,"'","''")
          Response.Write (Name) & " "
          Vorname = xmlProduct.selectSingleNode("Vorname").text
          Vorname = Replace(Vorname,"'","''")
          Response.Write (Vorname) & " "
          APID = xmlProduct.selectSingleNode("APID").text
          Response.Write (APID) & "<br> "        
                    
          if Trim(APID) & "" <> "" Then
            SET rs=Server.CreateObject("ADODB.recordset")
            rs.Open "SELECT APID FROM AdressenAP WHERE APID='" & APID & "'",o_cnn

            if rs.EOF Then
              sql = "INSERT INTO AdressenAP (Geschlecht, Name, Vorname) " &_
              "values ('" & Geschlecht & "','" & Name & "','" & Vorname & "')"            
            Else
              sql = "UPDATE AdressenAP " &_
              "SET Name='" & Name & "', Vorname='" & Vorname &  "', Geschlecht='" & Geschlecht & "'" &_
              "WHERE APID='" & APID & "'"
            End if
          Else
              sql = "INSERT INTO AdressenAP (Geschlecht, Name, Vorname) " &_
              "values ('" & Geschlecht & "','" & Name & "','" & Vorname & "')"  
          End if
          Response.Write "<br>" & sql
          o_cnn.Execute sql
      Next
     o_cnn.close

     Response.Redirect("../index.asp")
      
    Else
        Response.Write v_ContentType
        Response.Write "Wrong Filetype! Only .xls uploads are supported"
    End If
End If

%>

 