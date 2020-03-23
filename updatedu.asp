
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8">	

<%
	Dim  conn,strSQL1,rsX1
	Set 	conn = server.CreateObject("ADODB.Connection")
			conn.Provider="Microsoft.Jet.OLEDB.4.0"
			'DK0HP.connectionstring = "Provider = Microsoft.Jet.OLEDB.4.0; Data source="&server.MapPath("../ADAS_Database/Databasep.mdb; Jet OLEDB:Database Password="#thoan30"""")		
			'DK0HP.connectionstring = "Provider = Microsoft.ACE.OLEDB.12.0; Data source="&server.MapPath("../ADAS_Database/DataHyouka.accdb")
			conn.CursorLocation = 3
			'DK0HP.ConnectionTimeout = 30 
			'DK0HP.CommandTimeout = 80 
	 on error resume next
	conn.Open server.MapPath("../src/database/Database_manual.mdb")
	
'	If q ="" Then 
'	q=Session("q")
'else
'	Session("q")=q
'End if 
	action=Trim(Request.QueryString("action"))
	if (action="openfile") then
	 'set rsX1=Server.CreateObject("ADODB.recordset")
	ten_file=Trim(Request.QueryString("ten_file"))

	'strURL = Request.QueryString("gr")
	'link1=Trim(Request.QueryString("links"))
	 
	 'set rsX1=Server.CreateObject("ADODB.recordset")
	 strSQL1="SELECT * FROM MATERIAL WHERE MaterialID="&ten_file
	 set rsX1=Server.CreateObject("ADODB.recordset")
	 rsX1.CursorLocation = 3
	 rsX1.Open strSQL1,conn,3,3
	 Response.Write("lấy đc file cần mở. ID=<b>"&ten_file& "<b> LINK="&rsX1("MaterialLink") &"<br>")
	 'If rs.RecordCount <=0 then
		'strSQL1="INSERT INTO MATERIAL (Dem) values(0);"
	'else
		strSQL1="UPDATE MATERIAL SET Dem=Dem+1 WHERE MaterialID="&ten_file
	'end if
	  'Response.Write(strSQL1)
	 conn.Execute strSQL1,recaffected
			
			
			' rsX1("Dem") = rsX1("Dem")+1
			' rsX1.Update
			' rsX1.ReQuery
		
	
	 

	' If rsX1.RecordCount <=0 then
	 ' sql="INSERT INTO Table1 (Name,Tool) values('"&ten_file&"',0);"
	 ' else
	 ' LUOTXEM=rsX1("Tool")+1
	 ' sql="UPDATE Table1 SET Tool="&LUOTXEM&" WHERE Name='"&ten_file&"'"
	 ' end if
	 ' Response.Write(sql)
	 ' on error resume next
	' conn.Execute sql,recaffected
if err<>0 then
	  Response.Write("No update permissions!" &err.Description&"\n")
	 else
	  Response.Write("<h3>" & recaffected & " record added</h3>")
	end if
	'GỌI LỆNH MỞ FILE
 end if
 link1=rsX1.Fields(3).Value
 link2=Replace(link1,"\","/")
 'Response.Write(linkk1)
 'Response.Write("<script>window.open('file:"&link1&"','',true)</script>")
 
%>	
	<a id="link_page" href="file://vn-ntv-fs004ada/NG-DATA/⑥ MC機能テスト/99_マニュアル基準書類/BAT/Hardware/HARN/01.標準書/KD2-74106_3.pdf%>"><%=link1%></a>
	
<%
Response.Write("<script>window.open('file:"&link2&"','',true)</script>")
%>

