<%
@codepage=65001
%>
<%
set conn=Server.CreateObject("ADODB.Connection")
conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open "D:/ASP/testsite/tai_lieu_db.mdb"
table = Request.QueryString("table")


groupid=Request.QueryString("groupid")

Keyword = Trim(Request.QueryString("Keyword"))

action = Trim(Request.QueryString("action"))
if ( action="openfile") then
	ten_file = Trim(Request.QueryString("ten_file"))
	link = Trim(Request.QueryString("link"))
	Response.Write("Thưa mẹ con đã lấy đc file cần mở. TÊN=<b>"&ten_file& "<b> LINK="&link &"<br>")

	
	sql="SELECT * FROM tai_lieu1 WHERE name = '"&ten_file&"'" 
	set rs1=Server.CreateObject("ADODB.recordset")
	rs1.CursorLocation = 3 
	rs1.Open sql, conn 
	If rs1.RecordCount <=0 then
		sql="INSERT INTO tai_lieu1 (dem) values(0);"
	else
		sql="UPDATE tai_lieu1 SET `dem`=`dem`+1 WHERE name='"&ten_file&"'"
	end if
	  Response.Write(sql)
	on error resume next
	conn.Execute sql,recaffected
	if err<>0 then
	  Response.Write("No update permissions!" &err.Description&"\n")
	else
	  Response.Write("<h3>" & recaffected & " record added</h3>")
	end if

	'GỌI LỆNH MỞ FILE
end if

Response.Redirect ""&link &""
%>
