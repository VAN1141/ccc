
  
<!DOCTYPE html>
<html>
  <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js">
    </script>
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
	
	sql="SELECT * FROM Table1 WHERE name = '"&ten_file&"'" 
	set rs=Server.CreateObject("ADODB.recordset")
	rs.CursorLocation = 3 
	rs.Open sql, conn 
	If rs.RecordCount <=0 then
		sql="INSERT INTO Table1 (name,dem) values('"&ten_file&"',0);"
	else
		sql="UPDATE Table1 SET `dem`=`dem`+1 WHERE name='"&ten_file&"'"
	end if
	  Response.Write(sql)
	on error resume next
	conn.Execute sql,recaffected
	if err<>0 then
	  Response.Write("No update permissions!" &err.Description&"\n")
	else
	  Response.Write("<h3>" & recaffected & " record added</h3>")
	end if
	 link=Replace(link,"\","/")
		Response.Write("<script>window.open('file:"&link&"','',true)</script>") 
end if
 
	
set rs=Server.CreateObject("ADODB.recordset")
'sql="SELECT * from  "&table
' column="name"
' if( table= "tai_lieu2") then
	' column="ten"
' end if
'sql="SELECT * FROM "&table &" WHERE name LIKE '%" & Replace(Keyword, "'", "''") & "%' "
sql="SELECT * FROM tai_lieu1 WHERE name LIKE '%" & Replace(Keyword, "'", "''") & "%'  "
Response.Write(sql)
rs.CursorLocation = 3 ' adUseClient
rs.Open sql, conn
rs.PageSize = 2
nPageCount = rs.PageCount
nPage = CLng(Request.QueryString("page"))
If nPage < 1 Or nPage > nPageCount Then
	 nPage = 1
End If
%>
<head>
<!--#include file="navit.asp"-->
<script>
	function clickOnLink(file,link) {
	  //alert('fuck'+file+"   "+link);
	  $.post( "/updateCount.asp?action=updateCount&ten_file="+file, function( data ) {
		 // alert( data );
		  window.open('file://'+link,'_blank', 'location=yes');
		});
	};
</script>
</head>
<body>
<!--#include file="nb.asp"-->

<%
If rs.RecordCount > 0 then
  rs.AbsolutePage = nPage
  Do While Not ( rs.Eof Or RS.AbsolutePage <> nPage )
	%>
	<div class="card" style="width: 18rem;">
	  <div class="card-body">
		<h5 class="card-title"><%Response.Write(rs.Fields(1).Value)%></h5>
		<h6 class="card-subtitle mb-2 text-muted">Card subtitle</h6>
		<p class="card-text"><%Response.Write(rs.Fields(2).Value)%></p>
		<a href="#" class="card-link"><%Response.Write(rs.Fields(3).Value)%></a>
		<a  href="#" onClick = "clickOnLink('<%Response.Write(rs.Fields(1).Value)%>','<%Response.Write( Replace(rs.Fields(3).Value,"\","/")  )%>');" class="card-link">Mở tài liệu</a>
	  </div>
	</div>

	  <%
		rs.MoveNext
		Loop
	  %>

	<nav aria-label="Page navigation example">
	  <ul class="pagination">
		<li class="page-item"><a class="page-link" href="?page=<%Response.Write(nPage-1)%>">Previous</a></li>
		<li class="page-item"><a class="page-link" href="?page=<%Response.Write(nPage+1)%>">Next</a></li>
	  </ul>
	</nav>
	<%
	rs.Close
end if	
	  %>
</BODY>
</HTML>
