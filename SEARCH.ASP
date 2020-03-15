<!DOCTYPE html>
<html>

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



set rs=Server.CreateObject("ADODB.recordset")
'sql="SELECT * from  "&table
' column="name"
' if( table= "tai_lieu2") then
	' column="ten"
' end if
'sql="SELECT * FROM "&table &" WHERE name LIKE '%" & Replace(Keyword, "'", "''") & "%' "

sql="SELECT * FROM tai_lieu1 ORDER BY dem ASC WHERE name LIKE '%" & Replace(Keyword, "'", "''") & "%' and groupid ='"&groupid&"' "
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
		<a href="<%Response.Write(rs.Fields(3).Value)%>" class="card-link"><%Response.Write(rs.Fields(3).Value)%></a>
		<a href="./update1.asp?action=openfile&ten_file=<%Response.Write(rs.Fields(1).Value)%>&link=<%Response.Write(rs.Fields(3).Value)%>" class="card-link">Mở tài liệu</a>
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
