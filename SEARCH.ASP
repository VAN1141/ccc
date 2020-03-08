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
link = Request.QueryString("link")
desc=Request.QueryString("desc")
Keyword = Trim(Request.QueryString("Keyword"))
set rs=Server.CreateObject("ADODB.recordset")
'sql="SELECT * from  "&table
' column="name"
' if( table= "tai_lieu2") then
	' column="ten"
' end if
' sql="SELECT * FROM "&table &" WHERE " &column& " LIKE '%" & Replace(Keyword, "'", "''") & "%'"


sql="SELECT * FROM "&table&" WHERE name LIKE '%" & Replace(Keyword, "'", "''") & "%' and link like '%"&link&"%' and desc like '%"&desc&"%' "
Response.Write(sql)
rs.Open sql, conn
' nPageCount = RS.PageCount
' nPage = CLng(Request.QueryString("Page"))
' If nPage < 1 Or nPage > nPageCount Then
	' nPage = 1

' End If
%>
<head>
<!--#include file="navit.asp"-->

</head>
<body>
<!--#include file="nb.asp"-->
  <%
  do until rs.EOF
  %>

<div class="card" style="width: 18rem;">
  <div class="card-body">
    <h5 class="card-title"><%Response.Write(rs.Fields(1).Value)%></h5>
    <h6 class="card-subtitle mb-2 text-muted">Card subtitle</h6>
    <p class="card-text"><%Response.Write(rs.Fields(2).Value)%></p>
    <a href="#" class="card-link"><%Response.Write(rs.Fields(3).Value)%></a>
    <a href="#" class="card-link">Another link</a>
  </div>
</div>
    <%
    rs.MoveNext%>



  <%loop
  rs.close
  conn.close%>
</BODY>
</HTML>