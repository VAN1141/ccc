


<!DOCTYPE html>
<html>

<%
@codepage=65001
%>

<head>
<meta charset="utf-8">
	<meta name="viewport" content="width=device-width, initial-scale=1">
	
     <meta http-equiv="X-UA-Compatible" content="IE=edge">

	<title>Tra cứu tài liệu</title>
	<!-- Import Boostrap css, js, font awesome here -->
	
<script type="text/javascript" src="./js/i18n/i18next-jquery.min.js"></script>
<script type="text/javascript" src="./js/i18n/i18next.min.js"></script>
<script type="text/javascript" src="./js/jquery/jquery-3.4.1.min.js"></script>
<script type="text/javascript" src="./js/popper/popper.js"></script>
<script type="text/javascript" src="./js/bootstrap/bootstrap.js"></script>
    <link rel="stylesheet" href="./css/bootstrap/bootstrap.min.css">
	<link rel="stylesheet" href="./css/fontawesome/all.css">
	<!-- <link rel="stylesheet" href="./css/style.css"> -->
    <link href="./css/search-style.css" rel="stylesheet">
	<!--<link href="./css/styletest.css" rel="stylesheet">-->
	<link rel="stylesheet" href="./css/seach-style.css">
</head>
<body style="background-color: rgba(222, 226, 230, 1);">
<!--#Include file="./includes/navibarDU.asp"-->
<div class="container bg-light">
 
            <!-- ' <h1>Kết quả tìm được</h1> -->
			<!-- ' </div> -->
<%
	Dim  conn,sqltext,oConn,sql,rsImage
	dim arrcha(20)
	Set 	conn = server.CreateObject("ADODB.Connection")
			conn.Provider="Microsoft.Jet.OLEDB.4.0"
			'DK0HP.connectionstring = "Provider = Microsoft.Jet.OLEDB.4.0; Data source="&server.MapPath("../ADAS_Database/Databasep.mdb; Jet OLEDB:Database Password="#thoan30"""")		
			'DK0HP.connectionstring = "Provider = Microsoft.ACE.OLEDB.12.0; Data source="&server.MapPath("../ADAS_Database/DataHyouka.accdb")
			conn.CursorLocation = 3
			'DK0HP.ConnectionTimeout = 30 
			'DK0HP.CommandTimeout = 80 
	 on error resume next
	conn.Open server.MapPath("../src/database/Database_manual.mdb")
%>

<%
	
	q = trim(request.QueryString("q"))
	TypeID=Request.QueryString("TypeID")
	ECUID=Request.QueryString("ECUID")
	GroupID=Request.QueryString("GroupID")
	PartID=Request.QueryString("PartID")
'	If q ="" Then 
'	q=Session("q")
'else
'	Session("q")=q
'End if 
	
	set rsX = server.CreateObject("ADODB.Recordset")
	if q <> "" then
	'bien tat ca ky tu nhap vao thanh dau +
	
	q=Replace(q,"ㇿ","□")
	q=Replace(q,"ロ","□")
	sqlstr = "SELECT MATERIAL.*, Department.Department  FROM MATERIAL inner join Department on MATERIAL.DepartmentID=Department.DepartmentID  WHERE MATERIAL.MaterialName LIKE '%"& q &"%' or MATERIAL.ManualOutline LIKE '%"& q &"%' or MATERIAL.MaterialKeyword LIKE '%"& q &"%' "
	' arrcha(0)="【"
	' arrcha(1)="、"
	
	' for i=0 to 1
	' q=Replace(q,arrcha(i),"*")
	' next
	
	 q=Replace(q," ","*")
	 q=Replace(q,"+","*")
	 'q=Replace(q,"【","*",2)
	q=Replace(q,"】","*")
	 q=Replace(q,"、","*")
	 q=Replace(q,",","*")
	 q=Replace(q,".","*")
	 q=Replace(q,"/","*")
	 'q=Replace(q,"(","*")
	 'q=Replace(q,")","*")
	 'q=Replace(q,"_","*")
	 q=Replace(q,"-","*")
	 q=Replace(q,"　","*")
	 q=Replace(q,"（","*")
	 q=Replace(q,"＿","*")
	 'q=Replace(q,"）","*")
	
	searchquery = split(q,"*")
	
	for each part in searchquery
	
		sqlstr = sqlstr &  " or MATERIAL.MaterialName LIKE '%"& part &"%' or MATERIAL.ManualOutline LIKE '%"& part &"%' or MATERIAL.MaterialKeyword LIKE '%"& part &"%' "
	next
	
	sqlstr = sqlstr & " ORDER BY dem DESC"
	'Response.Write(sqlstr)
else
		if ECUID <>"" then
			if GroupID <> "" then
				if PartID <> "" then
					if TypeID <> "" then
						sqlstr = "SELECT MATERIAL.*, Department.Department  FROM MATERIAL inner join Department on MATERIAL.DepartmentID=Department.DepartmentID  WHERE TypeID = "&TypeID&" and GroupID ="&GroupID&" and PartID = "&PartID&" and ECUID = "&ECUID&" ORDER BY dem DESC"
					else 
						sqlstr = "SELECT MATERIAL.*, Department.Department  FROM MATERIAL inner join Department on MATERIAL.DepartmentID=Department.DepartmentID  WHERE GroupID ="&GroupID&" and PartID = "&PartID&" and ECUID = "&ECUID&" ORDER BY dem DESC"
					end if
				else
					sqlstr = "SELECT MATERIAL.*, Department.Department  FROM MATERIAL inner join Department on MATERIAL.DepartmentID=Department.DepartmentID  WHERE GroupID ="&GroupID&" and ECUID = "&ECUID&" ORDER BY dem DESC"
				end if
			else
				sqlstr = "SELECT MATERIAL.*, Department.Department  FROM MATERIAL inner join Department on MATERIAL.DepartmentID=Department.DepartmentID  WHERE ECUID = "&ECUID&" ORDER BY dem DESC"
			end if
		else
			
		end if
	end if
	on error resume next
	rsX.Open sqlstr,conn,3,3



%>
<br>
<br>
<br>
<br>
<%
 rsX.PageSize = 10
 nPageCount = rsX.PageCount
 nPage = CLng(Request.QueryString("Page"))
 If nPage < 1 Or nPage > nPageCount Then
	nPage = 1
End If


%>

<div class="container  text-white "style="
    background-color: #0062cc">

<p class="p-3 mt-9">Hiển thị kết quả <%=(nPage - 1)*rsX.PageSize + 1%> - <%=nPage*rsX.PageSize%> trong tổng số <%=rsX.recordcount%> kết quả</p>
 
 
      </div>
<script>
	function clickOnLink(file,link) {
	 //alert('fuck'+file+"   "+link);
	  $.post( "./updateCount.asp?action=updateCount&ten_file="+file, function( data ) {
		 // alert( data );
		  window.open('file://'+link,'_blank', 'location=yes');
		});
	};
</script>
<%
'Response.Write(sql)
'Response.Write(sqlstr)
'Response.Write(rsImage("Image_DefaultLink"))
'Response.Write("lấy đc file cần mở. TÊN=<b>"&ten_file& "<b> LINK="&link1 &"<br>")
	rsX.AbsolutePage = nPage
	if rsX.recordcount > 0 then
		 Do While Not ( rsX.Eof Or rsX.AbsolutePage <> nPage )
%>
<style>
.border-5 {
     border-width:1px !important;
}
</style>

<%
set rsImage = server.CreateObject("ADODB.Recordset")
	sql ="SELECT MaterialType.Image_DefaultLink FROM 	MATERIAL inner join MaterialType on MaterialType.TypeID=MATERIAL.TypeID where  MaterialType.TypeID = " & rsX("TypeID")
	rsImage.Open sql,conn,3,3
	
if cstr(rsX("ImageLink")) = "" then
	
	link = rsImage("Image_DefaultLink")

else
link = rsX("ImageLink")
end if

%>
	<!-- <div class="p-3 mb-2 bg-danger text-white"></div> -->
<div class="container">
	

		<div class="row no-gutters border-top border-secondary">
			<div class="col-2">
			
				<img class="card-img-top pull-left px-2 py-3" style="height: 9rem;" src="<%=link%>" alt="Card image cap">
			</div>
			<div class="col">
					  <div class="card-block  px-3 py-3 bg-light" style="margin-bottom: -30px;">
				<h5 class="card-title" ><%=rsX("MaterialName") %></h5>
				<p class="card-subtitle mb-2 text-muted"><%=rsX("Department") %></p>
				<p class="card-text text-muted "><%=rsX("ManualOutline") %></p>
				
				<a href="javascript:void();" onClick="clickOnLink('<%=rsX("MaterialID")%>','<%Response.Write( Replace(rsX.Fields(3).Value,"\","/")  )%>');"class="card-link"><%=rsX("MaterialLink") %></a>
				<p class="card-text  text-muted text-right">View:<%=rsX("dem") %></p>
				
			
				
					</div>
				
					
			</div>
			
	 
		</div>
 </div>
 
  <br>
 

  
<%  
rsImage.Movenext
	rsX.Movenext
	Loop
	rsX.close
	rsImage.close
	else
%>
	<div class="p-3 mb-2 bg-danger text-white">No result is found!</div>
<%	
	
	end if
	
%>

<script>
function next() {
  var x = window.location.href;
  x = x.slice(x.indexOf("?")+1);
  var vt = x.indexOf("page");
  if (vt > 0) {
    x = x.substring(0, vt-1);
  };
  x = "?" + x + "&page=<%Response.Write(nPage+1)%>";
  return x;
};
function Previous() {
  var x = window.location.href;
  x = x.slice(x.indexOf("?")+1);
  var vt = x.indexOf("page");
  if (vt > 0) {
    x = x.substring(0, vt-1);
  };
  x = "?" + x + "&page=<%Response.Write(nPage-1)%>";
  return x;
};
function cr() {
  var x = window.location.href;
  x = x.slice(x.indexOf("?")+1);
  var vt = x.indexOf("page");
  if (vt > 0) {
    x = x.substring(0, vt-1);
  };
  x = "?" + x + "&page=<%Response.Write(nPageCount)%>";
  return x;
};
function cra() {
  var x = window.location.href;
  x = x.slice(x.indexOf("?")+1);
  var vt = x.indexOf("page");
  if (vt > 0) {
    x = x.substring(0, vt-1);
  };
  x = "?" + x + "&page=<%Response.Write(0)%>";
  return x;
};
function cra1() {
  var x = window.location.href;
  x = x.slice(x.indexOf("?")+1);
  var vt = x.indexOf("page");
  if (vt > 0) {
    x = x.substring(0, vt-1);
  };
  x = "?" + x + "&page=<%Response.Write(0)%>";
  return x;
};
function cra2() {
  var x = window.location.href;
  x = x.slice(x.indexOf("?")+1);
  var vt = x.indexOf("page");
  if (vt > 0) {
    x = x.substring(0, vt-1);
  };
  x = "?" + x + "&page=<%Response.Write(2)%>";
  return x;
};
function cra3() {
  var x = window.location.href;
  x = x.slice(x.indexOf("?")+1);
  var vt = x.indexOf("page");
  if (vt > 0) {
    x = x.substring(0, vt-1);
  };
  x = "?" + x + "&page=<%Response.Write(3)%>";
  return x;
};
</script>

<div class="d-flex justify-content-center">
<ul class="pagination ">
    <li class="page-item "><a class="page-link" onClick = "document.location=cra();"><<</a></li>
    <li class="page-item"><a class="page-link" onClick = "document.location=Previous();">Previous</a></li>
	 <li class="page-item"><a class="page-link" onClick = "document.location=cra1();"> 1 </a></li>
	 <li class="page-item"><a class="page-link" onClick = "document.location=cra2();"> 2 </a></li>
	 <li class="page-item"><a class="page-link" onClick = "document.location=cra3();"> 3 </a></li>
    <li class="page-item"><a class="page-link" onClick = "document.location=next();"> Next </a></li>
		<li class="page-item"><a class="page-link" onClick = "document.location=cr();">>></a></li>
	
</ul>
</div>



<!-- <div id="right"> -->
<!-- <nav aria-label="..."> -->
  <!-- <ul class="pagination page-custom"> -->
    <!-- <li class="page-item disabled"> -->
      <!-- <span class="page-link">Previous</span> -->
    <!-- </li> -->
    <!-- <li class="page-item"><a class="page-link" href="#">1</a></li> -->
    <!-- <li class="page-item active"> -->
      <!-- <span class="page-link"> 2 -->
        <!-- <span class="sr-only">(current)</span> -->
      <!-- </span> -->
    <!-- </li> -->
    <!-- <li class="page-item"><a class="page-link" href="#">3</a></li> -->
    <!-- <li class="page-item"> -->
      <!-- <a class="page-link" href="#">Next</a> -->
    <!-- </li> -->
  <!-- </ul> -->
<!-- </nav> -->
<!-- </div> -->
</div>
<!--#Include file="./includes/Footer.asp"-->
</body>

</HTML>
