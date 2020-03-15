
<!-- 
Nguyen Duc Hoang (Mr)
Youtube channel: https://www.youtube.com/c/nguyenduchoang
Programming tutorial channel 
-->
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

sql="SELECT * FROM tai_lieu1 ORDER BY dem DESC"
Response.Write(sql)
rs.CursorLocation = 3 ' adUseClient
rs.Open sql, conn
rs.PageSize = 8
nPageCount = rs.PageCount
nPage = CLng(Request.QueryString("page"))
If nPage < 1 Or nPage > nPageCount Then
	 nPage = 1
End If
%>
<!--#include file="navit.asp"-->
<body>
<!--#include file="nb.asp"-->

<!-- Carousel -->

<!-- jumbotron -->







<!--Carousel Wrapper-->
<div id="multi-item-example" class="carousel slide carousel-multi-item" data-ride="carousel">

  <!--Controls-->
  <div class="controls-top">
    <a class="btn-floating" href="#multi-item-example" data-slide="prev"><i class="fas fa-chevron-left"></i></a>
    <a class="btn-floating" href="#multi-item-example" data-slide="next"><i
        class="fas fa-chevron-right"></i></a>
  </div>
  <!--/.Controls-->

  <!--Indicators-->
  <ol class="carousel-indicators">
    <li data-target="#multi-item-example" data-slide-to="0" class="active"></li>
    <li data-target="#multi-item-example" data-slide-to="1"></li>
    <li data-target="#multi-item-example" data-slide-to="2"></li>
  </ol>
  <!--/.Indicators-->

  <!--Slides-->
			<div class="carousel-inner" role="listbox">
							<%
If rs.RecordCount > 0 then
  rs.AbsolutePage = nPage
  Do While Not ( rs.Eof Or RS.AbsolutePage <> nPage )
  if( rs.AbsolutePosition < 4) then

	%>
				<div class="row text-center">

    <!--First slide-->
							

								<div class="carousel-item active">

									  <div class="col-sm-4">
										<div class="card sm-3">
										  <img class="card-img-top"
											src="https://mdbootstrap.com/img/Photos/Horizontal/Nature/4-col/img%20(34).jpg"
											alt="Card image cap">
										  <div class="card-body">
											<h4 class="card-title"><%Response.Write(rs.Fields(1).Value)%></h4>
											<p class="card-text"><%Response.Write(rs.Fields(2).Value)%></p>
											<a class="btn btn-primary">Button</a>
										  </div>
										</div>
									  </div>
								</div>

							
    <!--/.First slide-->
 <%
							 else
							 %>
				 
    <!--Second slide-->
							<div class="carousel-item active">

									  <div class="col-sm-4">
										<div class="card sm-3">
										  <img class="card-img-top"
											src="https://mdbootstrap.com/img/Photos/Horizontal/Nature/4-col/img%20(34).jpg"
											alt="Card image cap">
										  <div class="card-body">
											<h4 class="card-title"><%Response.Write(rs.Fields(1).Value)%></h4>
											<p class="card-text"><%Response.Write(rs.Fields(2).Value)%></p>
											<a class="btn btn-primary">Button</a>
										  </div>
										</div>
									  </div>
								</div>
    <!--/.Second slide-->

				</div>
    <!--/.Third slide-->
			</div>
			<%
								 end if
								 
										rs.MoveNext
										Loop
									  %>
									  

										<%
									rs.Close

								end if	
									  %>
  </div>
  <!--/.Slides-->


<!--/.Carousel Wrapper-->
	


<script>
/*
Nguyen Duc Hoang (Mr)
Youtube channel: https://www.youtube.com/c/nguyenduchoang
Programming tutorial channel
*/
function isEmail(inputEmail) {
    var regex = /^([a-zA-Z0-9_.+-])+\@(([a-zA-Z0-9-])+\.)+([a-zA-Z0-9]{2,4})+$/;
	return regex.test(inputEmail);
}
function validatePassword(inputPassword) {
	return inputPassword.length > 4;
}

$(document).ready(function(){
    $('#email').change(function(){
        var email = $(this).val().trim();
        // alert(`email = ${JSON.stringify(email)}`)
        if(!isEmail(email)) {
            //Error ?
            $(".emailError").html("Email is not valid format");
        } else {
            $(".emailError").html("");
        }
    });
    $('#password').change(function(){
        var password = $(this).val();	
        if(!validatePassword(password)) {
			$(".passwordError").html("Password must be at least 5 characters");
		} else {
			$(".passwordError").html("");
		}
    });
});
</script>
</body>
</html>	









