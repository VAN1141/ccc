
<!DOCTYPE html>
<html lang="vi">
<%
@codepage=65001
%>
<%
	Dim  conn,sqltext,oConn,sql,rsImage
	Set 	conn = server.CreateObject("ADODB.Connection")
			conn.Provider="Microsoft.Jet.OLEDB.4.0"
			'DK0HP.connectionstring = "Provider = Microsoft.Jet.OLEDB.4.0; Data source="&server.MapPath("../ADAS_Database/Databasep.mdb; Jet OLEDB:Database Password="#thoan30"""")		
			'DK0HP.connectionstring = "Provider = Microsoft.ACE.OLEDB.12.0; Data source="&server.MapPath("../ADAS_Database/DataHyouka.accdb")
			conn.CursorLocation = 3
			'DK0HP.ConnectionTimeout = 30 
			'DK0HP.CommandTimeout = 80 
	 on error resume next
	conn.Open server.MapPath("../src/database/Database_manual.mdb")
	set rs=Server.CreateObject("ADODB.recordset")
	sql="SELECT top 8 *  FROM ( SELECT * FROM MATERIAL  ORDER BY dem DESC ) WHERE ECUID=1"
	rs.CursorLocation = 3 ' adUseClient
rs.Open sql, conn
'rs.PageSize = 4
set rs1=Server.CreateObject("ADODB.recordset")
sql1="SELECT top 8 *  FROM ( SELECT * FROM MATERIAL  ORDER BY dem DESC ) WHERE ECUID=3"
rs1.CursorLocation = 3 ' adUseClient
rs1.Open sql1, conn
set rs2=Server.CreateObject("ADODB.recordset")
sql2="SELECT top 8 *  FROM ( SELECT * FROM MATERIAL  ORDER BY dem DESC ) WHERE ECUID=2"
rs2.CursorLocation = 3 ' adUseClient
rs2.Open sql2, conn
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
	
	<body  >
	<script>
	function clickOnLink(file,link) {
	 //alert('fuck'+file+"   "+link);
	  $.post( "./updateCount.asp?action=updateCount&ten_file="+file, function( data ) {
		 // alert( data );
		  window.open('file://'+link,'_blank', 'location=yes');
		});
	};
</script>
	<script>
	$(function() {
  // ------------------------------------------------------- //
  // Multi Level dropdowns
  // ------------------------------------------------------ //
  $("ul.dropdown-menu [data-toggle='dropdown']").on("click", function(event) {
    event.preventDefault();
    event.stopPropagation();

    $(this).siblings().toggleClass("show");


    if (!$(this).next().hasClass('show')) {
      $(this).parents('.dropdown-menu').first().find('.show').removeClass("show");
    }
    $(this).parents('li.nav-item.dropdown.show').on('hidden.bs.dropdown', function(e) {
      $('.dropdown-submenu .show').removeClass("show");
    });

  });
});
$(".btn-group, .dropdown").hover(
                        function () {
                            $('>.dropdown-menu', this).stop(true, true).fadeIn("fast");
                            $(this).addClass('open');
                        },
                        function () {
                            $('>.dropdown-menu', this).stop(true, true).fadeOut("fast");
                            $(this).removeClass('open');
                        });
	</script>



		<!--#Include file="./includes/navibarDU.asp"-->
	
	<!--#Include file="./includes/bodyDU.asp"-->

	<!--#Include file="./includes/Footer.asp"-->
	</body>
</html>
