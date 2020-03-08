<style>
/*
Nguyen Duc Hoang (Mr)
Youtube channel: https://www.youtube.com/c/nguyenduchoang
Programming tutorial channel: React Native, Swift, Nodejs, Angular
*/
@import url('https://fonts.googleapis.com/css?family=Raleway');
html, body {
	width: 100%;
	height: 100%;
	font-family: 'Raleway';
	font-size: 16px;	
	font-weight: bold;
}
.bg {
	background: url('../images/background.jpg') no-repeat;
	background-size: cover;
	width: 100%;
	height: 100%;
}
.row-container {
	border: 1px solid #fff;
	border-radius: 20px;
	margin-top: 20vh; /* vh= Viewport's height */
	padding: 30px;
	-webkit-box-shadow: 0px 1px 22px 5px rgba(0,0,0,0.75);
	-moz-box-shadow: 0px 1px 22px 5px rgba(0,0,0,0.75);
	box-shadow: 0px 1px 22px 5px rgba(0,0,0,0.75);
}
label, .form-check-label, h1 {
	color: white;	
	text-shadow: 2px 2px 10px #000;
}
.emailError, .passwordError {
	color: red;
	background-color: blanchedalmond;
	padding-left: 10px;
	border-radius: 4px;
	margin-top: 1px;
}
</style>

<!-- Navigation -->
<nav class="navbar navbar-expand-md navbar-light bg-light sticky-top">
	<div class="container-fluid">
		<a class="navbar-branch" href="#">
			<img src="./images/logo.png" height="50">
		</a>
		<button class="navbar-toggler" type="button" data-toggle="collapse" 
			data-target="#navbarResponsive">
			<span class="navbar-toggler-icon"></span>
		</button>
		<div class="collapse navbar-collapse" id="navbarResponsive">
			<ul class="navbar-nav ml-auto">
				<li class="nav-item">
					<a class="nav-link active" href="./item1.asp?table=tai_lieu3">BAT</a>
				</li>
				<li class="nav-item">
					<a class="nav-link" href="./index.asp">Home</a>
				</li>
				
			</ul>
			
		</div>
		<!-- Search form -->
<FORM ACTION="SEARCH.asp" METHOD="GET">
Item Name: <INPUT TYPE="text" NAME="Keyword"> <INPUT TYPE="submit" VALUE=" Find ">
<input type="hidden" id="table" name="table" value="tai_lieu1">
</FORM>

		<div class="dropdown show">
  <a class="btn btn-secondary dropdown-toggle" href="#" role="button" id="dropdownMenuLink" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
    Dropdown link
  </a>

  <div class="dropdown-menu" aria-labelledby="dropdownMenuLink">
    <a class="dropdown-item" href="./SEARCH.asp?table=tai_lieu1">NUT1</a>
    <a class="dropdown-item" href="./SEARCH.asp?table=tai_lieu2">NUT2</a>
	<a class="dropdown-item" href="./SEARCH.asp?link=link4&table=tai_lieu1 tai_lieu2">NUT3</a>
	<a class="dropdown-item" href="./SEARCH.asp?desc=dd&table=tai_lieu1">NUT4</a>
    <a class="dropdown-item" href="#">Something else here</a>
  </div>
	</div>
		<!-- <div class="dropdown show"> -->
  <!-- <a id="thang_bo" class="btn btn-secondary dropdown-toggle" href="#" role="button" id="dropdownMenuLink" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false"> -->
    <!-- Loại tài liệu -->
  <!-- </a> -->

  <!-- <div id="demolist" class="dropdown-menu" aria-labelledby="loại tài liệu"> -->
    <!-- <a class="dropdown-item" value="tai_lieu1" href="#">tài liệu 1</a> -->
    <!-- <a class="dropdown-item" value="tai_lieu2"  href="#">tài liệu 2</a> -->
  <!-- </div> -->
	<!-- </div> -->
	<script>
  $(document).ready(function () {
    $('#demolist a').on('click', function () {
       var value= ($(this).attr('value'));
	  var txt= ($(this).text());
		$("#table").val(value);
		$('#thang_bo').text(txt);
    });
  });
</script>
</nav>