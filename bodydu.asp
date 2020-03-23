
<!-- Carousel -->
<div class="container">
		<div id="slides" class="carousel slide" data-ride="carousel">
				<ul class="carousel-indicators">
					<li data-target="#slides" data-slide-to="0" class="active"></li>
					<li data-target="#slides" data-slide-to="1"></li>
					<li data-target="#slides" data-slide-to="2"></li>		
					
				</ul>
				<div class="carousel-inner">
					<div class="carousel-item active">
						<img src="./images/cover.jpg  "style="width: 100%" >
						<div class="carousel-caption">
							<h1 class="display-2">BAT</h1>
							<h3>Intrduce BAT</h3>
						<a href="indexint.asp" 
							<button type="button" class="btn btn-outline-light btn-lg">
								VIEW TUTORIALS
							</button></a>
							<a href="./SEARCH.asp?ECUID=1&table=MATERIAL" 
							<button type="button" class="btn btn-primary btn-lg ">Get started</button></a>
						</div>
					</div>
					<div class="carousel-item">
						<img src="./images/868-1920x1080.jpg" style="width: 100%">
						<div class="carousel-caption">
							<h1 class="display-2">MTR</h1>
							<h3>Intrduce MTR</h3>
							<button type="button" class="btn btn-outline-light btn-lg">
								VIEW TUTORIALS
							</button>
							<a href="./SEARCH.asp?ECUID=3&table=MATERIAL" 
							<button type="button" class="btn btn-primary btn-lg ">Get started</button></a>
						</div>
					</div>
					<div class="carousel-item">
						<img src="./images/1018-1920x1080.jpg"style="width: 100%">
						<div class="carousel-caption">
							<h1 class="display-2">INV</h1>
							<h3>Intrduce INV</h3>
							<button type="button" class="btn btn-outline-light btn-lg">
								VIEW TUTORIALS
							</button>
							<a href="./SEARCH.asp?ECUID=2&table=MATERIAL" 
							<button type="button" class="btn btn-primary btn-lg ">Get started</button></a>
						</div>
					</div>
					
				</div>
		</div>
</div>

		
		<!-- <div class="container"> -->
	<!-- <div class="jumbotron"> -->
		<!-- <div class="col-xs-12 col-sm-12 col-md-9 col-lg-9 col-xl-10"> -->
			<!-- <p>This is an example of using Bootstrap to make a responsive Website. This is a tutorial</p> -->
		<!-- </div> -->
		<!-- </div> -->
<div class="container">
	<div class="row">
		<div class="col-md-12 ">
						
									<!-- Nav tabs -->
									
											<nav>
												<div class="nav nav-tabs nav-fill" id="nav-tab" role="tablist">
												  <a class="nav-item nav-link active" id="nav-home-tab" data-toggle="tab" href="#nav-home" role="tab" aria-controls="nav-home" aria-selected="true">BAT</a>
												  <a class="nav-item nav-link" id="nav-profile-tab" data-toggle="tab" href="#nav-profile" role="tab" aria-controls="nav-profile" aria-selected="false">MTR</a>
												  <a class="nav-item nav-link" id="nav-contact-tab" data-toggle="tab" href="#nav-contact" role="tab" aria-controls="nav-contact" aria-selected="false">INV</a>
												</div>
											</nav>
								<!-- Nav tabs -->
							
							<!--slide-->
				<div class="tab-content py-3 px-3 px-sm-0" id="nav-tabContent">
									 <!--id-->
												<div class="tab-pane fade show active" id="nav-home" role="tabpanel" aria-labelledby="nav-home-tab">
														<!--Carousel Wrapper-->
														<div id="multi-item-example1" class="carousel slide carousel-multi-item" data-ride="carousel">

																		<div class="controls-top">
																			<a class="btn-floating" href="#multi-item-example1" data-slide="prev"><i ><<</i></a>
																		
																			<a class="btn-floating" href="#multi-item-example1" data-slide="next"><i >>></i></a>
																		</div>
														<!--/.Controls-->

																	  <ol class="carousel-indicators">
																		<li data-target="#multi-item-example1" data-slide-to="0" class="active"></li>
																		<li data-target="#multi-item-example1" data-slide-to="1"></li>
																		<li data-target="#multi-item-example1" data-slide-to="2"></li>
																		
																	  </ol>
																	<div class="carousel-inner" role="listbox">

																						<%
																						i_dem = 0
																						do until rs.EOF
																						i_dem = i_dem +1
																						if (i_dem mod 4 = 1) then 
																							if i_dem = 1 then%>

																			<div class="carousel-item active">
																					<div class="row">
																						<%else%>
																									<div class="carousel-item">
																											<div class="row">
																										<%end if
																										end if
																										%>
<%
set rsImage1 = server.CreateObject("ADODB.Recordset")
	sql3 ="SELECT MaterialType.Image_DefaultLink FROM 	MATERIAL inner join MaterialType on MaterialType.TypeID=MATERIAL.TypeID where  MaterialType.TypeID = " & rs("TypeID")
	rsImage1.Open sql3,conn,3,3
	
if cstr(rs("ImageLink")) = "" then
	
	link3 = rsImage1("Image_DefaultLink")

else
link3 = rs("ImageLink")
end if

%>
																											<div class="col-md-4 col-lg-3 col-sm-6 "style="height: 40rem;" >
																													<div class="card mb-2 ">
																														<img class="card-img-top " style="height: 8rem;" src="<%=link3%>" alt="Card image cap">
																														<div class="card-body bg-light">
																														<h5 class="card-title cb"><%Response.Write(rs.Fields(1).Value)%></h5>
																														<p class="card-text ca text-muted"><%Response.Write(rs.Fields(4).Value)%></p>
																														<p class="card-text  text-muted">View:<%=rs("dem") %></p>
																														<a href="#"  class="btn btn-primary" onClick="clickOnLink('<%=rs("MaterialID")%>','<%Response.Write( Replace(rs.Fields(3).Value,"\","/")  )%>');">GET</a>
																														</div>
																													</div>
																												</div>
																										<%if (i_dem mod 4 = 0) then %>
																											</div>
																									</div>
																						<%end if
																						rs.MoveNext%>
																						<%loop
																						rs.close
																						
																						%>
																					</div>
																			</div>
														
																		</div><!--inner-->
									
														
												<!--taphome-->
								<!--slide mtr-->
							<div class="tab-pane fade" id="nav-profile" role="tabpanel" aria-labelledby="nav-profile-tab">
									<div class="tab-pane fade show active" id="nav-profile" role="tabpanel" aria-labelledby="nav-profile-tab">
										<!--Carousel Wrapper-->
										<div id="multi-item-example3" class="carousel slide carousel-multi-item" data-ride="carousel">

													<div class="controls-top">
														<a class="btn-floating" href="#multi-item-example3" data-slide="prev"><i ><<</i></a>
													
														<a class="btn-floating" href="#multi-item-example3" data-slide="next"><i >>></i></a>
													</div>
										
												  <ol class="carousel-indicators">
													<li data-target="#multi-item-example3" data-slide-to="0" class="active"></li>
													<li data-target="#multi-item-example3" data-slide-to="1"></li>
													
													
												  </ol>
								<div class="carousel-inner" role="listbox">
															<%
															i_dem1 = 0
															do until rs1.EOF
															i_dem1 = i_dem1 +1
															if (i_dem1 mod 4 = 1) then 
																if i_dem1 = 1 then%>

												<div class="carousel-item active">

														<div class="row">
															<%else%>

															<div class="carousel-item">

																<div class="row">

															<%end if
															end if

															%>
															<%
set rsImage2 = server.CreateObject("ADODB.Recordset")
	sql4 ="SELECT MaterialType.Image_DefaultLink FROM 	MATERIAL inner join MaterialType on MaterialType.TypeID=MATERIAL.TypeID where  MaterialType.TypeID = " & rs1("TypeID")
	rsImage2.Open sql4,conn,3,3
	
if cstr(rs1("ImageLink")) = "" then
	
	link4 = rsImage2("Image_DefaultLink")

else
link4 = rs1("ImageLink")
end if

%>
																		<div class="col-md-4 col-lg-3 col-sm-6 ">
																		<div class="card mb-2 ">
																			<img class="card-img-top " style="height: 8rem;" src="<%=link4%>" alt="Card image cap">
																		
																			<div class="card-body bg-light">
																			<h5 class="card-title cb"><%Response.Write(rs1.Fields(1).Value)%></h5>
																			<p class="card-text ca text-muted"><%Response.Write(rs1.Fields(2).Value)%></p>
																			<p class="card-text  text-muted">View:<%=rs1("dem") %></p>
																			<a href= "#" onClick="clickOnLink('<%=rs1("MaterialID")%>','<%Response.Write( Replace(rs1.Fields(3).Value,"\","/")  )%>');" class="btn btn-primary">GET</a>
																			</div>
																		</div>
																		</div>
															<%if (i_dem1 mod 4 = 0) then %>
																	</div>
															</div>


															<%end if
															rs1.MoveNext%>
															<%loop
															rs1.close
															%>

														</div>
												<!--/.Slides-->
												</div>
									<!--Carousel Wrapper-->

								</div>
												<!--/.Slides-->
							</div>			
									
							
								 
							<div class="tab-pane fade" id="nav-contact" role="tabpanel" aria-labelledby="nav-contact-tab">
								
										<!--Carousel Wrapper-->
										<div id="multi-item-example2" class="carousel slide carousel-multi-item" data-ride="carousel">

										<!--Controls-->
												<div class="controls-top">
													<a class="btn-floating" href="#multi-item-example2" data-slide="prev"><i ><<</i></a>
												
													<a class="btn-floating" href="#multi-item-example2" data-slide="next"><i >>></i></a>
												</div>
										<!--/.Controls-->

								<!--Indicators-->
											  <ol class="carousel-indicators">
												<li data-target="#multi-item-example2" data-slide-to="0" class="active"></li>
												<li data-target="#multi-item-example2" data-slide-to="1"></li>
												
												
											  </ol>
								<!--/.Indicators-->
							
								 
								
								<!--Slides-->
												<div class="carousel-inner" role="listbox">

															<!--First slide-->
														<%
															i_dem2 = 0
															do until rs2.EOF
															i_dem2 = i_dem2 +1
															if (i_dem2 mod 4 = 1) then 
																if i_dem2 = 1 then%>

												<div class="carousel-item active">

														<div class="row">
															<%else%>

															<div class="carousel-item">

																<div class="row">

															<%end if
															end if

															%>
																														<%
																set rsImage3 = server.CreateObject("ADODB.Recordset")
																	sql5 ="SELECT MaterialType.Image_DefaultLink FROM 	MATERIAL inner join MaterialType on MaterialType.TypeID=MATERIAL.TypeID where  MaterialType.TypeID = " & rs2("TypeID")
																	rsImage3.Open sql4,conn,3,3
																	
																if cstr(rs2("ImageLink")) = "" then
																	
																	link5 = rsImage3("Image_DefaultLink")

																else
																link5 = rs2("ImageLink")
																end if

																%>
																		<div class="col-md-4 col-lg-3 col-sm-6 ">
																		<div class="card mb-2  ">
																			<img class="card-img-top " style="height: 8rem;" src="<%=link5%>" alt="Card image cap">
																			
																			<div class="card-body bg-light" >
																			<h5 class="card-title cb"><%Response.Write(rs2.Fields(1).Value)%></h5>
																			<p class="card-text ca text-muted"><%Response.Write(rs2.Fields(2).Value)%></p>
																			<p class="card-text  text-muted">View:<%=rs2("dem") %></p>
																			<a href= "#" onClick="clickOnLink('<%=rs2("MaterialID")%>','<%Response.Write( Replace(rs2.Fields(3).Value,"\","/")  )%>');"class="btn btn-primary">GET</a>
																			</div>
																		</div>
																		</div>
															<%if (i_dem2 mod 4 = 0) then %>
																	</div>
															</div>


															<%end if
															rs2.MoveNext%>
															<%loop
															rs2.close
															conn.close%>

														</div>
												<!--/.Slides-->
												</div>
												</div>
												<!--/.Slides-->
										</div>
								
								
							</div>
				</div><!--content-12-->
			</div>	<!--col-12-->		
	 </div><!--roww-->
</div><!--container-->
