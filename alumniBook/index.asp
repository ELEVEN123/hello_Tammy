<!doctype html>
<html lang="en">
<head>
	<title>ͬѧ¼</title>
  <link rel="stylesheet" href="styles.css" type="text/css" media="screen" />
  <link rel="stylesheet" href="css/movingbox.css" type="text/css" media="screen">
  <style type="text/css">
  body {
	background-color: #E6E6E6;
}
 
  </style>
  <script src="js/jquery-1.3.1.min.js" type="text/javascript" charset="utf-8"></script>
	<script src="js/slider.js" type="text/javascript" charset="utf-8"></script>
	<script type='text/javascript' src='http://ajax.googleapis.com/ajax/libs/jquery/1.4/jquery.min.js?ver=1.3.2'></script>
        <script type='text/javascript' src='js/jquery.color-RGBa-patch.js'></script>
        <script type='text/javascript' src='js/example.js'></script>
<meta charset="gb2312">
</head>
<body>

<%
Dim conn,strConn 
Set conn=Server.CreateObject("ADODB.Connection")
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../�Źܰ�����.mdb")
conn.Open strConn 

dim rs, sql, id
id=1
set rs=server.CreateObject("adodb.recordset")
sql="select ѧ��,����,��Ƭ from �Źܰ��Ա��"
rs.open sql,conn,1,2
%>
<section id="main"><!-- #main content area -->
	  <p>&nbsp;</p>
	  
<p align="center">Ԥ֪���飬����"�����ҡ�"��<br></p>
	  <div id="slider">    
     <img class="scrollButtons left" src="images/leftarrow.png">
  	  <div style="overflow: hidden;" class="scroll">
	
				<div class="scrollContainer">
<%
	 do while not rs.eof 
	    response.write"<div class='panel' id='panel_"&id&"'><div class='inside'><h1 align='center'><a href='detail.asp?requestedUserName="&rs("ѧ��")&"' target='_blank'>�����ҡ�</a></h1><img src='"&replace(rs("��Ƭ"),"C:\inetpub\wwwroot\�����","..")&"' alt='picture' /></div></div>"
		rs.movenext
		id=id+1
	  loop
%>
       
                                        
				
         </div>

				<div id="left-shadow"></div>
				<div id="right-shadow"></div>

       </div>

			<img class="scrollButtons right" src="images/rightarrow.png">

     </div>
        

	
</div><!-- #wrapper -->
<hr>
<div id="detail"></div>
</body>
</html>
