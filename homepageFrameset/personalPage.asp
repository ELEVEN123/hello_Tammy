<!--#Include File="../messageBoard/function.asp"-->
<!--#Include File="ajax.asp"-->
<%
'连接数据库
Dim conn,strConn 
Set conn=Server.CreateObject("ADODB.Connection")
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../信管班纪念册.mdb")
conn.Open strConn 

%>
<html>
<style type="text/css">
<!--
#diarys {
	position:absolute;
	left:32%;
	top:549px;
	width:394px;
	height:536px;
	z-index:5;
}
-->
</style>
<head>

<!--<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
--><title>个人主页</title>
<link rel="stylesheet" href="../class.css">
<script type="text/javascript">
function showClassmate(str)
{
var xmlhttp;    
if (str=="")
  {
  document.getElementById("txtHint").innerHTML="";
  return;
  }
if (window.XMLHttpRequest)
  {// code for IE7+, Firefox, Chrome, Opera, Safari
  xmlhttp=new XMLHttpRequest();
  }
else
  {// code for IE6, IE5
  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
  }
xmlhttp.onreadystatechange=function()
  {
  if (xmlhttp.readyState==4 && xmlhttp.status==200)
    {
    document.getElementById("txtHint").innerHTML=xmlhttp.responseText;
    }
  }
xmlhttp.open("GET","../alumniBook/basicInfor.asp?ID="+str,true);
xmlhttp.send();
}



function showCare()
{
var xmlhttp;
if (window.XMLHttpRequest)
  {// code for IE7+, Firefox, Chrome, Opera, Safari
  xmlhttp=new XMLHttpRequest();
  }
else
  {// code for IE6, IE5
  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
  }
xmlhttp.onreadystatechange=function()
  {
  if (xmlhttp.readyState==4 && xmlhttp.status==200)
    {
    document.getElementById("diarys").innerHTML=xmlhttp.responseText;
    }
  }
xmlhttp.open("GET","../personalDiary/caresSelect.asp",true);

xmlhttp.send();
}

</script>
<style type="text/css">
<!--
#txtHint {
	position:absolute;
	left:14px;
	top:435px;
	width:295px;
	height:208px;
	z-index:1;
}
body {
	background-image: url();
}
#latestMessage {
	position:absolute;
	left:16px;
	top:824px;
	width:298px;
	height:182px;
	z-index:2;
}
#myWishes {
	position:absolute;
	left:67%;
	top:443px;
	width:286px;
	height:250px;
	z-index:4;
}

-->
</style>
</head>

<body>
<div id="latestMessage">
<%
dim rs2,sql2
set rs2=server.createobject("adodb.recordset")
sql2=" select * from 姓名留言表 where 留言时间 between #"& dateadd("ww",-1,now())&"# and #"&now()&"# Order By 留言时间 Desc"
rs2.open sql2,conn,1,2
if rs2.recordcount=0 then response.write "...可惜，这周还没有最新留言<br>快去踩一踩吧，亲..." end if
 
do while not rs2.eof
response.write"<table width='290px' ><tr><td>留言人：</td><td>"&rs2("姓名")&"</td></tr><tr><td>留言时间：</td><td>"&rs2("留言时间")&"</td></tr><tr><td>留言内容：</td><td>"&myHTMLEncode(rs2("留言内容"))&"</td></tr></table><br>"
rs2.movenext
loop
rs2.close
set rs2=nothing
%>
</div>


<div id="diarys">
  <p>..今日的随机同学声色</p>
  <p>将会在这里展现哦...</p>
</div>
<table width="100%" height="1334" border="2" cellpadding="2" cellspacing="2" bordercolor="#D0E180">
  <!--DWLayoutDefaultTable-->
  <tr>
    <td width="298" rowspan="5" valign="top"><form name="formSignature" method="post" action="../personalDiary/diaryInto.asp">
      <p><img src="../images/心情.png" width="82" height="81">今日，你的声色记录：</p>
      <p>
          <textarea name="signature" cols="33" rows="6" id="signature"></textarea>
          <input name="signatureSub" type="submit" id="signature" value="记录">
        </p>
    </form></td>
    <td width="400" height="23" valign="top"></td>
    <td width="328" rowspan="6" valign="top" ><form name="formWish" method="post" action="../wishes/wishInto.asp">
      <p><img src="../images/个人主页按钮.png" width="82" height="81">今日，你的祝福送给：
        <select name="wishTo" id="wishTo">
  <%
dim rs1, sql1
set rs1=server.CreateObject("adodb.recordset")
sql1="select 学号,姓名 from 信管班成员表"
rs1.open sql1,conn,1,2
do while not rs1.eof
response.write"<option value="&rs1("学号")&">"&rs1("姓名")&"</option>"
rs1.movenext
loop

%>
          </select>
        </p>
        <p>
          <textarea name="wishes" cols="33" rows="6" id="wishes"></textarea>
          <input type="submit" name="Submit" value="祝福">
        </p>
    </form></td>
    <td width="12"></td>
  </tr>
  <tr>
    <td height="48"><p>点我点我，知晓飘到你这儿的今日随机声色哦...

     <!--
        <button name="button" type="button" onClick="showCare()">
        随缘<img src="../images/心情1.png" width="82" height="81"></button>-->
		
        <!--<button name="button" type="button" onClick="showCare()">
        随缘</button>-->
		<button type="button" onClick="showCare()">随缘<img src="../images/心情1.png" width="82" height="81"></button></td>
    <td></td>
  </tr>
  <tr>
    <td height="18" valign="top"><!--DWLayoutEmptyCell-->&nbsp;</td>
    <td></td>
  </tr>
  <tr>
    <td height="18">&nbsp;</td>
    <td></td>
  </tr>
  <tr>
    <td rowspan="7" valign="top"><p align="center"><img   src="../images/同学录图标.jpg" alt="同学录" width="80%" height="285"></p></td>
    <td height="2"></td>
  </tr>
  
  
  <tr>
    <td height="18">&nbsp;</td>
    <td></td>
  </tr>
  <tr>
    <td height="18">&nbsp;</td>
    <td >&nbsp;</td>
    <td></td>
  </tr>
  
  
  <tr>
    <td rowspan="3" valign="top">
	<form name="formSelect" method="post" action="">
      
      <p><img src="../images/搜索.png" width="82" height="81">想要看看同学的基本资料？那么：</p>
      <p>选择你要查询的同学名 ：
          <select name="classmateSelect" id="classmateSelect" onChange="showClassmate(this.value)" >
<%

rs1.movefirst
do while not rs1.eof
response.write"<option value="&rs1("学号")&">"&rs1("姓名")&"</option>"
rs1.movenext
loop
rs1.close
set rs1=nothing
%>
          </select>
        </p>
  </form></td>
    <td height="33" valign="top"><img src="../images/最近记录.png" width="82" height="81">看一看近一周，你收获的祝福吧：</td>
    <td></td>
  </tr>
  <tr>
    <td height="11"></td>
    <td></td>
  </tr>
  <tr>
    <td rowspan="8" valign="top"><div id="myWishes">
<%
dim rs3,sql3
set rs3=server.createobject("adodb.recordset")
sql3="select * from 姓名祝福表 where 被祝福人='"&session("userID")&"' and 祝福时间 between #"& dateadd("ww",-1,now())&"# and #"&now()&"# Order By 祝福时间 Desc"
rs3.open sql3,conn,1,2

if rs3.recordcount=0 then response.write "...可惜，这周还没有最新祝福哦...<br>" end if
 
do while not rs3.eof
response.write"<table width='290px' ><tr><td>祝福人：</td><td>"&rs3("姓名")&"</td></tr><tr><td>祝福时间：</td><td>"&rs3("祝福时间")&"</td></tr><tr><td>祝福内容：</td><td>"&myHTMLEncode(rs3("祝福内容"))&"</td></tr></table><br>"
rs3.movenext
loop
rs3.close
set rs3=nothing
%>
</div></td>
    <td height="4"></td>
  </tr>
  <tr>
    <td rowspan="3" valign="top"><div  id="txtHint">同学的基本信息将在此列出...</div></td>
    <td height="67"></td>
  </tr>
  <tr>
    <td height="18">&nbsp;</td>
    <td></td>
  </tr>
  <tr>
    <td rowspan="5" valign="top"><!--DWLayoutEmptyCell-->&nbsp;</td>
    <td height="154"></td>
  </tr>
  <tr>
    <td height="11"></td>
    <td></td>
  </tr>
  <tr>
    <td height="18" valign="top"><p><img src="../images/最近记录.png" width="82" height="81">近一周留言：</p>    </td>
    <td></td>
  </tr>
  <tr>
    <td height="12"></td>
    <td></td>
  </tr>
  <tr>
    <td height="512" valign="top"><!--DWLayoutEmptyCell-->&nbsp;</td>
    <td></td>
  </tr>
</table>
</body>
</html>
