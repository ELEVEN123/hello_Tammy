
<html>
<head>
<title>照片选择</title>
<link rel="stylesheet" href="class.css">
<style>
#LayerForm {
	position:absolute;
	left:60%;
	top:20%;
	width:500px;
	height:700px;
	z-index:2;
	background-color: #D2BADA;
	visibility: visible;
}
body {
	background-image: url(../images/%E6%9C%A8%E8%B4%A8404.jpg);
}
#LayerForm p {
	font-family: "新宋体";
	font-size: large;
	font-style: normal;
	font-weight: bolder;
	text-decoration: blink;
	
}
</style>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>

<body>
<%
Dim conn,strConn 
Set conn=Server.CreateObject("ADODB.Connection")
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../信管班纪念册.mdb")
conn.Open strConn 

dim rs,sql
set rs=server.CreateObject("adodb.recordset")
sql="select * from 信管班成员表 where 学号='"&session("userID")&"'"


rs.open sql,conn,1,2

%>

<div id="LayerForm">

  <table width="260" border="1">
    <tr>
      <td width="135" height="132"><img src="../images/笔记.png" alt="笔记本回忆" width="128" height="128"></td>
      <td width="109"><p align="center">信管班的</p>
      <p align="center">“我”</p></td>
	 
    </tr>
  </table>
  
  
<table width="479" border="0" cellpadding="0" cellspacing="0">
  <!--DWLayoutTable-->
  <tr>
    <td width="110" height="55" valign="top">学号:<br>  <%=rs("学号")%></td>
    <td width="110" valign="top">姓名
      :<br>  <%=rs("姓名")%></td>
    <td width="110" valign="top"></td>
    <td width="148" rowspan="7" valign="top">个人座右铭:<br>
        <textarea cols="15" rows="21" style= "background-color:#D2BADA; ">
      <%=rs("个人座右铭")%></textarea
      ></td>
  <td width="1">&nbsp;</td>
  </tr>
  <tr>
    <td style= "background-color: #FFE1FF; " colspan="2" rowspan="5" valign="top"><br>
      <br><br><br>
      <form  name="form" method="post" action="photoUP.asp" enctype="multipart/form-data" >  
  <p>
  <input type="file" name="file1" size=><p>
 <input type="submit" name="Submit" value="上传">
  </p>
</form> <br></br></td>
    <td height="56" valign="top">性别:<br>  <%=rs("性别")%> </td>
    <td>&nbsp;</td>
    </tr>
  <tr>
    <td height="53" valign="top">生日:<br>  <%=rs("生日")%> </td>
    <td>&nbsp;</td>
    </tr>
  
  
  
  <tr>
    <td height="13">&nbsp;</td>
    <td>&nbsp;</td>
    </tr>
  <tr>
    <td height="44" valign="top">QQ号:<br>  <%=rs("QQ号")%></td>
    <td>&nbsp;</td>
    </tr>
  
  
  <tr>
    <td height="45" valign="top">手机号:<br>  <%=rs("手机号码")%></td>
    <td>&nbsp;</td>
    </tr>
  
  
  
  
  <tr>
    <td height="45" valign="top">最喜欢的偶像 :<br>
       <%=rs("最喜欢的偶像")%></td>
    <td valign="top">最喜欢的老师:<br>
        <%=rs("最喜欢的老师")%></td>
    <td valign="top">最喜欢的颜色:<br>
        <%=rs("最喜欢的颜色")%></td>
    <td>&nbsp;</td>
    </tr>
  <tr>
    <td height="123" colspan="4" valign="top">大学里印象最深的一件事 :<br>     
      <textarea cols="66" rows="8" style= "background-color:#D2BADA; ">  <%=rs("大学里印象最深的一件事")%></textarea></td>
    <td>&nbsp;</td>
    </tr>
</table> 

</div>
</body>
</html>

<body> 


</body> 
</html> 
