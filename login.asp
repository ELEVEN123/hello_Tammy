<!--#Include File="conn.asp"-->
<html>
<head>
<title>登陆界面</title>
<link rel="stylesheet" href="class.css">
<style>
#LayerForm {
	position:absolute;
	left:61%;
	top:41%;
	width:259px;
	height:328px;
	z-index:1;
	background-color: #D2BADA;
}
body {
	background-image: url(images/%E6%9C%A8%E8%B4%A8404.jpg);
}
#LayerForm p {
	font-family: "新宋体";
	font-size: large;
	font-style: normal;
	font-weight: bolder;
	text-decoration: blink;
}
#Layer1 {
	position:absolute;
	left:61%;
	top:20%;
	width:229px;
	height:232px;
	z-index:2;
}
#Layer1 p {
	font-family: "新宋体";
	font-color:#D2BADA;
	font-size: large;
	font-style: normal;
	font-weight: bolder;
	text-decoration: blink;
}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="JavaScript">
		//该函数用来进行客户端验证
		function check_Null(){
			if (document.form1.userID.value==""){
				alert("用户名不能为空!");
				return false;
			}
			if (document.form1.password.value==""){
				alert("密码不能为空!");
				return false;
			}
			return true;
		}
	</script>
</head>

<body>

<div id="LayerForm">

  <table width="260" border="1">
    <tr>
      <td width="135" height="132"><img src="images/笔记.png" alt="笔记本回忆" width="128" height="128"></td>
      <td width="109"><p align="center"> 信管,</p>
      <p> 我们的回忆</p>        </td>
    </tr>
  </table>
  <form id="form1" name="form1" method="post" action="" onSubmit="JavaScript:return check_Null();">
    <p>
      <label for="userID">用户名：<br>
      </label>
      <input type="text" name="userID" id="userID" />
    </p>
	<p>
    <label for="password">密 码:<br></label>

	  <input type="password" name="password" id="password" />
    </p>
	<p>
	<a href="informationmodify.asp" target="_self" >完善个人信息</a>&nbsp;&nbsp;&nbsp;
	<input name="登录" type="submit" value="提交" class="STYLE1"/>
    <input name="重置" type="reset" id="重置" value="重置" class="STYLE1"/>
</form>

</div>


<div id="Layer1">
<%
dim userID,password
userID=request.form("userID")
password=request.form("password")
dim rs,sql
set rs=server.CreateObject("adodb.recordset")
sql="select 学号,密码 from 信管班成员表"
rs.open sql,conn,1,2
if userID<>"" and password<>"" then 
 do while not rs.eof
   if rs("学号")=userID and rs("密码")=password then
     Session("userID")=userID
     Response.Redirect "homepageFrameset/homepageFrame.html" 
   else 
   rs.movenext
   end if
  loop
if rs.eof then response.write"亲，你的账号不能用耶！可能有以下解释：<br>*你还没有完善你的个人资料哦！<br>*你的账号名与密码没能通过认证！" end if 
end if
%>
</div>
</body>
</html>

