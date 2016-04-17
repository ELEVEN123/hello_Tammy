<!--#Include File="conn.asp"-->
<html>
<head>
<title>修改密码</title>
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
			if (document.form1.oldPwd.value==""){
				alert("请输入旧密码!");
				return false;
			}
			if (document.form1.pwd1.value==""){
				alert("新密码不能为空!");
				return false;
			}
			if (document.form1.pwd2.value==""){
				alert("新密码确认不能为空!");
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
      <label for="userID">旧密码：<br>
      </label>
      <input type="password" name="oldPwd" id="oldPwd" />
    </p>
	<p>
    <label for="password">新密码:<br>
    </label>

	  <input type="password" name="pwd1" id="pwd1" />
    </p>
	<p>
    <label for="password">新密码确认:<br>
    </label>

	  <input type="password" name="pwd2" id="pwd2" />
    </p>
	<p align="right">
	<input name="登录" type="submit" value="提交" class="STYLE1"/>
    <input name="重置" type="reset" id="重置" value="重置" class="STYLE1"/></p>
</form>

</div>
<div id="Layer1">
<%
dim oldPwd,pwd1,pwd2
oldPwd=request.form("oldPwd")
pwd1=request.form("pwd1")
pwd2=request.form("pwd2")
if pwd1<>pwd2 and oldPwd<>"" then 
response.write"两次密码输入不一致！<br>请重新输入。"
response.end
end if

dim rs,sql
set rs=server.CreateObject("adodb.recordset")
sql="select 密码 from 信管班成员表 where 学号='"&session("userID")&"'"
rs.open sql,conn,1,2
if oldPwd<>"" and oldPwd<>rs("密码") then
response.write "你的旧密码输入不正确！<br>请确认后重输。"
end if 
if oldPwd=rs("密码") then 
rs("密码")=pwd1
rs.update
rs.close
response.write"密码修改成功！"
end if


%>
</div>
</body>
</html>


