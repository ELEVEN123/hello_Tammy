<!--#Include File="conn.asp"-->
<html>
<head>
<title>完善个人信息</title>
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
	background-image: url(images/%E6%9C%A8%E8%B4%A8404.jpg);
}
#LayerForm p {
	font-family: "新宋体";
	font-size: large;
	font-style: normal;
	font-weight: bolder;
	text-decoration: blink;
}
</style>
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
			if (document.form1.password1.value==""){
				alert("请再次确认密码!");
				return false;
			}
			if (document.form1.password.value!=document.form1.password1.value){
				alert("两次密码必须输入一致!");
				return false;
			}
			return true;
		}
	</script>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>

<body>

<div id="LayerForm">

  <table width="260" border="1">
    <tr>
      <td width="135" height="132"><img src="images/笔记.png" alt="笔记本回忆" width="128" height="128"></td>
      <td width="109"><p align="center">信管班的</p>
      <p align="center">“我”</p></td>
	 
    </tr>
  </table>
  <form action="" method="post"  name="form1" id="form1" onSubmit="JavaScript:return check_Null();">
  
<table width="479" border="0" cellpadding="0" cellspacing="0">
  <!--DWLayoutTable-->
  <tr>
    <td width="110" height="55" valign="top">学号
      <input name="userID" type="text" id="userID" value="" size="13"></td>
    <td width="110" valign="top">姓名
      <input name="userName" type="text" id="userName" size="13"></td>
    <td width="110" valign="top">密码
      <input name="password" type="password" id="password" size="13">
      再次确认密码
      <input name="password1" type="password" size="13"></td>
    <td width="148" rowspan="7" valign="top">个人座右铭
      <textarea name="motto" cols="13" rows="21" id="motto">
      </textarea></td>
  <td width="1">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" rowspan="5" valign="top"><br>
      <br></br></td>
    <td height="56" valign="top">性别<br>
        <input type="radio" name="sex" value="男">
        男 <br>
        
        <input type="radio" name="sex" value="女">
        女 </p>      </td>
    <td>&nbsp;</td>
    </tr>
  <tr>
    <td height="53" valign="top">生日<br>
      <select name="birthMonth" id="birthMonth">
        <option value="01">01</option>
        <option value="02">02</option>
        <option value="03">03</option>
        <option value="04">04</option>
        <option value="05">05</option>
        <option value="06">06</option>
        <option value="07">07</option>
        <option value="08">08</option>
        <option value="09">09</option>
        <option value="10">10</option>
        <option value="11">11</option>
        <option value="12">12</option>
      </select>月<br>
      <select name="birthDay" id="birthDay">
        <option value="01">01</option>
        <option value="02">02</option>
        <option value="03">03</option>
        <option value="04">04</option>
        <option value="05">05</option>
        <option value="06">06</option>
        <option value="07">07</option>
        <option value="08">08</option>
        <option value="09">09</option>
        <option value="10">10</option>
        <option value="11">11</option>
        <option value="12">12</option>
        <option value="13">13</option>
        <option value="14">14</option>
        <option value="15">15</option>
        <option value="16">16</option>
        <option value="17">17</option>
        <option value="18">18</option>
        <option value="19">19</option>
        <option value="20">20</option>
        <option value="21">21</option>
        <option value="22">22</option>
        <option value="23">23</option>
        <option value="24">24</option>
        <option value="25">25</option>
        <option value="26">26</option>
        <option value="27">27</option>
        <option value="28">28</option>
        <option value="29">29</option>
        <option value="30">30</option>
        <option value="31">31</option>
      </select> 日      </td>
    <td>&nbsp;</td>
    </tr>
  
  
  
  <tr>
    <td height="13">&nbsp;</td>
    <td>&nbsp;</td>
    </tr>
  <tr>
    <td height="44" valign="top">QQ号
      <input name="QQnumber" type="text" id="QQnumber" size="13"></td>
    <td>&nbsp;</td>
    </tr>
  
  
  <tr>
    <td height="45" valign="top">手机号
      <input name="telephoneNumber" type="text" id="telephoneNumber" size="13"></td>
    <td>&nbsp;</td>
    </tr>
  
  
  
  
  <tr>
    <td height="45" valign="top">最喜欢的偶像
      <input name="idol" type="text" id="idol" size="13"></td>
    <td valign="top">最喜欢的老师
      <input name="favoriteTeacher" type="text" id="favoriteTeacher" size="13"></td>
    <td valign="top">最喜欢的颜色
      <input name="favoriteColor" type="text" id="favoriteColor" size="13"></td>
    <td>&nbsp;</td>
    </tr>
  <tr>
    <td height="123" colspan="4" valign="top">大学里印象最深的一件事      
      <textarea name="deepestMemoery" cols="66" rows="8" id="deepestMemoery"></textarea></td>
    <td>&nbsp;</td>
    </tr>
</table>   
	<p> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	  <input name="登录" type="submit" value="提交" class="STYLE1"/>
	  <input name="重置" type="reset" id="重置" value="重置" class="STYLE1"/>
  </form>
   <%
dim rs,sql
set rs=server.CreateObject("adodb.recordset")
sql="select * from 信管班成员表"
rs.open sql,conn,1,2
do while not rs.eof
   if rs("学号")=request.form("userID") then
     Response.write "此同学已完善了个人信息！" 
	 response.end
   else 
   rs.movenext
   end if
loop
if request.form("userID")<>"" then
rs.addnew
rs("学号")=request.form("userID")
rs("姓名")=request.form("userName")
rs("密码")=request.form("password")
rs("个人座右铭")=request.form("motto")
rs("性别")=request.form("sex")
rs("生日")=request.form("birthMonth")&request.form("birthDay")
rs("QQ号")=request.form("QQnumber")
rs("手机号码")=request.form("telephoneNumber")
rs("最喜欢的偶像")=request.form("idol")
rs("最喜欢的老师")=request.form("favoriteTeacher")
rs("最喜欢的颜色")=request.form("favoriteColor")
rs("大学里印象最深的一件事")=request.form("deepestMemoery")
rs.update

session("userID")=request.form("userID")
response.redirect "photoUP/photoSelect.asp"
end if
%>
</div>
</body>
</html>
