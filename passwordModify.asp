<!--#Include File="conn.asp"-->
<html>
<head>
<title>�޸�����</title>
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
	font-family: "������";
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
	font-family: "������";
	font-color:#D2BADA;
	font-size: large;
	font-style: normal;
	font-weight: bolder;
	text-decoration: blink;
}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="JavaScript">
		//�ú����������пͻ�����֤
		function check_Null(){
			if (document.form1.oldPwd.value==""){
				alert("�����������!");
				return false;
			}
			if (document.form1.pwd1.value==""){
				alert("�����벻��Ϊ��!");
				return false;
			}
			if (document.form1.pwd2.value==""){
				alert("������ȷ�ϲ���Ϊ��!");
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
      <td width="135" height="132"><img src="images/�ʼ�.png" alt="�ʼǱ�����" width="128" height="128"></td>
      <td width="109"><p align="center"> �Ź�,</p>
      <p> ���ǵĻ���</p>        </td>
    </tr>
  </table>
  <form id="form1" name="form1" method="post" action="" onSubmit="JavaScript:return check_Null();">
    <p>
      <label for="userID">�����룺<br>
      </label>
      <input type="password" name="oldPwd" id="oldPwd" />
    </p>
	<p>
    <label for="password">������:<br>
    </label>

	  <input type="password" name="pwd1" id="pwd1" />
    </p>
	<p>
    <label for="password">������ȷ��:<br>
    </label>

	  <input type="password" name="pwd2" id="pwd2" />
    </p>
	<p align="right">
	<input name="��¼" type="submit" value="�ύ" class="STYLE1"/>
    <input name="����" type="reset" id="����" value="����" class="STYLE1"/></p>
</form>

</div>
<div id="Layer1">
<%
dim oldPwd,pwd1,pwd2
oldPwd=request.form("oldPwd")
pwd1=request.form("pwd1")
pwd2=request.form("pwd2")
if pwd1<>pwd2 and oldPwd<>"" then 
response.write"�����������벻һ�£�<br>���������롣"
response.end
end if

dim rs,sql
set rs=server.CreateObject("adodb.recordset")
sql="select ���� from �Źܰ��Ա�� where ѧ��='"&session("userID")&"'"
rs.open sql,conn,1,2
if oldPwd<>"" and oldPwd<>rs("����") then
response.write "��ľ��������벻��ȷ��<br>��ȷ�Ϻ����䡣"
end if 
if oldPwd=rs("����") then 
rs("����")=pwd1
rs.update
rs.close
response.write"�����޸ĳɹ���"
end if


%>
</div>
</body>
</html>


