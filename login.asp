<!--#Include File="conn.asp"-->
<html>
<head>
<title>��½����</title>
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
			if (document.form1.userID.value==""){
				alert("�û�������Ϊ��!");
				return false;
			}
			if (document.form1.password.value==""){
				alert("���벻��Ϊ��!");
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
      <label for="userID">�û�����<br>
      </label>
      <input type="text" name="userID" id="userID" />
    </p>
	<p>
    <label for="password">�� ��:<br></label>

	  <input type="password" name="password" id="password" />
    </p>
	<p>
	<a href="informationmodify.asp" target="_self" >���Ƹ�����Ϣ</a>&nbsp;&nbsp;&nbsp;
	<input name="��¼" type="submit" value="�ύ" class="STYLE1"/>
    <input name="����" type="reset" id="����" value="����" class="STYLE1"/>
</form>

</div>


<div id="Layer1">
<%
dim userID,password
userID=request.form("userID")
password=request.form("password")
dim rs,sql
set rs=server.CreateObject("adodb.recordset")
sql="select ѧ��,���� from �Źܰ��Ա��"
rs.open sql,conn,1,2
if userID<>"" and password<>"" then 
 do while not rs.eof
   if rs("ѧ��")=userID and rs("����")=password then
     Session("userID")=userID
     Response.Redirect "homepageFrameset/homepageFrame.html" 
   else 
   rs.movenext
   end if
  loop
if rs.eof then response.write"�ף�����˺Ų�����Ү�����������½��ͣ�<br>*�㻹û��������ĸ�������Ŷ��<br>*����˺���������û��ͨ����֤��" end if 
end if
%>
</div>
</body>
</html>

