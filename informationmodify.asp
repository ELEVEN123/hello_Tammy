<!--#Include File="conn.asp"-->
<html>
<head>
<title>���Ƹ�����Ϣ</title>
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
	font-family: "������";
	font-size: large;
	font-style: normal;
	font-weight: bolder;
	text-decoration: blink;
}
</style>
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
			if (document.form1.password1.value==""){
				alert("���ٴ�ȷ������!");
				return false;
			}
			if (document.form1.password.value!=document.form1.password1.value){
				alert("���������������һ��!");
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
      <td width="135" height="132"><img src="images/�ʼ�.png" alt="�ʼǱ�����" width="128" height="128"></td>
      <td width="109"><p align="center">�Źܰ��</p>
      <p align="center">���ҡ�</p></td>
	 
    </tr>
  </table>
  <form action="" method="post"  name="form1" id="form1" onSubmit="JavaScript:return check_Null();">
  
<table width="479" border="0" cellpadding="0" cellspacing="0">
  <!--DWLayoutTable-->
  <tr>
    <td width="110" height="55" valign="top">ѧ��
      <input name="userID" type="text" id="userID" value="" size="13"></td>
    <td width="110" valign="top">����
      <input name="userName" type="text" id="userName" size="13"></td>
    <td width="110" valign="top">����
      <input name="password" type="password" id="password" size="13">
      �ٴ�ȷ������
      <input name="password1" type="password" size="13"></td>
    <td width="148" rowspan="7" valign="top">����������
      <textarea name="motto" cols="13" rows="21" id="motto">
      </textarea></td>
  <td width="1">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" rowspan="5" valign="top"><br>
      <br></br></td>
    <td height="56" valign="top">�Ա�<br>
        <input type="radio" name="sex" value="��">
        �� <br>
        
        <input type="radio" name="sex" value="Ů">
        Ů </p>      </td>
    <td>&nbsp;</td>
    </tr>
  <tr>
    <td height="53" valign="top">����<br>
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
      </select>��<br>
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
      </select> ��      </td>
    <td>&nbsp;</td>
    </tr>
  
  
  
  <tr>
    <td height="13">&nbsp;</td>
    <td>&nbsp;</td>
    </tr>
  <tr>
    <td height="44" valign="top">QQ��
      <input name="QQnumber" type="text" id="QQnumber" size="13"></td>
    <td>&nbsp;</td>
    </tr>
  
  
  <tr>
    <td height="45" valign="top">�ֻ���
      <input name="telephoneNumber" type="text" id="telephoneNumber" size="13"></td>
    <td>&nbsp;</td>
    </tr>
  
  
  
  
  <tr>
    <td height="45" valign="top">��ϲ����ż��
      <input name="idol" type="text" id="idol" size="13"></td>
    <td valign="top">��ϲ������ʦ
      <input name="favoriteTeacher" type="text" id="favoriteTeacher" size="13"></td>
    <td valign="top">��ϲ������ɫ
      <input name="favoriteColor" type="text" id="favoriteColor" size="13"></td>
    <td>&nbsp;</td>
    </tr>
  <tr>
    <td height="123" colspan="4" valign="top">��ѧ��ӡ�������һ����      
      <textarea name="deepestMemoery" cols="66" rows="8" id="deepestMemoery"></textarea></td>
    <td>&nbsp;</td>
    </tr>
</table>   
	<p> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	  <input name="��¼" type="submit" value="�ύ" class="STYLE1"/>
	  <input name="����" type="reset" id="����" value="����" class="STYLE1"/>
  </form>
   <%
dim rs,sql
set rs=server.CreateObject("adodb.recordset")
sql="select * from �Źܰ��Ա��"
rs.open sql,conn,1,2
do while not rs.eof
   if rs("ѧ��")=request.form("userID") then
     Response.write "��ͬѧ�������˸�����Ϣ��" 
	 response.end
   else 
   rs.movenext
   end if
loop
if request.form("userID")<>"" then
rs.addnew
rs("ѧ��")=request.form("userID")
rs("����")=request.form("userName")
rs("����")=request.form("password")
rs("����������")=request.form("motto")
rs("�Ա�")=request.form("sex")
rs("����")=request.form("birthMonth")&request.form("birthDay")
rs("QQ��")=request.form("QQnumber")
rs("�ֻ�����")=request.form("telephoneNumber")
rs("��ϲ����ż��")=request.form("idol")
rs("��ϲ������ʦ")=request.form("favoriteTeacher")
rs("��ϲ������ɫ")=request.form("favoriteColor")
rs("��ѧ��ӡ�������һ����")=request.form("deepestMemoery")
rs.update

session("userID")=request.form("userID")
response.redirect "photoUP/photoSelect.asp"
end if
%>
</div>
</body>
</html>
