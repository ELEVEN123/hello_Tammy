
<html>
<head>
<title>ͬѧ��ϸ����</title>
<link rel="stylesheet" href="../class.css">
<link rel="stylesheet" href="css/movingbox.css" type="text/css" media="screen">

<style>
#LayerForm {
	position:absolute;
	left:30%;
	top:10%;
	width:500px;
	height:800px;
	z-index:2;
	background-color: #D2BADA;
	visibility: visible;
}

#LayerForm p {
	font-family: "������";
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
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../�Źܰ�����.mdb")
conn.Open strConn 

dim rs,sql,requestedUserName
requestedUserName=request.querystring("requestedUserName")
set rs=server.CreateObject("adodb.recordset")
sql="select * from �Źܰ��Ա�� where ѧ��='"&requestedUserName&"'"



rs.open sql,conn,1,2

%>

<div id="LayerForm">

  <table width="260" border="1">
    <tr>
      <td width="135" height="132"><img src="../images/�ʼ�.png" alt="�ʼǱ�����" width="128" height="128"></td>
      <td width="109"><p align="center">�Źܰ��</p>
      <p align="center">���ҡ�</p></td>
	 
    </tr>
  </table>
  
  
<table width="479" border="0" cellpadding="0" cellspacing="0">
  <!--DWLayoutTable-->
  <tr>
    <td width="110" height="55" valign="top">ѧ��:<br>  <%=rs("ѧ��")%></td>
    <td width="110" valign="top">����
      :<br>  <%=rs("����")%></td>
    <td width="110" valign="top"></td>
    <td width="148" rowspan="7" valign="top">����������:<br>
        <textarea cols="15" rows="21" style= "background-color:#D2BADA; ">
      <%=rs("����������")%></textarea
      ></td>
  <td width="1">&nbsp;</td>
  </tr>
  <tr>
    <td style= "background-color: #FFE1FF; " colspan="2" rowspan="5" valign="top"><div class="inside"><img src=<%=replace(rs("��Ƭ"),"C:\inetpub\wwwroot\�����","..")%> alt="picture" </div></td>
    <td height="56" valign="top">�Ա�:<br>  <%=rs("�Ա�")%> </td>
    <td>&nbsp;</td>
    </tr>
  <tr>
    <td height="53" valign="top">����:<br>  <%=rs("����")%> </td>
    <td>&nbsp;</td>
    </tr>
  
  
  
  <tr>
    <td height="13">&nbsp;</td>
    <td>&nbsp;</td>
    </tr>
  <tr>
    <td height="44" valign="top">QQ��:<br>  <%=rs("QQ��")%></td>
    <td>&nbsp;</td>
    </tr>
  
  
  <tr>
    <td height="45" valign="top">�ֻ���:<br>  <%=rs("�ֻ�����")%></td>
    <td>&nbsp;</td>
    </tr>
  
  
  
  
  <tr>
    <td height="45" valign="top">��ϲ����ż�� :<br>
       <%=rs("��ϲ����ż��")%></td>
    <td valign="top">��ϲ������ʦ:<br>
        <%=rs("��ϲ������ʦ")%></td>
    <td valign="top">��ϲ������ɫ:<br>
        <%=rs("��ϲ������ɫ")%></td>
    <td>&nbsp;</td>
    </tr>
  <tr>
    <td height="123" colspan="4" valign="top">��ѧ��ӡ�������һ���� :<br>     
      <textarea cols="66" rows="8" style= "background-color:#D2BADA; ">  <%=rs("��ѧ��ӡ�������һ����")%></textarea></td>
    <td>&nbsp;</td>
    </tr>
</table> 

</div>
</body>
</html>

<body> 


</body> 
</html> 

