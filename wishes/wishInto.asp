<!--#Include File="../messageBoard/function.asp"-->
<%
Dim conn,strConn 
Set conn=Server.CreateObject("ADODB.Connection")
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../�Źܰ�����.mdb")
conn.Open strConn 

dim rs,sql,wishTo,wishes
wishTo=request.form("wishTo")
wishes=myDangerEncode(request.form("wishes"))
set rs=server.createobject("adodb.recordset")
sql="select * from ף����"
rs.open sql,conn,1,2
rs.addnew
rs("ף����")=session("userID")
rs("��ף����")=wishTo
rs("ף��ʱ��")=now()
rs("ף������")=wishes
rs.update
rs.close
set rs=nothing
response.redirect"../homepageFrameset/personalPage.asp"
%>
