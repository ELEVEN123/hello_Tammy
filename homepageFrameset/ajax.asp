
<title>ajax�ĺ���</title>
<script type="text/javascript">
function selectInterval()
{
setInterval("selectWish()",10*1000);  //ÿ���10������һ����Ϣ
}
function selectWish(){
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
    document.getElementById("myWishes").innerHTML=xmlhttp.responseText;
    }
  }
xmlhttp.open("GET","../wishes/wishShow.asp,true);
xmlhttp.send();
}
</script>