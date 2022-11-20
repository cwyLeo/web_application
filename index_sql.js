var username;
var pwd;
var database;
var objdbConn = new ActiveXObject("ADODB.Connection");
var mes = manyValues();
function manyValues(){
    var url=window.location.search;
    if(url.indexOf("?")!=-1){
        var str = url.substr(1);
        strs = str.split("&"); 
        var key=new Array(strs.length);
        var value=new Array(strs.length);
        for(i=0;i<strs.length;i++){
            key[i]=strs[i].split("=")[0]
            value[i]=unescape(strs[i].split("=")[1]); 
        }
    }
    username = value[0];
    pwd = value[1];
    database = value[2];
    var strdsn = "Driver={SQL Server};SERVER=(local);UID="+value[0]+";PWD="+value[1]+";DATABASE="+value[2]; 
    objdbConn.Open(strdsn);
};
var a1 = document.getElementById("a1");
var a2 = document.getElementById("a2");
var a3 = document.getElementById("a3");
var a4 = document.getElementById("a4");
a1.href = "a.html?username="+username+"&pwd="+pwd+"&database="+database;
a2.href = "b.html?username="+username+"&pwd="+pwd+"&database="+database;
a3.href = "c.html?username="+username+"&pwd="+pwd+"&database="+database;
a4.href = "d.html?username="+username+"&pwd="+pwd+"&database="+database;
let totop = document.querySelector('.totop')
// 页面滚动窗口监听事件
window.onscroll = function() {
    // 获取浏览器卷去的高度
    let high = document.documentElement.scrollTop || document.body.scrollTop
    if (high >= 100) {
        totop.style.display = 'block'
    } else {
        totop.style.display = 'none'
    }
}
var options = "";
var sqls = "select dept_name from department";
var add = objdbConn.Execute(sqls);
var len = add.Fields.Count - 1;
while(!add.EOF)
{
    for(var i = 0;i <= len;i++){
        options += "<option value=\""+add.Fields(i).Value+"\">"+add.Fields(i).Value+"</option>";
    }
    add.moveNext();
}
document.getElementById("keywords").innerHTML = options;
function query1(){
    var stu_id = document.getElementById("stu_id").value;
    // var table_name = document.getElementById("table_name").value;
    var sql = "select student.name,student.dept_name,takes.course_id,course.title,course.credits,takes.grade from student,takes,course where student.ID = "+stu_id+" and takes.ID = "+stu_id+" and course.course_id = takes.course_id";
    var judge = objdbConn.Execute("select name,dept_name from student where ID="+stu_id);
    if(judge.EOF)
    {
        alert("no this person!");
        return ;
    }
    var rs = objdbConn.Execute(sql);
    var c = rs.Fields.Count-1;
    var str = "<table border=1><tr>";
    for(var i = 0; i <=c; i++)
    {
        str += "<td>" + rs.Fields(i).Name + "</td>";
    }
    str += "</tr>";
    if(rs.EOF)
    {
        str +="<tr>"
        for(var i = 0;i <= judge.Fields.Count-1;i++)
        {
            str += "<td>" + judge.Fields(i).Value + "</td>";
        }
        str += "</tr>";
    }
    var flag = 0;
    while(!rs.EOF)
    {
        str += "<tr>";
        for(var i = 0;i <= c; i++)
        {
            if(flag==1 && (i == 0 ||i == 1))
            {
                str += "<td>" + "</td>"
            }
            else{
                str += "<td>" + rs.Fields(i).Value + "</td>";
            }
        }
        str += "</tr>";
        flag = 1;
        rs.moveNext();
    }
    str += "</table>";
    var yy = document.getElementById("sp");
    yy.innerHTML = str;
}
function query2(){
    // var stu_id = document.getElementById("stu_id").value;
    var keywords = document.getElementById("keywords").value;
    var sql = "select * from course where title like \'%"+keywords+"%\'";
    var rs = objdbConn.Execute(sql);
    var c = rs.Fields.Count-1;
    //拼接表的字段名称
    var str = "<table border=1><tr>";
    for(var i = 0; i <=c; i++)
    {
        str += "<td>" + rs.Fields(i).Name + "</td>";
    }
    str += "</tr>";
    //拼接表的数据
    while(!rs.EOF)
    {
        str += "<tr>";
        for(var i = 0;i <= c; i++)
        {
            str += "<td>" + rs.Fields(i).Value + "</td>";
        }
        str += "</tr>";
        rs.moveNext();
    }
    str += "</table>";
    var yy = document.getElementById("sp");
    yy.innerHTML = str;
}
function query3(){
    // var stu_id = document.getElementById("stu_id").value;
    var keywords = document.getElementById("keywords").value;
    var sql = "select takes.ID,student.name,count(takes.ID) as cnt from takes,student where student.ID=takes.ID and (takes.grade ='C-' or takes.grade = 'C') group by takes.ID,student.name having count(takes.ID)"+keywords;
    var rs = objdbConn.Execute(sql);
    var c = rs.Fields.Count-1;
    //拼接表的字段名称
    var str = "<table border=1><tr>";
    for(var i = 0; i <=c; i++)
    {
        str += "<td>" + rs.Fields(i).Name + "</td>";
    }
    str += "</tr>";
    //拼接表的数据
    while(!rs.EOF)
    {
        str += "<tr>";
        for(var i = 0;i <= c; i++)
        {
            str += "<td>" + rs.Fields(i).Value + "</td>";
        }
        str += "</tr>";
        rs.moveNext();
    }
    str += "</table>";
    var yy = document.getElementById("sp");
    yy.innerHTML = str;
}
function insert(){
    var stu_id = document.getElementById("stu_id").value;
    var stu_name = document.getElementById("stu_name").value;
    var keywords = document.getElementById("keywords").value;
    var sql = "insert into student values("+stu_id+",'"+stu_name+"','"+keywords+"',0)";
    var judge1 = objdbConn.Execute("select * from student where ID='"+stu_id+"'");
    // var judge2 = objdbConn.Execute("select * from department where dept_name='"+keywords+"'")
    if(!judge1.EOF){
        alert("wrong var!");
    }
    else{
        
        var rs = objdbConn.Execute(sql);
        alert("inserted!");
    }   
}
function exit(){
objdbConn.Close();
}
function toa(){
    location.href="a.html?username="+username+"&pwd="+pwd+"&database="+database;
}
function tob(){
    location.href="b.html?username="+username+"&pwd="+pwd+"&database="+database;
}
function toc(){
    location.href="c.html?username="+username+"&pwd="+pwd+"&database="+database;
}
function tod(){
    location.href="d.html?username="+username+"&pwd="+pwd+"&database="+database;
}