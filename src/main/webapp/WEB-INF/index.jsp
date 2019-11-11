<%--
  Created by IntelliJ IDEA.
  User: admin
  Date: 2019/10/25
  Time: 9:40
  To change this template use File | Settings | File Templates.
--%>
<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<html>
<head>
    <title>首页</title>
</head>
<body style="background-color: aliceblue;">
<form action="/user/batchSave" method="post"  enctype="multipart/form-data">
    <div style="text-align: center;text-align: center;width: 50%;margin-left: 25%;">
        <h2>请上传.XSL格式文件：</h2>
        <input type="file" name="excelFile" style="width: 50%;height: 30px;border: 1px solid #ee9238;">
        <input type="submit" value="提交" style=" width: 50%;height: 35px;">
    </div>
</form>
</body>
</html>
