<%@ Language="VBScript" %>
<html>
<head>
    <title>Add Employee</title>
    <style>
        body {
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            margin: 0;
            font-family: Arial, sans-serif;
            background: #f4f4f4;
        }
        .container { text-align: center; }
        h2 { color: #2c3e50; }
        form {
            background: #fff;
            padding: 20px 30px;
            border-radius: 8px;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
            text-align: left;
            width: 400px;
        }
        input[type=text], input[type=date] {
            width: 100%;
            padding: 8px;
            margin: 5px 0 12px 0;
            border: 1px solid #ccc;
            border-radius: 4px;
        }
        input[type=submit] {
            background-color: #3498db;
            color: white;
            padding: 10px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            width: 100%;
            font-size: 16px;
        }
        .success { color: green; }
        .error { color: red; }
    </style>
</head>

<body>
<div class="container">

<h2>Insert Employee Details</h2>

<form method="post" action="">
    Employee No*: <input type="text" name="emp_no" required>
    Name*: <input type="text" name="name" required>
    Unit: <input type="text" name="unit">
    Designation: <input type="text" name="designation">
    Reason Leave: <input type="text" name="reason_leave">
    Leaving Date: <input type="date" name="dtlvng">
    Flag: <input type="text" name="emp_flag">
    <input type="submit" value="Submit">
</form>

<%
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

    ' ---- Read values ----
    Dim emp_no, emp_name, unit, designation, reason_leave, dtlvng, emp_flag
    emp_no = Trim(Request.Form("emp_no"))
    emp_name = Trim(Request.Form("name"))
    unit = Trim(Request.Form("unit"))
    designation = Trim(Request.Form("designation"))
    reason_leave = Trim(Request.Form("reason_leave"))
    dtlvng = Trim(Request.Form("dtlvng"))
    emp_flag = Trim(Request.Form("emp_flag"))

    If emp_no = "" Or emp_name = "" Then
        Response.Write "<p class='error'>Employee No and Name are required</p>"
    Else

        ' ---- Database connection ----
        Dim conn
        Set conn = Server.CreateObject("ADODB.Connection")
        conn.Open "Provider=SQLOLEDB;Data Source=localhost\SQLEXPRESS;Initial Catalog=newpro;User ID=stefina;Password=Stefina@1111;"

        ' ---- Optional fields ----
        If unit = "" Then unit = "NULL" Else unit = "'" & Replace(unit,"'","''") & "'"
        If designation = "" Then designation = "NULL" Else designation = "'" & Replace(designation,"'","''") & "'"
        If reason_leave = "" Then reason_leave = "NULL" Else reason_leave = "'" & Replace(reason_leave,"'","''") & "'"
        If emp_flag = "" Then emp_flag = "NULL" Else emp_flag = "'" & Replace(emp_flag,"'","''") & "'"

        ' ---- Date handling ----
        If dtlvng = "" Then
            dtlvng = "NULL"
        Else
            dtlvng = "'" & dtlvng & "'"
        End If

        ' ---- SQL INSERT (FIXED) ----
        Dim sql, rows
        sql = "INSERT INTO dbo.EMP_MST " & _
              "(EMP_NO, [NAME], UNIT, DESIGNATION, REASON_LEAVE, DTLVNG, EMP_FLAG) VALUES (" & _
              emp_no & ", '" & Replace(emp_name,"'","''") & "', " & unit & ", " & designation & ", " & _
              reason_leave & ", " & dtlvng & ", " & emp_flag & ")"

        conn.Execute sql, rows

        If rows = 1 Then
            Response.Write "<p class='success'>Employee inserted successfully!</p>"
        Else
            Response.Write "<p class='error'>Insert failed. No rows affected.</p>"
        End If

        conn.Close
        Set conn = Nothing

    End If
End If
%>

</div>
</body>
</html>

