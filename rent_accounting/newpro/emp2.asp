<%@ Language="VBScript" %>
<%
Option Explicit
Response.Buffer = True

Function IIf(condition, truePart, falsePart)
    If condition Then IIf = truePart Else IIf = falsePart
End Function
%>

<!DOCTYPE html>
<html>
<head>
    <title>View Employee | HR Portal</title>
<style>
        body {
            background: linear-gradient(rgba(0,0,0,0.6), rgba(0,0,0,0.6)),
                        url('https://images.unsplash.com/photo-1497366216548-37526070297c?auto=format&fit=crop&w=1350&q=80');
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
            min-height: 100vh;
            margin: 0;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            display: flex;
            justify-content: center;
            align-items: center;
        }

        /* Home Button */
        .home-nav {
            position: absolute;
            top: 20px;
            left: 20px;
        }

        .home-nav a {
            display: flex;
            align-items: center;
            text-decoration: none;
            color: white;
            background: rgba(255,255,255,0.2);
            padding: 10px 20px;
            border-radius: 30px;
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255,255,255,0.3);
            transition: 0.3s;
        }

        .home-nav a:hover {
            background: rgba(255,255,255,0.4);
            transform: scale(1.05);
        }

        .home-nav img {
            width: 20px;
            margin-right: 8px;
            filter: invert(1);
        }

        /* Card */
        .card {
            background: rgba(255,255,255,0.95);
            padding: 30px;
            border-radius: 15px;
            width: 100%;
            max-width: 900px;
            box-shadow: 0 15px 35px rgba(0,0,0,0.3);
        }

        h2 {
            text-align: center;
            color: #2c3e50;
            margin-bottom: 25px;
        }

        label {
            font-weight: bold;
            color: #555;
            font-size: 0.9rem;
        }

        select, input[type=submit] {
            width: 30%;
            padding: 10px;
            margin-top: 6px;
            margin-bottom: 15px;
            border-radius: 6px;
            border: 1px solid #ccc;
        }

        input[type=submit] {
            background: linear-gradient(135deg, #3498db, #2980b9);
            color: white;
            font-weight: bold;
            cursor: pointer;
            border: none;
        }

        input[type=submit]:hover {
            opacity: 0.9;
            transform: translateY(-2px);
        }

        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 25px;
        }

        th {
            background: #3498db;
            color: white;
            padding: 10px;
            text-align: left;
        }

        td {
            padding: 10px;
            border-bottom: 1px solid #ddd;
        }

        tr:hover {
            background: #f4f6f8;
        }

        .error {
            color: red;
            font-weight: bold;
            text-align: center;
        }
        </style>

    
</head>

<body>

<nav class="home-nav">
    <a href="index.html">
        <img src="https://cdn-icons-png.flaticon.com/512/1946/1946436.png">
        Home
    </a>
</nav>

<div class="card">
<h2> Details</h2>

<%
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLOLEDB;Data Source=localhost\SQLEXPRESS;Initial Catalog=newpro;User ID=stefina;Password=Stefina@1111;"

Dim searchBy, searchValue
searchBy = Trim(Request.Form("search_by"))
searchValue = Trim(Request.Form("search_value"))
%>

<form method="post" action="emp2.asp">
    <label>Search By</label>
    <select name="search_by" onchange="this.form.submit()">
        <option value="">-- Select --</option>
        <option value="EMP_NO" <%=IIf(searchBy="EMP_NO","selected","")%>>Employee No</option>
        <option value="NAME" <%=IIf(searchBy="NAME","selected","")%>>Name</option>
        <option value="ADDRESS" <%=IIf(searchBy="ADDRESS","selected","")%>>Address</option>
    </select>

<%
If searchBy <> "" Then
%>
    <label>Select <%=searchBy%></label>
    <select name="search_value">
        <option value="">-- Select --</option>

<%
    Dim rsList, sqlList
    Select Case searchBy
        Case "EMP_NO": sqlList = "SELECT EMP_NO FROM EMP_MST ORDER BY EMP_NO"
        Case "NAME": sqlList = "SELECT DISTINCT NAME FROM EMP_MST ORDER BY NAME"
        Case "ADDRESS": sqlList = "SELECT DISTINCT RAS_STREETDESC FROM EMP_ADDRESS ORDER BY RAS_STREETDESC"
    End Select

    Set rsList = conn.Execute(sqlList)
    Do While Not rsList.EOF
%>
        <option value="<%=rsList(0)%>" <%=IIf(searchValue=rsList(0),"selected","")%>>
            <%=rsList(0)%>
        </option>
<%
        rsList.MoveNext
    Loop
    rsList.Close
%>
    </select>

    <input type="submit" value="View Employee">
<%
End If
%>
</form>

<%
If Request.ServerVariables("REQUEST_METHOD")="POST" And searchValue<>"" Then
    Dim rsEmp, sqlEmp

    Select Case searchBy
        Case "EMP_NO"
            sqlEmp = "SELECT e.*, a.RAS_STREETDESC ADDRESS FROM EMP_MST e LEFT JOIN EMP_ADDRESS a ON e.EMP_NO=a.EMP_NO WHERE e.EMP_NO='" & Replace(searchValue,"'","''") & "'"
        Case "NAME"
            sqlEmp = "SELECT e.*, a.RAS_STREETDESC ADDRESS FROM EMP_MST e LEFT JOIN EMP_ADDRESS a ON e.EMP_NO=a.EMP_NO WHERE e.NAME='" & Replace(searchValue,"'","''") & "'"
        Case "ADDRESS"
            sqlEmp = "SELECT e.*, a.RAS_STREETDESC ADDRESS FROM EMP_MST e INNER JOIN EMP_ADDRESS a ON e.EMP_NO=a.EMP_NO WHERE a.RAS_STREETDESC='" & Replace(searchValue,"'","''") & "'"
    End Select

    Set rsEmp = conn.Execute(sqlEmp)

    If Not rsEmp.EOF Then
%>
<table>
<tr>
    <th>Emp No</th>
    <th>Name</th>
    <th>Unit</th>
    <th>Designation</th>
    <th>Address</th>
</tr>
<%
    Do While Not rsEmp.EOF
%>
<tr>
    <td><%=rsEmp("EMP_NO")%></td>
    <td><%=rsEmp("NAME")%></td>
    <td><%=rsEmp("UNIT")%></td>
    <td><%=rsEmp("DESIGNATION")%></td>
    <td><%=IIf(IsNull(rsEmp("ADDRESS")),"N/A",rsEmp("ADDRESS"))%></td>
</tr>
<%
        rsEmp.MoveNext
    Loop
%>
</table>
<%
    Else
        Response.Write "<p class='error'>No employees found</p>"
    End If

    rsEmp.Close
End If

conn.Close
%>

</div>
</body>
</html>
