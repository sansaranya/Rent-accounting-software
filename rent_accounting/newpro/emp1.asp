<%@ Language="VBScript" %>
<html>
<head>
    <title>Add Employee | HR Portal</title>
    <style>
        /* Modern Background */
        body {
            background: linear-gradient(rgba(0, 0, 0, 0.6), rgba(0, 0, 0, 0.6)), 
                        url('https://images.unsplash.com/photo-1497366216548-37526070297c?auto=format&fit=crop&w=1350&q=80');
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
            margin: 0;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }

        /* --- Top Left Home Navigation --- */
        .home-nav {
            position: absolute;
            top: 20px;
            left: 20px;
            z-index: 1000;
        }

        .home-nav a {
            display: flex;
            align-items: center;
            text-decoration: none;
            color: white;
            background: rgba(255, 255, 255, 0.2);
            padding: 10px 20px;
            border-radius: 30px;
            backdrop-filter: blur(10px); /* Blur effect */
            border: 1px solid rgba(255, 255, 255, 0.3);
            transition: all 0.3s ease;
            font-weight: 500;
        }

        .home-nav a:hover {
            background: rgba(255, 255, 255, 0.4);
            transform: scale(1.05);
        }

        .home-nav img {
            width: 20px;
            margin-right: 8px;
            filter: invert(1); /* Makes the black icon white */
        }

        /* Glassmorphism Form Container */
        .card {
            background: rgba(255, 255, 255, 0.95);
            padding: 30px;
            border-radius: 15px;
            box-shadow: 0 15px 35px rgba(0,0,0,0.3);
            width: 100%;
            max-width: 450px;
            text-align: center;
            margin-top: 50px; /* Space for the top nav on mobile */
        }

        .profile-header {
            width: 80px;
            height: 80px;
            background-color: #3498db;
            border-radius: 50%;
            margin: -70px auto 20px auto;
            border: 5px solid white;
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: 0 4px 10px rgba(0,0,0,0.1);
        }

        h2 { color: #2c3e50; margin-bottom: 25px; }
        form { text-align: left; }
        label { font-size: 0.85rem; font-weight: bold; color: #555; display: block; margin-bottom: 5px; }

        input[type=text], input[type=date] {
            width: 100%;
            padding: 10px;
            margin-bottom: 15px;
            border: 1px solid #ddd;
            border-radius: 6px;
            box-sizing: border-box;
        }

        input[type=submit] {
            background: linear-gradient(135deg, #3498db, #2980b9);
            color: white;
            padding: 12px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            width: 100%;
            font-size: 16px;
            font-weight: bold;
            text-transform: uppercase;
            transition: 0.3s;
        }

        input[type=submit]:hover { opacity: 0.9; transform: translateY(-2px); }
        .success { color: green; font-weight: bold; }
        .error { color: red; font-weight: bold; }
    </style>
</head>

<body>

<nav class="home-nav">
    <a href="index.html">
        <img src="https://cdn-icons-png.flaticon.com/512/1946/1946436.png" alt="Home">
        Home
    </a>
</nav>

<div class="card">
    <div class="profile-header">
        <img src="https://cdn-icons-png.flaticon.com/512/9131/9131529.png" alt="User" style="width: 45px; filter: invert(1);">
    </div>

    <h2> Employee Registration</h2>

    <form method="post" action="">
        <label>Employee No*</label>
        <input type="text" name="emp_no" required>

        <label>Name*</label>
        <input type="text" name="name" required>

        <div style="display: flex; gap: 10px;">
            <div style="flex: 1;">
                <label>Unit</label>
                <input type="text" name="unit">
            </div>
            <div style="flex: 1;">
                <label>Designation</label>
                <input type="text" name="designation">
            </div>
        </div>

        <label>Leaving Date</label>
        <input type="date" name="dtlvng">

        <label>Reason Leave</label>
        <input type="text" name="reason_leave">

        <label>Flag</label>
        <input type="text" name="emp_flag">

        <input type="submit" value="Register">
    </form>

<%
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' --- Your existing VBScript code starts here ---
    Dim emp_no, emp_name, unit, designation, reason_leave, dtlvng, emp_flag
    emp_no = Trim(Request.Form("emp_no"))
    emp_name = Trim(Request.Form("name"))
    unit = Trim(Request.Form("unit"))
    designation = Trim(Request.Form("designation"))
    reason_leave = Trim(Request.Form("reason_leave"))
    dtlvng = Trim(Request.Form("dtlvng"))
    emp_flag = Trim(Request.Form("emp_flag"))

    If emp_no = "" Or emp_name = "" Then
        Response.Write "<p class='error'>Required fields missing</p>"
    Else
        Dim conn, sql, rows
        Set conn = Server.CreateObject("ADODB.Connection")
        ' Ensure your connection string is correct for your environment
        conn.Open "Provider=SQLOLEDB;Data Source=localhost\SQLEXPRESS;Initial Catalog=newpro;User ID=stefina;Password=Stefina@1111;"

        If unit = "" Then unit = "NULL" Else unit = "'" & Replace(unit,"'","''") & "'"
        If designation = "" Then designation = "NULL" Else designation = "'" & Replace(designation,"'","''") & "'"
        If reason_leave = "" Then reason_leave = "NULL" Else reason_leave = "'" & Replace(reason_leave,"'","''") & "'"
        If emp_flag = "" Then emp_flag = "NULL" Else emp_flag = "'" & Replace(emp_flag,"'","''") & "'"
        If dtlvng = "" Then dtlvng = "NULL" Else dtlvng = "'" & dtlvng & "'"

        sql = "INSERT INTO dbo.EMP_MST (EMP_NO, [NAME], UNIT, DESIGNATION, REASON_LEAVE, DTLVNG, EMP_FLAG) VALUES (" & _
              emp_no & ", '" & Replace(emp_name,"'","''") & "', " & unit & ", " & designation & ", " & _
              reason_leave & ", " & dtlvng & ", " & emp_flag & ")"

        conn.Execute sql, rows

        If rows = 1 Then
            Response.Write "<p class='success'>Saved Successfully!</p>"
        Else
            Response.Write "<p class='error'>Error saving record.</p>"
        End If

        conn.Close
        Set conn = Nothing
    End If
End If
%>

</div>
</body>
</html>