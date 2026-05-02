<%@ Language="VBScript" %>
<%
Option Explicit
Response.Buffer = True
%>

<!DOCTYPE html>
<html>
<head>
<title>Employee Demand & Rent Details</title>

<style>
body {
    background: linear-gradient(rgba(0,0,0,0.6), rgba(0,0,0,0.6)),
                url('n9.jpg') no-repeat center center fixed;
    background-size: cover;
    margin: 0;
    min-height: 100vh;
    font-family: 'Segoe UI', Arial, sans-serif;
    display: flex;
    justify-content: center;
    align-items: center;
}
.home-nav {
    position: absolute;
    top: 20px;
    left: 20px;
}
.home-nav a {
    text-decoration: none;
    color: white;
    background: rgba(255,255,255,0.2);
    padding: 10px 18px;
    border-radius: 25px;
}
.home-nav img {
            width: 20px;
            margin-right: 8px;
            filter: invert(1);
        }
.main {
    background: #fff;
    padding: 30px;
    border-radius: 15px;
    width: 95%;
    max-width: 1100px;
}
h2, h3 {
    color: #2c3e50;
}
.row {
    display: flex;
    gap: 15px;
    flex-wrap: wrap;
    margin-bottom: 15px;
}
input, select, button {
    padding: 8px 12px;
}
button {
    background: #3498db;
    color: white;
    border: none;
    border-radius: 20px;
    cursor: pointer;
}
table {
    width: 100%;
    border-collapse: collapse;
    margin-top: 15px;
}
th {
    background: #3498db;
    color: white;
    padding: 10px;
}
td {
    padding: 10px;
    border-bottom: 1px solid #ddd;
    text-align: center;
}
.paging {
    text-align: center;
    margin-top: 15px;
}
</style>

<script>
function validateSearch(){
    if(document.f1.selectedDate.value==""){
        alert("Please select a date");
        return false;
    }
    return true;
}
</script>
</head>

<body>

<div class="home-nav">
    <a href="index.html"> 
     <img src="https://cdn-icons-png.flaticon.com/512/1946/1946436.png">Home</a>
</div>

<div class="main">
<h2>Employee Demand & Rent Details</h2>

<%
'========================
' INPUTS & PAGING
'========================
Dim empNo, selectedDate, actionType, pageNo, pageSize
empNo        = Trim(Request("emp_no"))
selectedDate = Trim(Request("selectedDate"))
actionType   = Trim(Request("actionType"))

pageSize = 5
pageNo = Request("page")
If pageNo="" Or Not IsNumeric(pageNo) Then pageNo = 1
pageNo = CInt(pageNo)
If pageNo < 1 Then pageNo = 1

'========================
' DB CONNECTION
'========================
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLOLEDB;Data Source=localhost\SQLEXPRESS;Initial Catalog=newpro;User ID=stefina;Password=Stefina@1111;"
%>

<form method="post" name="f1">
<div class="row">
    Date:
    <input type="date" name="selectedDate" value="<%=selectedDate%>">

    Employee:
    <select name="emp_no">
        <option value="">-- All Employees --</option>
        <%
        Dim rsEmp
        Set rsEmp = conn.Execute("SELECT EMP_NO FROM EMP_MST ORDER BY EMP_NO")
        Do Until rsEmp.EOF
        %>
        <option value="<%=rsEmp("EMP_NO")%>" <%If empNo=rsEmp("EMP_NO") Then Response.Write("selected")%>>
            <%=rsEmp("EMP_NO")%>
        </option>
        <%
            rsEmp.MoveNext
        Loop
        rsEmp.Close
        %>
    </select>

    <button type="submit" name="actionType" value="search" onclick="return validateSearch()">Search</button>
    <button type="submit" name="actionType" value="proceed" onclick="return validateSearch()">Proceed</button>
</div>
</form>

<hr>

<%
'==================================================
' EMPLOYEE DETAILS WITH PAGING (SEARCH ONLY)
'==================================================
If actionType="search" And selectedDate<>"" Then

    Dim rs, sql, totalPages, i
    sql = _
        "SELECT m.EMP_NO,m.NAME,m.UNIT,m.DESIGNATION," & _
        "o.QTRS_KEY,o.QTR_BLOCK,a.RAS_STREETDESC,o.OCCN_DATE " & _
        "FROM EMP_MST m " & _
        "JOIN EMP_OCCU o ON m.EMP_NO=o.CPF_NO " & _
        "JOIN EMP_ADDRESS a ON m.EMP_NO=a.EMP_NO "

    If empNo<>"" Then sql = sql & " WHERE m.EMP_NO=" & empNo

    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3
    rs.PageSize = pageSize
    rs.Open sql, conn, 1, 1

    If Not (rs.BOF And rs.EOF) Then
        totalPages = rs.PageCount
        If pageNo > totalPages Then pageNo = totalPages
        rs.AbsolutePage = pageNo
%>

<h3>Employee Details</h3>

<table>
<tr>
<th>EMP NO</th><th>NAME</th><th>UNIT</th><th>DESIGNATION</th>
<th>QTR KEY</th><th>BLOCK</th><th>ADDRESS</th><th>OCCN DATE</th>
</tr>

<%
i = 0
Do Until rs.EOF Or i = pageSize
%>
<tr>
<td><%=rs("EMP_NO")%></td>
<td><%=rs("NAME")%></td>
<td><%=rs("UNIT")%></td>
<td><%=rs("DESIGNATION")%></td>
<td><%=rs("QTRS_KEY")%></td>
<td><%=rs("QTR_BLOCK")%></td>
<td><%=rs("RAS_STREETDESC")%></td>
<td><%=FormatDateTime(rs("OCCN_DATE"),vbShortDate)%></td>
</tr>
<%
rs.MoveNext
i = i + 1
Loop
%>
</table>

<div class="paging">
<% If pageNo > 1 Then %>
<form method="post" style="display:inline;">
<input type="hidden" name="selectedDate" value="<%=selectedDate%>">
<input type="hidden" name="emp_no" value="<%=empNo%>">
<input type="hidden" name="actionType" value="search">
<input type="hidden" name="page" value="<%=pageNo-1%>">
<button>⬅ Prev</button>
</form>
<% End If %>

 Page <%=pageNo%> of <%=totalPages%>

<% If pageNo < totalPages Then %>
<form method="post" style="display:inline;">
<input type="hidden" name="selectedDate" value="<%=selectedDate%>">
<input type="hidden" name="emp_no" value="<%=empNo%>">
<input type="hidden" name="actionType" value="search">
<input type="hidden" name="page" value="<%=pageNo+1%>">
<button>Next ➡</button>
</form>
<% End If %>
</div>

<%
    Else
        Response.Write "<p>No employees found</p>"
    End If
    rs.Close
End If
%>

<%
'==================================================
' PROCEED → INSERT + RENT SUMMARY
'==================================================
If actionType="proceed" And selectedDate<>"" And empNo<>"" Then

    Dim rsChk
    Set rsChk = conn.Execute( _
        "SELECT 1 FROM RENT_DMND_TXN WHERE CPF_NO=" & empNo & _
        " AND DMND_DATE='" & selectedDate & "'")

    If rsChk.EOF Then
        Dim rsRent
        Set rsRent = conn.Execute( _
            "SELECT r.QTR_EMP_AMT,r.QTR_EBRATE,r.QTR_WARATE,r.QTR_CONSUL_AMT " & _
            "FROM EMP_OCCU o JOIN RAS_RENT r ON o.QTR_TYPE_CD=r.QTR_TYPE_CD " & _
            "WHERE o.CPF_NO=" & empNo)

        If Not rsRent.EOF Then
            conn.Execute "INSERT INTO RENT_DMND_TXN VALUES (" & empNo & ",506,'" & selectedDate & "'," & rsRent(0) & ")"
            conn.Execute "INSERT INTO RENT_DMND_TXN VALUES (" & empNo & ",507,'" & selectedDate & "'," & rsRent(1) & ")"
            conn.Execute "INSERT INTO RENT_DMND_TXN VALUES (" & empNo & ",508,'" & selectedDate & "'," & rsRent(2) & ")"
            conn.Execute "INSERT INTO RENT_DMND_TXN VALUES (" & empNo & ",509,'" & selectedDate & "'," & rsRent(3) & ")"
        End If
        rsRent.Close
    End If
    rsChk.Close

    Dim rsTxn
    Set rsTxn = conn.Execute( _
        "SELECT CPF_NO,DEDN_CODE,DMND_DATE,DMND_AMT FROM RENT_DMND_TXN " & _
        "WHERE CPF_NO=" & empNo & " AND DMND_DATE='" & selectedDate & "' ORDER BY DEDN_CODE")

    If Not rsTxn.EOF Then
%>

<h3>Rent Summary</h3>
<table>
<tr><th>CPF NO</th><th>DEDN CODE</th><th>DATE</th><th>AMOUNT</th></tr>
<%
Do Until rsTxn.EOF
%>
<tr>
<td><%=rsTxn("CPF_NO")%></td>
<td><%=rsTxn("DEDN_CODE")%></td>
<td><%=FormatDateTime(rsTxn("DMND_DATE"),vbShortDate)%></td>
<td><%=rsTxn("DMND_AMT")%></td>
</tr>
<%
rsTxn.MoveNext
Loop
%>
</table>

<%
    End If
    rsTxn.Close
End If

conn.Close
%>

</div>
</body>
</html>

