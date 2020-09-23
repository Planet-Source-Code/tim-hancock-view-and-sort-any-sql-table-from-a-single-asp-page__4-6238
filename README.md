<div align="center">

## View and Sort any SQL table from a single ASP page


</div>

### Description

This page allows you to view and sort all of your tables in an SQL database
 
### More Info
 
You will only need to add the server name and database name to the connection string to adapt it to work on your server.

The value of 133575514 entered in the sysobjects SQL string is the id for the table dtproperties. This table is a system table and you will need to substitute the value for the relating id found in the sysobjects table in your database.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tim Hancock](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tim-hancock.md)
**Level**          |Advanced
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tim-hancock-view-and-sort-any-sql-table-from-a-single-asp-page__4-6238/archive/master.zip)





### Source Code

```
' ADO variables and SQL strings
Dim adoConn, adoRS, strSQL, strTableSQL
' Create the Connection Object
Set adoConn = Server.CreateObject("ADODB.Connection")
' Open the Connection using the SQL OLE ODBC adoConn.Open "Provider=SQLOLEDB; Data Source=<Your Server Name>; Initial Catalog=<Your Database Name>; User Id=sa; Password=;"
' Create the Recordset Object
Set adoRS = Server.CreateObject("ADODB.Recordset")
' Retrieve the table name from the submitting form, When the form first loads, there is no table name passed
TableName = Request.Form("Tables")
If TableName <> "" Then
    ' Store the value for later use if sorting the table
	Response.Cookies("TableSort")("TableName") = TableName
else
    ' The form has been refreshed, grab the table name from the cookie
	TableName = Request.Cookies("TableSort")("TableName")
end if
' Retrieve the sort preference
SortBy = Request.Form("SortBy")
' If this is not the first time the page has loaded
If TableName <> "" Then
	' Retrieve the id from the sysobjects table
	strSQL = "SELECT id FROM sysobjects WHERE name = '" & TableName & "'"
	Set adoRS = adoConn.Execute(strSQL)
	TableID = adoRS("id")
	adoRS.Close
	' Create the new SQL string to retrieve the columns
	strTableSQL = "SELECT * FROM " & TableName
	' If the form has refreshed with a sort by preference
    If SortBy <> "" Then
		strTableSQL = strTableSQL & " ORDER BY " & SortBy
	end if
End If
%>
<HTML>
<BODY>
<!-- Create the Information Row and Sort By row of the table -->
<form name="SortTable" action="Tables.asp" target="_self" method="post">
<table border="0">
	<tr>
		<td><font face="verdana,arial,helvetica" size="1"><b>Table Name: </b></font>
			<select name="Tables">
			<% strSQL = "SELECT name FROM sysobjects WHERE xtype = 'U' AND (Not (id) = 133575514) ORDER BY name"
			Set adoRS = adoConn.Execute(strSQL)
			Do While Not adoRS.EOF
				if adoRS("name") = TableName Then %>
					<option selected value="<%= adoRS("name") %>"><%= adoRS("name") %>
				<% else %>
					<option value="<%= adoRS("name") %>"><%= adoRS("name") %>
				<% end if %>
			<% adoRS.MoveNext
			Loop
			adoRS.Close %>
			</select>
		</td>
		<td><font face="verdana,arial,helvetica" size="1"><b>Sort By:</b></font>
				<select name="SortBy"><option value="">
					<% If TableID <> "" Then
						strSQL = "SELECT name FROM syscolumns WHERE id = " & TableID
						Set adoRS = adoConn.Execute(strSQL)
						Do While Not adoRS.EOF %>
							<option value="<%= adoRS("name") %>"><%= adoRS("name") %>
						<% adoRS.MoveNext
						Loop
						adoRS.Close
					End If %>
				</select>
			</td>
		<td><input type="submit"></td>
	</tr>
</table>
</form>
<% If TableName <> "" Then %>
<BR>
<font face="verdana,arial,helvetica" size="2"><b>Table Name: <%= TableName %></b>
<BR>
<BR>
<!-- Create the table itself -->
<table>
	<!-- The header row -->
	<tr>
		<% strSQL = "SELECT name FROM syscolumns WHERE id = " & TableID & " ORDER BY colid"
		Set adoRS = adoConn.Execute(strSQL)
		Do While Not adoRS.EOF %>
			<td><font face="verdana,arial,helvetica" size="1"><b><%= adoRS("name") %></b></td>
		<% adoRS.MoveNext
		Loop
		adoRS.Close %>
	</tr>
	<!-- The data -->
	<% Set adoRS.ActiveConnection = adoConn
		adoRS.Source = strTableSQL
		adoRS.LockType = 3
		adoRS.Open
		Do While Not adoRS.EOF %>
		<tr>
			<td>
				<form name="TableInfo" method="post" action="UpdateTable.asp?TableName=<%= TableName %>" target="_self">
						<tr>
						<% For Each oField in adoRS.Fields %>
							<td><input type="text" name="<%= oField.name %>" value="<%= oField.Value %>"></td>
						<% Next %>
						</tr>
				</form>
			</td>
		</tr>
		<% adoRS.MoveNext
		Loop
		adoRS.Close %>
</table>
<% End If %>
</BODY>
</HTML>
```

