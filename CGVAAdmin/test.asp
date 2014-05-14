<!-- #include virtual = "/incs/dbConnection.inc" -->
<%

	SQL = "SELECT * FROM db_accessadmin.PERSON_TBL ORDER BY LAST_NAME, FIRST_NAME"
''	SQL = "SELECT * FROM FIRST_CONTACT_TBL"

	set rs = cn.Execute(SQL)

	if not rs.EOF then
		rsData = rs.GetRows
		rsDataCols = UBound(rsData,1)
		rsDataRows = UBound(rsData,2)
	else
		rsDataRows = -1
	end if

''	rw(SQL)
''	Response.End

	rw("<table border='1'>")
	rw("<tr>")

	For i = 0 to rsDataCols
		rw("<th><font size='2' face='arial'>" & rs(i).Name & "</font></th>")
	Next

	rw("</tr>")

	For i = 0 to rsDataRows
		rw("<tr>")

		For j = 0 to rsDataCols
			rw("<td><font size='2' face='arial'>" & rsData(j,i) & "&nbsp;</font></td>")
		Next

		rw("</tr>")
	Next

	rw("</table>")
	closeRSCNConnection








	Sub rw(ToPrint)
		Response.Write(ToPrint)
	End Sub

	sub closeRSCNConnection()
		rs.Close
		set rs = Nothing
		cn.Close
		set cn = Nothing
	end sub

	sub closeCNConnection()
		cn.Close
		set cn = Nothing
	end sub

 %>