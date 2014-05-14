Imports System.Data.SqlClient
Partial Class Home
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim SQL As String = ""
        Dim cn As SqlConnection
        Dim rs As SqlCommand
        Dim rsreply As SqlDataReader
        Dim connStr As String = ConfigurationManager.AppSettings("connectionString")
        cn = New SqlConnection
        cn.ConnectionString = connStr
        cn.Open()

        SQL = "SELECT PERSON_ID, " _
          & "LAST_NAME + ', ' + FIRST_NAME as 'NAME' " _
          & "FROM 		db_accessadmin.PERSON_TBL " _
          & "WHERE LOGICAL_DELETE_IND = 'N' " _
          & "AND SUPPRESS_EMAIL_IND = 'N' " _
          & "ORDER BY UPPER(LAST_NAME),UPPER(FIRST_NAME)"
        'rw("<!-- SQL: " & SQL & " -->")
        rs = New SqlCommand(SQL, cn)
        rsreply = rs.ExecuteReader
        'set rs = cn.(SQL)

        While rsreply.Read
            Dim li As ListItem = New ListItem(rsreply("NAME"), rsreply("PERSON_ID"))
            Me.PERSONNEL.Items.Add(li)
        End While

        rsreply.Close()

        SQL = "SELECT ID,DISTRO_NAME FROM EMAILDISTRO_TBL ORDER BY UPPER(DISTRO_NAME)"
        rs = New SqlCommand(SQL, cn)
        rsreply = rs.ExecuteReader
        'set rs = cn.(SQL)

        While rsreply.Read
            Dim li As ListItem = New ListItem(rsreply("DISTRO_NAME"), rsreply("ID"))
            Me.DISTRO.Items.Add(li)
        End While

        'Me.PERSONNEL.Width

    End Sub


    Protected Sub SAVEDISTRO_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ''Handles SAVEDISTRO.CheckedChanged
        Response.Write("HERE")
        'If Me.SAVEDISTRO.Checked Then
        'Response.Write("IF")
        'Me.DISTRO_NAME.Visible = True
        'Else
        'Response.Write("ELSE")
        'Me.DISTRO_NAME.Text = ""
        'Me.DISTRO_NAME.Visible = False
        'End If
    End Sub

    Protected Sub submitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles submitButton.Click
        Server.Transfer("test.aspx", True)
        'Response.Redirect("test.aspx", True)
    End Sub

End Class
