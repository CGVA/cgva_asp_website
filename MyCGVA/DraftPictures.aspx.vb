Imports System.Data.SqlClient
Imports System
Imports System.IO

Partial Class DraftPictures
    Inherits System.Web.UI.Page

    Private SQL As String
    Private connStr As String = ConfigurationManager.AppSettings("ConnectionString")

    Private imgPath As String = ""
    Private tTable As New Table
    Private tRow As New TableRow
    Private tCell As New TableCell
    Private nameLabel As New Label

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        imgPath = Request.ServerVariables("APPL_PHYSICAL_PATH") & "personnel_images\"

        Dim cn As SqlConnection
        Dim rs As SqlCommand
        Dim rsReply As SqlDataReader
        cn = New SqlConnection
        cn.ConnectionString = connStr
        cn.Open()
        'get information for this PERSON_ID
        SQL = "SELECT p.PERSON_ID, IsNull(LAST_NAME,'Unknown') as LAST_NAME, IsNull(FIRST_NAME,'Unknown') as FIRST_NAME,EVENT_CD " _
            & "FROM db_accessadmin.[PERSON_TBL] p " _
            & "LEFT JOIN REGISTRATION_TBL r " _
            & "ON p.PERSON_ID = r.PERSON_ID " _
            & "WHERE r.EVENT_CD IN (SELECT EVENT_CD FROM EVENT_TBL WHERE getdate() < EVENT_END_DATE) " _
            & "AND REGISTRATION_IND='Y' " _
            & "ORDER BY EVENT_CD, LAST_NAME,FIRST_NAME"

        'JPC 9/1/12
        'SQL = "SELECT PERSON_ID, LAST_NAME, FIRST_NAME " _
        '    & "FROM db_accessadmin.[PERSON_TBL] " _
        '    & "WHERE PERSON_ID IN (SELECT PERSON_ID FROM REGISTRATION_TBL " _
        '                        & "WHERE EVENT_CD='2011FL' AND REGISTRATION_IND = 'Y') " _
        '    & "ORDER BY EVENT_CD, LAST_NAME,FIRST_NAME"

        'Response.Write(SQL)
        'Response.End()

        rs = New SqlCommand(SQL, cn)
        rsReply = rs.ExecuteReader


        'tTable.Width = "100%"


        While rsReply.Read
            Dim tRow As New TableRow
            Dim tCell As New TableCell
            Dim nameLabel As New Label
            'name row
            nameLabel.Text = rsReply(1) + ", " + rsReply(2)
            tCell.Controls.Add(nameLabel)
            tRow.Controls.Add(tCell)
            tTable.Controls.Add(tRow)

            'see if there is an image for this user
            getImage(rsReply(0), rsReply(1), rsReply(2))
        End While

        Me.picturePanel.Controls.Add(tTable)

        rsReply.Close()
        cn.Close()

    End Sub

    Public Sub getImage(ByVal personID As Integer, ByVal lastName As String, ByVal firstName As String)
        'check for image file
        Dim str As String
        Dim imgFound As Boolean = False
        Dim imgFile As New Image
        Dim tRow As New TableRow
        Dim tCell As New TableCell
        Dim nameLabel As New Label

        For Each str In Directory.GetFiles(imgPath, personID.ToString + "_*")
            imgFile.ImageUrl = "http://cgva.org/personnel_images/" + Path.GetFileName(str)
            imgFile.Height = 300
            imgFile.Width = 200
            tCell.Controls.Add(imgFile)
            tRow.Controls.Add(tCell)
            tTable.Controls.Add(tRow)

            imgFound = True
            Exit For
        Next

        'if they didnt upload a file from MyCGVA, check to see if a file exists from past draft photos
        If Not imgFound Then
            For Each str In Directory.GetFiles(imgPath, lastName + "_" & firstName + "*")
                imgFile.ImageUrl = "http://cgva.org/personnel_images/" + Path.GetFileName(str)
                imgFile.Height = 300
                imgFile.Width = 200
                tCell.Controls.Add(imgFile)
                tRow.Controls.Add(tCell)
                tTable.Controls.Add(tRow)

                imgFound = True
                Exit For
            Next
        End If

        'no image found, add blank row
        If Not imgFound Then
            nameLabel.Text = "NO IMAGE FOUND FOR THIS PERSON"
            nameLabel.Height = 300
            tCell.Controls.Add(nameLabel)
            tRow.Controls.Add(tCell)
            tTable.Controls.Add(tRow)
        End If

    End Sub


End Class
