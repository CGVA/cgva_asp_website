Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Xml
Imports System.IO
Imports System.Globalization

Imports System.Data.SqlClient

Namespace PWGenerator

    Public Class PasswordGenerator


        Public Shared Sub CreatePasswords()
            Dim SQL As String
            'Dim connStr As String = ConfigurationManager.AppSettings("ConnectionString")
            Dim connStr As String = "Data Source=whsql-v20.prod.mesa1.secureserver.net;Initial Catalog=cgva;User Id=cgva;Password=Sideout123;"
            Dim tempPassword As String = ""

            Dim cn As SqlConnection
            Dim cn2 As SqlConnection
            Dim rs As SqlCommand
            Dim rs2 As SqlCommand
            Dim rsreply As SqlDataReader
            cn = New SqlConnection
            cn.ConnectionString = connStr
            cn.Open()
            cn2 = New SqlConnection
            cn2.ConnectionString = connStr
            cn2.Open()

            SQL = "SELECT PERSON_ID FROM USER_LOGIN_TBL"

            rs = New SqlCommand(SQL, cn)
            rsreply = rs.ExecuteReader

            While rsreply.Read
                SQL = "TEMP_PASSWORD_CREATION @PERSON_ID = '" & rsreply(0) & "', @TEMP_PASSWORD='" & PasswordGenerator(8) & "'"
                rs2 = New SqlCommand(SQL, cn2)
                rs2.ExecuteNonQuery()
            End While
            ''& "SET TEMP_PASSWORD = '" & PasswordGenerator(8) & "' " _

        End Sub
        Public Shared Function PasswordGenerator(ByVal lngLength As Long) _
      As String

            ' Description: Generate a random password of 'user input' length
            ' Parameters : lngLength - the length of the password to be 
            'generated
            ' Returns    : String    - Randomly generated password
            ' Created    : 08/21/1999 Andrew Ells-O'Brien 
            '(andrew@ellsobrien@msn.com)

            On Error GoTo Err_Proc

            Dim iChr As Integer
            Dim c As Long
            Dim strResult As String = ""
            Dim iAsc As String

            Randomize(Timer)

            For c = 1 To lngLength

                ' Randomly decide what set of ASCII chars we will use
                iAsc = Int(3 * Rnd() + 1)

                'Randomly pick a char from the random set
                Select Case iAsc
                    Case 1
                        iChr = Int((Asc("Z") - Asc("A") + 1) * Rnd() + Asc("A"))
                    Case 2
                        iChr = Int((Asc("z") - Asc("a") + 1) * Rnd() + Asc("a"))
                    Case 3
                        iChr = Int((Asc("9") - Asc("0") + 1) * Rnd() + Asc("0"))
                    Case Else
                        Err.Raise(20000, , "PasswordGenerator has a problem.")
                End Select

                strResult = strResult & Chr(iChr)

            Next c

            PasswordGenerator = strResult

Exit_Proc:
            Exit Function

Err_Proc:
            MsgBox(Err.Number & ": " & Err.Description, _
               vbOKOnly + vbCritical)
            PasswordGenerator = vbNullString
            Resume Exit_Proc

        End Function
    End Class

End Namespace

