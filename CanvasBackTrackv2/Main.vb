Imports System.Net
Imports System.IO
Imports System.Configuration
Imports System.Xml
Imports System.Data.SqlClient


Public Class Properties
    Public Shared FCDBConnstring As String
    Public Shared NFLDBConnString As String
    Public Shared Username As String
    Public Shared Password As String
    Public Shared IsSaveSignature As Boolean
    Public Shared SignaturePath As String
    Public Shared FarmParam As String
    Public Shared WeekParam As Integer
    Public Shared YearParam As Integer
End Class


Public Class FarmCHem
    Private Shared start As DateTime
    Private Shared conn As SqlConnection
    Private Shared com As SqlCommand
    Private Shared dt As DataTable
    Private Shared ds As DataSet
    Private Shared adap As SqlDataAdapter


    Public Function ProcessFarmChem(ByVal DateFrom As Date, ByVal DateTo As Date) As String

        Dim Status As String = String.Empty
        start = DateTime.Now

        Try
            Dim web As New WebClient()
            Dim ds As New DataSet()
            Dim url As String = String.Format("https://www.gocanvas.com/apiv2/submissions.xml?username=" + Properties.Username +
                                              "&password=" + Properties.Password +
                                              "&form_name=FARMCHEM" +
                                              "&begin_date=" + DateFrom +
                                              "&end_date=" + DateTo)

            Dim response As String = web.DownloadString(url)

            Using stringReader As New StringReader(response)
                ds = New DataSet()
                ds.ReadXml(stringReader)
            End Using
            Status += RunBacktrack(ds)
        Catch ex As Exception
            Status += ex.Message + Environment.NewLine
            Status += "Error at method RunBacktrack()" + Environment.NewLine
            Try
                Status += ex.InnerException.Message + Environment.NewLine
                Status += ex.InnerException.InnerException.Message + Environment.NewLine
            Catch e As Exception

            End Try
        Finally
            Status += "Process started at " + start + " and ended at " + DateTime.Now + Environment.NewLine
            Status += "---" + Environment.NewLine
        End Try
        Return Status

    End Function

    Public Function ProcessAllFarmChem(ByVal DateFrom As Date, ByVal DateTo As Date) As String

        Dim Status As String = String.Empty
        start = DateTime.Now

        Try
            Dim web As New WebClient()
            Dim ds As New DataSet()
            Dim url As String = String.Format("https://www.gocanvas.com/apiv2/submissions.xml?username=" + Properties.Username +
                                              "&password=" + Properties.Password +
                                              "&form_name=FARMCHEM" +
                                              "&begin_date=" + DateFrom +
                                              "&end_date=" + DateTo)

            Dim response As String = web.DownloadString(url)

            Using stringReader As New StringReader(response)
                ds = New DataSet()
                ds.ReadXml(stringReader)
            End Using
            Status += RunAllBacktrack(ds)
        Catch ex As Exception
            Status += ex.Message + Environment.NewLine
            Status += "Error at method RunAllBacktrack()" + Environment.NewLine
            Try
                Status += ex.InnerException.Message + Environment.NewLine
                Status += ex.InnerException.InnerException.Message + Environment.NewLine
            Catch e As Exception

            End Try
        Finally
            Status += "Process started at " + start + " and ended at " + DateTime.Now + Environment.NewLine
            Status += "---" + Environment.NewLine
        End Try
        Return Status

    End Function

    Private Function RunBacktrack(ByVal ds As DataSet) As String
        Dim Status As String = String.Empty
        Dim dt As New DataTable
        Dim trueCounter As Integer = 0, falseCounter As Integer = 0

        Dim IsSaved As Boolean

        Try
            If ds.Tables.Count > 1 Then

                For Each a As DataRow In ds.Tables("Submission").Rows

                    IsSaved = False 'reference for signature
                    Dim tempRecord() As DataRow = ds.Tables("Section").Select("Sections_Id = " + a("Submission_Id").ToString())
                    Dim dttempMain As DataTable = ds.Tables("Response").Select("Responses_Id = " + tempRecord(0)("Section_Id").ToString()).CopyToDataTable()
                    Dim dttempBI As DataTable = ds.Tables("Response").Select("Responses_Id = " + tempRecord(1)("Section_Id").ToString()).CopyToDataTable()
                    Dim dttempBBSF As DataTable = ds.Tables("Response").Select("Responses_Id = " + tempRecord(2)("Section_Id").ToString()).CopyToDataTable()
                    Dim dttempBBSI As DataTable = ds.Tables("Response").Select("Responses_Id = " + tempRecord(3)("Section_Id").ToString()).CopyToDataTable()
                    Dim dttempMSSI As DataTable = ds.Tables("Response").Select("Responses_Id = " + tempRecord(4)("Section_Id").ToString()).CopyToDataTable()
                    Dim dttempBTEI As DataTable = ds.Tables("Response").Select("Responses_Id = " + tempRecord(5)("Section_Id").ToString()).CopyToDataTable()
                    Dim dttempTotal As DataTable = ds.Tables("Response").Select("Responses_Id = " + tempRecord(6)("Section_Id").ToString()).CopyToDataTable()
                    Dim dttempEnd As DataTable = ds.Tables("Response").Select("Responses_Id = " + tempRecord(7)("Section_Id").ToString()).CopyToDataTable()

                    'Check if submission id is new
                    If CheckIfNew(a("Id").ToString()) Then

                        Dim farmtemp = dttempMain(4)(1).ToString()
                        Dim weeknotemp = Convert.ToInt32(dttempMain(2)(1).ToString())
                        Dim yeartemp = Convert.ToInt32(dttempMain(0)(1).ToString().Split("/")(2))

                        If farmtemp = Properties.FarmParam And weeknotemp = Properties.WeekParam And yeartemp = Properties.YearParam Then

                            conn = New SqlConnection(Properties.FCDBConnstring)
                            conn.Open()
                            Dim trans As SqlTransaction = conn.BeginTransaction()

                            Try
                                com = New SqlCommand("sp_UploadTransHeader", conn, trans)
                                com.CommandType = CommandType.StoredProcedure
                                com.CommandTimeout = 999999
                                com.Parameters.Add("@Transdate", SqlDbType.VarChar).Value = dttempMain(0)(1).ToString()
                                Try
                                    com.Parameters.Add("@Transtime", SqlDbType.VarChar).Value = dttempMain(1)(1).ToString()
                                Catch ex As Exception
                                    com.Parameters.Add("@Transtime", SqlDbType.VarChar).Value = ""
                                End Try

                                Try
                                    com.Parameters.Add("@Week", SqlDbType.Int).Value = dttempMain(2)(1).ToString()
                                Catch ex As Exception
                                    com.Parameters.Add("@Week", SqlDbType.Int).Value = Nothing
                                End Try

                                Try
                                    com.Parameters.Add("@Farm", SqlDbType.VarChar).Value = dttempMain(4)(1).ToString()
                                Catch ex As Exception
                                    com.Parameters.Add("@Farm", SqlDbType.VarChar).Value = ""
                                End Try

                                Try
                                    com.Parameters.Add("@Grower", SqlDbType.VarChar).Value = dttempMain(5)(1).ToString()
                                Catch ex As Exception
                                    com.Parameters.Add("@Grower", SqlDbType.VarChar).Value = ""
                                End Try

                                Try
                                    com.Parameters.Add("@PHCode", SqlDbType.VarChar).Value = dttempMain(6)(1).ToString()
                                Catch ex As Exception
                                    com.Parameters.Add("@PHCode", SqlDbType.VarChar).Value = ""
                                End Try

                                Try
                                    com.Parameters.Add("@Agronomist", SqlDbType.VarChar).Value = dttempMain(7)(1).ToString()
                                Catch ex As Exception
                                    com.Parameters.Add("@Agronomist", SqlDbType.VarChar).Value = ""
                                End Try

                                Try
                                    com.Parameters.Add("@Rep", SqlDbType.VarChar).Value = dttempEnd(1)(1).ToString()
                                Catch ex As Exception
                                    com.Parameters.Add("@Rep", SqlDbType.VarChar).Value = ""
                                End Try

                                Try
                                    com.Parameters.Add("@Others", SqlDbType.VarChar).Value = dttempTotal(5)(1).ToString()
                                Catch ex As Exception
                                    com.Parameters.Add("@Others", SqlDbType.VarChar).Value = ""
                                End Try

                                Try
                                    com.Parameters.Add("@Remarks", SqlDbType.VarChar).Value = dttempTotal(5)(1).ToString()
                                Catch ex As Exception
                                    com.Parameters.Add("@Remarks", SqlDbType.VarChar).Value = ""
                                End Try

                                com.Parameters.Add("@Id", SqlDbType.BigInt).Value = a("Id").ToString()
                                com.Parameters.Add("@SubmissionDateTime", SqlDbType.DateTime).Value = a("Date").ToString()
                                com.ExecuteNonQuery()

                                trueCounter += 1
                                trans.Commit()

                            Catch ex As Exception
                                trans.Rollback()
                                Status += ex.Message + Environment.NewLine
                                Try
                                    Status += ex.InnerException.Message + Environment.NewLine
                                    Status += ex.InnerException.InnerException.Message + Environment.NewLine
                                Catch e As Exception

                                End Try
                            Finally
                                conn.Close()
                            End Try

                        End If

                    End If

                    'Check if submission id exists
                    If Not CheckIfNew(a("Id").ToString()) Then


                        Dim farmtemp = dttempMain(4)(1).ToString()
                        Dim weeknotemp = Convert.ToInt32(dttempMain(2)(1).ToString())
                        Dim yeartemp = Convert.ToInt32(dttempMain(0)(1).ToString().Split("/")(2))

                        Dim activities() As String = {"BI", "BBS - FUNGICIDE", "BBS - INSECTICIDE", "MSSI", "BTEI"}
                        Dim transheaddetailList As New List(Of String)()

                        Dim dtHeaderSysid As DataTable = FetchDataTable("SELECT hdrsysid FROM tblTransHeader WHERE SubmissionID = '" & a("Id").ToString() & "'", Properties.FCDBConnstring)
                        Dim submissionIdTemp As String = dtHeaderSysid(0)(0).ToString()

                        'Transaction and Header detail join reference to retrieve activity missing
                        Dim dtTransHeadDetail As DataTable = FetchDataTable("SELECT th.hdrsysid,td.activity FROM tblTransHeader th " _
                                                                            & "INNER JOIN tblTransDetail td ON " _
                                                                            & "td.hdrsysid = th.hdrsysid" _
                                                                             & " WHERE td.hdrsysid= '" & submissionIdTemp & "'",
                                                                            Properties.FCDBConnstring)

                        conn = New SqlConnection(Properties.FCDBConnstring)
                        conn.Open()
                        Dim trans As SqlTransaction = conn.BeginTransaction()

                        Try

                            If farmtemp = Properties.FarmParam And weeknotemp = Properties.WeekParam And yeartemp = Properties.YearParam Then

                                'loop through transheaddetail and add it into array for checking activity
                                For Each row As DataRow In dtTransHeadDetail.Rows
                                    transheaddetailList.Add(row("activity"))
                                Next

                                'check activity that does not exists
                                For Each activity As String In activities
                                    If Not transheaddetailList.Contains(activity) Then

                                        If activity = "BI" Then

                                            com = New SqlCommand("sp_UploadTransDetails_v2", conn, trans)
                                            com.CommandType = CommandType.StoredProcedure
                                            com.CommandTimeout = 999999
                                            com.Parameters.Add("@hdrsysid", SqlDbType.VarChar).Value = submissionIdTemp
                                            com.Parameters.Add("@activity", SqlDbType.VarChar).Value = "BI"
                                            com.Parameters.Add("@brand", SqlDbType.VarChar, 50).Value = dttempBI(0)(1).ToString()
                                            com.Parameters.Add("@active", SqlDbType.VarChar, 50).Value = dttempBI(1)(1).ToString()
                                            com.Parameters.Add("@recomm", SqlDbType.VarChar).Value = dttempBI(3)(1).ToString()
                                            com.Parameters.Add("@arate", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(4)(1).ToString())
                                            com.Parameters.Add("@arateUOM", SqlDbType.VarChar, 50).Value = dttempBI(2)(1).ToString()
                                            com.Parameters.Add("@rate", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(5)(1).ToString())
                                            com.Parameters.Add("@brs", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(6)(1).ToString())
                                            com.Parameters.Add("@premix", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(7)(1).ToString())
                                            com.Parameters.Add("@agitation", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(8)(1).ToString())
                                            com.Parameters.Add("@volcal", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(9)(1).ToString())
                                            com.Parameters.Add("@waste", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(10)(1).ToString())
                                            com.Parameters.Add("@mixing", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(11)(1).ToString())
                                            com.Parameters.Add("@cleaning", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(13)(1).ToString())
                                            com.Parameters.Add("@signages", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(14)(1).ToString())
                                            com.Parameters.Add("@ppe", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(15)(1).ToString())
                                            com.Parameters.Add("@remarks", SqlDbType.VarChar).Value = dttempBI(16)(1).ToString()
                                            com.Parameters.Add("@chembodega", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(12)(1).ToString())
                                            com.ExecuteNonQuery()

                                            If ConvertToDoubleElseZero(dttempBI(4)(1).ToString()) = "0" And dttempBI(0)(1).ToString() = "BUD BAGGING" Then
                                                trueCounter += 1
                                            ElseIf ConvertToDoubleElseZero(dttempBI(4)(1).ToString()) <> "0" Then
                                                trueCounter += 1
                                            ElseIf ConvertToDoubleElseZero(dttempBI(4)(1).ToString()) = "0" Then

                                            End If



                                        ElseIf activity = "BBS - FUNGICIDE" Then

                                            com = New SqlCommand("sp_UploadTransDetails_v2", conn, trans)
                                            com.CommandType = CommandType.StoredProcedure
                                            com.CommandTimeout = 999999
                                            com.Parameters.Add("@hdrsysid", SqlDbType.VarChar).Value = submissionIdTemp
                                            com.Parameters.Add("@activity", SqlDbType.VarChar).Value = "BBS - FUNGICIDE"
                                            com.Parameters.Add("@brand", SqlDbType.VarChar, 50).Value = dttempBBSF(0)(1).ToString()
                                            com.Parameters.Add("@active", SqlDbType.VarChar, 50).Value = dttempBBSF(1)(1).ToString()
                                            com.Parameters.Add("@recomm", SqlDbType.VarChar).Value = dttempBBSF(3)(1).ToString()
                                            com.Parameters.Add("@arate", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(4)(1).ToString())
                                            com.Parameters.Add("@arateUOM", SqlDbType.VarChar, 50).Value = dttempBBSF(2)(1).ToString()
                                            com.Parameters.Add("@rate", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(5)(1).ToString())
                                            com.Parameters.Add("@brs", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(6)(1).ToString())
                                            com.Parameters.Add("@premix", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(7)(1).ToString())
                                            com.Parameters.Add("@agitation", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(8)(1).ToString())
                                            com.Parameters.Add("@volcal", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(9)(1).ToString())
                                            com.Parameters.Add("@waste", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(10)(1).ToString())
                                            com.Parameters.Add("@mixing", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(11)(1).ToString())
                                            com.Parameters.Add("@cleaning", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(13)(1).ToString())
                                            com.Parameters.Add("@signages", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(14)(1).ToString())
                                            com.Parameters.Add("@ppe", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(15)(1).ToString())
                                            com.Parameters.Add("@remarks", SqlDbType.VarChar).Value = dttempBBSF(16)(1).ToString()
                                            com.Parameters.Add("@chembodega", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(12)(1).ToString())
                                            com.ExecuteNonQuery()

                                            If ConvertToDoubleElseZero(dttempBBSF(4)(1).ToString()) <> "0" Then
                                                trueCounter += 1
                                            End If

                                        ElseIf activity = "BBS - INSECTICIDE" Then

                                            com = New SqlCommand("sp_UploadTransDetails_v2", conn, trans)
                                            com.CommandType = CommandType.StoredProcedure
                                            com.CommandTimeout = 999999
                                            com.Parameters.Add("@hdrsysid", SqlDbType.VarChar).Value = submissionIdTemp
                                            com.Parameters.Add("@activity", SqlDbType.VarChar).Value = "BBS - INSECTICIDE"
                                            com.Parameters.Add("@brand", SqlDbType.VarChar, 50).Value = dttempBBSI(0)(1).ToString()
                                            com.Parameters.Add("@active", SqlDbType.VarChar, 50).Value = dttempBBSI(1)(1).ToString()
                                            com.Parameters.Add("@recomm", SqlDbType.VarChar).Value = dttempBBSI(3)(1).ToString()
                                            com.Parameters.Add("@arate", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(4)(1).ToString())
                                            com.Parameters.Add("@arateUOM", SqlDbType.VarChar, 50).Value = dttempBBSI(2)(1).ToString()
                                            com.Parameters.Add("@rate", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(5)(1).ToString())
                                            com.Parameters.Add("@brs", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(6)(1).ToString())
                                            com.Parameters.Add("@premix", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(7)(1).ToString())
                                            com.Parameters.Add("@agitation", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(8)(1).ToString())
                                            com.Parameters.Add("@volcal", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(9)(1).ToString())
                                            com.Parameters.Add("@waste", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(10)(1).ToString())
                                            com.Parameters.Add("@mixing", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(11)(1).ToString())
                                            com.Parameters.Add("@cleaning", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(13)(1).ToString())
                                            com.Parameters.Add("@signages", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(14)(1).ToString())
                                            com.Parameters.Add("@ppe", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(15)(1).ToString())
                                            com.Parameters.Add("@remarks", SqlDbType.VarChar).Value = dttempBBSI(16)(1).ToString()
                                            com.Parameters.Add("@chembodega", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(12)(1).ToString())
                                            com.ExecuteNonQuery()

                                            If ConvertToDoubleElseZero(dttempBBSI(4)(1).ToString()) <> "0" Then
                                                trueCounter += 1
                                            End If

                                        ElseIf activity = "MSSI" Then

                                            com = New SqlCommand("sp_UploadTransDetails_v2", conn, trans)
                                            com.CommandType = CommandType.StoredProcedure
                                            com.CommandTimeout = 999999
                                            com.Parameters.Add("@hdrsysid", SqlDbType.VarChar).Value = submissionIdTemp
                                            com.Parameters.Add("@activity", SqlDbType.VarChar).Value = "MSSI"
                                            com.Parameters.Add("@brand", SqlDbType.VarChar, 50).Value = dttempMSSI(0)(1).ToString()
                                            com.Parameters.Add("@active", SqlDbType.VarChar, 50).Value = dttempMSSI(1)(1).ToString()
                                            com.Parameters.Add("@recomm", SqlDbType.VarChar).Value = dttempMSSI(3)(1).ToString()
                                            com.Parameters.Add("@arate", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(4)(1).ToString())
                                            com.Parameters.Add("@arateUOM", SqlDbType.VarChar, 50).Value = dttempMSSI(2)(1).ToString()
                                            com.Parameters.Add("@rate", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(5)(1).ToString())
                                            com.Parameters.Add("@brs", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(6)(1).ToString())
                                            com.Parameters.Add("@premix", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(7)(1).ToString())
                                            com.Parameters.Add("@agitation", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(8)(1).ToString())
                                            com.Parameters.Add("@volcal", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(9)(1).ToString())
                                            com.Parameters.Add("@waste", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(10)(1).ToString())
                                            com.Parameters.Add("@mixing", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(11)(1).ToString())
                                            com.Parameters.Add("@cleaning", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(13)(1).ToString())
                                            com.Parameters.Add("@signages", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(14)(1).ToString())
                                            com.Parameters.Add("@ppe", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(15)(1).ToString())
                                            com.Parameters.Add("@remarks", SqlDbType.VarChar).Value = dttempMSSI(16)(1).ToString()
                                            com.Parameters.Add("@chembodega", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(12)(1).ToString())
                                            com.ExecuteNonQuery()

                                            If ConvertToDoubleElseZero(dttempMSSI(4)(1).ToString()) <> "0" Then
                                                trueCounter += 1
                                            End If

                                        ElseIf activity = "BTEI" Then

                                            com = New SqlCommand("sp_UploadTransDetails_v2", conn, trans)
                                            com.CommandType = CommandType.StoredProcedure
                                            com.CommandTimeout = 999999
                                            com.Parameters.Add("@hdrsysid", SqlDbType.VarChar).Value = submissionIdTemp
                                            com.Parameters.Add("@activity", SqlDbType.VarChar).Value = "BTEI"
                                            com.Parameters.Add("@brand", SqlDbType.VarChar, 50).Value = dttempBTEI(0)(1).ToString()
                                            com.Parameters.Add("@active", SqlDbType.VarChar, 50).Value = dttempBTEI(1)(1).ToString()
                                            com.Parameters.Add("@recomm", SqlDbType.VarChar).Value = dttempBTEI(3)(1).ToString()
                                            com.Parameters.Add("@arate", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(4)(1).ToString())
                                            com.Parameters.Add("@arateUOM", SqlDbType.VarChar, 50).Value = dttempBTEI(2)(1).ToString()
                                            com.Parameters.Add("@rate", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(5)(1).ToString())
                                            com.Parameters.Add("@brs", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(6)(1).ToString())
                                            com.Parameters.Add("@premix", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(7)(1).ToString())
                                            com.Parameters.Add("@agitation", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(8)(1).ToString())
                                            com.Parameters.Add("@volcal", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(9)(1).ToString())
                                            com.Parameters.Add("@waste", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(10)(1).ToString())
                                            com.Parameters.Add("@mixing", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(11)(1).ToString())
                                            com.Parameters.Add("@cleaning", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(13)(1).ToString())
                                            com.Parameters.Add("@signages", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(14)(1).ToString())
                                            com.Parameters.Add("@ppe", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(15)(1).ToString())
                                            com.Parameters.Add("@remarks", SqlDbType.VarChar).Value = dttempBTEI(16)(1).ToString()
                                            com.Parameters.Add("@chembodega", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(12)(1).ToString())
                                            com.ExecuteNonQuery()

                                            If ConvertToDoubleElseZero(dttempBTEI(4)(1).ToString()) <> "0" Then
                                                trueCounter += 1
                                            End If

                                        End If

                                    End If


                                Next
                                IsSaved = True
                                trans.Commit()

                            End If

                        Catch ex As Exception
                            trans.Rollback()

                            Status += ex.InnerException.Message + Environment.NewLine
                            Status += ex.InnerException.InnerException.Message + Environment.NewLine

                        Finally
                            conn.Close()
                        End Try

                        If (Properties.IsSaveSignature) Then
                            Dim Client As New WebClient
                            Try
                                Client.DownloadFile("https://www.gocanvas.com/apiv2/images.xml?image_id=" + dttempEnd.Rows(2)(1).ToString() + "&username=" + Properties.Username + "&password=" + Properties.Password, Properties.SignaturePath + "/" + dttempEnd(1)(1).ToString() + ".jpg.")
                                Client.DownloadFile("https://www.gocanvas.com/apiv2/images.xml?image_id=" + dttempEnd.Rows(4)(1).ToString() + "&username=" + Properties.Username + "&password=" + Properties.Password, Properties.SignaturePath + "/" + dttempEnd(3)(1).ToString() + ".jpg.")
                                Client.Dispose()
                            Catch ex As Exception
                            End Try
                        End If

                    End If
                Next

            End If
        Catch ex As Exception
            Status += ex.Message + Environment.NewLine
            Try
                Status += ex.InnerException.Message + Environment.NewLine
                Status += ex.InnerException.InnerException.Message + Environment.NewLine
            Catch e As Exception

            End Try
        End Try

        'Status += "FCMS: Data from " + DateTime.Now.AddDays(-1).ToShortDateString() + " to " + DateTime.Now.ToString() + "." + Environment.NewLine
        Status += trueCounter.ToString() + " record/s succeed." + Environment.NewLine
        Status += falseCounter.ToString() + " record/s failed." + Environment.NewLine
        Return Status

        Return Status


    End Function

    Private Function RunAllBacktrack(ByVal ds As DataSet) As String
        Dim Status As String = String.Empty
        Dim dt As New DataTable
        Dim trueCounter As Integer = 0, falseCounter As Integer = 0

        Dim IsSaved As Boolean

        Try
            If ds.Tables.Count > 1 Then

                For Each a As DataRow In ds.Tables("Submission").Rows

                    IsSaved = False 'reference for signature

                    Dim tempRecord() As DataRow = ds.Tables("Section").Select("Sections_Id = " + a("Submission_Id").ToString())
                    Dim dttempMain As DataTable = ds.Tables("Response").Select("Responses_Id = " + tempRecord(0)("Section_Id").ToString()).CopyToDataTable()
                    Dim dttempBI As DataTable = ds.Tables("Response").Select("Responses_Id = " + tempRecord(1)("Section_Id").ToString()).CopyToDataTable()
                    Dim dttempBBSF As DataTable = ds.Tables("Response").Select("Responses_Id = " + tempRecord(2)("Section_Id").ToString()).CopyToDataTable()
                    Dim dttempBBSI As DataTable = ds.Tables("Response").Select("Responses_Id = " + tempRecord(3)("Section_Id").ToString()).CopyToDataTable()
                    Dim dttempMSSI As DataTable = ds.Tables("Response").Select("Responses_Id = " + tempRecord(4)("Section_Id").ToString()).CopyToDataTable()
                    Dim dttempBTEI As DataTable = ds.Tables("Response").Select("Responses_Id = " + tempRecord(5)("Section_Id").ToString()).CopyToDataTable()
                    Dim dttempTotal As DataTable = ds.Tables("Response").Select("Responses_Id = " + tempRecord(6)("Section_Id").ToString()).CopyToDataTable()
                    Dim dttempEnd As DataTable = ds.Tables("Response").Select("Responses_Id = " + tempRecord(7)("Section_Id").ToString()).CopyToDataTable()

                    'Check if submission id is new
                    If CheckIfNew(a("Id").ToString()) Then

                        conn = New SqlConnection(Properties.FCDBConnstring)
                        conn.Open()
                        Dim trans As SqlTransaction = conn.BeginTransaction()

                        Try
                            com = New SqlCommand("sp_UploadTransHeader", conn, trans)
                            com.CommandType = CommandType.StoredProcedure
                            com.CommandTimeout = 999999
                            com.Parameters.Add("@Transdate", SqlDbType.VarChar).Value = dttempMain(0)(1).ToString()
                            Try
                                com.Parameters.Add("@Transtime", SqlDbType.VarChar).Value = dttempMain(1)(1).ToString()
                            Catch ex As Exception
                                com.Parameters.Add("@Transtime", SqlDbType.VarChar).Value = ""
                            End Try

                            Try
                                com.Parameters.Add("@Week", SqlDbType.Int).Value = dttempMain(2)(1).ToString()
                            Catch ex As Exception
                                com.Parameters.Add("@Week", SqlDbType.Int).Value = Nothing
                            End Try

                            Try
                                com.Parameters.Add("@Farm", SqlDbType.VarChar).Value = dttempMain(4)(1).ToString()
                            Catch ex As Exception
                                com.Parameters.Add("@Farm", SqlDbType.VarChar).Value = ""
                            End Try

                            Try
                                com.Parameters.Add("@Grower", SqlDbType.VarChar).Value = dttempMain(5)(1).ToString()
                            Catch ex As Exception
                                com.Parameters.Add("@Grower", SqlDbType.VarChar).Value = ""
                            End Try

                            Try
                                com.Parameters.Add("@PHCode", SqlDbType.VarChar).Value = dttempMain(6)(1).ToString()
                            Catch ex As Exception
                                com.Parameters.Add("@PHCode", SqlDbType.VarChar).Value = ""
                            End Try

                            Try
                                com.Parameters.Add("@Agronomist", SqlDbType.VarChar).Value = dttempMain(7)(1).ToString()
                            Catch ex As Exception
                                com.Parameters.Add("@Agronomist", SqlDbType.VarChar).Value = ""
                            End Try

                            Try
                                com.Parameters.Add("@Rep", SqlDbType.VarChar).Value = dttempEnd(1)(1).ToString()
                            Catch ex As Exception
                                com.Parameters.Add("@Rep", SqlDbType.VarChar).Value = ""
                            End Try

                            Try
                                com.Parameters.Add("@Others", SqlDbType.VarChar).Value = dttempTotal(5)(1).ToString()
                            Catch ex As Exception
                                com.Parameters.Add("@Others", SqlDbType.VarChar).Value = ""
                            End Try

                            Try
                                com.Parameters.Add("@Remarks", SqlDbType.VarChar).Value = dttempTotal(5)(1).ToString()
                            Catch ex As Exception
                                com.Parameters.Add("@Remarks", SqlDbType.VarChar).Value = ""
                            End Try

                            com.Parameters.Add("@Id", SqlDbType.BigInt).Value = a("Id").ToString()
                            com.Parameters.Add("@SubmissionDateTime", SqlDbType.DateTime).Value = a("Date").ToString()
                            com.ExecuteNonQuery()

                            trueCounter += 1
                            trans.Commit()

                        Catch ex As Exception
                            trans.Rollback()
                            Status += ex.Message + Environment.NewLine
                            Try
                                Status += ex.InnerException.Message + Environment.NewLine
                                Status += ex.InnerException.InnerException.Message + Environment.NewLine
                            Catch e As Exception

                            End Try
                        Finally
                            conn.Close()
                        End Try

                    End If


                    'If submission id exists
                    If Not CheckIfNew(a("Id").ToString()) Then

                        Dim farmtemp = dttempMain(4)(1).ToString()
                        Dim weeknotemp = Convert.ToInt32(dttempMain(2)(1).ToString())
                        Dim yeartemp = Convert.ToInt32(dttempMain(0)(1).ToString().Split("/")(2))

                        Dim activities() As String = {"BI", "BBS - FUNGICIDE", "BBS - INSECTICIDE", "MSSI", "BTEI"}
                        Dim transheaddetailList As New List(Of String)()

                        Dim dtHeaderSysid As DataTable = FetchDataTable("SELECT hdrsysid FROM tblTransHeader WHERE SubmissionID = '" & a("Id").ToString() & "'", Properties.FCDBConnstring)
                        Dim submissionIdTemp As String = dtHeaderSysid(0)(0).ToString()

                        'Transaction and Header detail join reference to retrieve activity missing
                        Dim dtTransHeadDetail As DataTable = FetchDataTable("SELECT th.hdrsysid,td.activity FROM tblTransHeader th " _
                                                                            & "INNER JOIN tblTransDetail td ON " _
                                                                            & "td.hdrsysid = th.hdrsysid" _
                                                                             & " WHERE td.hdrsysid= '" & submissionIdTemp & "'",
                                                                            Properties.FCDBConnstring)

                        conn = New SqlConnection(Properties.FCDBConnstring)
                        conn.Open()
                        Dim trans As SqlTransaction = conn.BeginTransaction()

                        Try

                            'loop through transheaddetail and add it into array for checking activity
                            For Each row As DataRow In dtTransHeadDetail.Rows
                                transheaddetailList.Add(row("activity"))
                            Next

                            'check activity that does not exists
                            For Each activity As String In activities
                                If Not transheaddetailList.Contains(activity) Then

                                    If activity = "BI" Then

                                        com = New SqlCommand("sp_UploadTransDetails_v2", conn, trans)
                                        com.CommandType = CommandType.StoredProcedure
                                        com.CommandTimeout = 999999
                                        com.Parameters.Add("@hdrsysid", SqlDbType.VarChar).Value = submissionIdTemp
                                        com.Parameters.Add("@activity", SqlDbType.VarChar).Value = "BI"
                                        com.Parameters.Add("@brand", SqlDbType.VarChar, 50).Value = dttempBI(0)(1).ToString()
                                        com.Parameters.Add("@active", SqlDbType.VarChar, 50).Value = dttempBI(1)(1).ToString()
                                        com.Parameters.Add("@recomm", SqlDbType.VarChar).Value = dttempBI(3)(1).ToString()
                                        com.Parameters.Add("@arate", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(4)(1).ToString())
                                        com.Parameters.Add("@arateUOM", SqlDbType.VarChar, 50).Value = dttempBI(2)(1).ToString()
                                        com.Parameters.Add("@rate", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(5)(1).ToString())
                                        com.Parameters.Add("@brs", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(6)(1).ToString())
                                        com.Parameters.Add("@premix", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(7)(1).ToString())
                                        com.Parameters.Add("@agitation", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(8)(1).ToString())
                                        com.Parameters.Add("@volcal", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(9)(1).ToString())
                                        com.Parameters.Add("@waste", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(10)(1).ToString())
                                        com.Parameters.Add("@mixing", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(11)(1).ToString())
                                        com.Parameters.Add("@cleaning", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(13)(1).ToString())
                                        com.Parameters.Add("@signages", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(14)(1).ToString())
                                        com.Parameters.Add("@ppe", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(15)(1).ToString())
                                        com.Parameters.Add("@remarks", SqlDbType.VarChar).Value = dttempBI(16)(1).ToString()
                                        com.Parameters.Add("@chembodega", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBI(12)(1).ToString())
                                        com.ExecuteNonQuery()

                                        If ConvertToDoubleElseZero(dttempBI(4)(1).ToString()) = "0" And dttempBI(0)(1).ToString() = "BUD BAGGING" Then
                                            trueCounter += 1
                                        ElseIf ConvertToDoubleElseZero(dttempBI(4)(1).ToString()) <> "0" Then
                                            trueCounter += 1
                                        ElseIf ConvertToDoubleElseZero(dttempBI(4)(1).ToString()) = "0" Then

                                        End If

                                    ElseIf activity = "BBS - FUNGICIDE" Then

                                        com = New SqlCommand("sp_UploadTransDetails_v2", conn, trans)
                                        com.CommandType = CommandType.StoredProcedure
                                        com.CommandTimeout = 999999
                                        com.Parameters.Add("@hdrsysid", SqlDbType.VarChar).Value = submissionIdTemp
                                        com.Parameters.Add("@activity", SqlDbType.VarChar).Value = "BBS - FUNGICIDE"
                                        com.Parameters.Add("@brand", SqlDbType.VarChar, 50).Value = dttempBBSF(0)(1).ToString()
                                        com.Parameters.Add("@active", SqlDbType.VarChar, 50).Value = dttempBBSF(1)(1).ToString()
                                        com.Parameters.Add("@recomm", SqlDbType.VarChar).Value = dttempBBSF(3)(1).ToString()
                                        com.Parameters.Add("@arate", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(4)(1).ToString())
                                        com.Parameters.Add("@arateUOM", SqlDbType.VarChar, 50).Value = dttempBBSF(2)(1).ToString()
                                        com.Parameters.Add("@rate", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(5)(1).ToString())
                                        com.Parameters.Add("@brs", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(6)(1).ToString())
                                        com.Parameters.Add("@premix", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(7)(1).ToString())
                                        com.Parameters.Add("@agitation", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(8)(1).ToString())
                                        com.Parameters.Add("@volcal", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(9)(1).ToString())
                                        com.Parameters.Add("@waste", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(10)(1).ToString())
                                        com.Parameters.Add("@mixing", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(11)(1).ToString())
                                        com.Parameters.Add("@cleaning", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(13)(1).ToString())
                                        com.Parameters.Add("@signages", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(14)(1).ToString())
                                        com.Parameters.Add("@ppe", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(15)(1).ToString())
                                        com.Parameters.Add("@remarks", SqlDbType.VarChar).Value = dttempBBSF(16)(1).ToString()
                                        com.Parameters.Add("@chembodega", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSF(12)(1).ToString())
                                        com.ExecuteNonQuery()

                                        If ConvertToDoubleElseZero(dttempBBSF(4)(1).ToString()) <> "0" Then
                                            trueCounter += 1
                                        End If



                                    ElseIf activity = "BBS - INSECTICIDE" Then

                                        com = New SqlCommand("sp_UploadTransDetails_v2", conn, trans)
                                        com.CommandType = CommandType.StoredProcedure
                                        com.CommandTimeout = 999999
                                        com.Parameters.Add("@hdrsysid", SqlDbType.VarChar).Value = submissionIdTemp
                                        com.Parameters.Add("@activity", SqlDbType.VarChar).Value = "BBS - INSECTICIDE"
                                        com.Parameters.Add("@brand", SqlDbType.VarChar, 50).Value = dttempBBSI(0)(1).ToString()
                                        com.Parameters.Add("@active", SqlDbType.VarChar, 50).Value = dttempBBSI(1)(1).ToString()
                                        com.Parameters.Add("@recomm", SqlDbType.VarChar).Value = dttempBBSI(3)(1).ToString()
                                        com.Parameters.Add("@arate", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(4)(1).ToString())
                                        com.Parameters.Add("@arateUOM", SqlDbType.VarChar, 50).Value = dttempBBSI(2)(1).ToString()
                                        com.Parameters.Add("@rate", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(5)(1).ToString())
                                        com.Parameters.Add("@brs", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(6)(1).ToString())
                                        com.Parameters.Add("@premix", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(7)(1).ToString())
                                        com.Parameters.Add("@agitation", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(8)(1).ToString())
                                        com.Parameters.Add("@volcal", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(9)(1).ToString())
                                        com.Parameters.Add("@waste", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(10)(1).ToString())
                                        com.Parameters.Add("@mixing", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(11)(1).ToString())
                                        com.Parameters.Add("@cleaning", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(13)(1).ToString())
                                        com.Parameters.Add("@signages", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(14)(1).ToString())
                                        com.Parameters.Add("@ppe", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(15)(1).ToString())
                                        com.Parameters.Add("@remarks", SqlDbType.VarChar).Value = dttempBBSI(16)(1).ToString()
                                        com.Parameters.Add("@chembodega", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBBSI(12)(1).ToString())
                                        com.ExecuteNonQuery()

                                        If ConvertToDoubleElseZero(dttempBBSI(4)(1).ToString()) <> "0" Then
                                            trueCounter += 1
                                        End If

                                    ElseIf activity = "MSSI" Then

                                        com = New SqlCommand("sp_UploadTransDetails_v2", conn, trans)
                                        com.CommandType = CommandType.StoredProcedure
                                        com.CommandTimeout = 999999
                                        com.Parameters.Add("@hdrsysid", SqlDbType.VarChar).Value = submissionIdTemp
                                        com.Parameters.Add("@activity", SqlDbType.VarChar).Value = "MSSI"
                                        com.Parameters.Add("@brand", SqlDbType.VarChar, 50).Value = dttempMSSI(0)(1).ToString()
                                        com.Parameters.Add("@active", SqlDbType.VarChar, 50).Value = dttempMSSI(1)(1).ToString()
                                        com.Parameters.Add("@recomm", SqlDbType.VarChar).Value = dttempMSSI(3)(1).ToString()
                                        com.Parameters.Add("@arate", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(4)(1).ToString())
                                        com.Parameters.Add("@arateUOM", SqlDbType.VarChar, 50).Value = dttempMSSI(2)(1).ToString()
                                        com.Parameters.Add("@rate", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(5)(1).ToString())
                                        com.Parameters.Add("@brs", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(6)(1).ToString())
                                        com.Parameters.Add("@premix", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(7)(1).ToString())
                                        com.Parameters.Add("@agitation", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(8)(1).ToString())
                                        com.Parameters.Add("@volcal", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(9)(1).ToString())
                                        com.Parameters.Add("@waste", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(10)(1).ToString())
                                        com.Parameters.Add("@mixing", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(11)(1).ToString())
                                        com.Parameters.Add("@cleaning", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(13)(1).ToString())
                                        com.Parameters.Add("@signages", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(14)(1).ToString())
                                        com.Parameters.Add("@ppe", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(15)(1).ToString())
                                        com.Parameters.Add("@remarks", SqlDbType.VarChar).Value = dttempMSSI(16)(1).ToString()
                                        com.Parameters.Add("@chembodega", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempMSSI(12)(1).ToString())
                                        com.ExecuteNonQuery()

                                        If ConvertToDoubleElseZero(dttempMSSI(4)(1).ToString()) <> "0" Then
                                            trueCounter += 1
                                        End If


                                    ElseIf activity = "BTEI" Then

                                        com = New SqlCommand("sp_UploadTransDetails_v2", conn, trans)
                                        com.CommandType = CommandType.StoredProcedure
                                        com.CommandTimeout = 999999
                                        com.Parameters.Add("@hdrsysid", SqlDbType.VarChar).Value = submissionIdTemp
                                        com.Parameters.Add("@activity", SqlDbType.VarChar).Value = "BTEI"
                                        com.Parameters.Add("@brand", SqlDbType.VarChar, 50).Value = dttempBTEI(0)(1).ToString()
                                        com.Parameters.Add("@active", SqlDbType.VarChar, 50).Value = dttempBTEI(1)(1).ToString()
                                        com.Parameters.Add("@recomm", SqlDbType.VarChar).Value = dttempBTEI(3)(1).ToString()
                                        com.Parameters.Add("@arate", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(4)(1).ToString())
                                        com.Parameters.Add("@arateUOM", SqlDbType.VarChar, 50).Value = dttempBTEI(2)(1).ToString()
                                        com.Parameters.Add("@rate", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(5)(1).ToString())
                                        com.Parameters.Add("@brs", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(6)(1).ToString())
                                        com.Parameters.Add("@premix", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(7)(1).ToString())
                                        com.Parameters.Add("@agitation", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(8)(1).ToString())
                                        com.Parameters.Add("@volcal", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(9)(1).ToString())
                                        com.Parameters.Add("@waste", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(10)(1).ToString())
                                        com.Parameters.Add("@mixing", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(11)(1).ToString())
                                        com.Parameters.Add("@cleaning", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(13)(1).ToString())
                                        com.Parameters.Add("@signages", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(14)(1).ToString())
                                        com.Parameters.Add("@ppe", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(15)(1).ToString())
                                        com.Parameters.Add("@remarks", SqlDbType.VarChar).Value = dttempBTEI(16)(1).ToString()
                                        com.Parameters.Add("@chembodega", SqlDbType.VarChar).Value = ConvertToDoubleElseZero(dttempBTEI(12)(1).ToString())
                                        com.ExecuteNonQuery()

                                        If ConvertToDoubleElseZero(dttempBTEI(4)(1).ToString()) <> "0" Then
                                            trueCounter += 1
                                        End If

                                    End If

                                End If

                            Next
                            IsSaved = True
                            trans.Commit()


                        Catch ex As Exception
                            trans.Rollback()

                            Status += ex.InnerException.Message + Environment.NewLine
                            Status += ex.InnerException.InnerException.Message + Environment.NewLine

                        Finally
                            conn.Close()
                        End Try

                        If (Properties.IsSaveSignature) Then
                            Dim Client As New WebClient
                            Try
                                Client.DownloadFile("https://www.gocanvas.com/apiv2/images.xml?image_id=" + dttempEnd.Rows(2)(1).ToString() + "&username=" + Properties.Username + "&password=" + Properties.Password, Properties.SignaturePath + "/" + dttempEnd(1)(1).ToString() + ".jpg.")
                                Client.DownloadFile("https://www.gocanvas.com/apiv2/images.xml?image_id=" + dttempEnd.Rows(4)(1).ToString() + "&username=" + Properties.Username + "&password=" + Properties.Password, Properties.SignaturePath + "/" + dttempEnd(3)(1).ToString() + ".jpg.")
                                Client.Dispose()
                            Catch ex As Exception
                            End Try
                        End If

                    End If
                Next

            End If
        Catch ex As Exception
            Status += ex.Message + Environment.NewLine
            Try
                Status += ex.InnerException.Message + Environment.NewLine
                Status += ex.InnerException.InnerException.Message + Environment.NewLine
            Catch e As Exception

            End Try
        End Try

        'Status += "FCMS: Data from " + DateTime.Now.AddDays(-1).ToShortDateString() + " to " + DateTime.Now.ToString() + "." + Environment.NewLine
        Status += trueCounter.ToString() + " record/s succeed." + Environment.NewLine
        Status += falseCounter.ToString() + " record/s failed." + Environment.NewLine
        Return Status

        Return Status



    End Function

#Region "Methods"

    Private Function FetchDataTable(ByVal query As String, ByVal dbConn As String) As DataTable
        Try
            conn = New SqlConnection(dbConn)
            dt = New DataTable()
            conn.Open()
            com = New SqlCommand(query, conn)
            com.CommandTimeout = 999999
            adap = New SqlDataAdapter(com)
            adap.Fill(dt)
            conn.Close()
        Catch
            conn.Close()
            dt = Nothing
        End Try

        Return dt
    End Function

    Private Function ConvertToDoubleElseZero(ByVal param As String) As String
        Try
            Return Convert.ToDouble(param).ToString()
        Catch

        End Try

        Return "0"
    End Function


    Private Function CheckIfNew(ByVal param) As Boolean

        Dim dt As DataTable = FetchDataTable("SELECT COUNT(*) FROM tblTransHeader WHERE SubmissionID = '" + param + "'", Properties.FCDBConnstring)
        If (Convert.ToInt32(dt.Rows(0)(0).ToString()) > 0) Then
            Return False
        End If

        Return True
    End Function

#End Region

End Class


Public Class NFLSurvey
    Private Shared conn As SqlConnection
    Private Shared dt As DataTable
    Private Shared ds As DataSet
    Private Shared adap As SqlDataAdapter
    Private Shared com As SqlCommand
    Private Shared start As DateTime

    Public Function ProcessNFLSurvey(ByVal datefrom As Date, ByVal dateto As Date) As String
        Dim Status As String = String.Empty
        start = DateTime.Now
        Try
            Dim web As New WebClient()

            Dim url As String = String.Format("https://www.gocanvas.com/apiv2/submissions.xml?username=" + Properties.Username +
                                              "&password=" + Properties.Password +
                                              "&form_name=NFL%20SURVEY%20r3&begin_date=" + datefrom +
                                              "&end_date=" + dateto)

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls Or SecurityProtocolType.Ssl3
            Dim response As String = web.DownloadString(url)
            Dim ds As New DataSet()
            Using stringReader As New StringReader(response)
                ds = New DataSet()
                ds.ReadXml(stringReader)
            End Using
            Status += FormatToSaveNFLSurvey(ds)

        Catch ex As Exception
            Status += ex.Message + Environment.NewLine
            Status += "Error at method ProcessNFLSurvey()" + Environment.NewLine
            Try
                Status += ex.InnerException.Message + Environment.NewLine
                Status += ex.InnerException.InnerException.Message + Environment.NewLine
            Catch e As Exception

            End Try
        Finally
            Status += "Process started at " + start + " and ended at " + DateTime.Now + Environment.NewLine
            Status += "---" + Environment.NewLine
        End Try
        Return Status
    End Function

    Public Function ProcessAllNFLSurvey(ByVal datefrom As Date, ByVal dateto As Date) As String
        Dim Status As String = String.Empty
        start = DateTime.Now
        Try
            Dim web As New WebClient()

            Dim url As String = String.Format("https://www.gocanvas.com/apiv2/submissions.xml?username=" + Properties.Username +
                                              "&password=" + Properties.Password +
                                              "&form_name=NFL%20SURVEY%20r3&begin_date=" + datefrom +
                                              "&end_date=" + dateto)

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls Or SecurityProtocolType.Ssl3
            Dim response As String = web.DownloadString(url)
            Dim ds As New DataSet()
            Using stringReader As New StringReader(response)
                ds = New DataSet()
                ds.ReadXml(stringReader)
            End Using
            Status += FormatToSaveAllNFLSurvey(ds)
        Catch ex As Exception
            Status += ex.Message + Environment.NewLine
            Status += "Error at method ProcessAllNFLSurvey()" + Environment.NewLine
            Try
                Status += ex.InnerException.Message + Environment.NewLine
                Status += ex.InnerException.InnerException.Message + Environment.NewLine
            Catch e As Exception

            End Try
        Finally
            Status += "Process started at " + start + " and ended at " + DateTime.Now + Environment.NewLine
            Status += "---" + Environment.NewLine
        End Try
        Return Status
    End Function

    Private Function FormatToSaveNFLSurvey(ByVal ds As DataSet) As String
        Dim Status As String = String.Empty
        Dim trueCounter As Integer = 0, falseCounter As Integer = 0
        Dim IsSaved As Boolean
        Try
            If ds.Tables.Count > 1 Then
                For Each a As DataRow In ds.Tables("Submission").Rows
                    If CheckIfNew(a("Id").ToString()) Then
                        IsSaved = False
                        Dim tempRecord As DataTable = ds.Tables("Section").Select("Sections_Id = " + a("Submission_Id").ToString()).CopyToDataTable()
                        Dim dttempMain As DataTable = ds.Tables("Response").Select("Responses_Id = " + tempRecord(0)("Section_Id").ToString()).CopyToDataTable()
                        Dim dttempEnd As DataTable = ds.Tables("Response").Select("Responses_Id = " + tempRecord(tempRecord.Rows.Count - 1)("Section_Id").ToString()).CopyToDataTable()
                        Dim dtDetail As DataTable

                        Dim farmtemp As String = dttempMain(6)(1).ToString()
                        Dim weektemp As Integer = Convert.ToInt32(dttempMain(3)(1).ToString())
                        Dim yeartemp = Convert.ToInt32(dttempMain(1)(1).ToString().Split("/")(2))

                        If farmtemp = Properties.FarmParam And weektemp = Properties.WeekParam And yeartemp = Properties.YearParam Then

                            conn = New SqlConnection(Properties.NFLDBConnString)
                            conn.Open()
                            Dim trans As SqlTransaction = conn.BeginTransaction()
                            Try
                                com = New SqlCommand("sp_UploadTransHeader", conn, trans)
                                com.CommandType = CommandType.StoredProcedure
                                com.CommandTimeout = 999999
                                com.Parameters.Add("@SubmissionID", SqlDbType.Int).Value = a("Id").ToString()
                                com.Parameters.Add("@SurveyDate", SqlDbType.DateTime).Value = dttempMain(1)(1).ToString() + " " + dttempMain(2)(1).ToString()
                                com.Parameters.Add("@Elevation", SqlDbType.VarChar).Value = dttempMain(4)(1).ToString()
                                Try
                                    com.Parameters.Add("@WeekNo", SqlDbType.Int).Value = dttempMain(3)(1).ToString()
                                Catch ex As Exception
                                    com.Parameters.Add("@WeekNo", SqlDbType.Int).Value = Nothing
                                End Try

                                Try
                                    com.Parameters.Add("@Plantation", SqlDbType.VarChar).Value = dttempMain(5)(1).ToString()
                                Catch ex As Exception
                                    com.Parameters.Add("@Plantation", SqlDbType.VarChar).Value = ""
                                End Try

                                Try
                                    com.Parameters.Add("@Farm", SqlDbType.VarChar).Value = dttempMain(6)(1).ToString()
                                Catch ex As Exception
                                    com.Parameters.Add("@Farm", SqlDbType.VarChar).Value = ""
                                End Try

                                Try
                                    com.Parameters.Add("@PHCode", SqlDbType.VarChar).Value = dttempMain(7)(1).ToString()
                                Catch ex As Exception
                                    com.Parameters.Add("@PHCode", SqlDbType.VarChar).Value = ""
                                End Try

                                Try
                                    com.Parameters.Add("@Block_Line", SqlDbType.VarChar).Value = dttempMain(9)(1).ToString()
                                Catch ex As Exception
                                    com.Parameters.Add("@Block_Line", SqlDbType.VarChar).Value = ""
                                End Try

                                Try
                                    com.Parameters.Add("@Remarks", SqlDbType.Text).Value = dttempEnd(0)(1).ToString()
                                Catch ex As Exception
                                    com.Parameters.Add("@Remarks", SqlDbType.Text).Value = ""
                                End Try

                                Try
                                    com.Parameters.Add("@CoordinatorName", SqlDbType.VarChar).Value = dttempEnd(1)(1).ToString()
                                Catch ex As Exception
                                    com.Parameters.Add("@CoordinatorName", SqlDbType.VarChar).Value = ""
                                End Try

                                Try
                                    com.Parameters.Add("@FarmRepresentative", SqlDbType.VarChar).Value = dttempEnd(3)(1).ToString()
                                Catch ex As Exception
                                    com.Parameters.Add("@FarmRepresentative", SqlDbType.VarChar).Value = ""
                                End Try
                                com.Parameters.Add("@SubmissionDateTime", SqlDbType.DateTime).Value = a("Date").ToString()
                                com.ExecuteNonQuery()

                                For b As Integer = 1 To 16 Step 1
                                    Try
                                        dtDetail = ds.Tables("Response").Select("Responses_Id = " + tempRecord(b)("Section_Id").ToString()).CopyToDataTable()
                                        If Not dtDetail(0)(1).ToString().Equals("") Then
                                            com = New SqlCommand("sp_UploadTransDetails", conn, trans)
                                            com.CommandType = CommandType.StoredProcedure
                                            com.CommandTimeout = 999999
                                            com.Parameters.Add("@WeekAge", SqlDbType.VarChar).Value = dtDetail(0)(1).ToString()
                                            com.Parameters.Add("@Color", SqlDbType.VarChar).Value = dtDetail(1)(1).ToString()
                                            com.Parameters.Add("@NL1", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(2)(1).ToString())
                                            com.Parameters.Add("@NL2", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(3)(1).ToString())
                                            com.Parameters.Add("@NL3", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(4)(1).ToString())
                                            com.Parameters.Add("@NL4", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(5)(1).ToString())
                                            com.Parameters.Add("@NL5", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(6)(1).ToString())
                                            com.Parameters.Add("@NL6", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(7)(1).ToString())
                                            com.Parameters.Add("@NL7", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(8)(1).ToString())
                                            com.Parameters.Add("@NL8", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(9)(1).ToString())
                                            com.Parameters.Add("@NL9", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(10)(1).ToString())
                                            com.Parameters.Add("@NL10", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(11)(1).ToString())
                                            com.Parameters.Add("@NL11", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(12)(1).ToString())
                                            com.Parameters.Add("@NL12", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(13)(1).ToString())
                                            com.Parameters.Add("@NL13", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(14)(1).ToString())
                                            com.Parameters.Add("@NL14", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(15)(1).ToString())
                                            com.Parameters.Add("@NL15", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(16)(1).ToString())
                                            com.ExecuteNonQuery()

                                            trueCounter += 1

                                        End If
                                    Catch ex As Exception
                                    End Try
                                Next

                                trans.Commit()
                                IsSaved = True
                            Catch ex As Exception
                                trans.Rollback()
                                Status += ex.Message + Environment.NewLine
                                Try
                                    Status += ex.InnerException.Message + Environment.NewLine
                                    Status += ex.InnerException.InnerException.Message + Environment.NewLine
                                Catch e As Exception

                                End Try
                            Finally
                                conn.Close()
                            End Try

                            If IsSaved Then
                                If (Properties.IsSaveSignature) Then
                                    Status += "Downloaded signature/s"
                                    Dim Client As New WebClient
                                    Try
                                        Client.DownloadFile("https://www.gocanvas.com/apiv2/images.xml?image_id=" + dttempEnd.Rows(2)(1).ToString() + "&username=" + Properties.Username + "&password=" + Properties.Password, Properties.SignaturePath + "/" + dttempEnd(1)(1).ToString() + ".jpg.")
                                        Client.DownloadFile("https://www.gocanvas.com/apiv2/images.xml?image_id=" + dttempEnd.Rows(4)(1).ToString() + "&username=" + Properties.Username + "&password=" + Properties.Password, Properties.SignaturePath + "/" + dttempEnd(3)(1).ToString() + ".jpg.")
                                        Client.Dispose()
                                    Catch ex As Exception
                                    End Try
                                End If

                                trueCounter += 1
                            Else
                                falseCounter += 1
                            End If


                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            Status += ex.Message + Environment.NewLine
            Try
                Status += ex.InnerException.Message + Environment.NewLine
                Status += ex.InnerException.InnerException.Message + Environment.NewLine
            Catch e As Exception

            End Try
        End Try

        'Status += "NFL: Data from " + DateTime.Now.AddDays(-1).ToShortDateString() + " to " + DateTime.Now.ToString() + "." + Environment.NewLine
        Status += trueCounter.ToString() + " record/s succeed." + Environment.NewLine
        Status += falseCounter.ToString() + " record/s failed." + Environment.NewLine
        Return Status
    End Function


    Private Function FormatToSaveAllNFLSurvey(ByVal ds As DataSet) As String
        Dim Status As String = String.Empty
        Dim trueCounter As Integer = 0, falseCounter As Integer = 0
        Dim IsSaved As Boolean
        Try
            If ds.Tables.Count > 1 Then
                For Each a As DataRow In ds.Tables("Submission").Rows
                    If CheckIfNew(a("Id").ToString()) Then
                        IsSaved = False
                        Dim tempRecord As DataTable = ds.Tables("Section").Select("Sections_Id = " + a("Submission_Id").ToString()).CopyToDataTable()
                        Dim dttempMain As DataTable = ds.Tables("Response").Select("Responses_Id = " + tempRecord(0)("Section_Id").ToString()).CopyToDataTable()
                        Dim dttempEnd As DataTable = ds.Tables("Response").Select("Responses_Id = " + tempRecord(tempRecord.Rows.Count - 1)("Section_Id").ToString()).CopyToDataTable()
                        Dim dtDetail As DataTable

                        Dim farmtemp As String = dttempMain(6)(1).ToString()
                        Dim weektemp As Integer = Convert.ToInt32(dttempMain(3)(1).ToString())
                        Dim yeartemp = Convert.ToInt32(dttempMain(1)(1).ToString().Split("/")(2))

                        conn = New SqlConnection(Properties.NFLDBConnString)
                        conn.Open()
                        Dim trans As SqlTransaction = conn.BeginTransaction()
                        Try
                            com = New SqlCommand("sp_UploadTransHeader", conn, trans)
                            com.CommandType = CommandType.StoredProcedure
                            com.CommandTimeout = 999999
                            com.Parameters.Add("@SubmissionID", SqlDbType.Int).Value = a("Id").ToString()
                            com.Parameters.Add("@SurveyDate", SqlDbType.DateTime).Value = dttempMain(1)(1).ToString() + " " + dttempMain(2)(1).ToString()
                            com.Parameters.Add("@Elevation", SqlDbType.VarChar).Value = dttempMain(4)(1).ToString()
                            Try
                                com.Parameters.Add("@WeekNo", SqlDbType.Int).Value = dttempMain(3)(1).ToString()
                            Catch ex As Exception
                                com.Parameters.Add("@WeekNo", SqlDbType.Int).Value = Nothing
                            End Try

                            Try
                                com.Parameters.Add("@Plantation", SqlDbType.VarChar).Value = dttempMain(5)(1).ToString()
                            Catch ex As Exception
                                com.Parameters.Add("@Plantation", SqlDbType.VarChar).Value = ""
                            End Try

                            Try
                                com.Parameters.Add("@Farm", SqlDbType.VarChar).Value = dttempMain(6)(1).ToString()
                            Catch ex As Exception
                                com.Parameters.Add("@Farm", SqlDbType.VarChar).Value = ""
                            End Try

                            Try
                                com.Parameters.Add("@PHCode", SqlDbType.VarChar).Value = dttempMain(7)(1).ToString()
                            Catch ex As Exception
                                com.Parameters.Add("@PHCode", SqlDbType.VarChar).Value = ""
                            End Try

                            Try
                                com.Parameters.Add("@Block_Line", SqlDbType.VarChar).Value = dttempMain(9)(1).ToString()
                            Catch ex As Exception
                                com.Parameters.Add("@Block_Line", SqlDbType.VarChar).Value = ""
                            End Try

                            Try
                                com.Parameters.Add("@Remarks", SqlDbType.Text).Value = dttempEnd(0)(1).ToString()
                            Catch ex As Exception
                                com.Parameters.Add("@Remarks", SqlDbType.Text).Value = ""
                            End Try

                            Try
                                com.Parameters.Add("@CoordinatorName", SqlDbType.VarChar).Value = dttempEnd(1)(1).ToString()
                            Catch ex As Exception
                                com.Parameters.Add("@CoordinatorName", SqlDbType.VarChar).Value = ""
                            End Try

                            Try
                                com.Parameters.Add("@FarmRepresentative", SqlDbType.VarChar).Value = dttempEnd(3)(1).ToString()
                            Catch ex As Exception
                                com.Parameters.Add("@FarmRepresentative", SqlDbType.VarChar).Value = ""
                            End Try
                            com.Parameters.Add("@SubmissionDateTime", SqlDbType.DateTime).Value = a("Date").ToString()
                            com.ExecuteNonQuery()

                            For b As Integer = 1 To 16 Step 1
                                Try
                                    dtDetail = ds.Tables("Response").Select("Responses_Id = " + tempRecord(b)("Section_Id").ToString()).CopyToDataTable()
                                    If Not dtDetail(0)(1).ToString().Equals("") Then
                                        com = New SqlCommand("sp_UploadTransDetails", conn, trans)
                                        com.CommandType = CommandType.StoredProcedure
                                        com.CommandTimeout = 999999
                                        com.Parameters.Add("@WeekAge", SqlDbType.VarChar).Value = dtDetail(0)(1).ToString()
                                        com.Parameters.Add("@Color", SqlDbType.VarChar).Value = dtDetail(1)(1).ToString()
                                        com.Parameters.Add("@NL1", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(2)(1).ToString())
                                        com.Parameters.Add("@NL2", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(3)(1).ToString())
                                        com.Parameters.Add("@NL3", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(4)(1).ToString())
                                        com.Parameters.Add("@NL4", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(5)(1).ToString())
                                        com.Parameters.Add("@NL5", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(6)(1).ToString())
                                        com.Parameters.Add("@NL6", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(7)(1).ToString())
                                        com.Parameters.Add("@NL7", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(8)(1).ToString())
                                        com.Parameters.Add("@NL8", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(9)(1).ToString())
                                        com.Parameters.Add("@NL9", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(10)(1).ToString())
                                        com.Parameters.Add("@NL10", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(11)(1).ToString())
                                        com.Parameters.Add("@NL11", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(12)(1).ToString())
                                        com.Parameters.Add("@NL12", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(13)(1).ToString())
                                        com.Parameters.Add("@NL13", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(14)(1).ToString())
                                        com.Parameters.Add("@NL14", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(15)(1).ToString())
                                        com.Parameters.Add("@NL15", SqlDbType.Int).Value = ConvertToDoubleElseZero(dtDetail(16)(1).ToString())
                                        com.ExecuteNonQuery()

                                        trueCounter += 1

                                    End If
                                Catch ex As Exception
                                End Try
                            Next

                            trans.Commit()
                            IsSaved = True
                        Catch ex As Exception
                            trans.Rollback()
                            Status += ex.Message + Environment.NewLine
                            Try
                                Status += ex.InnerException.Message + Environment.NewLine
                                Status += ex.InnerException.InnerException.Message + Environment.NewLine
                            Catch e As Exception

                            End Try
                        Finally
                            conn.Close()
                        End Try

                        If IsSaved Then
                            If (Properties.IsSaveSignature) Then
                                Status += "Downloaded signature/s"
                                Dim Client As New WebClient
                                Try
                                    Client.DownloadFile("https://www.gocanvas.com/apiv2/images.xml?image_id=" + dttempEnd.Rows(2)(1).ToString() + "&username=" + Properties.Username + "&password=" + Properties.Password, Properties.SignaturePath + "/" + dttempEnd(1)(1).ToString() + ".jpg.")
                                    Client.DownloadFile("https://www.gocanvas.com/apiv2/images.xml?image_id=" + dttempEnd.Rows(4)(1).ToString() + "&username=" + Properties.Username + "&password=" + Properties.Password, Properties.SignaturePath + "/" + dttempEnd(3)(1).ToString() + ".jpg.")
                                    Client.Dispose()
                                Catch ex As Exception
                                End Try
                            End If

                            trueCounter += 1
                        Else
                            falseCounter += 1
                        End If


                    End If
                Next
            End If
        Catch ex As Exception
            Status += ex.Message + Environment.NewLine
            Try
                Status += ex.InnerException.Message + Environment.NewLine
                Status += ex.InnerException.InnerException.Message + Environment.NewLine
            Catch e As Exception

            End Try
        End Try

        'Status += "NFL: Data from " + DateTime.Now.AddDays(-1).ToShortDateString() + " to " + DateTime.Now.ToString() + "." + Environment.NewLine
        Status += trueCounter.ToString() + " record/s succeed." + Environment.NewLine
        Status += falseCounter.ToString() + " record/s failed." + Environment.NewLine
        Return Status
    End Function



#Region "Methods"

    Private Function FetchDataTable(ByVal query As String) As DataTable
        Try
            conn = New SqlConnection(Properties.NFLDBConnString)
            dt = New DataTable()
            conn.Open()
            com = New SqlCommand(query, conn)
            com.CommandTimeout = 999999
            adap = New SqlDataAdapter(com)
            adap.Fill(dt)
            conn.Close()
        Catch
            conn.Close()
        End Try

        Return dt
    End Function

    Private Function ConvertToDoubleElseZero(ByVal param As String) As String
        Try
            Return Convert.ToDouble(param).ToString()
        Catch

        End Try

        Return "0"
    End Function

    Private Function CheckIfNew(ByVal param) As Boolean

        Dim dt As DataTable = FetchDataTable("SELECT COUNT(*) FROM tblTransHeader WHERE SubmissionID = '" + param + "'")
        If (Convert.ToInt32(dt.Rows(0)(0).ToString()) > 0) Then
            Return False
        End If

        Return True
    End Function

#End Region

End Class
