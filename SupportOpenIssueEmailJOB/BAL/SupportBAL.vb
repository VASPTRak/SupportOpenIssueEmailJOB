Imports System.Data.SqlClient
Imports log4net
Imports log4net.Config

Public Class SupportBAL
    Private Shared ReadOnly log As ILog = LogManager.GetLogger(GetType(SupportBAL))

    Shared Sub New()
        XmlConfigurator.Configure()
    End Sub

    Public Function GetSupportOpenDataAndSendEmail(ByVal condition As String) As DataTable
        Dim dal = New GeneralizedDAL()
        Try

            Dim ds As DataSet = New DataSet()

            Dim Param As SqlParameter() = New SqlParameter() _
                {New SqlParameter("@Conditions", SqlDbType.NVarChar) With {.Value = condition}}


            ds = dal.ExecuteStoredProcedureGetDataSet("usp_tt_Support_GetSupportOpenItemsByEmails", Param)

            Return ds.Tables(0)

        Catch ex As Exception
            log.Error("Error occurred in GetUnResolvedSupportItems Exception is :" + ex.ToString())
            Return Nothing
        Finally

        End Try
    End Function


End Class
