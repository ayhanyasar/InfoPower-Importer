<%
'====================================================================
' InfoPower Importer PRO v3.0
' File    : classes/class.history.asp
' Part    : 1
'====================================================================
'Option Explicit

Class History

    Private mDB
    Private mTable

    Private Sub Class_Initialize()

        Set mDB = New Database

        mTable = "import_history"

    End Sub



    Private Sub Class_Terminate()

        Set mDB = Nothing

    End Sub



    Public Property Set Database(ByRef DB)

        Set mDB = DB

    End Property



    Public Property Get Database()

        Set Database = mDB

    End Property



    Public Property Let TableName(Value)

        mTable = Trim(CStr(Value))

    End Property



    Public Property Get TableName()

        TableName = mTable

    End Property



    Public Function Exists(ProductID)

        Exists = mDB.RecordExists( _
            "SELECT 1 FROM " & mTable & _
            " WHERE product_id=" & CLng(ProductID))

    End Function



    Public Function LastImport(ProductID)

        LastImport = mDB.Scalar( _
            "SELECT MAX(import_date) FROM " & _
            mTable & _
            " WHERE product_id=" & CLng(ProductID))

    End Function



    Public Function ImportCount(ProductID)

        ImportCount = mDB.Scalar( _
            "SELECT COUNT(*) FROM " & _
            mTable & _
            " WHERE product_id=" & CLng(ProductID))

    End Function



    Public Function Add(ProductID, Status, Message)

        Dim SQL

        SQL = ""

        SQL = SQL & "INSERT INTO " & mTable & "("
        SQL = SQL & "product_id,"
        SQL = SQL & "status,"
        SQL = SQL & "message,"
        SQL = SQL & "import_date"
        SQL = SQL & ") VALUES ("

        SQL = SQL & CLng(ProductID) & ","
        SQL = SQL & mDB.Escape(Status) & ","
        SQL = SQL & mDB.Escape(Message) & ","
        SQL = SQL & "GETDATE()"
        SQL = SQL & ")"

        Add = mDB.Execute(SQL)

    End Function



    Public Function Success(ProductID,Message)

        Success = Add(ProductID,"SUCCESS",Message)

    End Function



    Public Function Failed(ProductID,Message)

        Failed = Add(ProductID,"FAILED",Message)

    End Function



    Public Function Warning(ProductID,Message)

        Warning = Add(ProductID,"WARNING",Message)

    End Function
 
     Public Function Delete(ProductID)

        Delete = mDB.Execute( _
            "DELETE FROM " & _
            mTable & _
            " WHERE product_id=" & CLng(ProductID))

    End Function



    Public Function Clear()

        Clear = mDB.Execute( _
            "TRUNCATE TABLE " & mTable)

    End Function



    Public Function Get(ProductID)

        Set Get = mDB.GetRow( _
            "SELECT TOP 1 * FROM " & _
            mTable & _
            " WHERE product_id=" & CLng(ProductID) & _
            " ORDER BY import_date DESC")

    End Function



    Public Function GetAll()

        GetAll = mDB.GetRows( _
            "SELECT * FROM " & _
            mTable & _
            " ORDER BY import_date DESC")

    End Function



    Public Function GetFailed()

        GetFailed = mDB.GetRows( _
            "SELECT * FROM " & _
            mTable & _
            " WHERE status='FAILED'" & _
            " ORDER BY import_date DESC")

    End Function



    Public Function GetSuccess()

        GetSuccess = mDB.GetRows( _
            "SELECT * FROM " & _
            mTable & _
            " WHERE status='SUCCESS'" & _
            " ORDER BY import_date DESC")

    End Function



    Public Function TotalCount()

        TotalCount = mDB.Scalar( _
            "SELECT COUNT(*) FROM " & mTable)

    End Function



    Public Function SuccessCount()

        SuccessCount = mDB.Scalar( _
            "SELECT COUNT(*) FROM " & _
            mTable & _
            " WHERE status='SUCCESS'")

    End Function



    Public Function FailedCount()

        FailedCount = mDB.Scalar( _
            "SELECT COUNT(*) FROM " & _
            mTable & _
            " WHERE status='FAILED'")

    End Function

    Public Function WarningCount()

        WarningCount = mDB.Scalar( _
            "SELECT COUNT(*) FROM " & _
            mTable & _
            " WHERE status='WARNING'")

    End Function



    Public Function LastID()

        LastID = mDB.Scalar( _
            "SELECT MAX(id) FROM " & mTable)

    End Function



    Public Function RemoveOlderThan(DayCount)

        RemoveOlderThan = mDB.Execute( _
            "DELETE FROM " & _
            mTable & _
            " WHERE import_date < DATEADD(day,-" & _
            CLng(DayCount) & ",GETDATE())")

    End Function



    Public Function Summary()

        Dim D

        Set D = Server.CreateObject("Scripting.Dictionary")

        D.CompareMode = 1

        D.Add "total", TotalCount()
        D.Add "success", SuccessCount()
        D.Add "failed", FailedCount()
        D.Add "warning", WarningCount()

        Set Summary = D

    End Function



    Public Function HasHistory(ProductID)

        HasHistory = Exists(ProductID)

    End Function



    Public Function LastStatus(ProductID)

        LastStatus = mDB.Scalar( _
            "SELECT TOP 1 status FROM " & _
            mTable & _
            " WHERE product_id=" & CLng(ProductID) & _
            " ORDER BY import_date DESC")

    End Function



    Public Function LastMessage(ProductID)

        LastMessage = mDB.Scalar( _
            "SELECT TOP 1 message FROM " & _
            mTable & _
            " WHERE product_id=" & CLng(ProductID) & _
            " ORDER BY import_date DESC")

    End Function

End Class
%>