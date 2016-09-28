Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports ADODB


Public Class clsFits
    Private vServer As String
    Private vPassword As String
    Private vUser As String
    Private vDatabase As String

    Dim cn As New ADODB.Connection()

    Public Property server As String
        Get
            server = vServer
        End Get
        Set(value As String)
            vServer = value
        End Set
    End Property

    Public Property user As String
        Get
            user = vUser
        End Get
        Set(value As String)
            vUser = value
        End Set
    End Property

    Public Property password As String
        Get
            password = vPassword
        End Get
        Set(value As String)
            vPassword = value
        End Set
    End Property

    Public Property database As String
        Get
            database = vDatabase
        End Get
        Set(value As String)
            vDatabase = value
        End Set
    End Property

    Public Function connect() As ADODB.Connection

        Dim connectionString As String = "Provider=SQLOLEDB;Initial Catalog=" & vDatabase & ";" & _
            "Data Source=" & vServer & ";User Id=" & vUser & ";" & _
            "Password=" & vPassword & ";Database=" & vDatabase & ";"

        cn.ConnectionString = connectionString
        cn.Open()
        cn.CursorLocation = CursorLocationEnum.adUseClient
        connect = cn

    End Function

    Public Sub disconnect()

    End Sub

    Public Function getEvents(vDateFrom As String, vDateTo As String) As ADODB.Recordset
        'Date format : 2016-07-01
        'Dim vSql As String = "select event.*,operation_map.description as operation_name " & _
        '                "from event INNER JOIN operation_map ON event.operation = operation_map.operation " & _
        '                "where event.buildtype in ('RMA','PRODUCTION','QUALIFICATION') and " & _
        '                "event.model=operation_map.model_type and event.timestamp between ? and ?  " & _
        '                "order by event.date_time"
        'Edit by Chutchai S on Sep 28,2016
        'To use Datetime_checkout -- Completed process.
        Dim vSql As String = "select event.*,operation_map.description as operation_name " & _
                "from event INNER JOIN operation_map ON event.operation = operation_map.operation " & _
                "where event.buildtype in ('RMA','PRODUCTION','QUALIFICATION') and " & _
                "event.model=operation_map.model_type and event.date_time_checkout between ? and ?  " & _
                "order by event.date_time"

        Dim cmd As New ADODB.Command()
        Dim sDateFromParam As ADODB.Parameter
        Dim sDateToParam As ADODB.Parameter

        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sDateFromParam = .CreateParameter("vDateFrom", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vDateFrom)
            sDateToParam = .CreateParameter("vDateTo", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vDateTo)
            .Parameters.Append(sDateFromParam)
            .Parameters.Append(sDateToParam)
            getEvents = .Execute
        End With

    End Function

    Public Function getEvents(vSerialnumber As String) As ADODB.Recordset
        'Date format : 2016-07-01
        Dim vSql As String = "select event.*,operation_map.description as operation_name " & _
                        "from event INNER JOIN operation_map ON event.operation = operation_map.operation " & _
                        "where serial_no = ?  and event.model=operation_map.model_type " & _
                        "order by date_time"

        Dim cmd As New ADODB.Command()
        Dim sSnParam As ADODB.Parameter
        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sSnParam = .CreateParameter("vSerialnumber", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vSerialnumber)
            .Parameters.Append(sSnParam)
            getEvents = .Execute
        End With
    End Function


    Public Function getNextStation(vSerialnumber As String,
                                   vStation As String,
                                   vTrans_seq As String, vModel As String) As String
        'Date format : 2016-07-01
        Dim vSql As String = "select * " & _
                            "from route_table " & _
                            "where model_type=? and prev_valid_operation=?"

        Dim cmd As New ADODB.Command()
        Dim sSnParam As ADODB.Parameter
        Dim sStation As ADODB.Parameter
        Dim vRouteRst As ADODB.Recordset
        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sSnParam = .CreateParameter("vModelType", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vModel)
            sStation = .CreateParameter("vModelType", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vStation)
            .Parameters.Append(sSnParam)
            .Parameters.Append(sStation)
            vRouteRst = .Execute
        End With
        getNextStation = vStation 'Default value
        Dim vCriteria As String
        Dim cmdCriteria As New ADODB.Command()
        With cmdCriteria
            .ActiveConnection = cn
            .CommandType = CommandTypeEnum.adCmdText
        End With
        Try
            With vRouteRst
                Do While Not .EOF
                    'Query on criteria field
                    If IsDBNull(.Fields("criteria").Value) Then
                        getNextStation = .Fields("cur_operation").Value
                        Exit Do
                    End If
                    vCriteria = .Fields("criteria").Value
                    If vCriteria = "where m.buildetype ='RMA' c" Then
                        GoTo NextLoop
                    End If
                    vCriteria = Replace(vCriteria, "m.max_trans_seq", "m.trans_seq")

                    vSql = "select * " & _
                            "from event as m " & _
                            " " & vCriteria & " " & _
                            "and m.serial_no='" & vSerialnumber & "' " & _
                            "and m.operation='" & vStation & "' " & _
                            "and m.trans_seq=" & vTrans_seq
                    '"and m.date_time =(select max(date_time) from event where serial_no='163256019' and operation=100 )"
                    cmdCriteria.CommandText = vSql
                    If cmdCriteria.Execute.RecordCount > 0 Then
                        getNextStation = .Fields("cur_operation").Value
                        Exit Do
                    End If
NextLoop:
                    .MoveNext()
                Loop
            End With
        Catch ex As Exception

        End Try
        

    End Function

    Public Function getEventMaster(vSerialnumber As String) As ADODB.Recordset
        'Date format : 2016-07-01
        Dim vSql As String = "select * " & _
                        "from event_master " & _
                        "where serial_no = ? " & _
                        "order by date_time"

        Dim cmd As New ADODB.Command()
        Dim sSnParam As ADODB.Parameter
        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sSnParam = .CreateParameter("vSerialnumber", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vSerialnumber)
            .Parameters.Append(sSnParam)
            getEventMaster = .Execute
        End With
    End Function

    Public Function getParameters(vSerialnumber As String, vAtrrCode As Integer, _
                                  vTransSeq As Integer) As ADODB.Recordset
        'Date format : 2016-07-01
        Dim vSql As String = "SELECT attribute.* ,attribute_map.* " & _
                        "FROM attribute_map INNER JOIN " & _
                         "attribute ON attribute_map.attribute_code = attribute.attribute_code " & _
                        "WHERE (attribute.serial_no =? " & _
                        "and attribute.trans_seq=? " & _
                        "and attribute.sn_attr_code=?)"

        Dim cmd As New ADODB.Command()
        Dim sSnParam As ADODB.Parameter
        Dim sTranSeqParam As ADODB.Parameter
        Dim sAttCodeParam As ADODB.Parameter

        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sSnParam = .CreateParameter("vSerialnumber", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vSerialnumber)
            sTranSeqParam = .CreateParameter("vSerialnumber", DataTypeEnum.adInteger, _
                                                 ParameterDirectionEnum.adParamInput, 50, vTransSeq)
            sAttCodeParam = .CreateParameter("vSerialnumber", DataTypeEnum.adInteger, _
                                                 ParameterDirectionEnum.adParamInput, 50, vAtrrCode)
            .Parameters.Append(sSnParam)
            .Parameters.Append(sTranSeqParam)
            .Parameters.Append(sAttCodeParam)
            getParameters = .Execute
        End With
    End Function














End Class


Public Class clsAutoTest
    Private vServer As String
    Private vPassword As String
    Private vUser As String
    Private vDatabase As String

    Dim cn As New ADODB.Connection()

    Public Property server As String
        Get
            server = vServer
        End Get
        Set(value As String)
            vServer = value
        End Set
    End Property

    Public Property user As String
        Get
            user = vUser
        End Get
        Set(value As String)
            vUser = value
        End Set
    End Property

    Public Property password As String
        Get
            password = vPassword
        End Get
        Set(value As String)
            vPassword = value
        End Set
    End Property

    Public Property database As String
        Get
            database = vDatabase
        End Get
        Set(value As String)
            vDatabase = value
        End Set
    End Property

    Public Function connect() As ADODB.Connection

        Dim connectionString As String = "Provider=SQLOLEDB;Initial Catalog=" & vDatabase & ";" & _
            "Data Source=" & vServer & ";User Id=" & vUser & ";" & _
            "Password=" & vPassword & ";Database=" & vDatabase & ";"

        cn.ConnectionString = connectionString
        cn.Open()
        cn.CursorLocation = CursorLocationEnum.adUseClient
        connect = cn

    End Function

    Public Sub disconnect()

    End Sub

    Public Function getEvents(vDateFrom As String, vDateTo As String) As ADODB.Recordset
        'Date format : 2016-07-01
        Dim vSql As String = "select event.*,operation_map.description as operation_name " & _
                        "from event INNER JOIN operation_map ON event.operation = operation_map.operation " & _
                        "where event.buildtype in ('RMA','PRODUCTION','QUALIFICATION') and " & _
                        "event.model=operation_map.model_type and event.timestamp between ? and ?  " & _
                        "order by event.date_time"

        Dim cmd As New ADODB.Command()
        Dim sDateFromParam As ADODB.Parameter
        Dim sDateToParam As ADODB.Parameter

        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sDateFromParam = .CreateParameter("vDateFrom", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vDateFrom)
            sDateToParam = .CreateParameter("vDateTo", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vDateTo)
            .Parameters.Append(sDateFromParam)
            .Parameters.Append(sDateToParam)
            getEvents = .Execute
        End With

    End Function

    Public Function getEvents(vSerialnumber As String) As ADODB.Recordset
        'Date format : 2016-07-01
        Dim vSql As String = "select event.*,operation_map.description as operation_name " & _
                        "from event INNER JOIN operation_map ON event.operation = operation_map.operation " & _
                        "where serial_no = ?  and event.model=operation_map.model_type " & _
                        "order by date_time"

        Dim cmd As New ADODB.Command()
        Dim sSnParam As ADODB.Parameter
        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sSnParam = .CreateParameter("vSerialnumber", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vSerialnumber)
            .Parameters.Append(sSnParam)
            getEvents = .Execute
        End With
    End Function


    Public Function getNextStation(vSerialnumber As String,
                                   vStation As String,
                                   vTrans_seq As String, vModel As String) As String
        'Date format : 2016-07-01
        Dim vSql As String = "select * " & _
                            "from route_table " & _
                            "where model_type=? and prev_valid_operation=?"

        Dim cmd As New ADODB.Command()
        Dim sSnParam As ADODB.Parameter
        Dim sStation As ADODB.Parameter
        Dim vRouteRst As ADODB.Recordset
        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sSnParam = .CreateParameter("vModelType", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vModel)
            sStation = .CreateParameter("vModelType", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vStation)
            .Parameters.Append(sSnParam)
            .Parameters.Append(sStation)
            vRouteRst = .Execute
        End With
        getNextStation = vStation 'Default value
        Dim vCriteria As String
        Dim cmdCriteria As New ADODB.Command()
        With cmdCriteria
            .ActiveConnection = cn
            .CommandType = CommandTypeEnum.adCmdText
        End With
        Try
            With vRouteRst
                Do While Not .EOF
                    'Query on criteria field
                    If IsDBNull(.Fields("criteria").Value) Then
                        getNextStation = .Fields("cur_operation").Value
                        Exit Do
                    End If
                    vCriteria = .Fields("criteria").Value
                    If vCriteria = "where m.buildetype ='RMA' c" Then
                        GoTo NextLoop
                    End If
                    vCriteria = Replace(vCriteria, "m.max_trans_seq", "m.trans_seq")

                    vSql = "select * " & _
                            "from event as m " & _
                            " " & vCriteria & " " & _
                            "and m.serial_no='" & vSerialnumber & "' " & _
                            "and m.operation='" & vStation & "' " & _
                            "and m.trans_seq=" & vTrans_seq
                    '"and m.date_time =(select max(date_time) from event where serial_no='163256019' and operation=100 )"
                    cmdCriteria.CommandText = vSql
                    If cmdCriteria.Execute.RecordCount > 0 Then
                        getNextStation = .Fields("cur_operation").Value
                        Exit Do
                    End If
NextLoop:
                    .MoveNext()
                Loop
            End With
        Catch ex As Exception

        End Try


    End Function


    Public Function getTestData(vSerialnumber As String, vProcess As String, _
                                vTester As String, vDatetime As String) As ADODB.Recordset
        'Date format : 2016-07-01
        Dim vSql As String = "select id,station_id,start_date_time,uut_status,process,tps_name " & _
                            "from uut_result " & _
                            "where uut_serial_number=?  " & _
                            "and station_id=? " & _
                            "and process=? and start_date_time <= ? " & _
                            "order by id"

        Dim cmd As New ADODB.Command()
        Dim sSnParam, sStationParam, sProcessParam, sDdatetimeParam As ADODB.Parameter
        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sSnParam = .CreateParameter("vSerialnumber", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vSerialnumber)
            sStationParam = .CreateParameter("vSerialnumber", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vTester)
            sProcessParam = .CreateParameter("vSerialnumber", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vProcess)
            sDdatetimeParam = .CreateParameter("vSerialnumber", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vDatetime)
            .Parameters.Append(sSnParam)
            .Parameters.Append(sStationParam)
            .Parameters.Append(sProcessParam)
            .Parameters.Append(sDdatetimeParam)
            getTestData = .Execute
        End With

        If getTestData.RecordCount > 0 Then
            getTestData.MoveLast() 'Get last record.
            Dim vUutResult As String
            vUutResult = getTestData.Fields("id").Value
            vSql = "SELECT STEP_RESULT.UUT_RESULT, STEP_RESULT.STEP_NAME, " & _
                "STEP_RESULT.STEP_GROUP, STEP_RESULT.STATUS, PROP_RESULT.DATA, " & _
                "PROP_NUMERICLIMIT.HIGH_LIMIT, PROP_NUMERICLIMIT.LOW_LIMIT, " & _
                "PROP_NUMERICLIMIT.UNITS, PROP_NUMERICLIMIT.COMP_OPERATOR,STEP_RESULT.ID " & _
                "FROM STEP_RESULT INNER JOIN " & _
                         "PROP_RESULT ON STEP_RESULT.ID = PROP_RESULT.STEP_RESULT INNER JOIN " & _
                         "PROP_NUMERICLIMIT ON PROP_RESULT.ID = PROP_NUMERICLIMIT.PROP_RESULT " & _
                "WHERE(STEP_RESULT.UUT_RESULT = ?)"
            Dim cmd2 As New ADODB.Command()
            Dim sUUTparam As ADODB.Parameter
            With cmd2
                .ActiveConnection = cn
                .CommandText = vSql
                .CommandType = CommandTypeEnum.adCmdText
                sUUTparam = .CreateParameter("vUUTserial", DataTypeEnum.adVarChar, _
                                                     ParameterDirectionEnum.adParamInput, 50, vUutResult)
                .Parameters.Append(sUUTparam)
                getTestData = .Execute
            End With

        End If


    End Function

    'Public Function getParameters(vSerialnumber As String, vAtrrCode As Integer, _
    '                              vTransSeq As Integer) As ADODB.Recordset
    '    'Date format : 2016-07-01
    '    Dim vSql As String = "SELECT attribute.* ,attribute_map.* " & _
    '                    "FROM attribute_map INNER JOIN " & _
    '                     "attribute ON attribute_map.attribute_code = attribute.attribute_code " & _
    '                    "WHERE (attribute.serial_no =? " & _
    '                    "and attribute.trans_seq=? " & _
    '                    "and attribute.sn_attr_code=?)"

    '    Dim cmd As New ADODB.Command()
    '    Dim sSnParam As ADODB.Parameter
    '    Dim sTranSeqParam As ADODB.Parameter
    '    Dim sAttCodeParam As ADODB.Parameter

    '    With cmd
    '        .ActiveConnection = cn
    '        .CommandText = vSql
    '        .CommandType = CommandTypeEnum.adCmdText
    '        sSnParam = .CreateParameter("vSerialnumber", DataTypeEnum.adVarChar, _
    '                                             ParameterDirectionEnum.adParamInput, 50, vSerialnumber)
    '        sTranSeqParam = .CreateParameter("vSerialnumber", DataTypeEnum.adInteger, _
    '                                             ParameterDirectionEnum.adParamInput, 50, vTransSeq)
    '        sAttCodeParam = .CreateParameter("vSerialnumber", DataTypeEnum.adInteger, _
    '                                             ParameterDirectionEnum.adParamInput, 50, vAtrrCode)
    '        .Parameters.Append(sSnParam)
    '        .Parameters.Append(sTranSeqParam)
    '        .Parameters.Append(sAttCodeParam)
    '        getParameters = .Execute
    '    End With
    'End Function














End Class
