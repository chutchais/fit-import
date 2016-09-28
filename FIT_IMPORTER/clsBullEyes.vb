Imports System.Xml
Imports System
Imports System.Reflection
Imports ADODB
Imports System.IO

Public Class clsBullEyes
    Dim vSerialNumber As String
    Dim vWorkOrder As String
    Dim vModel As String
    Dim vPartNumber As String
    Dim vOperation As String
    Dim vBuildType As String
    Dim vRunType As Boolean
    Dim vEmployee As String
    Dim vShift As String
    Dim vSnAtrrCode As Integer
    Dim vTransSeq As Integer
    Dim vDateTimeIn As String
    Dim vDdateTimeOut As String
    Private vOutputPath As String
    Private vResult As String
    Private vOutputFile As String
    Private vDisposeCode As String
    Private vTester As String
    Private vOperationName As String
    Private vNextOperation As String

    Public Property serialnumber As String
        Get
            serialnumber = vSerialNumber
        End Get
        Set(value As String)
            vSerialNumber = value
        End Set
    End Property

    Public Property trans_seq As Integer
        Get
            trans_seq = vTransSeq
        End Get
        Set(value As Integer)
            vTransSeq = value
        End Set
    End Property

    Public Property sn_attr_code As Integer
        Get
            sn_attr_code = vSnAtrrCode
        End Get
        Set(value As Integer)
            vSnAtrrCode = value
        End Set
    End Property

    Public Property operation As String
        Get
            operation = vOperation
        End Get
        Set(value As String)
            vOperation = value
        End Set
    End Property

    Public Property employee As String
        Get
            employee = vEmployee
        End Get
        Set(value As String)
            vEmployee = value
        End Set
    End Property

    Public Property shift As String
        Get
            shift = vShift
        End Get
        Set(value As String)
            vShift = value
        End Set
    End Property

    Public Property buildtype As String
        Get
            buildtype = vBuildType
        End Get
        Set(value As String)
            vBuildType = value
        End Set
    End Property

    Public Property runtype As String
        Get
            runtype = vRunType
        End Get
        Set(value As String)
            vRunType = value
        End Set
    End Property

    Public Property partnumber As String
        Get
            partnumber = vPartNumber
        End Get
        Set(value As String)
            vPartNumber = value
        End Set
    End Property

    Public Property model As String
        Get
            model = vModel
        End Get
        Set(value As String)
            vModel = value
        End Set
    End Property

    Public Property datetimein As String
        Get
            datetimein = vDateTimeIn
        End Get
        Set(value As String)
            vDateTimeIn = value
        End Set
    End Property

    Public Property datetimeout As String
        Get
            datetimeout = vDdateTimeOut
        End Get
        Set(value As String)
            vDdateTimeOut = value
        End Set
    End Property

    Public Property workorder As String
        Get
            workorder = vWorkOrder
        End Get
        Set(value As String)
            vWorkOrder = value
        End Set
    End Property

    Public Property outputPath As String
        Get
            outputPath = vOutputPath
        End Get
        Set(value As String)
            vOutputPath = value
        End Set
    End Property

    Public Property result As String
        Get
            result = vResult
        End Get
        Set(value As String)
            vResult = value
        End Set
    End Property

    Public Property outputfile As String
        Get
            outputfile = vOutputFile
        End Get
        Set(value As String)
            vOutputFile = value
        End Set
    End Property

    Public Property disposecode As String
        Get
            disposecode = vDisposeCode
        End Get
        Set(value As String)
            vDisposeCode = value
        End Set
    End Property


    Public Property tester As String
        Get
            tester = vTester
        End Get
        Set(value As String)
            vTester = value
        End Set
    End Property

    Public Property operationname As String
        Get
            operationname = vOperationName
        End Get
        Set(value As String)
            vOperationName = value
        End Set
    End Property

    Public Property next_operation As String
        Get
            next_operation = vNextOperation
        End Get
        Set(value As String)
            vNextOperation = value
        End Set
    End Property

    Public Sub makeXML(rstParameter As ADODB.Recordset, Optional rstTestData As ADODB.Recordset = Nothing)
        Dim vXMLFile As String = vSerialNumber & "_" & vTransSeq & "_" & vSnAtrrCode & ".xml"
        Try
            Dim vLocation As String = vOutputPath
            Dim doc As XmlDocument = New XmlDocument()
            Dim dataNode As XmlNode
            With doc
                Dim rootNode As XmlNode = .CreateElement("performing")
                .AppendChild(rootNode)

                Dim _type As Type = Me.GetType()
                Dim properties() As PropertyInfo = _type.GetProperties()  'line 3
                For Each _property As PropertyInfo In properties
                    dataNode = .CreateElement(_property.Name)
                    dataNode.InnerText = _property.GetValue(Me, Nothing)
                    rootNode.AppendChild(dataNode)
                Next

                Dim paramsNode As XmlNode = .CreateElement("parameters")
                rootNode.AppendChild(paramsNode)


                With rstParameter
                    Do While Not .EOF
                        Dim paramNode As XmlNode = doc.CreateElement("parameter")
                        Dim codeAttr As XmlNode = doc.CreateAttribute("code")
                        codeAttr.Value = .Fields("attribute_code").Value
                        paramNode.Attributes.Append(codeAttr)

                        Dim descAttr As XmlNode = doc.CreateAttribute("desc")
                        descAttr.Value = .Fields("description").Value
                        paramNode.Attributes.Append(descAttr)

                        Dim minAttr As XmlNode = doc.CreateAttribute("min")
                        minAttr.Value = ""
                        paramNode.Attributes.Append(minAttr)

                        Dim maxAttr As XmlNode = doc.CreateAttribute("max")
                        maxAttr.Value = ""
                        paramNode.Attributes.Append(maxAttr)

                        Dim resultAttr As XmlNode = doc.CreateAttribute("result")
                        resultAttr.Value = "Passed"
                        paramNode.Attributes.Append(resultAttr)

                        paramNode.InnerText = .Fields("attribute_value").Value

                        paramsNode.AppendChild(paramNode)
                        .MoveNext()
                    Loop
                End With

                'For test parameter -- add by Chutchai S on Sep 27,2016
                If rstTestData Is Nothing Then
                    GoTo NoTestData
                End If

                'If rstTestData.State Then
                '    GoTo NoTestData
                'End If

                With rstTestData
                    Do While Not .EOF
                        Dim paramNode As XmlNode = doc.CreateElement("parameter")

                        Dim codeAttr As XmlNode = doc.CreateAttribute("code")
                        codeAttr.Value = .Fields("step_name").Value
                        paramNode.Attributes.Append(codeAttr)

                        Dim descAttr As XmlNode = doc.CreateAttribute("desc")
                        descAttr.Value = IIf(IsDBNull(.Fields("step_name").Value), "", .Fields("step_name").Value)
                        paramNode.Attributes.Append(descAttr)

                        Dim minAttr As XmlNode = doc.CreateAttribute("min")
                        minAttr.Value = IIf(IsDBNull(.Fields("low_limit").Value), "", .Fields("low_limit").Value)
                        paramNode.Attributes.Append(minAttr)

                        Dim maxAttr As XmlNode = doc.CreateAttribute("max")
                        maxAttr.Value = IIf(IsDBNull(.Fields("high_limit").Value), "", .Fields("high_limit").Value)
                        paramNode.Attributes.Append(maxAttr)

                        Dim resultAttr As XmlNode = doc.CreateAttribute("result")
                        resultAttr.Value = IIf(IsDBNull(.Fields("status").Value), "Passed", .Fields("status").Value)
                        paramNode.Attributes.Append(resultAttr)

                        paramNode.InnerText = IIf(IsDBNull(.Fields("data").Value), "", .Fields("data").Value)

                        paramsNode.AppendChild(paramNode)
                        .MoveNext()
                    Loop
                End With
NoTestData:

                ' Save the document to a file. White space is
                ' preserved (no white space).
                .PreserveWhitespace = True


                '----create main folder (date)-----
                Dim vPerformDate As Date = doc.SelectSingleNode("/performing/datetimein").InnerText
                If IsDBNull(vPerformDate) Then
                    vPerformDate = Now
                End If
                Dim vOutFolder As String = vOutputPath & vPerformDate.ToString("yyyyMMdd") & "\"
                If Not Directory.Exists(vOutFolder) Then
                    ' This path is a directory.
                    Directory.CreateDirectory(vOutFolder)
                End If
                '----create sucess folder -----
                '----create failure folder ----

                If File.Exists(vOutFolder & vXMLFile) Then
                    File.Delete(vOutFolder & vXMLFile)
                End If
                vOutputFile = vOutFolder & vXMLFile
                .Save(vOutputFile)
            End With
        Catch ex As Exception
            writeLog("Unable to import XML file :" & vXMLFile & "  " & ex.Message)
        End Try

    End Sub

    Sub writeLog(vMsg As String)
        Dim strFile As String = vOutputPath & "_error.txt"
        Dim fileExists As Boolean = File.Exists(strFile)
        Using sw As New StreamWriter(File.Open(strFile, FileMode.OpenOrCreate))
            sw.WriteLine( _
                IIf(fileExists, _
                    "Error Message in  Occured at-- " & DateTime.Now, _
                    "Start Error Log for today"))
            sw.WriteLine(Now & "--" & vMsg)
        End Using
    End Sub
End Class
