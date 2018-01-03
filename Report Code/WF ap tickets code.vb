/* CIRESON SSRS VB CODE */


Const LocTableStringQuery As String = "R/S[@N='{0}']/text()"

Const TimeZoneParameterName As String = "TimeZone"
Const SD_BaseTypeParameterName As String = "StartDate_BaseType"
Const SD_BaseValueParameterName As String = "StartDate_BaseValue"
Const SD_OffsetTypeParameterName As String = "StartDate_OffsetType"
Const SD_OffsetValueParameterName As String = "StartDate_OffsetValue"
Const ED_BaseTypeParameterName As String = "EndDate_BaseType"
Const ED_BaseValueParameterName As String = "EndDate_BaseValue"
Const ED_OffsetTypeParameterName As String = "EndDate_OffsetType"
Const ED_OffsetValueParameterName As String = "EndDate_OffsetValue"
Const IsRelativeTimeSupported As Boolean = False
Const TimeTypeParameterName As String = "TimeType"
Const TimeWeekMapParameterName As String = "TimeWeekMap"

Dim LocTables As System.Collections.Generic.Dictionary(Of String, Microsoft.EnterpriseManagement.Reporting.XmlStringTable)
Dim ReportTimeZone As Microsoft.EnterpriseManagement.Reporting.TimeZoneCoreInformation
Dim ReportStartDate As DateTime
Dim ReportEndDate As DateTime
Dim ReportTime As Microsoft.EnterpriseManagement.Reporting.ParameterProcessor.RelativeTime
Dim ReportCulture As System.Globalization.CultureInfo
Dim ParameterProcessor As Microsoft.EnterpriseManagement.Reporting.ParameterProcessor
Dim TargetList As String

Protected Overrides Sub OnInit()
  LocTables = new System.Collections.Generic.Dictionary(Of String, Microsoft.EnterpriseManagement.Reporting.XmlStringTable)()
  ReportTimeZone =Nothing
  ReportStartDate = DateTime.MinValue
  ReportEndDate = DateTime.MinValue
  ReportTime = Nothing
  ReportCulture = System.Globalization.CultureInfo.GetCultureInfo(Report.User("Language"))
  ParameterProcessor = New Microsoft.EnterpriseManagement.Reporting.ParameterProcessor(ReportCulture)
  TargetList =Nothing
End Sub

'Public Function GetCallingManagementGroupId() As String
'   Return 'Microsoft.EnterpriseManagement.Reporting.ReportingConfiguration.M'anagementGroupId
'End Function

Public Function GetReportLocLanguageCode() As String
   Return ReportCulture.ThreeLetterWindowsLanguageName
End Function

Public Function GetReportLCID() As Integer
    Return ReportCulture.LCID
End Function

Public Function GetLocTable(Name As String) As Microsoft.EnterpriseManagement.Reporting.XmlStringTable
   Dim LocTable As Microsoft.EnterpriseManagement.Reporting.XmlStringTable

   If Not LocTables.TryGetValue(Name, LocTable) Then
      LocTable = New Microsoft.EnterpriseManagement.Reporting.XmlStringTable(LocTableStringQuery, Report.Parameters(Name).Value)
      LocTables.Add(Name, LocTable)
   End If

   Return LocTable
End Function

Public Function GetReportTimeZone() As Microsoft.EnterpriseManagement.Reporting.TimeZoneCoreInformation
   If IsNothing(ReportTimeZone) Then ReportTimeZone = Microsoft.EnterpriseManagement.Reporting.TimeZoneCoreInformation.FromValueString(Report.Parameters(TimeZoneParameterName).Value)
   If IsNothing(ReportTimeZone) Then ReportTimeZone = Microsoft.EnterpriseManagement.Reporting.TimeZoneInformation.Current
   Return ReportTimeZone
End Function

Public Function ToDbDate(ByVal DateValue As DateTime) As DateTime
   return GetReportTimeZone.ToUniversalTime(DateValue)
End Function

Public Function ToReportDate(ByVal DateValue As DateTime) As DateTime
   return GetReportTimeZone.ToLocalTime(DateValue)
End Function

Public Function GetReportStartDate() As DateTime
  If (ReportStartDate = DateTime.MinValue) Then
    If (IsRelativeTimeSupported) Then
      ReportStartDate = ParameterProcessor.GetDateTime(ToReportDate(DateTime.UtcNow), Report.Parameters(SD_BaseTypeParameterName).Value, Report.Parameters(SD_BaseValueParameterName).Value, Report.Parameters(SD_OffsetTypeParameterName).Value, Report.Parameters(SD_OffsetValueParameterName).Value, Report.Parameters(TimeTypeParameterName).Value)
    Else
      ReportStartDate = ParameterProcessor.GetDateTime(ToReportDate(DateTime.UtcNow), Report.Parameters(SD_BaseTypeParameterName).Value, Report.Parameters(SD_BaseValueParameterName).Value, Report.Parameters(SD_OffsetTypeParameterName).Value, Report.Parameters(SD_OffsetValueParameterName).Value)
    End if
  End If
  return ReportStartDate
End Function

Public Function GetReportEndDate() As DateTime
  If (ReportEndDate = DateTime.MinValue) Then
    If (IsRelativeTimeSupported) Then
      ReportEndDate = ParameterProcessor.GetDateTime(ToReportDate(DateTime.UtcNow), Report.Parameters(ED_BaseTypeParameterName).Value, Report.Parameters(ED_BaseValueParameterName).Value, Report.Parameters(ED_OffsetTypeParameterName).Value, Report.Parameters(ED_OffsetValueParameterName).Value, Report.Parameters(TimeTypeParameterName).Value)
      If IsBusinessHours(GetReportTimeFilter()) Then ReportEndDate = ReportCulture.Calendar.AddDays(ReportEndDate, 1)
    Else
      ReportEndDate = ParameterProcessor.GetDateTime(ToReportDate(DateTime.UtcNow), Report.Parameters(ED_BaseTypeParameterName).Value, Report.Parameters(ED_BaseValueParameterName).Value, Report.Parameters(ED_OffsetTypeParameterName).Value, Report.Parameters(ED_OffsetValueParameterName).Value)
    End if
  End If
  return ReportEndDate
End Function

Public Function GetReportTimeFilter() As Microsoft.EnterpriseManagement.Reporting.ParameterProcessor.RelativeTime
  If IsNothing(ReportTime) Then ReportTime = New Microsoft.EnterpriseManagement.Reporting.ParameterProcessor.RelativeTime(Report.Parameters(TimeTypeParameterName).Value, Report.Parameters(SD_BaseValueParameterName).Value, Report.Parameters(ED_BaseValueParameterName).Value, CStr(Join(Report.Parameters(TimeWeekMapParameterName).Value, ",")))
  return ReportTime
End Function

Public Function IsBusinessHours(Value As Microsoft.EnterpriseManagement.Reporting.ParameterProcessor.RelativeTime) As Boolean
  return (Not IsNothing(Value)) And (Value.TimeType = Microsoft.EnterpriseManagement.Reporting.ParameterProcessor.RelativeTimeType.Business)
End Function

Public Function FormatDateTime(Format As String, Value As DateTime) As String
  return Value.ToString(Format, ReportCulture)
End Function

Public Function FormatNumber(Format As String, Value As Decimal) As String
  return Value.ToString(Format, ReportCulture)
End Function

Public Function FormatString(Format As String, ParamArray Values() as  Object) As String
  return String.Format(ReportCulture, Format, Values)
End Function

Public Function NullFormatString(Format As String, Value as  String) As String
  return IIF(String.IsNullOrEmpty(Value), String.Empty, String.Format(ReportCulture, Format, Value))
End Function

Public Function FormatBusinessHours(Format As String, Value As Microsoft.EnterpriseManagement.Reporting.ParameterProcessor.RelativeTime) As String
  Dim result As String

  If IsBusinessHours(Value) Then
 
    Dim firstDay As DayOfWeek
    Dim days As System.Collections.Generic.List(Of String)

    firstDay = ReportCulture.DateTimeFormat.FirstDayOfWeek
    days = new System.Collections.Generic.List(Of String)()

    For loopDay As DayOfWeek = DayOfWeek.Sunday To DayOfWeek.Saturday
      Dim day As DayOfWeek
      day = CType((CInt(loopDay) + CInt(firstDay)) Mod 7, DayOfWeek)

      If value.WeekMap.Contains(day) Then days.Add(ReportCulture.DateTimeFormat.GetAbbreviatedDayName(day))
    Next loopDay

     result = FormatString(Format, DateTime.Today.Add(Value.StartTime).ToString(ReportCulture.DateTimeFormat.ShortTimePattern), DateTime.Today.Add(Value.EndTime).ToString(ReportCulture.DateTimeFormat.ShortTimePattern), String.Join(",", days.ToArray()))

  Else
     result = String.Empty
  End if
  
  return result
End Function

Public Function BuildXmlValueList(ByVal ValueList() As Object) As String
    Return Microsoft.EnterpriseManagement.Reporting.MultiValueParameter.ToXml("Data", "Value", ValueList)
End Function


REM -------------------------------------------------

Dim ReportFormatList As System.Collections.Generic.Dictionary(Of String, IDataFormatter)
Const FormatListParameterName As String = "DataFormat"

REM -------------------------------------------------

    Public Interface IDataFormatter
        Function FormatData(ByVal Value As String) As String
    End Interface

    Public Class RankDataFormatter
        Implements IDataFormatter

        Private Series As System.Collections.Generic.SortedDictionary(Of Decimal, String)
        Private Culture As System.Globalization.CultureInfo

        Public Sub New(ByVal Config As System.Xml.XmlNode, ByVal ReportCulture As System.Globalization.CultureInfo)
            Culture = ReportCulture
            Series = New System.Collections.Generic.SortedDictionary(Of Decimal, String)()

            For Each SeriesXml As System.Xml.XmlNode In Config.ChildNodes
                Series.Add(Decimal.Parse(SeriesXml.Attributes("Rank").Value), SeriesXml.Attributes("Format").Value)
            Next
        End Sub

        Public Function FormatData(ByVal Value As String) As String Implements IDataFormatter.FormatData
            Dim Result As String = CDec(Value).ToString("G", Culture)

            For Each FormatItem As System.Collections.Generic.KeyValuePair(Of Decimal, String) In Series
                Dim RunningValue As Decimal = Math.Round(Value / FormatItem.Key)
                If RunningValue > 0 Then
                    Result = String.Format(Culture, FormatItem.Value, RunningValue.ToString("G", Culture))
                Else
                    Exit For
                End If
            Next

            Return Result
        End Function
    End Class

    Public Class LookupDataFormatter
        Implements IDataFormatter

        Private MappingTable As System.Collections.Generic.IDictionary(Of String, String)

        Public Sub New(ByVal Config As System.Xml.XmlNode)
            MappingTable = New System.Collections.Generic.Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            For Each SeriesXml As System.Xml.XmlNode In Config.ChildNodes
                MappingTable.Add(Decimal.Parse(SeriesXml.Attributes("Value").Value), SeriesXml.Attributes("Label").Value)
            Next
        End Sub

        Public Function FormatData(ByVal Value As String) As String Implements IDataFormatter.FormatData
            Dim Result As String
            If MappingTable.ContainsKey(Value) Then
                Result = MappingTable(Value)
            Else
                Result = Value
            End If
            Return Result
        End Function
    End Class

    Public Function GetReportFormatList() As System.Collections.Generic.Dictionary(Of String, IDataFormatter)
        If IsNothing(ReportFormatList) Then
            Dim Xml As System.Xml.XmlDocument = New System.Xml.XmlDocument()
            Xml.LoadXml(Report.Parameters(FormatListParameterName).Value)

            ReportFormatList = New System.Collections.Generic.Dictionary(Of String, IDataFormatter)(StringComparer.OrdinalIgnoreCase)
            For Each Node As System.Xml.XmlNode In Xml.DocumentElement.ChildNodes
                Dim Formatter As IDataFormatter = Nothing

                Select Case Node.Attributes("Type").Value.ToUpper()
                    Case "RANK"
                        Formatter = New RankDataFormatter(Node, ReportCulture)
                    Case "LOOKUP"
                        Formatter = New LookupDataFormatter(Node)
                End Select

                If Not IsNothing(Formatter) Then
                    ReportFormatList.Add(Node.Attributes("Name").Value, Formatter)
                End If
            Next
        End If
        Return ReportFormatList
    End Function

    Public Function FormatData(ByVal FormatName As String, ByVal DataType As String, ByVal Value As String)
        Dim FormatList As System.Collections.Generic.Dictionary(Of String, IDataFormatter) = GetReportFormatList()
        If Not String.IsNullOrEmpty(FormatName) And Not IsNothing(FormatList) Then
            If FormatList.ContainsKey(FormatName) Then
                Return FormatList(FormatName).FormatData(Value)
            End If
        End If

        If Not String.IsNullOrEmpty(DataType) Then
            If DataType = "DateTime" Then
                Return FormatDateTime("g", CDate(Value))
            ElseIf DataType.StartsWith("UInt") Then
                Return FormatNumber("G", Value)
            End If
        End If

        Return Value
    End Function

REM ----------------------------------------

Public Dim CurrentObjectList As String

    Public Function InitList(ByRef List As String) As String
        List  = String.Empty
        Return List
    End Function

    Public Function AddListItem(ByRef List As String, Item as String) As String
        List = List + Item
        Return List
    End Function

    Public Function GetObjectList(ByVal OptionsXml As String) As String()
        Dim Xml As System.Xml.XmlDocument
        Xml = New System.Xml.XmlDocument()
        Xml.LoadXml(OptionsXml)

        Dim Result As System.Collections.Generic.List(Of String)
        Result = New System.Collections.Generic.List(Of String)
        For Each ObjectNode As System.Xml.XmlNode In Xml.SelectNodes("/Value/Object")
            If Not Result.Contains(ObjectNode.InnerText) Then
                Result.Add(ObjectNode.InnerText)
            End If
        Next

        Return Result.ToArray()
    End Function
