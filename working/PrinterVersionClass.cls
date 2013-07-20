VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrinterVersionClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
'Private Main As MainClass
Private Main As Object
'
'========================================================================
'   v3.3 Add-on Compatibility
'       To make an Add-on that works the same in v3.3 and v3.4, use this adapter instead of the execute above
'========================================================================
'
Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    'Set Csv = CsvObject
    Call Init(MainObject)
    Execute = GetContent(OptionString)
End Function
'
'========================================================================
'   Init()
'========================================================================
'
Public Sub Init(MainObject As Object)
    Set Main = MainObject
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("PrinterVersionClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
'=================================================================================
'   Aggregate Object Interface
'=================================================================================
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    Const printerVersionFpo = "<!-- printer-version-icon-fpo -->"
    '
    Dim QueryString As String
    Dim PageStructure As String
    Dim ATag As String
    Dim StrucRows() As String
    Dim RowPtr As Long
    Dim Colsplit() As String
    Dim Descriptor As Long
    Dim RootPageID As Long
    Dim AllowChildPageList As Boolean
    Dim AllowAutoPrint As Boolean
    Dim s As String
    Dim Body As String
    '
    ' ----- Printer Friendly Version
    '
    Body = Main.FilterInput
    If (Main.ServerPagePrintVersion) Then
        '
        ' do not process for the print version itself
        '
    ElseIf InStr(1, Body, printerVersionFpo, vbTextCompare) <> 0 Then
        AllowChildPageList = kmaEncodeBoolean(Main.GetViewingProperty("AllowChildPageList"))
        AllowAutoPrint = kmaEncodeBoolean(Main.GetViewingProperty("AutoPrint"))
        'AllowChildPageList = kmaEncodeBoolean(Main.GetAggrOption("AllowChildPageList", OptionString))
        'AllowAutoPrint = kmaEncodeBoolean(Main.GetAggrOption("AutoPrint", OptionString))
        PageStructure = Main.ContentPageStructure
        StrucRows = Split(PageStructure, vbCrLf)
        If UBound(StrucRows) > -1 Then
            Colsplit = Split(StrucRows(0), vbTab)
            If UBound(Colsplit) > 1 Then
                Descriptor = kmaEncodeInteger(Colsplit(1))
                If RootPageID = 0 Then
                    RootPageID = kmaEncodeInteger(Colsplit(2))
                End If
                If Descriptor = 0 Then
                    RootPageID = kmaEncodeInteger(Colsplit(2))
                End If
            End If
        End If
        QueryString = Main.RefreshQueryString
        QueryString = ModifyQueryString(QueryString, RequestNameRootPageID, CStr(RootPageID), True)
        QueryString = ModifyQueryString(QueryString, "method", HardCodedPagePrinterVersion, True)
        QueryString = ModifyQueryString(QueryString, RequestNameInterceptpage, "", False)
        'QueryString = ModifyQueryString(QueryString, RequestNameInterceptpage, InterceptPagePageContentPrinterVersion, True)
        QueryString = ModifyQueryString(QueryString, RequestNameAllowChildPageList, CStr(AllowChildPageList), True)
        QueryString = ModifyQueryString(QueryString, "AutoPrint", CStr(AllowAutoPrint), True)
        ATag = "<A href=""" & Main.ServerPage & "?" & QueryString & """ target=""_blank"">"
        s = ATag & "<IMG src=""/ccLib/images/IconPrintVersion2.gif"" width=""13"" height=""13"" border=""0"" align=""absmiddle""/></a>&nbsp;" & ATag & "Printer Version</A>"
        Main.FilterInput = Replace(Body, printerVersionFpo, s, , , vbTextCompare)
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("PrinterVersionClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function




