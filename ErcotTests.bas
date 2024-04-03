Attribute VB_Name = "ErcotTests"
'@IgnoreModule UseMeaningfulName, EmptyMethod, VariableNotUsed, IndexedUnboundDefaultMemberAccess
'@TestModule
'@Folder("Tests")

Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
End Sub

'@TestCleanup
Private Sub TestCleanup()
End Sub

'@TestMethod("Uncategorized")
Private Sub ParseCsv_ValidCsvReturnsSuccess()
    On Error GoTo TestFail
    'Arrange:
    Dim filename As String
    filename = "cdr.00012301.0000000000000000.20240329.010211.SPPHLZNP6905_20240329_0100.csv"
    
    Dim file As Scripting.FileSystemObject
    Set file = New Scripting.FileSystemObject
    
    Dim csvStream As TextStream
    Set csvStream = file.OpenTextFile(Application.ActiveWorkbook.path & "\" & filename, ForReading)
    
    Dim csv As String
    csv = csvStream.ReadAll
    
    'Act:
    Dim csvd As Object
    Set csvd = CSVUtils.ParseCSVToDictionary(csv, 4)
    
    'Assert:
    Debug.Assert StrComp(csvd("AMOCO_PUN1")(2), "1", vbTextCompare) = 0
    Assert.Succeed
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Failed to parse CSV file: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub ParseCsvToPriceTable_GetCollectionFromCsvSuccess()
    On Error GoTo TestFail
    'Arrange:
    
    'Act:
    Dim priceTable As Collection
    Set priceTable = ErcotAutomation.ParseCsvToPriceTable("CVC_CC1")
    
    Dim priceRecord As ErcotPriceRecord
    For Each priceRecord In priceTable
        Debug.Print priceRecord.SettlementPointName
        Debug.Print priceRecord.SettlementPointPrice
    Next priceRecord
    
    'Assert:
    Assert.Succeed
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Failed to parse CSV file: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub ErcotAutomation_FullRunReturnsSuccess()
    On Error GoTo TestFail
    'Arrange:
    
    'Act:
    ErcotAutomation.UpdatePrices
    
    'Assert:
    Assert.Succeed
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Failed to parse CSV file: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

