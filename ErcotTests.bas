Attribute VB_Name = "ErcotTests"
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
    filename = "cdr.00012301.0000000000000000.20240324.164708.SPPHLZNP6905_20240324_1645.csv"
    
    Dim path As String
    path = Environ$("AppData") & "\ErcotDocumentCache\"
    
    Dim file As Scripting.FileSystemObject
    Set file = New Scripting.FileSystemObject
    
    'TODO figure this out. read all lines? read as stream?
    Dim csvStream As TextStream
    csvStream = file.OpenTextFile(path & filename, ForReading)
    'Act:
    Dim csvd As Object
    Set csvd = CSVUtils.ParseCSVToDictionary(path & filename, 4)
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
