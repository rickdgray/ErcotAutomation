VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErcotPriceRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("ErcotAutomation")
Option Explicit

Public DeliveryDate As Date
Public DeliveryHour As Long
Public DeliveryInterval As Long
Public SettlementPointName As String
Public SettlementPointPrice As Variant 'Decimal

