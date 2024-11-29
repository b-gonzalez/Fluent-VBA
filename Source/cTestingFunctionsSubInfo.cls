VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTestingFunctionsSubInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ITestingFunctionsSubInfo

Private pCount As Long
Private pPassed As Long
Private pFailed As Long
Private pUnexpected As Long

Private Property Let ITestingFunctionsSubInfo_Count(ByVal value As Long)
    pCount = value
End Property

Private Property Get ITestingFunctionsSubInfo_Count() As Long
    ITestingFunctionsSubInfo_Count = pCount
End Property

Private Property Let ITestingFunctionsSubInfo_Passed(ByVal value As Long)
    pPassed = value
End Property

Private Property Get ITestingFunctionsSubInfo_Passed() As Long
    ITestingFunctionsSubInfo_Passed = pPassed
End Property

Private Property Let ITestingFunctionsSubInfo_Failed(ByVal value As Long)
    pFailed = value
End Property

Private Property Get ITestingFunctionsSubInfo_Failed() As Long
    ITestingFunctionsSubInfo_Failed = pFailed
End Property

Private Property Let ITestingFunctionsSubInfo_Unexpected(ByVal value As Long)
    pUnexpected = value
End Property

Private Property Get ITestingFunctionsSubInfo_Unexpected() As Long
    ITestingFunctionsSubInfo_Unexpected = pUnexpected
End Property
