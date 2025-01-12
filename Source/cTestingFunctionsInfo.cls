VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTestingFunctionsInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ITestingFunctionsInfo

Private pCount As Long
Private pPassed As Long
Private pFailed As Long
Private pUnexpected As Long
Private pName As String

Private Property Let ITestingFunctionsInfo_Count(ByVal value As Long)
    pCount = value
End Property

Private Property Get ITestingFunctionsInfo_Count() As Long
    ITestingFunctionsInfo_Count = pCount
End Property

Private Property Let ITestingFunctionsInfo_Passed(ByVal value As Long)
    pPassed = value
End Property

Private Property Get ITestingFunctionsInfo_Passed() As Long
    ITestingFunctionsInfo_Passed = pPassed
End Property

Private Property Let ITestingFunctionsInfo_Failed(ByVal value As Long)
    pFailed = value
End Property

Private Property Get ITestingFunctionsInfo_Failed() As Long
    ITestingFunctionsInfo_Failed = pFailed
End Property

Private Property Let ITestingFunctionsInfo_Unexpected(ByVal value As Long)
    pUnexpected = value
End Property

Private Property Get ITestingFunctionsInfo_Unexpected() As Long
    ITestingFunctionsInfo_Unexpected = pUnexpected
End Property

Private Property Let ITestingFunctionsInfo_Name(ByVal value As String)
    pName = value
End Property

Private Property Get ITestingFunctionsInfo_Name() As String
    ITestingFunctionsInfo_Name = pName
End Property
