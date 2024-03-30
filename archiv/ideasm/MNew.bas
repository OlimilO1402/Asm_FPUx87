Attribute VB_Name = "MNew"
Option Explicit

Public Function FunctionHook(ByVal pFnc As Long) As FunctionHook
    Set FunctionHook = New FunctionHook: FunctionHook.New_ pFnc
End Function
