Attribute VB_Name = "GlobalDeffinitions"
#If VBA7 Then
    Declare PtrSafe Function FindWindowA Lib "USER32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Declare PtrSafe Function GetWindowLongA Lib "USER32" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
    Declare PtrSafe Function SetWindowLongA Lib "USER32" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#Else
    Declare Function FindWindowA Lib "USER32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Declare Function GetWindowLongA Lib "USER32" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Declare Function SetWindowLongA Lib "USER32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If

Public Const TaskIndex As Long = 1
Public Const EnvironmentIndex As Long = 2
Public Const MissionIndex As Long = 3
Public Const InterfaceIndex As Long = 4
Public Const AutomationIndex As Long = 5
Public Const OtherIndex As Long = 6

Public Const ActExecuting As String = "actExecuting"
Public Const ActDone As String = "actDone"

Public Steps As Collection
Public Variables As Collection

Public Function ChangeBetweenSteps(VariableCollection As Collection, Step As Long) As Boolean
    Dim j As Long
    ChangeBetweenSteps = False
    For j = 1 To VariableCollection(Step).Count
        If VariableCollection(Step)(j).Changed Then
            ChangeBetweenSteps = True
            Exit Function
        End If
    Next j
End Function

Public Function ActionExecuting(VariableCollection As Collection, Step As Long) As Boolean
    Dim j As Long
    ActionExecuting = False
    For j = 1 To VariableCollection(Step).Count
        If Left(VariableCollection(Step)(j).GetName, 1) = "h" And VariableCollection(Step)(j).GetValue = ActExecuting Then
            ActionExecuting = True
            Exit Function
        End If
    Next j
End Function
