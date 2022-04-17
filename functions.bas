Attribute VB_Name = "functions"
Option Explicit

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean

    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)

End Function
