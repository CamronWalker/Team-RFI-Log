Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveSheet.ListObjects("Rfi__2").Range.AutoFilter Field:=3, Criteria1:= _
        ">1/31/2020", Operator:=xlAnd
    ActiveWindow.SmallScroll Down:=252
End Sub
