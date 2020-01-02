Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

Public Class NativeMethods
    <DllImport("ole32.dll")>
    Public Shared Function GetRunningObjectTable(ByVal reserved As UInteger, <Out> ByRef pprot As System.Runtime.InteropServices.ComTypes.IRunningObjectTable) As Integer
    End Function

    <DllImport("ole32.dll")>
    Public Shared Function CreateFileMoniker(<MarshalAs(UnmanagedType.LPWStr)> ByVal lpszPathName As String, <Out> ByRef ppmk As System.Runtime.InteropServices.ComTypes.IMoniker) As Integer
    End Function

    <DllImport("oleaut32.dll")>
    Public Shared Function RevokeActiveObject(ByVal register As Integer, ByVal reserved As IntPtr) As Integer
    End Function
    <DllImport("ole32.dll")>
    Public Shared Function CreateBindCtx(ByVal reserved As UInteger, <Out> ByRef ppbc As IBindCtx) As Integer

    End Function
End Class
