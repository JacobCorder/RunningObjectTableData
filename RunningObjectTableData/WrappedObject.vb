Imports System.Runtime.InteropServices.ComTypes

Public Class WrappedObject
    Implements IDisposable
    Public Shared Function GetRunningInstances(ByVal ProgID As String) As Generic.List(Of Object)
        Return GetRunningInstances(New String() {ProgID})
    End Function
    Public Shared Function GetRunningInstances(ByVal progIds As String()) As Generic.List(Of Object)
        Dim clsIds As Generic.List(Of String) = New Generic.List(Of String)()

        For Each progId As String In progIds
            Dim type As Type = Type.GetTypeFromProgID(progId)
            If type IsNot Nothing Then clsIds.Add(type.GUID.ToString().ToUpper())
        Next
        Dim Rot As IRunningObjectTable = Nothing
        NativeMethods.GetRunningObjectTable(0, Rot)
        If Rot Is Nothing Then Return Nothing
        Dim monikerEnumerator As IEnumMoniker = Nothing
        Rot.EnumRunning(monikerEnumerator)
        If monikerEnumerator Is Nothing Then Return Nothing
        monikerEnumerator.Reset()
        Dim instances As Generic.List(Of Object) = Nothing
        Dim pNumFetched As IntPtr = New IntPtr()
        Dim monikers As IMoniker() = New IMoniker(0) {}
        Dim Rotct As Integer = 0
        While monikerEnumerator.Next(1, monikers, pNumFetched) = 0
            Dim bindCtx As IBindCtx = Nothing
            NativeMethods.CreateBindCtx(0, bindCtx)
            If bindCtx Is Nothing Then Continue While
            Dim displayName As String = ""
            monikers(0).GetDisplayName(bindCtx, Nothing, displayName)
            Rotct += 1
            For Each clsId As String In clsIds
                Debug.Print("ROT ITEM: " & clsId)
                If UCase(displayName).Contains(clsId) Then 'displayName.ToUpper().IndexOf(clsId) > 0 Then
                    Dim ComObject As Object = Nothing
                    Rot.GetObject(monikers(0), ComObject)
                    If ComObject Is Nothing Then Continue For
                    If instances Is Nothing Then instances = New Generic.List(Of Object)
                    instances.Add(ComObject)
                    Exit For
                End If
                Debug.Print(displayName)
            Next

        End While
        Debug.Print("ROT Object Count: " & Rotct)
        Return instances
    End Function

    Public Function RegisterObject(ByVal obj As Object, ByVal stringId As String) As Integer
        Dim CurrentObj As Generic.List(Of Object) = GetRunningInstances(stringId)
        If CurrentObj Is Nothing Then
            Dim regId As Integer = -1
            Dim pROT As System.Runtime.InteropServices.ComTypes.IRunningObjectTable = Nothing
            Dim pMoniker As System.Runtime.InteropServices.ComTypes.IMoniker = Nothing
            Dim hr As Integer = NativeMethods.GetRunningObjectTable(CUInt(0), pROT)
            If hr <> 0 Then Return (hr)
            hr = NativeMethods.CreateFileMoniker(stringId, pMoniker)
            If hr <> 0 Then Return (hr)
            Const ROTFLAGS_REGISTRATIONKEEPSALIVE As Integer = 1
            Const ROTFLAGS_ALLOWANYCLIENT As Integer = 2 'Find a way to use this value


            regId = pROT.Register(ROTFLAGS_REGISTRATIONKEEPSALIVE, obj, pMoniker)
            _RegisteredObjects.Add(New ObjectInROT(obj, regId))
            Return regId

        End If
        Return 0
    End Function

    Private _RegisteredObjects As List(Of ObjectInROT) = New List(Of ObjectInROT)()

    ' <Obsolete("Please refactor code that uses this function, it is a simple work-around to simulate inline assignment in VB!")>
    'Private Shared Function __InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
    '    target = value
    '    Return value
    'End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                If Not _RegisteredObjects Is Nothing Then
                    If _RegisteredObjects.Count > 0 Then
                        Try
                            Dim pROT As System.Runtime.InteropServices.ComTypes.IRunningObjectTable = Nothing
                            Dim hr As Integer = NativeMethods.GetRunningObjectTable(CUInt(0), pROT)
                            For i = 0 To _RegisteredObjects.Count - 1
                                Try
                                    If Not _RegisteredObjects.Item(i) Is Nothing Then
                                        Try
                                            pROT.Revoke(_RegisteredObjects.Item(i).regId)
                                            _RegisteredObjects.Item(i) = Nothing
                                        Catch ex As Exception
                                            Debug.Print("Failed to unregister the object in the ROT")
                                        End Try
                                    End If
                                Catch ex As Exception

                                End Try
                            Next
                        Catch ex As Exception

                        End Try
                    End If
                End If
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class
