Imports ROTWrapper
Imports RunningObjectTableData
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Public Class Form1

    Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Function GetItemsInRot(ByVal name As String) As Object
        Dim instances As Generic.List(Of Object) = WrappedObject.GetRunningInstances(name)
        Try
            Debug.Print("Instance count: " & instances.Count)
            Dim swapp As SldWorks = instances.Item(0)
            If Not swapp Is Nothing Then
                swapp.SendMsgToUser2("Got solidworks!", swMessageBoxIcon_e.swMbInformation, swMessageBoxBtn_e.swMbOk)
            End If

        Catch ex As Exception

        End Try

    End Function

    Private Sub BTNRead_Click(sender As Object, e As EventArgs) Handles BTNRead.Click
        GetItemsInRot(TextBox1.Text)

    End Sub
End Class
