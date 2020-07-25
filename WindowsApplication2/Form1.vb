Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.IO

Public Class Form1
    Public Shared Function ExtractIcon(ByVal path As String, Optional ByVal index As Integer = 0) As Icon
        'http://stackoverflow.com/questions/1988393/need-working-example-of-shell32s-extractassociatedicon-function-in-net

        Dim handle As IntPtr = ExtractAssociatedIcon(IntPtr.Zero, path, index)
        If handle = IntPtr.Zero Then Throw New Win32Exception(Marshal.GetLastWin32Error())
        Dim retval As Icon = Nothing
        Using temp As Icon = Icon.FromHandle(handle)
            retval = CType(temp.Clone(), Icon)
            DestroyIcon(handle)
        End Using
        Return retval
    End Function

    Private Declare Auto Function ExtractAssociatedIcon Lib "shell32" ( _
        ByVal hInst As IntPtr, ByVal path As String, ByRef index As Integer) As IntPtr
    Private Declare Auto Function DestroyIcon Lib "user32" (ByVal hIcon As IntPtr) As Boolean

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        Dim Icon As Icon = ExtractIcon("G:\fernandobhz\BoletoAgogeWindowsForms\BoletoAgogeWindowsForms\bin\Debug\Boletos Agoge Sistemas.exe")

        Dim MS As New MemoryStream

        Icon.Save(MS)

        My.Computer.FileSystem.WriteAllBytes("g:\icon.ico", MS.ToArray, False)

        MsgBox("OK")


    End Sub

End Class