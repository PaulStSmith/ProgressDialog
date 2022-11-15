Imports System.Runtime.InteropServices

Public Class frmTest

    <Flags()> _
    Private Enum LoadLibraryExFlags As UInteger
        DontResolveDllReferences = &H1
        LoadLibraryAsDatafile = &H2
        LoadWithAlteredSearchPath = &H8
        LoadIgnoreCodeAuthzLevel = &H10
    End Enum

    <DllImport("kernel32", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function LoadLibraryEx( _
         <MarshalAs(UnmanagedType.LPTStr)> ByVal lpFileName As String, _
                                           ByVal hFile As IntPtr, _
                                           ByVal dwFlags As LoadLibraryExFlags) As IntPtr
    End Function

    Private Sub btnTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTest.Click

        Call DoSomethingLong()

    End Sub

    Private Sub DoSomethingLong()

        '*
        '* Create a new instance of the ProgressDialog Class
        '*
        Dim pd As New ProgressDialog(Me)

        '*
        '* Initialize the Dialog Properties
        '*
        With pd
            .Complete = 0
            .Total = 10000
            .Title = "Title"
            .Animation = ProgressDialog.dlgAnimations.FileMove
        End With

        '*
        '* Starts the Dialog
        '*
        pd.Start()

        '*
        '* Do something long
        '*
        For i As UInteger = 1 To 10000
            pd.Complete += 10
            pd.SetLineText(ProgressDialog.dlgLines.LineTwo, "Item No." & i)

            '*
            '* Checks for user cancellation
            '*
            If (pd.UserCancelled) Then
                Exit For
            End If

            System.Threading.Thread.Sleep(50)
        Next

        '*
        '* Closes the dialog
        '*
        pd.Stop()

    End Sub

End Class
