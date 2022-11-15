' Windows.Controls.ProgressDialog
' 
' Copyright (c) 2004-2005
' by Paulo Santos. (http://pjondevelopment.blogspot.com)
' All Rights Reserved
' 
' The above copyright notice and this permission notice shall  be
' included in all copies or substantial portions of the Software.
' 
' DISCLAIMER
' ==========
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
' EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
' OF  MERCHANTABILITY,  FITNESS  FOR  A  PARTICULAR  PURPOSE  AND
' NONINFRINGEMENT. IN NO EVENT SHALL  THE  AUTHORS  OR  COPYRIGHT 
' HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES  OR  OTHER  LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT  OR  OTHERWISE,  ARISING 
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE  USE  OR
' OTHER DEALINGS IN THE SOFTWARE.
'  

Imports System.Runtime.InteropServices

Namespace System.Windows.Forms

    <ComImport()> _
    <Guid("F8383852-FCD3-11D1-A6B9-006097DF5BD4")> _
    Friend Class shellProgressDialog
    End Class

    <ComImport()> _
    <Guid("00000114-0000-0000-C000-000000000046")> _
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
    Friend Interface IOleWindow

        <PreserveSig()> _
        Sub GetWindow(<Out(), MarshalAs(UnmanagedType.SysInt)> ByVal pHwnd As IntPtr)

    End Interface

    <ComImport()> _
    <Guid("EBBC7C04-315E-11D2-B62F-006097DF5BD4")> _
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
    Friend Interface IProgressDialog

        <PreserveSig()> _
        Sub StartProgressDialog( _
                                    ByVal hwndParent As IntPtr, _
<MarshalAs(UnmanagedType.IUnknown)> ByVal punkEnableModless As Object, _
                                    ByVal dwFlags As UInteger, _
                                    ByVal pvReserved As IntPtr)

        <PreserveSig()> _
        Sub StopProgressDialog()

        <PreserveSig()> _
        Sub SetTitle( _
<MarshalAs(UnmanagedType.LPWStr)> ByVal pwzTitle As String)

        <PreserveSig()> _
        Sub SetAnimation(ByVal hInstAnimation As IntPtr, ByVal idAnimation As UShort)

        <PreserveSig()> _
        Function HasUserCancelled() As <MarshalAs(UnmanagedType.Bool)> Boolean

        <PreserveSig()> _
        Sub SetProgress(ByVal dwCompleted As UInteger, ByVal dwTotal As UInteger)

        <PreserveSig()> _
        Sub SetProgress64(ByVal dwCompleted As ULong, ByVal dwTotal As ULong)

        <PreserveSig()> _
        Sub SetLine( _
                                  ByVal dwLineNum As Integer, _
<MarshalAs(UnmanagedType.LPWStr)> ByVal pwzString As String, _
                                  ByVal fCompactPath As Integer, _
                                  ByVal pvReserved As IntPtr)

        <PreserveSig()> _
        Sub SetCancelMsg( _
<MarshalAs(UnmanagedType.LPWStr)> ByVal pwzCancelMsg As String, _
                                  ByVal pvReserved As Object)

        <PreserveSig()> _
        Sub Timer(ByVal dwTimerAction As UInteger, ByVal pvReserved As Object)

    End Interface

    Public Class ProgressDialog

        Private Const WM_USER As Integer = &H400
        Private Const ACM_OPEN As Integer = WM_USER + 103 ' -> ACM_OPENW : Unicode Version

        <DllImport("user32.dll")> _
        Private Overloads Shared Function SendMessage( _
                ByVal hWnd As IntPtr, _
                ByVal uMsg As Integer, _
                ByVal lParam As Integer, _
                ByVal lpData As Integer) As Integer
        End Function

        <DllImport("user32.dll")> _
        Private Overloads Shared Function SendMessage( _
                ByVal hWnd As IntPtr, _
                ByVal uMsg As Integer, _
                ByVal lParam As Integer, _
                ByVal lpData As String) As Integer
        End Function

        ''' <summary>
        ''' Lists all available default animations.
        ''' </summary>
        Public Enum dlgAnimations As UShort
            Custom = 0
            SearchFlashlight = 150
            SearchDocument = 151
            SearchComputer = 152
            FileMove = 160
            FileCopy = 161
            ToRecycleBinDelete = 162
            FromRecycleBinDelete = 163
            PermanentDelete = 164
            FlyingPapers = 165
            SearchGlobe = 166
            FileMove95 = 167
            FileCopy95 = 168
            FileDelete95 = 169
            InternetCopy = 170
            NoAnimation = UShort.MaxValue
        End Enum

        ''' <summary>
        ''' Defines all available lines
        ''' </summary>
        Public Enum dlgLines As Byte
            LineOne = 1
            LineTwo
            LineThree
        End Enum

        <Flags()> _
        Private Enum IPD_Flags As UInteger
            Normal = &H0
            Modal = &H1
            AutoTime = &H2
            NoTime = &H4
            NoMinimize = &H8
            NoProgressbar = &H10
        End Enum

        <Flags()> _
        Private Enum LoadLibraryExFlags As UInteger
            DontResolveDllReferences = &H1
            LoadLibraryAsDatafile = &H2
            LoadWithAlteredSearchPath = &H8
            LoadIgnoreCodeAuthzLevel = &H10
        End Enum

        Private Const PDTIMER_RESET As UInteger = 1
        Private Const MAX_CAPACITY As Integer = 45

        <DllImport("shlwapi.dll", CharSet:=CharSet.Auto)> _
        Private Shared Function PathCompactPathEx( _
                                    <Out()> ByVal pszOut As Text.StringBuilder, _
                                            ByVal szPath As String, _
                                            ByVal cchMax As Integer, _
                                            ByVal dwFlags As Integer) As Integer
        End Function

        <DllImport("kernel32", CharSet:=CharSet.Auto, SetLastError:=True)> _
        Private Shared Function LoadLibraryEx( _
             <MarshalAs(UnmanagedType.LPTStr)> ByVal lpFileName As String, _
                                               ByVal hFile As IntPtr, _
                                               ByVal dwFlags As LoadLibraryExFlags) As IntPtr
        End Function

        Private __Dialog As IProgressDialog = Nothing

        Private __Flags As IPD_Flags = IPD_Flags.Normal
        Private __Animation As dlgAnimations = dlgAnimations.NoAnimation

        Private __hShellAnimation As IntPtr = IntPtr.Zero
        Private __UserCancelled As Boolean = False

        Private __hWndParent As IntPtr = IntPtr.Zero
        Private __Title As String = "Please wait..."
        Private __CancelMessage As String = ""

        Private __Total As ULong = 0
        Private __Complete As ULong = 0

        Private __DlgClosed As Boolean = True

        Private Sub New()
            __Dialog = New shellProgressDialog
        End Sub

        Public Sub New(ByVal hWnd As IntPtr)
            Me.New()
            __hWndParent = hWnd
        End Sub

        Public Sub New(ByVal f As Form)
            Me.New(f.Handle)
        End Sub

        ''' <summary>
        ''' Defines the animation played by the dialog window while processing.
        ''' </summary>
        ''' <remarks>Setting this property after the dialog is being shown has no effect.</remarks>
        ''' <value>NoAnimation</value>
        Public Property Animation() As dlgAnimations
            Get
                Return __Animation
            End Get
            Set(ByVal value As dlgAnimations)
                If (__Animation = value) Then Exit Property

                If (value <> dlgAnimations.NoAnimation) AndAlso _
                   (value <> dlgAnimations.Custom) Then
                    __Animation = value

                    If (__hShellAnimation = IntPtr.Zero) Then
                        __hShellAnimation = LoadLibraryEx("shell32.dll" & Chr(0), IntPtr.Zero, LoadLibraryExFlags.DontResolveDllReferences Or LoadLibraryExFlags.LoadLibraryAsDatafile)
                    End If

                    If (__hShellAnimation = IntPtr.Zero) Then
                        Throw New Exception("Could not load ""shell32.dll"".")
                    End If

                    __Dialog.SetAnimation(__hShellAnimation, __Animation)
                Else
                    __Dialog.SetAnimation(__hShellAnimation, dlgAnimations.NoAnimation)
                End If
            End Set
        End Property

        ''' <summary>
        ''' Defines the custom animation to be used.
        ''' </summary>
        Public Sub SetCustomAnimation(ByVal hLibrary As IntPtr, ByVal resNumber As UShort)

            If (__DlgClosed) AndAlso (__Animation = dlgAnimations.Custom) Then
                If (hLibrary = IntPtr.Zero) Then
                    Throw New InvalidExpressionException("hLibrary can not be a IntPtr.Zero")
                End If
                __Dialog.SetAnimation(hLibrary, resNumber)
            End If

        End Sub

        ''' <summary>
        ''' The message to be displayed if the cancelation process needs to take time.
        ''' </summary>
        ''' <value>String.Empty</value>
        Public Property CancelMessage() As String
            Get
                Return __CancelMessage
            End Get
            Set(ByVal value As String)
                If (__CancelMessage = value) Then Exit Property
                __CancelMessage = value
                If (value.EndsWith(Chr(0))) Then
                    __Dialog.SetCancelMsg(value, Nothing)
                Else
                    __Dialog.SetCancelMsg(value & Chr(0), Nothing)
                End If
            End Set
        End Property

        ''' <summary>
        ''' Gets or Sets the amount of the job done.
        ''' </summary>
        Public Property Complete() As ULong
            Get
                Return __Complete
            End Get
            Set(ByVal value As ULong)
                __Complete = value
                If (Not __DlgClosed) Then
                    __Dialog.SetProgress64(__Complete, __Total)
                End If
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the flag that forces the progress window to be modal.
        ''' </summary>
        ''' <remarks>Setting this property after the dialog is being shown has no effect.</remarks>
        ''' <value>False.</value>
        Public Property IsModal() As Boolean
            Get
                Return CBool(__Flags And IPD_Flags.Modal)
            End Get
            Set(ByVal value As Boolean)
                If (Not __DlgClosed) Then Exit Property
                If (value) Then
                    __Flags = __Flags Or IPD_Flags.Modal
                Else
                    __Flags = __Flags And (Not IPD_Flags.Modal)
                End If
            End Set
        End Property

        ''' <summary>
        ''' Gets or Sets the automatic calculation of the remaining time.
        ''' </summary>
        ''' <remarks>Setting this property after the dialog is being shown has no effect.</remarks>
        ''' <value>True</value>
        Public Property AutomaticRemainingTime() As Boolean
            Get
                Return CBool(__Flags And IPD_Flags.AutoTime)
            End Get
            Set(ByVal value As Boolean)
                If (Not __DlgClosed) Then Exit Property
                If (value) Then
                    __Flags = __Flags Or IPD_Flags.AutoTime
                Else
                    __Flags = __Flags And (Not IPD_Flags.AutoTime)
                End If
            End Set
        End Property

        ''' <summary>
        ''' Gets of sets the flag that displays the remaining time.
        ''' </summary>
        ''' <remarks>Setting this property after the dialog is being shown has no effect.</remarks>
        ''' <value>False</value>
        Public Property HideRemainingTime() As Boolean
            Get
                Return CBool(__Flags And IPD_Flags.NoTime)
            End Get
            Set(ByVal value As Boolean)
                If (Not __DlgClosed) Then Exit Property
                If (value) Then
                    __Flags = __Flags Or IPD_Flags.NoTime
                Else
                    __Flags = __Flags And (Not IPD_Flags.NoTime)
                End If
            End Set
        End Property

        ''' <summary>
        ''' Gets or Sets the flag that displays the progress bar.
        ''' </summary>
        ''' <remarks>Setting this property after the dialog is being shown has no effect.</remarks>
        ''' <value>False</value>
        Public Property HideProgressBar() As Boolean
            Get
                Return CBool(__Flags And IPD_Flags.NoProgressbar)
            End Get
            Set(ByVal value As Boolean)
                If (Not __DlgClosed) Then Exit Property
                If (value) Then
                    __Flags = __Flags Or IPD_Flags.NoTime
                Else
                    __Flags = __Flags And (Not IPD_Flags.NoProgressbar)
                End If
            End Set
        End Property

        ''' <summary>
        ''' The Title to be displayes on the title bar of the dialog.
        ''' </summary>
        ''' <value>String.Empty</value>
        Public Property Title() As String
            Get
                Return __Title
            End Get
            Set(ByVal value As String)
                If (__Title = value) Then Exit Property
                __Title = value
                If (__Title.EndsWith(Chr(0))) Then
                    __Dialog.SetTitle(__Title)
                Else
                    __Dialog.SetTitle(__Title & Chr(0))
                End If
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the total size of the job.
        ''' </summary>
        Public Property Total() As ULong
            Get
                Return __Total
            End Get
            Set(ByVal value As ULong)
                __Total = value
                If (Not __DlgClosed) Then
                    __Dialog.SetProgress64(__Complete, __Total)
                End If
            End Set
        End Property

        ''' <summary>
        ''' A System.Boolean flag that indicates if the user has cancelled the process.
        ''' </summary>
        ''' <value>False</value>
        Public ReadOnly Property UserCancelled() As Boolean
            Get
                Return (__Dialog.HasUserCancelled)
            End Get
        End Property

        ''' <summary>
        ''' Sets the text to be displayed in one of the three available lines.
        ''' </summary>
        Public Sub SetLineText(ByVal Line As dlgLines, ByVal Text As String, Optional ByVal CompactPath As Boolean = True)
            If (CompactPath) Then
                Dim sb As New Text.StringBuilder(MAX_CAPACITY)
                PathCompactPathEx(sb, Text, MAX_CAPACITY, 0)
                Text = sb.ToString
            End If

            If (Text.EndsWith(Chr(0))) Then
                __Dialog.SetLine(Line, Text, 0, IntPtr.Zero)
            Else
                __Dialog.SetLine(Line, Text & Chr(0), 0, IntPtr.Zero)
            End If
        End Sub

        ''' <summary>
        ''' Display the dialog and resets the internal timer.
        ''' </summary>
        Public Sub Start()
            If (__DlgClosed) Then
                __UserCancelled = False
                __Dialog.StartProgressDialog(__hWndParent, Nothing, __Flags, IntPtr.Zero)
                __Dialog.Timer(PDTIMER_RESET, Nothing)
                __DlgClosed = False
            End If
        End Sub

        ''' <summary>
        ''' Closes the dialog.
        ''' </summary>
        Public Sub [Stop]()
            If (__DlgClosed) Then Exit Sub
            __Dialog.StopProgressDialog()
            __DlgClosed = True
        End Sub

        Protected Overrides Sub Finalize()
            If (Not __DlgClosed) Then
                Call [Stop]()
            End If
            __Dialog = Nothing
            MyBase.Finalize()
        End Sub

        ''' <summary>
        ''' Resets the internal timer.
        ''' </summary>
        Public Sub ResetTimer()
            __Dialog.Timer(PDTIMER_RESET, Nothing)
        End Sub
    End Class

End Namespace