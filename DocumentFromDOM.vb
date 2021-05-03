Imports System.Runtime.InteropServices
Imports System.Text
Imports mshtml

Public Class DocumentFromDOM


    <DllImport("user32.dll", EntryPoint:="GetClassNameA")>
    Public Shared Function GetClassName(ByVal hwnd As IntPtr, ByVal lpClassName As StringBuilder, ByVal nMaxCount As Integer) As Integer

    End Function
    Public Delegate Function EnumProc(ByVal hWnd As IntPtr, ByRef lParam As IntPtr) As Integer
    <DllImport("user32.dll")>
    Public Shared Function EnumChildWindows(ByVal hWndParent As IntPtr, ByVal lpEnumFunc As EnumProc, ByRef lParam As IntPtr) As Integer

    End Function
    <DllImport("user32.dll", EntryPoint:="RegisterWindowMessageA")>
    Public Shared Function RegisterWindowMessage(ByVal lpString As String) As Integer

    End Function
    <DllImport("user32.dll", EntryPoint:="SendMessageTimeoutA")>
    Public Shared Function SendMessageTimeout(ByVal hwnd As IntPtr, ByVal msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer, ByVal fuFlags As Integer, ByVal uTimeout As Integer, <Out> ByRef lpdwResult As Integer) As Integer

    End Function
    <DllImport("OLEACC.dll")>
    Public Shared Function ObjectFromLresult(ByVal lResult As Integer, ByRef riid As Guid, ByVal wParam As Integer, ByRef ppvObject As IHTMLDocument2) As Integer

    End Function

    Public Const SMTO_ABORTIFHUNG As Integer = &H2
    Public Shared IID_IHTMLDocument As Guid = New Guid("626FC520-A41E-11CF-A731-00A0C9082637")
    Public Shared document As IHTMLDocument2

    Public Shared Function EnumWindows(ByVal hWnd As IntPtr, ByRef lParam As IntPtr) As Integer
        Dim retVal As Integer = 1
        Try
            Dim classname As StringBuilder = New StringBuilder(128)
            GetClassName(hWnd, classname, classname.Capacity)

            If CBool((String.Compare(classname.ToString(), "Internet Explorer_Server") = 0)) Then
                lParam = hWnd
                retVal = 0
            End If
        Catch ex As Exception

        End Try
        Return retVal
    End Function


    Public Shared Function documentFromDOM(HwndUnico As IntPtr) As IHTMLDocument2
        Try

            If HwndUnico <> IntPtr.Zero Then
                Dim hWnd As IntPtr = HwndUnico
                Dim lngMsg As Integer = 0
                Dim lRes As Integer
                Dim proc As EnumProc = New EnumProc(AddressOf EnumWindows)
                EnumChildWindows(hWnd, proc, hWnd)

                If Not hWnd.Equals(IntPtr.Zero) Then
                    lngMsg = RegisterWindowMessage("WM_HTML_GETOBJECT")

                    If lngMsg <> 0 Then
                        SendMessageTimeout(hWnd, lngMsg, 0, 0, SMTO_ABORTIFHUNG, 1000, lRes)

                        If Not CBool((lRes = 0)) Then
                            Dim hr As Integer = ObjectFromLresult(lRes, IID_IHTMLDocument, 0, document)

                            If CBool((document Is Nothing)) Then
                                Return Nothing
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
        Return document
    End Function

    Shared Function AguardaInfoDOMDocument(Elemento As String, Hwnd As IntPtr, tempoSegundos As Integer) As Integer
        Dim retorno = 0

        Dim stopwatch As New Stopwatch
        stopwatch.Start()

        While stopwatch.Elapsed.Seconds < tempoSegundos
            Try

                Dim HTMLDom = documentFromDOM(Hwnd)
                Dim HTMLe = HTMLDom.getElementById(Elemento)
                If HTMLe IsNot Nothing Then
                    retorno = 1
                    Exit While
                End If

            Catch ex As Exception

            End Try
        End While

        Return retorno
    End Function

    Shared Function AguardaEnabledDOMDocument(Elemento As String, Hwnd As IntPtr, tempoSegundos As Integer) As Integer
        Dim retorno = 0

        Dim stopwatch As New Stopwatch
        stopwatch.Start()

        While stopwatch.Elapsed.Seconds < tempoSegundos
            Try

                Dim HTMLDom = documentFromDOM(Hwnd)
                Dim HTMLe = HTMLDom.getElementById(Elemento)
                If HTMLe IsNot Nothing Then
                    If HTMLe.outerHtml.contains("button disabled") = False Then
                        retorno = 1
                        Exit While
                    End If
                End If

            Catch ex As Exception

            End Try
        End While

        Return retorno
    End Function

    Shared Function AguardaCarregamentoBrowserVivoNext(Hwnd As IntPtr, tempoSegundos As Integer) As Integer
        Dim retorno = 0

        Dim stopwatch As New Stopwatch
        stopwatch.Start()

        While stopwatch.Elapsed.Seconds < tempoSegundos
            Try

                Dim HTMLDom = documentFromDOM(Hwnd)
                If HTMLDom IsNot Nothing Then
                    If HTMLDom.body.outerHTML.Contains("ux-loader__item") = False Then
                        retorno = 1
                        Exit While
                    End If
                End If

            Catch ex As Exception

            End Try
        End While

        Return retorno
    End Function

    Shared Function AguardaTextoHtmlDOMDocument(Texto As String, Hwnd As IntPtr, tempoSegundos As Integer) As Integer
        Dim retorno = 0

        Dim stopwatch As New Stopwatch
        stopwatch.Start()

        While stopwatch.Elapsed.Seconds < tempoSegundos
            Try

                Dim HTMLDom = documentFromDOM(Hwnd)
                Dim HTMLe = HTMLDom.body.innerHTML
                If HTMLe.Contains(Texto) = True Then
                    retorno = 1
                    Exit While
                End If

            Catch ex As Exception

            End Try
        End While

        Return retorno
    End Function

    Shared Function AguardaInfoClasseDOMDocument(Elemento As String, Hwnd As IntPtr, tempoSegundos As Integer) As Integer
        Dim retorno = 0

        Dim stopwatch As New Stopwatch
        stopwatch.Start()

        While stopwatch.Elapsed.Seconds < tempoSegundos
            Try

                Dim HTMLDom = documentFromDOM(Hwnd)
                Dim HTMLe = HTMLDom.GetElementsByClassName(Elemento)
                If HTMLe IsNot Nothing Then
                    retorno = 1
                    Exit While
                End If

            Catch ex As Exception

            End Try
        End While

        Return retorno
    End Function


    Shared Function AguardaDOMDocumentLoad(Hwnd As IntPtr, tempoSegundos As Integer) As Integer
        Dim retorno = 0

        Dim stopwatch As New Stopwatch
        stopwatch.Start()

        While stopwatch.Elapsed.Seconds < tempoSegundos
            Try

                Dim HTMLDom = documentFromDOM(Hwnd)
                If HTMLDom IsNot Nothing Then
                    retorno = 1
                    Exit While
                End If

            Catch ex As Exception

            End Try
        End While

        Return retorno
    End Function


End Class
