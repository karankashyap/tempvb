Imports System.Net
Imports System.IO
Imports System.Collections.Specialized
Module HttpListener
    Public ReadOnly Property QueryString As NameValueCollection
    Sub Main()

        Dim prefixes(0) As String
        prefixes(0) = "http://" & Constant.LOCALHOST & Constant.PORT & Constant.HTTPLISTENER
        ProcessRequests(prefixes)
    End Sub

    Private Sub ProcessRequests(ByVal prefixes() As String)
        If Not System.Net.HttpListener.IsSupported Then
            Console.WriteLine(
                "This system cannot run RetailVI application.")
            Exit Sub
        End If

        ' URI prefixes are required,
        If prefixes Is Nothing OrElse prefixes.Length = 0 Then
            Throw New ArgumentException("prefixes")
        End If

        ' Create a listener and add the prefixes.
        Dim listener As System.Net.HttpListener =
            New System.Net.HttpListener()
        For Each s As String In prefixes
            listener.Prefixes.Add(s)
        Next

        Try
            ' Start the listener to begin listening for requests.
            listener.Start()
            Console.WriteLine("Listening...")

            ' Set the number of requests this application will handle.
            Dim numRequestsToBeHandled As Integer = 20

            For i As Integer = 0 To numRequestsToBeHandled
                Dim response As HttpListenerResponse = Nothing
                Try
                    ' Note: GetContext blocks while waiting for a request.
                    Dim context As HttpListenerContext = listener.GetContext()

                    ' Create the response.
                    response = context.Response
                    Dim responseString As String
                    responseString = "success=false&msg=Internal Server Error"


                    context.Response.StatusCode = 200
                    context.Response.KeepAlive = False

                        Dim outputO = context.Response.OutputStream
                        Dim uiy = 0

                        Dim body = New StreamReader(context.Request.InputStream).ReadToEnd()

                    Dim fn, ItemAlias, CenterCode, sessId, userId
                    Dim words As String() = body.Split(New Char() {"&"c})
                        Dim word As String
                        For Each word In words
                            Dim params As String() = word.Split(New Char() {"="c})

                            Try
                            If params(0) = "fn" Then
                                fn = params(1)
                            ElseIf params(0) = "alias" Then
                                ItemAlias = params(1)
                            ElseIf params(0) = "centerCode" Then
                                CenterCode = params(1)
                            ElseIf params(0) = "sessId" Then
                                sessId = params(1)
                            ElseIf params(0) = "userId" Then
                                userId = params(1)

                            End If
                            Catch
                            End Try
                        Next


                    If fn = "getMaterialCentres" Then
                        responseString = Form3.GetMaterialCentres("")
                    ElseIf fn = "addToCart" Then
                        responseString = Form3.AddItemToCart(ItemAlias, sessId, userId)
                    ElseIf fn = "addQuantity" Then
                        responseString = Form3.editItemQuantity(ItemAlias, "add", sessId)
                    ElseIf fn = "removeQuantity" Then
                        responseString = Form3.editItemQuantity(ItemAlias, "remove", sessId)
                    ElseIf fn = "checkout" Then
                        responseString = Form3.Checkout(sessId, userId)
                    ElseIf fn = "checkBilling" Then
                        responseString = Form3.CheckBilling(sessId, userId)
                    ElseIf fn = "deleteItem" Then
                        responseString = Form3.DeleteItem(ItemAlias, sessId)
                    Else

                        responseString = "NULL"

                        End If

                    Console.WriteLine(sessId)
                    Console.WriteLine(responseString)


                    Dim buffer() As Byte =
                        System.Text.Encoding.UTF8.GetBytes(responseString)
                    response.ContentLength64 = buffer.Length
                    Dim output As System.IO.Stream = response.OutputStream
                    output.Write(buffer, 0, buffer.Length)

                Catch ex As HttpListenerException
                    Console.WriteLine(ex.Message)
                Finally
                    If response IsNot Nothing Then
                        response.Close()
                    End If
                End Try
            Next
        Catch ex As HttpListenerException
            Console.WriteLine(ex.Message)
        Finally
            ' Stop listening for requests.
            listener.Close()
            Console.WriteLine("Done Listening...")
        End Try
    End Sub
End Module