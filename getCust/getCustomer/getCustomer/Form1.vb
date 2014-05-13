
Imports QBPOSFC3Lib
'Imports QBFC12Lib

Public Class Form1
    Dim sessionBegun As Boolean = False
    Dim connectionOpen As Boolean = False
    Dim sessionManager As QBPOSSessionManager = Nothing

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            'Create the session Manager object
            sessionManager = New QBPOSSessionManager

            'Connect ot QB_POS and begin session
            sessionManager.OpenConnection("mapel", "test")
            sessionManager.BeginSession("")
            connectionOpen = True
            sessionBegun = True

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
            If (sessionBegun) Then
                sessionManager.EndSession()
            End If
            If (connectionOpen) Then
                sessionManager.CloseConnection()
            End If
        End Try
    End Sub

    Private Sub btnCustomers_Click(sender As Object, e As EventArgs) Handles btnCustomers.Click
        Try
            'Create the message set request object to hold the request
            Dim requestMsgSet As IMsgSetRequest
            requestMsgSet = sessionManager.CreateMsgSetRequest(3, 0)
            requestMsgSet.Attributes.OnError = QBPOSFC3Lib.ENRqOnError.roeContinue
            BuildCustomerQueryRq(requestMsgSet)

            'Send the request and get the response from QuickBooks
            Dim responseMsgSet As IMsgSetResponse
            responseMsgSet = sessionManager.DoRequests(requestMsgSet)

            WalkCustomerQueryRs(responseMsgSet)

        Catch er As Exception
            MessageBox.Show(er.Message, "Error")
            If (sessionBegun) Then
                sessionManager.EndSession()
            End If
            If (connectionOpen) Then
                sessionManager.CloseConnection()
            End If
        End Try
    End Sub
    Public Sub BuildCustomerQueryRq(requestMsgSet As IMsgSetRequest)
        Dim CustomerQueryRq As ICustomerQuery
        Dim dtFrom As DateTime = New DateTime(2014, 4, 26)
        Dim dtTo As DateTime = DateTime.Now

        CustomerQueryRq = requestMsgSet.AppendCustomerQueryRq()
        'CustomerQueryRq.MaxReturned.SetValue(6)

        CustomerQueryRq.OwnerIDList.Add(0)


        CustomerQueryRq.ORTimeCreatedFilters.TimeCreatedRangeFilter.FromTimeCreated.SetValue(dtFrom, False)
        CustomerQueryRq.ORTimeCreatedFilters.TimeCreatedRangeFilter.ToTimeCreated.SetValue(dtTo, False)

        'CustomerQueryRq.ORFirstNameFilters.FirstNameRangeFilter.FromFirstName.SetValue("Peter")
        'CustomerQueryRq.ORFirstNameFilters.FirstNameRangeFilter.ToFirstName.SetValue("Peter")

    End Sub
    Public Sub WalkCustomerQueryRs(responseMsgSet As IMsgSetResponse)
        If (responseMsgSet Is Nothing) Then
            Exit Sub
        End If
        Dim strReturn As String = responseMsgSet.ToXMLString()
        Dim responseList As IResponseList
        'MsgBox(strReturn)
        responseList = responseMsgSet.ResponseList
        If (responseList Is Nothing) Then
            Exit Sub
        End If

        Dim j As Integer = 0
        'if we sent only one request, there is only one response, we'll walk the list for this sample
        For j = 0 To responseList.Count - 1
            Dim response As IResponse
            response = responseList.GetAt(j)
            'check the status code of the response, 0=ok, >0 is warning
            If (response.StatusCode >= 0) Then
                'the request-specific response is in the details, make sure we have some
                If (Not response.Detail Is Nothing) Then
                    'make sure the response is the type we're expecting
                    Dim responseType As ENResponseType
                    responseType = CType(response.Type.GetValue(), ENResponseType)
                    If (responseType = ENResponseType.rtCustomerQueryRs) Then
                        '//upcast to more specific type here, this is safe because we checked with response.Type check above

                        Dim CustomerRet As ICustomerRetList = CType(response.Detail, ICustomerRetList)
                        WalkCustomerRet(CustomerRet)
                    End If
                End If
            End If
        Next j
    End Sub
    Public Sub WalkCustomerRet(CustomerRet As ICustomerRetList)
        If (CustomerRet Is Nothing) Then
            Exit Sub
        End If

        Dim intItems As Integer = CustomerRet.Count
        Dim StrFullNames As String = ""
        For i = 0 To CustomerRet.Count - 1
            'Dim ListID481 As String = CustomerRet.GetAt(0).ListID.GetValue()
            'ListID481 = CustomerRet.ListID.GetValue
            Dim strFirstName As String
            Dim strLastName As String
            'Go through all the elements of ICustomerRet
            If (Not CustomerRet.GetAt(i).FirstName Is Nothing) Then
                strFirstName = CustomerRet.GetAt(i).FirstName.GetValue()
            Else
                strFirstName = " "
            End If
            If (Not CustomerRet.GetAt(i).LastName Is Nothing) Then
                strLastName = CustomerRet.GetAt(i).LastName.GetValue()
            Else
                strLastName = " "
            End If

            StrFullNames = StrFullNames & Environment.NewLine & strFirstName & " " & strLastName
        Next i
        Label1.Text = StrFullNames
    End Sub


    Private Sub btnCompany_Click(sender As Object, e As EventArgs) Handles btnCompany.Click
        Try
            'Create the message set request object to hold the request
            Dim requestMsgSet As IMsgSetRequest
            requestMsgSet = sessionManager.CreateMsgSetRequest(3, 0)
            requestMsgSet.Attributes.OnError = QBPOSFC3Lib.ENRqOnError.roeContinue
            BuildCompanyQueryRq(requestMsgSet)

            'Send the request and get the response from QuickBooks
            Dim responseMsgSet As IMsgSetResponse
            responseMsgSet = sessionManager.DoRequests(requestMsgSet)

            WalkCompanyQueryRs(responseMsgSet)

        Catch er As Exception
            MessageBox.Show(er.Message, "Error")
            If (sessionBegun) Then
                sessionManager.EndSession()
            End If
            If (connectionOpen) Then
                sessionManager.CloseConnection()
            End If
        End Try
    End Sub
    Public Sub BuildCompanyQueryRq(requestMsgSet As IMsgSetRequest)
        Dim CompanyQueryRq As ICompanyQuery
        CompanyQueryRq = requestMsgSet.AppendCompanyQueryRq()

        CompanyQueryRq.OwnerIDList.Add(0)
    End Sub
    Private Sub WalkCompanyQueryRs(responseMsgSet As IMsgSetResponse)
        If (responseMsgSet Is Nothing) Then
            Exit Sub
        End If

        Dim responseList As IResponseList
        responseList = responseMsgSet.ResponseList
        If (responseList Is Nothing) Then
            Exit Sub
        End If

        'if we sent only one request, there is only one response, we'll walk the list for this sample
        For j = 0 To responseList.Count - 1
            Dim response As IResponse
            response = responseList.GetAt(j)
            'check the status code of the response, 0=ok, >0 is warning
            If (response.StatusCode >= 0) Then
                'the request-specific response is in the details, make sure we have some
                If (Not response.Detail Is Nothing) Then
                    'make sure the response is the type we're expecting
                    Dim responseType As ENResponseType
                    responseType = CType(response.Type.GetValue(), ENResponseType)
                    If (responseType = ENResponseType.rtCompanyQueryRs) Then
                        '//upcast to more specific type here, this is safe because we checked with response.Type check above
                        Dim CompanyRet As ICompanyRet
                        CompanyRet = CType(response.Detail, ICompanyRet)
                        WalkCompanyRet(CompanyRet)
                    End If
                End If
            End If
        Next j
    End Sub
    Public Sub WalkCompanyRet(CompanyRet As ICompanyRet)
        If (CompanyRet Is Nothing) Then
            Exit Sub
        End If

        'Go through all the elements of ICompanyRet
        'Get value of IsSampleCompany
        Dim IsSampleCompany82 As Boolean
        IsSampleCompany82 = CompanyRet.IsSampleCompany.GetValue()
        Dim CompanyName83 As String = CompanyRet.CompanyName.GetValue()
        Dim Street84 As String = CompanyRet.Address.Street.GetValue()
        Dim CityStateZIP85 As String = CompanyRet.Address.CityStateZIP.GetValue()
        Dim Misc186 As String = CompanyRet.Address.Misc1.GetValue()
        Dim StoreNumber90 As Integer = CompanyRet.StoreNumber.GetValue()

        'Get value of QuickBooksCompanyFile
        Dim QuickBooksCompanyFile89 As String = ""
        If (Not CompanyRet.QuickBooksCompanyFile Is Nothing) Then
            QuickBooksCompanyFile89 = CompanyRet.QuickBooksCompanyFile.GetValue()
        End If


        Dim fullAdd = CompanyName83 & Environment.NewLine & Street84 & Environment.NewLine & CityStateZIP85
        fullAdd = fullAdd & Environment.NewLine & Misc186 & Environment.NewLine & QuickBooksCompanyFile89
        fullAdd = fullAdd & Environment.NewLine & StoreNumber90

        Label1.Text = fullAdd

    End Sub


    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'End the session and close the connection to QuickBooks
        sessionManager.EndSession()
        sessionBegun = False
        sessionManager.CloseConnection()
        connectionOpen = False

        Me.Close()
    End Sub

End Class
