Public Class Form1
    Dim connString As String
    Dim filePath As String

    'Dim oFSO As New Scripting.FileSystemObject
    'Dim oTS As Scripting.TextStream

    Dim fcon As Boolean
    Dim serverName As String
    Dim coName As String
    Dim versionName As String
    Dim currentVersion As Integer
    Dim qbposXMLRP As QBPOSXMLRPLib.RequestProcessor

    Private Sub Form1_Load()
        ' Step-1: When the form is loaded, reading from
        ' a text file (or you could use registry) to get the last stored
        ' connection string
        connString = ""
        filePath = "cstr.txt"
        'oFSO = New Scripting.FileSystemObject
        'connString = FileRead
        'LabelConnString.Caption = connString
        'qbposXMLRP = New RequestProcessor
    End Sub


    Private Sub btnFindServer_Click(sender As Object, e As EventArgs) Handles btnFindServer.Click
        qbposXMLRP = getRequestProcessor()
        'cmdStart.Caption = "Searching. Please wait..."
        getServers(qbposXMLRP)
        'cmdStart.Caption = "Done"
        'cmdServer.Enabled = True
    End Sub

    Public Function getRequestProcessor() As QBPOSXMLRPLib.RequestProcessor
        getRequestProcessor = qbposXMLRP
    End Function


    Private Sub getServers(qbposXMLRP As QBPOSXMLRPLib.RequestProcessor)
        Dim posserversXML As String

        'Method signature for POSServers():
        'HRESULT POSServers([in] VARIANT_BOOL IsPractice, [out,retval] BSTR* ServersXML);
        '
        'Returns data about the available QBPOS instances, companies, and versions in the local network.
        '
        'Where,
        'IsPractice
        'Specify True if the company is the QBPOS practice company. This can decrease the lookup time by restricting the search to the test company. Useful for development. For real deployments, however, this value would be set to False to make sure all companies were listed.
        '
        'ServersXML
        'Pointer to the returned string containing an XML document containing the list of available servers.
        '
        'An example of what is returned as ServersXML from POSServers()
        '
        '<POSServers>
        '<POSServer>
        '    <ServerName>mtvd040202111</ServerName>
        '    <CompanyName>al&apos;s sports hut</CompanyName>
        '    <Version>4</Version>
        '</POSServer>
        '<POSServer>
        '    <ServerName>mtvl040202118</ServerName>
        '    <CompanyName>al&apos;s sports hut</CompanyName>
        '    <Version>4</Version>
        '</POSServer>
        '...
        '</POSServers>

        posserversXML = qbposXMLRP.POSServers(False)
        Dim xmlDoc As MSXML2.DOMDocument40
        Dim objNodeList As MSXML2.IXMLDOMNodeList
        Dim objChild As MSXML2.IXMLDOMNode
        Dim childNode As MSXML2.IXMLDOMNode

        Dim i As Integer
        Dim ret As Boolean
        Dim errorMsg As String

        errorMsg = ""
        xmlDoc = New MSXML2.DOMDocument40

        ret = xmlDoc.loadXML(posserversXML)
        If Not ret Then
            errorMsg = "loadXML failed, reason: " & xmlDoc.parseError.reason
            GoTo ErrHandler
        End If

        Dim server As String
        objNodeList = xmlDoc.getElementsByTagName("POSServer")
        For i = 0 To (objNodeList.length - 1)
            For Each objChild In objNodeList.Item(i).childNodes
                If objChild.nodeName = "ServerName" Then
                    server = objChild.Text
                End If
                If objChild.nodeName = "CompanyName" Then
                    server = server + " - " + objChild.Text
                End If
                If objChild.nodeName = "Version" Then
                    server = server + " - " + objChild.Text
                End If
                'ComboServer.List(i) = server
            Next
        Next
        Exit Sub

ErrHandler:
        MsgBox(Err.Description, vbExclamation, "Error")
        Exit Sub
    End Sub


End Class
