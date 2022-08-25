'Alternative way

Option Explicit
Public Function GetCoordinates(address As String) As String
    'Declaring the necessary variables.
    Dim apiKey              As String
    Dim xmlhttpRequest      As Object
    Dim xmlDoc              As Object
    Dim xmlStatusNode       As Object
    Dim xmlLatitudeNode     As Object
    Dim xmLongitudeNode     As Object

    apiKey = "The API Key"

    If apiKey = vbNullString Or apiKey = "The API Key" Then
        GetCoordinates = "Empty or invalid API Key"
        Exit Function
            
    End If
    On Error GoTo errorHandler

    Set xmlhttpRequest = CreateObject("MSXML2.ServerXMLHTTP")
    If xmlhttpRequest Is Nothing Then
        GetCoordinates = "Cannot create the request object"
        Exit Function
    End If

    xmlhttpRequest.Open "GET", "https://maps.googleapis.com/maps/api/geocode/xml?" _
    & "&address=" & Application.EncodeURL(address) & "&key=" & apiKey, False
    xmlhttpRequest.send

    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    If xmlDoc Is Nothing Then
        GetCoordinates = "Cannot create the DOM document object"
        Exit Function
    End If
    'Read the XML results from the request.
    xmlDoc.LoadXML xmlhttpRequest.responseText
    'Get the value from the status node.
    Set xmlStatusNode = xmlDoc.SelectSingleNode("//status")
    'Based on the status node result, proceed accordingly.
    Select Case UCase(xmlStatusNode.Text)
        Case "OK"                       
            Set xmlLatitudeNode = xmlDoc.SelectSingleNode("//result/geometry/location/lat")
            Set xmLongitudeNode = xmlDoc.SelectSingleNode("//result/geometry/location/lng")
            'Return the coordinates as a string (latitude, longitude).
            GetCoordinates = xmlLatitudeNode.Text & ", " & xmLongitudeNode.Text
        Case "ZERO_RESULTS"             
            GetCoordinates = "The address probably do not exist"
        Case "OVER_DAILY_LIMIT"         

            GetCoordinates = "Billing or payment problem"
        Case "OVER_QUERY_LIMIT"         'The requestor has exceeded the quota limit.
            GetCoordinates = "Quota limit exceeded"
        Case "REQUEST_DENIED"           'The API did not complete the request.
            GetCoordinates = "Server denied the request"
        Case "INVALID_REQUEST"           'The API request is empty or is malformed.
            GetCoordinates = "Request was empty or malformed"
        Case "UNKNOWN_ERROR"            'The request could not be processed due to a server error.
            GetCoordinates = "Unknown error"
        Case Else   'Just in case...
            GetCoordinates = "Error"
    End Select
    'Release the objects before exiting (or in case of error).
errorHandler:
    Set xmlStatusNode = Nothing
    Set xmlLatitudeNode = Nothing
    Set xmLongitudeNode = Nothing
    Set xmlDoc = Nothing
    Set xmlhttpRequest = Nothing
End Function

Public Function GetLatitude(address As String) As Double
    'Declaring the necessary variable.
    Dim coordinates As String
    'Get the coordinates for the given address.
    coordinates = GetCoordinates(address)
    'Return the latitude as a number (double).
    If coordinates <> vbNullString Then
        GetLatitude = CDbl(Left(coordinates, WorksheetFunction.Find(",", coordinates) - 1))
    Else
        GetLatitude = 0
    End If
End Function
Public Function GetLongitude(address As String) As Double
    'Declaring the necessary variable.
    Dim coordinates As String
    'Get the coordinates for the given address.
    coordinates = GetCoordinates(address)
    'Return the longitude as a number (double).
    If coordinates <> vbNullString Then
        GetLongitude = CDbl(Right(coordinates, Len(coordinates) - WorksheetFunction.Find(",", coordinates)))
    Else
        GetLongitude = 0
    End If
End Function
