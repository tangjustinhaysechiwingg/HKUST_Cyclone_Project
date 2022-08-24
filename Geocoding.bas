Option Explicit
Public Function GetCoordinates(address As String) As String
    '-----------------------------------------------------------------------------------------------------
    'This function returns the latitude and longitude of a given address using the Google Geocoding API.
    'The function uses the "simplest" form of Google Geocoding API (sending only the address parameter),
    'so, optional parameters such as bounds, language, region and components are NOT used.
    'In case of multiple xmlDoc (for example two cities sharing the same name), the function
    'returns the FIRST OCCURRENCE, so be careful in the input address
    'Tip: use the city name and the postal code if they are available).
    'NOTE: As Google points out, the use of the Google Geocoding API is subject to a limit of 40,000
    'requests per month, so be careful not to exceed this limit. For more info check:
    'https://cloud.google.com/maps-platform/pricing/sheet
    '2018 Update: In order to use this function you will now need a valid API key.
    'Check the next link that guides you on how to acquire a free API key:
    'https://www.myengineeringworld.net/2018/02/how-to-get-free-google-api-key.html
    '2018 Update 2 (July): The EncodeURL function was added to avoid problems with special characters.
    'This is a common problem with addresses that are from Greece, Serbia, Germany and other countries.
    'Note that this function was introduced in Excel 2013, so it will NOT work in older versions.
    '2020 Update: The code was switched to late binding, so no external reference is required.
    'Written By:    Christos Samaras
    'Date:          12/06/2014
    'Last Updated:  16/02/2020
    'E-mail:        xristos.samaras@gmail.com
    'Site:          https://www.myengineeringworld.net
    '-----------------------------------------------------------------------------------------------------
    'Declaring the necessary variables.
    Dim apiKey              As String
    Dim xmlhttpRequest      As Object
    Dim xmlDoc              As Object
    Dim xmlStatusNode       As Object
    Dim xmlLatitudeNode     As Object
    Dim xmLongitudeNode     As Object
    'Set your API key in this variable. Check this link for more info:
    'https://www.myengineeringworld.net/2018/02/how-to-get-free-google-api-key.html
    'Here is the ONLY place in the code where you have to put your API key.
    apiKey = "The API Key"
    'Check that an API key has been provided.
    If apiKey = vbNullString Or apiKey = "The API Key" Then
        GetCoordinates = "Empty or invalid API Key"
        Exit Function
    End If
    'Generic error handling.
    On Error GoTo errorHandler
    'Create the request object and check if it was created successfully.
    Set xmlhttpRequest = CreateObject("MSXML2.ServerXMLHTTP")
    If xmlhttpRequest Is Nothing Then
        GetCoordinates = "Cannot create the request object"
        Exit Function
    End If
    'Create the request based on Google Geocoding API. Parameters (from Google page):
    '- Address: The address that you want to geocode.
    'Note: The EncodeURL function was added to allow users from Greece, Poland, Germany, France and other countries
    'geocode address from their home countries without a problem. The particular function (EncodeURL),
    'returns a URL-encoded string without the special characters.
    'This function, however, was introduced in Excel 2013, so it will NOT work in older Excel versions.
    xmlhttpRequest.Open "GET", "https://maps.googleapis.com/maps/api/geocode/xml?" _
    & "&address=" & Application.EncodeURL(address) & "&key=" & apiKey, False
    'An alternative way, without the EncodeURL function, will be this:
    'xmlhttpRequest.Open "GET", "https://maps.googleapis.com/maps/api/geocode/xml?" & "&address=" & Address & "&key=" & ApiKey, False
    'Send the request to the Google server.
    xmlhttpRequest.send
    'Create the DOM document object and check if it was created successfully.
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
        Case "OK"                       'The API request was successful.
                                        'At least one result was returned.
            'Get the latitude and longitude node values of the first result.
            Set xmlLatitudeNode = xmlDoc.SelectSingleNode("//result/geometry/location/lat")
            Set xmLongitudeNode = xmlDoc.SelectSingleNode("//result/geometry/location/lng")
            'Return the coordinates as a string (latitude, longitude).
            GetCoordinates = xmlLatitudeNode.Text & ", " & xmLongitudeNode.Text
        Case "ZERO_RESULTS"             'The geocode was successful but returned no results.
            GetCoordinates = "The address probably do not exist"
        Case "OVER_DAILY_LIMIT"         'Indicates any of the following:
                                        '- The API key is missing or invalid.
                                        '- Billing has not been enabled on your account.
                                        '- A self-imposed usage cap has been exceeded.
                                        '- The provided method of payment is no longer valid
                                        '  (for example, a credit card has expired).
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
'------------------------------------------------------------------------------------------------------------------
'The next two functions use the GetCoordinates function to get the latitude and the longitude of a given address.
'------------------------------------------------------------------------------------------------------------------
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
