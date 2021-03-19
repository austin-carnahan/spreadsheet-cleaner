Attribute VB_Name = "CustomGeocodingFeatures"
Option Explicit

Function AddressToLatLong(Address As String) As String

    Dim Request         As New XMLHTTP60
    Dim Results         As New DOMDocument60
    Dim Json            As Object
    Dim Latitude        As String
    Dim Longitude       As String

    Request.Open "GET", "https://geocoding.geo.census.gov/geocoder/locations/onelineaddress?address=" _
    & Application.EncodeURL(Address) & "&benchmark=9&format=json", False

    Request.send

    Set Json = JsonConverter.ParseJson(Request.responseText)

    Latitude = Json("result")("addressMatches")(1)("coordinates")("y")
    Longitude = Json("result")("addressMatches")(1)("coordinates")("x")

    AddressToLatLong = Latitude & ", " & Longitude

End Function

Function NorthingEastingToLatLong(Northing As Double, Easting As Double) As String

    Dim Request         As New XMLHTTP60
    Dim Results         As New DOMDocument60
    Dim Json            As Object

    Request.Open "GET", "https://geodesy.noaa.gov/api/ncat/spc?" & _
            "spcZone=3702" & _
            "&inDatum=nad83(2011)" & _
            "&outDatum=nad83(2011)" & _
            "&northing=" & CStr(Northing) & _
            "&easting=" & CStr(Easting) & _
            "&units=usft", False

    Request.send

    Set Json = JsonConverter.ParseJson(Request.responseText)

    NorthingEastingToLatLong = Json("destLat") & ", " & Json("destLon")

End Function

Function GetLatFromLatLongString(Coordinates As String) As Double

    'Return the latitude as a number (double).
    If Coordinates <> vbNullString Then
        GetLatFromLatLongString = CDbl(Left(Coordinates, WorksheetFunction.Find(",", Coordinates) - 1))
    Else
        GetLatFromLatLongString = 0
    End If

End Function

Function GetLonFromLatLongString(Coordinates As String) As Double

    'Return the longitude as a number (double).
    If Coordinates <> vbNullString Then
        GetLonFromLatLongString = CDbl(Right(Coordinates, Len(Coordinates) - WorksheetFunction.Find(",", Coordinates)))
    Else
        GetLonFromLatLongString = 0
    End If

End Function

'-------------------------------------------------------------------------------------------------------------------
'The next two functions using the AddressToLatLong function to get the latitude and the longitude of a given address.
'-------------------------------------------------------------------------------------------------------------------

Function AddressToLat(Address As String) As Double

    'Declaring the necessary variable.
    Dim Coordinates As String

    'Get the coordinates for the given address.
    Coordinates = AddressToLatLong(Address)

    'Return the latitude as a number (double).
    If Coordinates <> vbNullString Then
        AddressToLat = CDbl(Left(Coordinates, WorksheetFunction.Find(",", Coordinates) - 1))
    Else
        AddressToLat = 0
    End If

End Function

Function AddressToLong(Address As String) As Double

    'Declaring the necessary variable.
    Dim Coordinates As String

    'Get the coordinates for the given address.
    Coordinates = AddressToLatLong(Address)

    'Return the longitude as a number (double).
    If Coordinates <> vbNullString Then
        AddressToLong = CDbl(Right(Coordinates, Len(Coordinates) - WorksheetFunction.Find(",", Coordinates)))
    Else
        AddressToLong = 0
    End If

End Function

'-------------------------------------------------------------------------------------------------------------------
'The next two functions using the NorthingEastingToLatLong function to get the latitude and the longitude of a given address.
'-------------------------------------------------------------------------------------------------------------------

Function NorthingEastingToLat(Northing As Double, Easting As Double) As Double

    'Declaring the necessary variable.
    Dim Coordinates As String

    'Get the coordinates for the given address.
    Coordinates = NorthingEastingToLatLong(Northing, Easting)

    'Return the latitude as a number (double).
    If Coordinates <> vbNullString Then
        NorthingEastingToLat = CDbl(Left(Coordinates, WorksheetFunction.Find(",", Coordinates) - 1))
    Else
        NorthingEastingToLat = 0
    End If

End Function

Function NorthingEastingToLong(Northing As Double, Easting As Double) As Double

    'Declaring the necessary variable.
    Dim Coordinates As String

    'Get the coordinates for the given address.
    Coordinates = NorthingEastingToLatLong(Northing, Easting)

    'Return the longitude as a number (double).
    If Coordinates <> vbNullString Then
        NorthingEastingToLong = CDbl(Right(Coordinates, Len(Coordinates) - WorksheetFunction.Find(",", Coordinates)))
    Else
        NorthingEastingToLong = 0
    End If

End Function