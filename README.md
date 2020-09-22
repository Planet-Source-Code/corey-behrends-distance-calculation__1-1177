<div align="center">

## Distance Calculation


</div>

### Description

This code shows how to use Visual Basic to calculate great circle distance (distance between 2 points using decimal latitudes and longitudes).
 
### More Info
 
see code

1. This code does not figure in differences in altiude

2. In order to use this code you must have the latitude and longitude in decmal form.

Returns the distance in the desired units between 2 points


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Corey Behrends](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/corey-behrends.md)
**Level**          |Unknown
**User Rating**    |4.8 (43 globes from 9 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/corey-behrends-distance-calculation__1-1177/archive/master.zip)

### API Declarations

NONE


### Source Code

```
Function LatLonDistance(ByVal dbLat1 As Double, _
             ByVal dbLon1 As Double, _
             ByVal dbLat2 As Double, _
             ByVal dbLon2 As Double, _
             ByVal stUnits As String) As Double
Dim loRadiusOfEarth As Long
Dim dbDeltaLat As Double
Dim dbDeltaLon As Double
Dim dbTemp As Double
Dim dbTemp2 As Double
  'Set the radius of the earth in the selected units
  Select Case UCase(stUnits)
    Case "MI" ' Miles
      loRadiusOfEarth = 3956
    Case "FT" ' Feet
      loRadiusOfEarth = 20887680
    Case "YD" ' Yards
      loRadiusOfEarth = 6962560
    Case "KM" ' Kilometers
      loRadiusOfEarth = 6367
    Case "M" ' Meters
      loRadiusOfEarth = 6367000
    Case Else ' Error
      LatLonDistance = -1
      Exit Function
  End Select
  'Calculate the Delta of the of the Longitudes and Latitudes and
  'subtract the destination point from the starting point
  dbDeltaLon = AsRadians(dbLon2) - AsRadians(dbLon1)
  dbDeltaLat = AsRadians(dbLat2) - AsRadians(dbLat1)
  'Intermediate values...
  dbTemp = Sin2(dbDeltaLat / 2) + _
    Cos(AsRadians(dbLat1)) * _
    Cos(AsRadians(dbLat2)) * _
    Sin2(dbDeltaLon / 2)
  'The temp value dbTemp2 is the great circle distance in radians
  dbTemp2 = 2 * Arcsin(GetMin(1, Sqr(dbTemp)))
  'Multiply the radians by the radius to get the distance in specified units
  LatLonDistance = loRadiusOfEarth * dbTemp2
End Function
Private Function Arcsin(ByVal X As Double) As Double
   Arcsin = Atn(X / Sqr(-X * X + 1))
End Function
Private Function AsRadians(ByVal pDb_Degrees As Double) As Double
Const vbPi = 3.14159265358979
  'To convert decimal degrees to radians, multiply
  'the number of degrees by pi/180 = 0.017453293 radians/degree
  AsRadians = pDb_Degrees * (vbPi / 180)
End Function
Private Function GetMin(ByVal X As Double, ByVal Y As Double) As Double
  If X <= Y Then
    GetMin = X
  Else
    GetMin = Y
  End If
End Function
Private Function Sin2(ByVal X As Double) As Double
   Sin2 = (1 - Cos(2 * X)) / 2
End Function
Function RoundNum(Num As Double) As Double
'This function rounds a floating point number to nearest whole
'number, a function which is sadly lacking from VB.
  If Int(Num + 0.5) > Num Then
    RoundNum = Int(Num + 0.5)
  Else
    RoundNum = Int(Num)
  End If
End Function
```

