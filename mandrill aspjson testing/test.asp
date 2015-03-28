<!--#include file="aspjson.asp" -->
<%
    jsonstring = ""

    If Request.TotalBytes > 0 Then 
        Dim lngBytesCount 
        lngBytesCount = Request.TotalBytes 
        jsonstring = BytesToStr(Request.BinaryRead(lngBytesCount)) 
    End If 

    Set oJSON = New aspJSON
    'Load JSON string
    jsonstring = replace(jsonstring, "mandrill_events=","")
    oJSON.loadJSON(URLDecode(jsonstring))

' Loop Through Records
for i = 0 to oJSON.data.count -1
    str = oJSON.data(i).item("event")
    straddress = oJSON.data(i).item("msg").item("email")
next
%>

<%
    Function URLDecode(sConvert)
    Dim aSplit
    Dim sOutput
    Dim I
    If IsNull(sConvert) Then
       URLDecode = ""
       Exit Function
    End If

    ' convert all pluses to spaces
    sOutput = REPLACE(sConvert, "+", " ")

    ' next convert %hexdigits to the character
    aSplit = Split(sOutput, "%")

    If IsArray(aSplit) Then
      sOutput = aSplit(0)
      For I = 0 to UBound(aSplit) - 1
        sOutput = sOutput & _
          Chr("&H" & Left(aSplit(i + 1), 2)) &_
          Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
      Next
    End If

    URLDecode = sOutput
End Function
%>

<%
Function BytesToStr(bytes) 
    Dim Stream 
    Set Stream = Server.CreateObject("Adodb.Stream") 
    Stream.Type = 1 
    'adTypeBinary 
    Stream.Open 
    Stream.Write bytes 
    Stream.Position = 0 
    Stream.Type = 2 'adTypeText 
    Stream.Charset = "iso-8859-1" 
    BytesToStr = Stream.ReadText 
    Stream.Close 
    Set Stream = Nothing 
End Function 

%>