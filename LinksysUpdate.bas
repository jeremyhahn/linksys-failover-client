Attribute VB_Name = "LinksysUpdate"
Public Function UpdateRouter(RouterIP, GetRequest, RouterPassword) As String


   Dim b64 As Base64Class
   Dim EncodedText As String
   
    Set b64 = New Base64Class


       Dim strBuffer As String
       Dim cchBuffer As Integer
       Dim strHeader As String
 
    ' TO DO WITH MODULE -- BIND SOCKET TO PROPER FORM WITH THE SOCKET CONTROL   < ~~~~~~
   
    
    frmService.Socket1.AutoResolve = False
    frmService.Socket1.Blocking = True
    frmService.Socket1.HostName = RouterIP
    frmService.Socket1.RemotePort = 80
    frmService.Socket1.Timeout = 600


    If frmService.Socket1.Connect() <> 0 Then
     
     Write2Log ("Unable to connect to " & RouterIP)
     Exit Function
   
End If



 EncodedText = b64.EncodeString(":" & RouterPassword)
       
    strHeader = "GET " & GetRequest & " HTTP/1.0" & Chr(13) & Chr(10) & _
                "User-Agent: Linksys_Failover_Client Build-1.1" & _
                "Authorization: Basic " & EncodedText & " & Chr(13) & Chr(10) & Chr(13) & Chr(10)"



If frmService.Socket1.Write(strHeader, 2048) Then

'Do Nothing

Else


Write2Log ("An error occurred at the protocol level while sending an update to the router. " & vbCrLf & _
           "The request could not be sent.")

Exit Function
End If



   Open App.Path & "/RouterResponse.dat" For Output As #1
          
    
    Do

           cchBuffer = frmService.Socket1.Read(strBuffer, 120)
        If cchBuffer = 0 Then
            
            ' The server has closed the connection and we have
            ' reached the end of the data stream
    
            Exit Do
        ElseIf cchBuffer = -1 Then
    
            ' An error has occurred while reading data from the
            ' server; this should be considered a fatal error
            
            frmService.Socket1.Disconnect
            MsgBox "An error occurred while reading the response from the HTTP server."
      
            Exit Function
    End If



            Write #1, strBuffer
    Loop


            Close #1







 Open App.Path & "/RouterResponse.dat" For Input As #2
 
      Input #2, LogFile
 
 Close #2
 
 



ResponseHEAD = Split(LogFile, " ")
  StatusCode = ResponseHEAD(1)




Select Case StatusCode
  
  
  
  'Successful
  Case "200"
  
    StatusDesc = "Update Successful"
    
  Case "201"

    StatusDesc = "Created"
    
  Case "202"
  
    StatusDesc = "Accepted"

  Case "204"
  
    StatusDesc = "No Content"
    
    
    
    
    
    
  'Redirection
  Case "300"
  
    StatusDesc = "Multiple Choices"
    
  Case "301"
  
    StatusDesc = "Moved Permanently"
    
  Case "302"
  
    StatusDesc = "Moved Temporarily"
    
  Case "303"
  
    StatusDesc = "Successful Connection"
    
  Case "304"
  
    StatusDesc = "Not Modified"
    
    
    
    
  'Client Request Errors
  Case "400"
  
    StatusDesc = "Bad Request"
    
  Case "401"
  
    StatusDesc = "Unauthorized Request"
    
  Case "403"
  
    StatusDesc = "Forbidden"
    
  Case "404"
  
    StatusDesc = "Not Found"
    
    
    
    
  'Server Errors
  Case "500"
  
    StatusDesc = "Internal Server Error"
    
  Case "501"
  
    StatusDesc = "Not Implemented"
    
  Case "502"
  
    StatusDesc = "Bad Gateway"
    
  Case "503"
  
    StatusDesc = "Service Unavailable"
    
    
    
    
 End Select
 
 
 
 
  

UpdateRouter = StatusCode & ",The router says: " & StatusDesc


frmService.Socket1.Disconnect
Set b64 = Nothing

End Function
