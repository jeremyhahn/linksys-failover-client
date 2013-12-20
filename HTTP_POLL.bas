Attribute VB_Name = "HTTP_POLL"
Dim strBuffer As String
Dim cchBuffer As Integer
Dim strHeader As String







Public Sub DO_HTTP_POLL()


VERIFY_SYSCONFIG
Open App.Path & "/sysconfig.dat" For Input As #1
  
  Input #1, dbPassword
  Input #1, Logging
  
Close #1


frmService.Crypto.Password = "PcTechnicsCryptoAPI"
AdminPassword = frmService.Crypto.Decrypt(dbPassword)







  VERIFY_IP_SETTINGS
  Open App.Path & "\IP.dat" For Input As #2

        Input #2, RouterIP, UpdateInterval
        
        LAN_ID = Split(RouterIP, ".")
        
        octet1 = LAN_ID(0)
        octet2 = LAN_ID(1)
        octet3 = LAN_ID(2)
        
        NET_ID = octet1 & "." & octet2 & "." & octet3 & "."

    Close #2
 
 
 









    VERIFY_PrimaryConfigs
    'Get The Primary Configuration Settings
    Open App.Path & "\PrimaryConfigs.dat" For Input As #3

      Input #3, dbPriHTTPport, dbPriHTTPip, dbPriHTTPenabled
      Input #3, dbPriDNSport, dbPriDNSip, dbPriDNSenabled
      Input #3, dbPriFTPport, dbPriFTPip, dbPriFTPenabled
      Input #3, dbPriSMTPport, dbPriSMTPip, dbPriSMTPenabled
      Input #3, dbPriPOP3port, dbPriPOP3ip, dbPriPOP3enabled


    Close #3











               VERIFY_FailoverConfigs
               'Get The Failover Settings
                Open App.Path & "\FailoverConfigs.dat" For Input As #4

                     Input #4, dbSecHTTPport, dbSecHTTPip, dbSecHTTPenabled
                     Input #4, dbSecDNSport, dbSecDNSip, dbSecDNSenabled
                     Input #4, dbSecFTPport, dbSecFTPip, dbSecFTPenabled
                     Input #4, dbSecSMTPport, dbSecSMTPip, dbSecSMTPenabled
                     Input #4, dbSecPOP3port, dbSecPOP3ip, dbSecPOP3enabled
                                 
                Close #4






 
 
 
 If dbPriHTTPenabled = "0" Then
      HTTP_ACTIVE = "&"
 Else
      HTTP_ACTIVE = "&VvG=on&"
End If
 
 
 
     If dbPriDNSenabled = "0" Then
         DNS_ACTIVE = "&"
   Else
         DNS_ACTIVE = "&VvD=on&"
 End If



       If dbPriFTPenabled = "0" Then
          FTP_ACTIVE = "&"
     Else
          FTP_ACTIVE = "&VvA=on&"
   End If

 
 
            If dbPriSMTPenabled = "0" Then
               SMTP_ACTIVE = "&"
          Else
               SMTP_ACTIVE = "&VvC=on&"
        End If
 
 
 
 
                  If dbPriPOP3enabled = "0" Then
                     POP3_ACTIVE = "&"
                Else
                     POP3_ACTIVE = "&VvH=on&"
              End If









   'Is There A Primary HTTP Server Enabled? If so, lets check it out...
   If dbPriHTTPenabled = "1" Then

    If Logging = "1" Then Write2Log (vbCrLf & vbCrLf & vbTab & Now & vbCrLf & _
                                     "Primary HTTP Enabled... Checking TCP socket status at " & NET_ID & dbPriHTTPip & ":" & dbPriHTTPport & "...")
                       
                           
    
 
                    'Check Status of primary HTTP server
                    If frmService.Socket1.IsBlocked = False Then


                            frmService.Socket1.AutoResolve = False
                            frmService.Socket1.Blocking = True
                            frmService.Socket1.Timeout = 400
                            frmService.Socket1.HostAddress = NET_ID & dbPriHTTPip
                            frmService.Socket1.RemotePort = dbPriHTTPport
                           
                                                      
                            strHeader = "HEAD / HTTP/1.0" & Chr(13) & Chr(10) & _
                                        "User-Agent: Linksys_Failover_Client Build-" & AppVersion & _
                                         Chr(13) & Chr(10) & Chr(13) & Chr(10)



                                   If frmService.Socket1.Connect = 0 Then
                                   'We made a successful connection to the web server

                                                             
                                           If frmService.Socket1.Write(strHeader, 2048) Then
                                              'We sent the request successfully
                                         Else
                                               Write2Log ("An error occurred at the protocol level while sending a HEAD request to " & _
                                                          NET_ID & dbPriHTTPip & ":" & dbPriHTTPport & ". HTTP Poll Aborted..." & vbCrLf)
                                               Exit Sub
                                       End If
                                           
                                           
                                               'Create a Loop to read in each line one at a time from the web server's response
                                               Do

                                                    cchBuffer = frmService.Socket1.Read(strBuffer, 1024)
                                                 If cchBuffer = 0 Then
                                                    'The server has closed the connection and we have
                                                    'reached the end of the data stream
                                                     Exit Do
        
                                                 ElseIf cchBuffer = -1 Then
                                                        'An error has occurred while reading data from the
                                                        'server... This is considered a fatal error
            
                               frmService.Socket1.Disconnect
                                                       Write2Log ("An internal error occurred at the protocol level while reading the HEAD request reply from " & _
                                                                            NET_ID & dbPriHTTPip & ":" & dbPriHTTPport & ". HTTP Poll Aborted..." & vbCrLf)
                                                       Exit Sub
                                             End If

                                                      'As the loop progresses and each line is read, it will populate
                                                      'this variable with the entire reply from the server.
                                                       HTTP_HEAD_RESPONSE = strBuffer
                                                       
                                                       
                                                       
                                                Loop
                                                
                               frmService.Socket1.Disconnect
                               
                                                      ResponseCode = Split(HTTP_HEAD_RESPONSE, " ")
                                                      StatusCode = ResponseCode(1)
                                                      
                                                
                                                          Select Case StatusCode
  
  
  
                                                                'Successful
                                                                Case "200"
  
                                                                      StatusDesc = "Successful"
    
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
    
    
    
    
                                                              'If the web server did not give us a valid reply, but allowed us to connect,
                                                              'we treat it the same as no connection at all. (The server is probably hung.)
                                                               Case Else
                                                              
                                                                    If Logging = "1" Then Write2Log ("The primary HTTP server allowed a successful connection, but did not return a valid response. The server is probably hung. Initiating an HTTP failover recovery attempt." & vbCrLf)
                                                                 
                                                        End Select
                                                             
                                                             
                                                              
                                                                    If Logging = "1" Then Write2Log (Now & vbTab & "The primary HTTP server responded with status code: " & StatusCode & " (" & StatusDesc & ")")


                             

                              Else
                               'Primary server is offline.
                               'Close Previous Socket Connection
                              
                                 frmService.Socket1.Disconnect
                               
                       
                       
                       
                       
                       
'Check To See If There Is A Failover Server Enabled For Primary HTTP Failed Server
If dbSecHTTPenabled = "1" Then
                                     
   If Logging = "1" Then Write2Log (Now & vbTab & "Failover server enabled.... Probing failover server at " & NET_ID & dbSecHTTPip & ":" & dbSecHTTPport & " for HTTP port status.")
                                                    
                              
                          
                              
                              
                              'Check HTTP Failover Server Status
                               If frmService.Socket1.IsBlocked = False Then

                                  frmService.Socket1.Blocking = True
                                  frmService.Socket1.Timeout = 400
                                  frmService.Socket1.HostAddress = NET_ID & dbSecHTTPip
                                  frmService.Socket1.RemotePort = dbSecHTTPport
                                  frmService.Socket1.Connect


                                  If frmService.Socket1.Connected Then
                                  'Failover Server is responding
                                                                   
                                     If Logging = "1" Then Write2Log (Now & vbTab & "HTTP acknowledgment response from failover server returned successful. Initiating HTTP recovery attempt." & vbCrLf)


                                     'WE NEED TO BE DISCONNECTED SO THAT THE LINKSYS
                                     'UPDATE FUNCTION CAN USE THE SOCKET !!!
                                      frmService.Socket1.Disconnect


                                                        
'---- Code to do the Linksys update here!!!
'----
'---
               
RouterResponse = (UpdateRouter(RouterIP, "/Gozila.cgi?Uvalid=" & _
                                         "&VpAint=" & dbPriFTPport & "&VipA3=" & dbPriFTPip & FTP_ACTIVE & _
                                         "VpBint=23&VipB3=0&" & _
                                         "VpCint=" & dbPriSMTPport & "&VipC3=" & dbPriSMTPip & SMTP_ACTIVE & _
                                         "VpDint=" & dbPriDNSport & "&VipD3=" & dbPriDNSip & DNS_ACTIVE & _
                                         "VpEint=69&VipE3=0&VpFint=79&VipF3=0&" & _
                                         "VpGint=" & dbSecHTTPport & "&VipG3=" & dbSecHTTPip & HTTP_ACTIVE & _
                                         "VpHint=" & dbPriPOP3port & "&VipH3=" & dbPriPOP3ip & POP3_ACTIVE & _
                                         "VpIint=119&VipI3=0&VpJint=161&VipJ3=0&ForwardEnd=1", AdminPassword))
                                     
                                     
          RouterArray = Split(RouterResponse, ",")
          
          ResponseCode = RouterArray(0)
          ResponseDescription = RouterArray(1)
                 
                 

          If ResponseCode = "200" Then
          
          
                 If Logging = "1" Then Write2Log ("Failover recovery was successful. Failover server at " & NET_ID & _
                                                   dbSecHTTPip & ":" & dbSecHTTPport & " is now your primary HTTP server.")
                                       
                        
                         
                         
                         
                         
                            
                                        'Update The PrimaryConfigs.dat file
                                         Open App.Path & "/PrimaryConfigs.dat" For Output As #11
                                                        
                                            Write #11, dbSecHTTPport, dbSecHTTPip, "1"
                                            Write #11, dbPriDNSport, dbPriDNSip, dbPriDNSenabled
                                            Write #11, dbPriFTPport, dbPriFTPip, dbPriFTPenabled
                                            Write #11, dbPriSMTPport, dbPriSMTPip, dbPriSMTPenabled
                                            Write #11, dbPriPOP3port, dbPriPOP3ip, dbPriPOP3enabled
                                                        
                                                        
                                         Close #11
                                         
                                         
                                         
                                         
                                          'Update The FailoverConfigs.dat file
                                           Open App.Path & "\FailoverConfigs.dat" For Output As #12

                                                Write #12, dbPriHTTPport, dbPriHTTPip, "1"
                                                Write #12, dbSecDNSport, dbSecDNSip, dbSecDNSenabled
                                                Write #12, dbSecFTPport, dbSecFTPip, dbSecFTPenabled
                                                Write #12, dbSecSMTPport, dbSecSMTPip, dbSecSMTPenabled
                                                Write #12, dbSecPOP3port, dbSecPOP3ip, dbSecPOP3enabled
                                 
                                           Close #12
                                         
           Else
           
           
           If Logging = "1" Then Write2Log ("----- ROUTER UPDATE FAILED ------")
           If Logging = "1" Then Write2Log (ResponseDescription)
           
           End If
           
           
           
           End If
                                                        
                                                        
         
         
    '---
    '---          End Linksys Update
    '-------------------------------------------



                                     
                                     
                                   
                                                                            
                                   
                                      
                                       
        
                            Else
                            'Close Previous Socket
                             frmService.Socket1.Disconnect
                
                            'The HTTP Failover Socket Is Blocked
                             If Logging = "1" Then Write2Log (Now & vbTab & "The failover HTTP socket connection was blocked. HTTP Poll was aborted.")
                                Exit Sub
                                                                       
                         End If
                              
                                  
                              
                              
                              
                              
                              
                Else
                'No Failover Server Is Configured
                 If Logging = "1" Then Write2Log (Now & vbTab & "Server failure at " & NET_ID & dbPriHTTPip & ":" & dbPriHTTPport & " was detected, however, there is no failover server configured. Ignoring failure.")
           End If
             
             
             
             
             
             
             
             
             
             
              End If
                                
               
                                       
                        
                        
                          Else
                             'The Primary HTTP Socket Is Blocked
                              If Logging = "1" Then Write2Log (Now & vbTab & "The primary HTTP socket connection was blocked. HTTP Poll was aborted.")
                                 Exit Sub
                          End If
 
 
 
 
 
    Else
    'PRIMARY HTTP IS NOT CONFIGURED...
End If
' ------------------------------------------ END HTTP Failover Update ----------------------------------------


End Sub





