Attribute VB_Name = "FTP_POLL"

Public Sub DO_FTP_POLL()


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









   'Is There A Primary FTP Server Enabled? If so, lets check it out...
   If dbPriFTPenabled = "1" Then

    If Logging = "1" Then Write2Log (vbCrLf & vbCrLf & vbTab & Now & vbCrLf & _
                                     "Primary FTP Enabled... Checking TCP socket status at " & NET_ID & dbPriFTPip & ":" & dbPriFTPport & "...")
                       
                           
    
 
                    'Check Status of primary FTP server
                    If frmService.Socket1.IsBlocked = False Then


                            'frmService.Socket1.AutoResolve = False
                            frmService.Socket1.Blocking = True
                            frmService.Socket1.Timeout = 400
                            frmService.Socket1.HostAddress = NET_ID & dbPriFTPip
                            frmService.Socket1.RemotePort = dbPriFTPport
                            frmService.Socket1.Connect
   
   
                                Dim FTP_RESPONSE As String
                                    frmService.Socket1.Read FTP_RESPONSE, 1024
                                
                                 
                                   
                                   If Not (FTP_RESPONSE = vbNullString) Then
                                   'We made a successful connection to the FTP server
                                                                                        
                                      frmService.Socket1.Disconnect
                               
                                    If Logging = "1" Then Write2Log (Now & vbTab & "The primary FTP server responded with : " & FTP_RESPONSE)


                             

                              Else
                               'Primary server is offline.
                               'Close Previous Socket Connection
                              
                                 frmService.Socket1.Disconnect
                               
                       
                       
                       
                       
'Check To See If There Is A Failover Server Enabled For Primary FTP Failed Server
If dbSecFTPenabled = "1" Then
                                     
   If Logging = "1" Then Write2Log (Now & vbTab & "Failover server enabled.... Probing failover server at " & NET_ID & dbSecFTPip & ":" & dbSecFTPport & " for FTP port status.")
                                                    
                              
                          
                              
                              
                              'Check FTP Failover Server Status
                               If frmService.Socket1.IsBlocked = False Then

                                  frmService.Socket1.Blocking = True
                                  frmService.Socket1.Timeout = 400
                                  frmService.Socket1.HostAddress = NET_ID & dbSecFTPip
                                  frmService.Socket1.RemotePort = dbSecFTPport
                                  frmService.Socket1.Connect
                                  frmService.Socket1.Connect

                                  Dim FO_FTP_RESPONSE As String
                                  
                                    frmService.Socket1.Read FO_FTP_RESPONSE, 1024
                                  
                                  If Not (FO_FTP_RESPONSE = vbNullString) Then
                                  'Failover Server is responding
                                                                   
                                     If Logging = "1" Then Write2Log (Now & vbTab & "FTP acknowledgment response from failover server returned successful. Initiating FTP recovery attempt." & vbCrLf)


                                     'WE NEED TO BE DISCONNECTED SO THAT THE LINKSYS
                                     'UPDATE FUNCTION CAN USE THE SOCKET !!!
                                      frmService.Socket1.Disconnect


                                                        
'---- Code to do the Linksys update here!!!
'----
'---
               
RouterResponse = (UpdateRouter(RouterIP, "/Gozila.cgi?Uvalid=" & _
                                         "&VpAint=" & dbSecFTPport & "&VipA3=" & dbSecFTPip & FTP_ACTIVE & _
                                         "VpBint=23&VipB3=0&" & _
                                         "VpCint=" & dbPriSMTPport & "&VipC3=" & dbPriSMTPip & SMTP_ACTIVE & _
                                         "VpDint=" & dbPriDNSport & "&VipD3=" & dbPriDNSip & DNS_ACTIVE & _
                                         "VpEint=69&VipE3=0&VpFint=79&VipF3=0&" & _
                                         "VpGint=" & dbPriHTTPport & "&VipG3=" & dbPriHTTPip & HTTP_ACTIVE & _
                                         "VpHint=" & dbPriPOP3port & "&VipH3=" & dbPriPOP3ip & POP3_ACTIVE & _
                                         "VpIint=119&VipI3=0&VpJint=161&VipJ3=0&ForwardEnd=1", AdminPassword))
                                     
                                     
          RouterArray = Split(RouterResponse, ",")
          
          ResponseCode = RouterArray(0)
          ResponseDescription = RouterArray(1)
                 
                 

          If ResponseCode = "200" Then
          
          
                 If Logging = "1" Then Write2Log ("Failover recovery was successful. Failover server at " & NET_ID & _
                                                   dbSecFTPip & ":" & dbSecFTPport & " is now your primary FTP server.")
                                       
                        
                         
                         
                         
                         
                            
                                        'Update The PrimaryConfigs.dat file
                                         Open App.Path & "/PrimaryConfigs.dat" For Output As #11
                                                        
                                            Write #11, dbPriHTTPport, dbPriHTTPip, dbPriHTTPenabled
                                            Write #11, dbPriDNSport, dbPriDNSip, dbPriDNSenabled
                                            Write #11, dbSecFTPport, dbSecFTPip, "1"
                                            Write #11, dbPriSMTPport, dbPriSMTPip, dbPriSMTPenabled
                                            Write #11, dbPriPOP3port, dbPriPOP3ip, dbPriPOP3enabled
                                                        
                                                        
                                         Close #11
                                         
                                         
                                         
                                         
                                          'Update The FailoverConfigs.dat file
                                           Open App.Path & "\FailoverConfigs.dat" For Output As #12

                                                Write #12, dbSecHTTPport, dbSecHTTPip, dbPriHTTPenabled
                                                Write #12, dbSecDNSport, dbSecDNSip, dbSecDNSenabled
                                                Write #12, dbPriFTPport, dbPriFTPip, "1"
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
                
                            'The FTP Failover Socket Is Blocked
                             If Logging = "1" Then Write2Log (Now & vbTab & "The failover FTP socket connection was blocked. FTP Poll was aborted.")
                                Exit Sub
                                                                       
                         End If
                              
                                  
                              
                              
                              
                              
                              
                Else
                'No Failover Server Is Configured
                 If Logging = "1" Then Write2Log (Now & vbTab & "Server failure at " & NET_ID & dbPriFTPip & ":" & dbPriFTPport & " was detected, however, there is no failover server configured. Ignoring failure.")
           End If
             
             
             
             
             
             
             
             
             
             
              End If
                                
               
                                       
                        
                        
                          Else
                             'The Primary FTP Socket Is Blocked
                              If Logging = "1" Then Write2Log (Now & vbTab & "The primary FTP socket connection was blocked. FTP Poll was aborted.")
                                 Exit Sub
                          End If
 
 
 
 
 
    Else
    'PRIMARY FTP IS NOT CONFIGURED...
End If
' ------------------------------------------ END FTP Failover Update ----------------------------------------


End Sub







