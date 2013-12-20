Attribute VB_Name = "POP3_POLL"

Public Sub DO_POP3_POLL()


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









   'Is There A Primary POP3 Server Enabled? If so, lets check it out...
   If dbPriPOP3enabled = "1" Then

    If Logging = "1" Then Write2Log (vbCrLf & vbCrLf & vbTab & Now & vbCrLf & _
                                     "Primary POP3 Enabled... Checking TCP socket status at " & NET_ID & dbPriPOP3ip & ":" & dbPriPOP3port & "...")
                       
                           
    
 
                    'Check Status of primary POP3 server
                    If frmService.Socket1.IsBlocked = False Then


                            'frmService.Socket1.AutoResolve = False
                            frmService.Socket1.Blocking = True
                            frmService.Socket1.Timeout = 400
                            frmService.Socket1.HostAddress = NET_ID & dbPriPOP3ip
                            frmService.Socket1.RemotePort = dbPriPOP3port
                            frmService.Socket1.Connect
   
   
                                Dim POP3_RESPONSE As String
                                    frmService.Socket1.Read POP3_RESPONSE, 1024
                                
                                 
                                   
                                   If Not (POP3_RESPONSE = vbNullString) Then
                                   'We made a successful connection to the POP3 server
                                                                                        
                                      frmService.Socket1.Disconnect
                               
                                    If Logging = "1" Then Write2Log (Now & vbTab & "The primary POP3 server responded with : " & POP3_RESPONSE)


                             

                              Else
                               'Primary server is offline.
                               'Close Previous Socket Connection
                              
                                 frmService.Socket1.Disconnect
                               
                       
                       
                       
                       
'Check To See If There Is A Failover Server Enabled For Primary POP3 Failed Server
If dbSecPOP3enabled = "1" Then
                                     
   If Logging = "1" Then Write2Log (Now & vbTab & "Failover server enabled.... Probing failover server at " & NET_ID & dbSecPOP3ip & ":" & dbSecPOP3port & " for POP3 port status.")
                                                    
                              
                          
                              
                              
                              'Check POP3 Failover Server Status
                               If frmService.Socket1.IsBlocked = False Then

                                  frmService.Socket1.Blocking = True
                                  frmService.Socket1.Timeout = 400
                                  frmService.Socket1.HostAddress = NET_ID & dbSecPOP3ip
                                  frmService.Socket1.RemotePort = dbSecPOP3port
                                  frmService.Socket1.Connect
                                  frmService.Socket1.Connect

                                  Dim FO_POP3_RESPONSE As String
                                  
                                    frmService.Socket1.Read FO_POP3_RESPONSE, 1024
                                  
                                  If Not (FO_POP3_RESPONSE = vbNullString) Then
                                  'Failover Server is responding
                                                                   
                                     If Logging = "1" Then Write2Log (Now & vbTab & "POP3 acknowledgment response from failover server returned successful. Initiating POP3 recovery attempt." & vbCrLf)


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
                                         "VpGint=" & dbPriHTTPport & "&VipG3=" & dbPriHTTPip & HTTP_ACTIVE & _
                                         "VpHint=" & dbSecPOP3port & "&VipH3=" & dbSecPOP3ip & POP3_ACTIVE & _
                                         "VpIint=119&VipI3=0&VpJint=161&VipJ3=0&ForwardEnd=1", AdminPassword))
                                     
                                     
          RouterArray = Split(RouterResponse, ",")
          
          ResponseCode = RouterArray(0)
          ResponseDescription = RouterArray(1)
                 
                 

          If ResponseCode = "200" Then
          
          
                 If Logging = "1" Then Write2Log ("Failover recovery was successful. Failover server at " & NET_ID & _
                                                   dbSecPOP3ip & ":" & dbSecPOP3port & " is now your primary POP3 server.")
                                       
                        
                         
                         
                         
                         
                            
                                        'Update The PrimaryConfigs.dat file
                                         Open App.Path & "/PrimaryConfigs.dat" For Output As #11
                                                        
                                            Write #11, dbPriHTTPport, dbPriHTTPip, dbPriHTTPenabled
                                            Write #11, dbPriDNSport, dbPriDNSip, dbPriDNSenabled
                                            Write #11, dbPriFTPport, dbPriFTPip, dbPriFTPenabled
                                            Write #11, dbPriSMTPport, dbPriSMTPip, dbPriSMTPenabled
                                            Write #11, dbSecPOP3port, dbSecPOP3ip, "1"
                                                        
                                                        
                                         Close #11
                                         
                                         
                                         
                                         
                                          'Update The FailoverConfigs.dat file
                                           Open App.Path & "\FailoverConfigs.dat" For Output As #12

                                                Write #12, dbSecHTTPport, dbSecHTTPip, dbPriHTTPenabled
                                                Write #12, dbSecDNSport, dbSecDNSip, dbSecDNSenabled
                                                Write #12, dbSecFTPport, dbSecFTPip, dbSecFTPenabled
                                                Write #12, dbSecSMTPport, dbSecSMTPip, dbSecSMTPenabled
                                                Write #12, dbPriPOP3port, dbPriPOP3ip, "1"
                                 
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
                
                            'The POP3 Failover Socket Is Blocked
                             If Logging = "1" Then Write2Log (Now & vbTab & "The failover POP3 socket connection was blocked. POP3 Poll was aborted.")
                                Exit Sub
                                                                       
                         End If
                              
                                  
                              
                              
                              
                              
                              
                Else
                'No Failover Server Is Configured
                 If Logging = "1" Then Write2Log (Now & vbTab & "Server failure at " & NET_ID & dbPriPOP3ip & ":" & dbPriPOP3port & " was detected, however, there is no failover server configured. Ignoring failure.")
           End If
             
             
             
             
             
             
             
             
             
             
              End If
                                
               
                                       
                        
                        
                          Else
                             'The Primary POP3 Socket Is Blocked
                              If Logging = "1" Then Write2Log (Now & vbTab & "The primary POP3 socket connection was blocked. POP3 Poll was aborted.")
                                 Exit Sub
                          End If
 
 
 
 
 
    Else
    'PRIMARY POP3 IS NOT CONFIGURED...
End If
' ------------------------------------------ END POP3 Failover Update ----------------------------------------


End Sub











