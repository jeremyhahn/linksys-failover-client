Attribute VB_Name = "Functions"
Global Const AppVersion = "1.1"  ' < ~~~ CHANGE WITH UPDATES !!!






'Vaildation Functions
Public Function VERIFY_SYSCONFIG()

Dim SysConfig
    SysConfig = App.Path & "/sysconfig.dat"
  
  
    If Dir$(SysConfig) = vbNullString Then

    Open App.Path & "/sysconfig.dat" For Output As #64
     
     frmService.Crypto.Password = "PcTechnicsCryptoAPI"
     Password = frmService.Crypto.Encrypt("admin")
     
     Write #64, Password
     Write #64, "1"
     
     
    Close #64
    
    
End If

Set SysConfig = Nothing

End Function



















Public Function VERIFY_IP_SETTINGS()

Dim IPsettings
    IPsettings = App.Path & "/IP.dat"
  
  
    If Dir$(IPsettings) = vbNullString Then

        Open App.Path & "/IP.dat" For Output As #65

             Write #65, "192.168.1.1", "10"

        Close #65
    
End If

Set IPsettings = Nothing

End Function

















Public Function VERIFY_PrimaryConfigs()

Dim PrimaryConfig
    PrimaryConfig = App.Path & "/PrimaryConfigs.dat"
  
  
    If Dir$(PrimaryConfig) = vbNullString Then

        Open App.Path & "/PrimaryConfigs.dat" For Output As #66

             Write #66, "80", "2", "0"
             Write #66, "53", "2", "0"
             Write #66, "21", "2", "0"
             Write #66, "25", "2", "0"
             Write #66, "110", "2", "0"

        Close #66
    
End If

End Function








Public Function VERIFY_FailoverConfigs()


Dim FailoverConfig
    FailoverConfig = App.Path & "/FailoverConfigs.dat"
  
  
    If Dir$(FailoverConfig) = vbNullString Then

        Open App.Path & "/FailoverConfigs.dat" For Output As #67

             Write #67, "80", "2", "0"
             Write #67, "53", "2", "0"
             Write #67, "21", "2", "0"
             Write #67, "25", "2", "0"
             Write #67, "110", "2", "0"

        Close #67
    
End If


End Function















Public Function VerifyPort(Port) As Boolean

If Not IsNumeric(Port) Then
  VerifyPort = False
Exit Function
End If


If Port <= 0 Or Port >= 65536 Then
VerifyPort = False
Exit Function
End If


VerifyPort = True

End Function












Public Function VerifyIP(IP) As Boolean

If Not IsNumeric(IP) Then
VerifyIP = False
Exit Function
End If

If IP <= 0 Or IP >= 256 Then
VerifyIP = False
Exit Function
End If

VerifyIP = True

End Function










Public Function POPULATE_LISTVIEW()


frmMain.txtLogDisplay.Clear


Dim LogFile
    LogFile = App.Path & "/LOG.txt"
  
  
    If Dir$(LogFile) = vbNullString Then

        Open App.Path & "/LOG.txt" For Output As #1

             Write #1, ""
   
        Close #1

End If





     Open App.Path & "/LOG.txt" For Input As #1
 
 i = 0
 
     Do While Not EOF(1)
   
     i = i + 1
     
     Input #1, strTemp
     
     strTemp = Replace(strTemp, vbCrLf, "")
     strTemp = Replace(strTemp, vbTab, "")
     strTemp = Replace(strTemp, "AM", " AM - ")
     strTemp = Replace(strTemp, "PM", " PM - ")
     
     
        frmMain.txtLogDisplay.AddItem strTemp
           
     DoEvents
     Loop
           
 
 Close #1
End Function













Public Function DO_GLOBAL_POLL()

DO_HTTP_POLL
DO_FTP_POLL
DO_SMTP_POLL
DO_POP3_POLL
DO_DNS_POLL

End Function















'Logging Function
Public Function Write2Log(Data2log)

Dim LogFile
    LogFile = App.Path & "/LOG.txt"
  
    If Dir$(LogFile) = vbNullString Then

        Open App.Path & "/LOG.txt" For Output As #1

             Write #1, ""

        Close #1
End If



Open App.Path & "/LOG.txt" For Append As #1

   Write #1, Data2log
   
Close #1


End Function


























'Updates the text controls on the config form
Public Function UpdateControls()



  VERIFY_PrimaryConfigs
  'Get The Primary Configuration Settings
    Open App.Path & "\PrimaryConfigs.dat" For Input As #20

      Input #20, dbPriHTTPport, dbPriHTTPip, dbPriHTTPenabled
      Input #20, dbPriDNSport, dbPriDNSip, dbPriDNSenabled
      Input #20, dbPriFTPport, dbPriFTPip, dbPriFTPenabled
      Input #20, dbPriSMTPport, dbPriSMTPip, dbPriSMTPenabled
      Input #20, dbPriPOP3port, dbPriPOP3ip, dbPriPOP3enabled


    Close #20



               VERIFY_FailoverConfigs
               'Get The Failover Settings
                Open App.Path & "\FailoverConfigs.dat" For Input As #21

                     Input #21, dbSecHTTPport, dbSecHTTPip, dbSecHTTPenabled
                     Input #21, dbSecDNSport, dbSecDNSip, dbSecDNSenabled
                     Input #21, dbSecFTPport, dbSecFTPip, dbSecFTPenabled
                     Input #21, dbSecSMTPport, dbSecSMTPip, dbSecSMTPenabled
                     Input #21, dbSecPOP3port, dbSecPOP3ip, dbSecPOP3enabled
                                 
                Close #21
                
                
                
                 'Update The Primary Configs Text Controls
                  frmMain.PriHTTPport.Text = dbPriHTTPport
                  frmMain.PriHTTPip.Text = dbPriHTTPip
                  frmMain.PriHTTPenabled = dbPriHTTPenabled
                  
                  frmMain.PriDNSport.Text = dbPriDNSport
                  frmMain.PriDNSip.Text = dbPriDNSip
                  frmMain.PriDNSenabled = dbPriDNSenabled
                  
                  frmMain.PriFTPport.Text = dbPriFTPport
                  frmMain.PriFTPip.Text = dbPriFTPip
                  frmMain.PriFTPenabled = dbPriFTPenabled
                  
                  frmMain.PriSMTPport.Text = dbPriSMTPport
                  frmMain.PriSMTPip.Text = dbPriSMTPip
                  frmMain.PriSMTPenabled = dbPriSMTPenabled
                  
                  frmMain.PriPOP3port.Text = dbPriPOP3port
                  frmMain.PriPOP3ip.Text = dbPriPOP3ip
                  frmMain.PriPOP3enabled = dbPriPOP3enabled
                  
                  
                  'Update The Failover Configs Text Controls
                  frmMain.secHTTPport.Text = dbSecHTTPport
                  frmMain.SecHTTPip.Text = dbSecHTTPip
                  frmMain.secHTTPenabled = dbSecHTTPenabled
                  
                  frmMain.secDNSport.Text = dbSecDNSport
                  frmMain.SecDNSip.Text = dbSecDNSip
                  frmMain.secDNSenabled = dbSecDNSenabled
                  
                  frmMain.secFTPport.Text = dbSecFTPport
                  frmMain.SecFTPip.Text = dbSecFTPip
                  frmMain.secFTPenabled = dbSecFTPenabled
                  
                  frmMain.secSMTPport.Text = dbSecSMTPport
                  frmMain.SecSMTPip.Text = dbSecSMTPip
                  frmMain.secSMTPenabled = dbSecSMTPenabled
                  
                  frmMain.secPOP3port.Text = dbSecPOP3port
                  frmMain.SecPOP3ip.Text = dbSecPOP3ip
                  frmMain.secPOP3enabled = dbSecPOP3enabled
                  
                  
                  POPULATE_LISTVIEW
                  
End Function
