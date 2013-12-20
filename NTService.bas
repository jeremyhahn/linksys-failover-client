Attribute VB_Name = "SubMain"
Option Explicit
Public Const SVCLOGFILE = "/LOG.txt"


Public Sub subLogCommand(strMessage As String, strCommand As String)


  '***  use frmService.NTService1.LogEvent svcEventInformation, svcMessageInfo, strCommand
  
  Dim lngFileNum As Long
  
  lngFileNum = FreeFile()
  Open App.Path & SVCLOGFILE For Append As #lngFileNum
  Print #lngFileNum, strMessage & " - " & Now, strCommand
  Close #lngFileNum

End Sub




Public Sub Main()
  
  'Write to the log
   subLogCommand "Service Started ", Command$
  
  'Select the right action to take
   Select Case Trim$(Command$)
    
    'Tell the OCX to install the service and quit
    Case "/install"
      
      
      'Set all defaults for the service here
      With frmService.NTService1
        
        'True if the service needs to interact with the user
        .Interactive = True
       
       
            'Use these to read / write to the registry at
            '***  HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\[SERVICE NAME]
            '.SaveSetting
            '.GetSetting
            '.DeleteSetting
            '.GetAllSettings
        
        
                'For example write the Description shown in the Services window.
                 .SaveSetting "", "Description", "This service updates your Linksys router when a server failure has occurred on your local network. This service will update your Linksys router to use a backup server to provide your network with redundancy and functionality 24 hours a day."
                
                
                    'Set the startmode to manual by default
                    .StartMode = svcStartAutomatic
                    
        
                        'Now install the service
                        .Install
        
        
                              'Write a line to the system log
                              .LogEvent svcEventInformation, svcMessageInfo, "Linksys Failover Client Installed Successfully."
                            End With
                            
                            subLogCommand Now & " - NT Service Installed Successfully", Command$
                            frmService.NTService1.StartService
                            
      
      
      
      
    'Tell the OCX to uninstall the service and quit
    Case "/uninstall"
    
    subLogCommand Now & " - NT Service Uninstalled", Command$
    
    frmService.NTService1.Uninstall
    Unload frmService
   
    
    


      
            '***  start the service. simply load the form
            Case ""
              frmService.NTService1.StartService

    
    
    
'###########################################################################################
'#                           Command Line Arguments For The Client                         #
'###########################################################################################
'# NOTE: Command Line Arguments Are Case Sensitive #
'###################################################
    
    
    
  'Start The Service
   Case "/start"
   Shell ("net start lfc"), vbHide
                                             
                                                                                         
                                                                                        
        'Stop The Service
         Case "/stop"
         Shell ("net stop lfc"), vbHide
                                             
                                             
                                                                                          
                                     
                                                   
                                                   
                                                   
                                                   
    
    
    'An Unrecognized Command Was Executed
    Case Else
      MsgBox " You have executed an unrecognized command. Please use one of the following." & vbNewLine & vbNewLine & _
      "/start - Starts The NT Service" & vbNewLine & _
      "/stop - Stops The NT Service" & vbNewLine & _
      "/install - Installs NT Service" & vbNewLine & _
      "/uninstall - Removes NT Service" & vbNewLine & _
      "NOTE: Command line arguments are CaSe SeNsItIvE!"

      Unload frmService
  
  End Select

End Sub

