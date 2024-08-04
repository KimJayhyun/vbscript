' https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ics/adding-an-outbound-exception
'  This VBScript file includes sample code that adds an  
'  outbound rule using the Microsoft Windows Firewall APIs.

option explicit

Dim CurrentProfiles

' Protocol
Const NET_FW_IP_PROTOCOL_TCP = 6
Const NET_FW_IP_PROTOCOL_UDP = 17

'Direction
Const NET_FW_RULE_DIR_IN = 1
Const NET_FW_RULE_DIR_OUT = 2

'Action
Const NET_FW_ACTION_ALLOW = 1

' Create the FwPolicy2 object.
Dim fwPolicy2
Set fwPolicy2 = CreateObject("HNetCfg.FwPolicy2")

' Get the Rules object
Dim RulesObject
Set RulesObject = fwPolicy2.Rules

CurrentProfiles = fwPolicy2.CurrentProfileTypes

'Create a Rule Object.
Dim NewRule
Set NewRule = CreateObject("HNetCfg.FWRule")
    
NewRule.Name = "Inbound Test"
NewRule.Description = "VBS Test"
NewRule.Applicationname = "%systemDrive%\Program Files\MyApplication.exe"
NewRule.Protocol = NET_FW_IP_PROTOCOL_TCP
NewRule.LocalPorts = 8443
'NewRule.LocalPorts = "8443,123"
'NewRule.LocalPorts = "8443-123"
NewRule.Direction = NET_FW_RULE_DIR_IN
NewRule.Enabled = TRUE  
'NewRule.Grouping = "@firewallapi.dll,-23255"
NewRule.Profiles = CurrentProfiles
NewRule.Action = NET_FW_ACTION_ALLOW
        
'Add a new rule
RulesObject.Add NewRule