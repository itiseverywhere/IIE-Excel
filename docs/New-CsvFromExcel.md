---
external help file: PSP-AzureInventory-help.xml
Module Name: PSP-AzureInventory
online version: https://www.powershellpr0mpt.com
schema: 2.0.0
---

# Get-PspAzManagedDiskInfo

## SYNOPSIS
Gets Azure Virtual Machines' Managed Disk information

## SYNTAX

```
Get-PspAzManagedDiskInfo [<CommonParameters>]
```

## DESCRIPTION
Provides an easy overview of Virtual Machines' Managed Disk information.
Consolidating information from various sources in to one output, such as LinkedVM, if it's Orphaned, Operating System information and more

## EXAMPLES

### EXAMPLE 1
```
Get-PspAzManagedDiskInfo
```

DiskName                  ResourceGroupName       State          LinkedVM         DiskSizeGB Orphaned
--------                  -----------------       -----          --------         ---------- --------
fs03_Cloned_E_L           RG-PRD-PSP-APP          Attached       vmPSPweuprdfs03  4096       False
vmpspweuprdah03-c         RG-PRD-PSP-APP          Attached       vmPSPweuprdah03  127        False
vmpspweuprdah03-e         RG-PRD-PSP-APP          Attached       vmPSPweuprdah03  512        False
vmpspweuprddp01-c         RG-PRD-PSP-APP          Attached       vmPSPweuprddp01  127        False
vmpspweuprddp01-e         RG-PRD-PSP-APP          Attached       vmPSPweuprddp01  64         False
vmpspweuprdep01_e         RG-PRD-PSP-APP          Attached       vmPSPweuprdep01  128        False
goldimage_disk1           RG-PRD-PSP-RDS          Reserved       goldimage        128        False
vmpspimage01_disk1        RG-PRD-PSP-RDS          Unattached                      128        True

Gets all Managed VM disks for the currently connected subscription and displays the default properties

### EXAMPLE 2
```
Get-PspAzManagedDiskInfo | Format-List
```

DiskName              : vmpspweuprdmg03_OsDisk_1
ResourceGroupName     : RG-PRD-PSP-APP
State                 : Attached
DiskSizeGB            : 127
LinkedVM              : vmpspweuprdmg03
LinkedVMResourceGroup : RG-PRD-PSP-APP
Orphaned              : False
CreateOption          : FromImage
OSDisk                : True
OSFamily              : Windows
OSOffer               : WindowsServer
OSSku                 : 2019-Datacenter
SourceDisk            :
Encryption            : EncryptionAtRestWithPlatformKey
IopsRW                : 500
MpsRW                 : 100
HyperVGen             : V1
TagsAvailable         : False
Tags                  : env=demo;createdby=ARM
Location              : westeurope
Subscription          : 1a2b3c4d-1234-5678-9101-5e6f7g8h9i0k
ResourceGuid          : 1a2b3c4d-1234-5678-9101-5e6f7g8h9i0k
ReportDateTime        : 2021-04-19 13-37

Gets all Managed VM disks for the currently connected subscription and displays the full properties

## PARAMETERS

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable.
For more information, see about_CommonParameters (http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### PSP.Azure.Inventory.ManagedDisk
## NOTES
Name: Get-PspAzManagedDiskInfo.ps1
Author: Robert Prüst
Module: PSP-AzureInventory
DateCreated: 12-04-2021
DateModified: 19-04-2021
Blog: https://www.powershellpr0mpt.com

## RELATED LINKS

[https://www.powershellpr0mpt.com](https://www.powershellpr0mpt.com)

