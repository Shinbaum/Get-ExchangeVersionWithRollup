###############################################################
#
#This Sample Code is provided for the purpose of illustration only
#and is not intended to be used in a production environment.  THIS
#SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS"
#WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,
#INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
#MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We
#grant You a nonexclusive, royalty-free right to use and modify
#the Sample Code and to reproduce and distribute the object code
#form of the Sample Code, provided that You agree: (i) to not use
#Our name, logo, or trademarks to market Your software product in
#which the Sample Code is embedded; (ii) to include a valid
#copyright notice on Your software product in which the Sample
#
#Code is embedded; and (iii) to indemnify, hold harmless, and
#defend Us and Our suppliers from and against any claims or
#lawsuits, including attorneys’ fees, that arise or result from
#the use or distribution of the Sample Code.
#Please note: None of the conditions outlined in the disclaimer
#above will supercede the terms and conditions contained within
#the Premier Customer Services Description.
#
###############################################################
#
# This script is to help collect the rollup info for Exchange 2007 and up
# Created by jashinba@microsoft.com
# Created on 10/10/13
# Last updated on 05/29/17
# Version 0.4
#
# Update: 10/15/13 - Added Alt reg key for Exchange 2007 - JRS
# Update: 10/24/13 - Fixed RUVer for 2010SP2 RU6 and 7
# Update: 04/22/14 - Added RUVer for 2007SP3 RU12-13, 2010SP3 RU3-5, 2013 RU3-4
# Update: 07/10/14 - Added RUVer for 2010SP3 RU6, 2013 CU5
# Update: 09/18/14 - Added 2013 CU6 
# Update: 02/05/15 - Added RUVer for 2007SP3 RU13-14,2010SP3 RU7-8, 2013 CU7
# Update: 04/07/15 - Added RUVer for 2007SP3 RU16, 2010SP3 RU9, 2013 CU8
# Update: 07/24/15 - Added RUVer for 2010SP3 RU10, 2013 CU9
# Update: 04/01/16 - Added RUVer for 2007SP3 RU17-19, 2010SP3 RU11-13, 2013 CU10-12, 2016 RTM-CU1
# Update: 09/23/16 - Added RUVer for 2007SP3 RU20-21, 2010SP3 RU14-15, 2013 CU13-14, 2016 CU2-3
# Update: 05/29/17 - Added RUVer for 2007SP3 RU22-23, 2010SP3 RU16-17, 2013 CU15-16, 2016 CU4-5
#
[CmdletBinding()]
param
  (
  [Parameter(Mandatory=$True,ValueFromPipeLine=$True)][String[]]$ExchangeServer
  )

Begin
  {
  $ScriptVersion = "0.4"
  $arrExServer   = @()

  Function Get-ExchangeServerInstallFolder([Microsoft.Exchange.Data.Directory.Management.ExchangeServer]$objExServer)
    {
    $HKCR = 2147483648 #HKEY_CLASSES_ROOT
    $HKCU = 2147483649 #HKEY_CURRENT_USER
    $HKLM = 2147483650 #HKEY_LOCAL_MACHINE
    $HKUS = 2147483651 #HKEY_USERS
    $HKCC = 2147483653 #HKEY_CURRENT_CONFIG

    $strExServer = $objExServer.Name
    $reg = [wmiclass]"\\$strExServer\root\default:StdRegprov"
    $intMajorVersion = $objExServer.AdminDisplayVersion.Major
    $key = "SOFTWARE\Microsoft\ExchangeServer\v$intMajorVersion\Setup"
    $value = "MsiInstallPath"

    $MsiPath = $reg.GetStringValue($HKLM, $key, $value).sValue

    #If MsiPath is null check alt reg key
    If (!$MsiPath)
      {
      $key = "SOFTWARE\Microsoft\Exchange\v$intMajorVersion.0\Setup"
      $MsiPath = $reg.GetStringValue($HKLM, $key, $value).sValue
      }

    Return $MsiPath
    }

  Function Get-FileVersion ([String]$strServer, [String]$strFolder)
    {
    $strFile = $strFolder.Replace("\","\\") + "Bin\\ExSetup.exe"
    Return (Get-WmiObject -Query "select Version from CIM_DataFile Where Name = '$strFile'" -ComputerName $strServer).Version
    }

  Function Convert-VersionToDetails([String]$Version)
    {
    $hashDetail = @{}

    switch ($Version)
      {
      #region Exchange 2007 RTM
      "8.0.685.24"
        {
        $hashDetail.Add("Rollup",0)
        $hashDetail.Add("VersionOfRollup",0)
        $hashDetail.Add("ReleaseDate","12/9/06")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.0.685.25"
        {
        $hashDetail.Add("Rollup",0)
        $hashDetail.Add("VersionOfRollup",0)
        $hashDetail.Add("ReleaseDate","12/9/06")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.0.708.3"
        {
        $hashDetail.Add("Rollup",1)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.0.711.2"
        {
        $hashDetail.Add("Rollup",2)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.0.730.1"
        {
        $hashDetail.Add("Rollup",3)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.0.744.0"
        {
        $hashDetail.Add("Rollup",4)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.0.754.0"
        {
        $hashDetail.Add("Rollup",5)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.0.783.2"
        {
        $hashDetail.Add("Rollup",6)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.0.813.0"
        {
        $hashDetail.Add("Rollup",0)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      #endregion
      
      #region Exchange 2007 SP1
      "8.1.240.6"
        {
        $hashDetail.Add("Rollup",0)
        $hashDetail.Add("VersionOfRollup",0)
        $hashDetail.Add("ReleaseDate","11/29/07")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.1.263.1"
        {
        $hashDetail.Add("Rollup",1)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","2/28/08")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.1.278.2"
        {
        $hashDetail.Add("Rollup",2)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","5/8/08")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.1.291.2"
        {
        $hashDetail.Add("Rollup",3)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","7/8/08")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.1.311.3"
        {
        $hashDetail.Add("Rollup",4)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","10/7/08")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.1.336.1"
        {
        $hashDetail.Add("Rollup",5)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","11/20/08")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.1.340.1"
        {
        $hashDetail.Add("Rollup",6)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","2/10/09")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.1.359.2"
        {
        $hashDetail.Add("Rollup",7)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","3/18/09")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.1.375.2"
        {
        $hashDetail.Add("Rollup",8)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","5/19/09")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.1.393.1"
        {
        $hashDetail.Add("Rollup",9)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","7/17/09")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.1.436.0"
        {
        $hashDetail.Add("Rollup",10)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","4/9/10")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      #endregion

      #region Exchange 2007 SP2
      "8.2.176.2"
        {
        $hashDetail.Add("Rollup",0)
        $hashDetail.Add("VersionOfRollup",0)
        $hashDetail.Add("ReleaseDate","8/24/09")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.2.217.3"
        {
        $hashDetail.Add("Rollup",1)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","11/19/09")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.2.234.1"
        {
        $hashDetail.Add("Rollup",2)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","1/22/10")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.2.247.2"
        {
        $hashDetail.Add("Rollup",3)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","3/17/10")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.2.254.0"
        {
        $hashDetail.Add("Rollup",4)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","4/9/10")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.2.305.3"
        {
        $hashDetail.Add("Rollup",5)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","12/7/10")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      #endregion

      #region Exchange 2007 SP3
      "8.3.083.6"
        {
        $hashDetail.Add("Rollup",0)
        $hashDetail.Add("VersionOfRollup",0)
        $hashDetail.Add("ReleaseDate","6/20/10")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.3.106.2"
        {
        $hashDetail.Add("Rollup",1)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","09/09/10")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.3.137.3"
        {
        $hashDetail.Add("Rollup",2)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","12/10/10")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.3.159.0"
        {
        $hashDetail.Add("Rollup",3)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","03/02/11")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.3.159.2"
        {
        $hashDetail.Add("Rollup",3)
        $hashDetail.Add("VersionOfRollup",2)
        $hashDetail.Add("ReleaseDate","3/30/11")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.3.192.1"
        {
        $hashDetail.Add("Rollup",4)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","7/7/11")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.3.213.1"
        {
        $hashDetail.Add("Rollup",5)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","9/21/11")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.3.245.2"
        {
        $hashDetail.Add("Rollup",6)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","1/25/12")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.3.264.0"
        {
        $hashDetail.Add("Rollup",7)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","4/16/12")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.3.279.3"
        {
        $hashDetail.Add("Rollup",8)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","8/13/12")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.3.279.5"
        {
        $hashDetail.Add("Rollup",8)
        $hashDetail.Add("VersionOfRollup",2)
        $hashDetail.Add("ReleaseDate","10/9/12")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.3.279.6"
        {
        $hashDetail.Add("Rollup",8)
        $hashDetail.Add("VersionOfRollup",3)
        $hashDetail.Add("ReleaseDate","11/13/12")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.3.297.2"
        {
        $hashDetail.Add("Rollup",9)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","12/10/12")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.3.298.3"
        {
        $hashDetail.Add("Rollup",10)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","2/11/13")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.3.327.1"
        {
        $hashDetail.Add("Rollup",11)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","8/13/13")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.3.342.4"
        {
        $hashDetail.Add("Rollup",12)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","12/10/13")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "8.3.348.2"
        {
        $hashDetail.Add("Rollup",13)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","02/25/14")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "8.3.379.2"
        {
        $hashDetail.Add("Rollup",14)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","08/26/14")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "8.3.389.2"
        {
        $hashDetail.Add("Rollup",15)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","12/09/14")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "8.3.406.0"
        {
        $hashDetail.Add("Rollup",16)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","03/17/15")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "8.3.417.1"
        {
        $hashDetail.Add("Rollup",17)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","06/16/15")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "8.3.445.0"
        {
        $hashDetail.Add("Rollup",18)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","12/10/15")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "8.3.459.0"
        {
        $hashDetail.Add("Rollup",19)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","03/14/16")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "8.3.468.0"
        {
        $hashDetail.Add("Rollup",20)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","06/21/16")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "8.3.485.1"
        {
        $hashDetail.Add("Rollup",21)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","09/20/16")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "8.3.502.0"
        {
        $hashDetail.Add("Rollup",22)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","12/13/16")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "8.3.517.0"
        {
        $hashDetail.Add("Rollup",23)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","03/21/17")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      #endregion

      #region Exchange 2010 RTM
      "14.0.639.21"
        {
        $hashDetail.Add("Rollup",0)
        $hashDetail.Add("VersionOfRollup",0)
        $hashDetail.Add("ReleaseDate","11/9/09")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "14.0.682.1"
        {
        $hashDetail.Add("Rollup",1)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","12/9/09")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "14.0.689.0"
        {
        $hashDetail.Add("Rollup",2)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","3/4/10")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "14.0.694.0"
        {
        $hashDetail.Add("Rollup",3)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","4/9/10")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "14.0.702.1"
        {
        $hashDetail.Add("Rollup",4)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","6/17/10")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "14.0.726.0"
        {
        $hashDetail.Add("Rollup",5)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","12/13/10")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      #endregion

      #region Exchange 2010 SP1
      "14.1.218.15"
        {
        $hashDetail.Add("Rollup",0)
        $hashDetail.Add("VersionOfRollup",0)
        $hashDetail.Add("ReleaseDate","8/24/10")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "14.1.255.2"
        {
        $hashDetail.Add("Rollup",1)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","10/4/10")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "14.1.270.1"
        {
        $hashDetail.Add("Rollup",2)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","12/9/10")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "14.1.289.3"
        {
        $hashDetail.Add("Rollup",3)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","3/7/11")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "14.1.289.7"
        {
        $hashDetail.Add("Rollup",3)
        $hashDetail.Add("VersionOfRollup",2)
        $hashDetail.Add("ReleaseDate","4/1/11")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "14.1.323.1"
        {
        $hashDetail.Add("Rollup",4)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","6/22/11")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "14.1.323.6"
        {
        $hashDetail.Add("Rollup",4)
        $hashDetail.Add("VersionOfRollup",2)
        $hashDetail.Add("ReleaseDate","7/27/11")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "14.1.339.1"
        {
        $hashDetail.Add("Rollup",5)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","8/23/11")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "14.1.355.2"
        {
        $hashDetail.Add("Rollup",6)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","10/27/11")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "14.1.421.0"
        {
        $hashDetail.Add("Rollup",7)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","8/13/12")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "14.1.421.2"
        {
        $hashDetail.Add("Rollup",7)
        $hashDetail.Add("VersionOfRollup",2)
        $hashDetail.Add("ReleaseDate","10/9/12")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "14.1.421.3"
        {
        $hashDetail.Add("Rollup",7)
        $hashDetail.Add("VersionOfRollup",3)
        $hashDetail.Add("ReleaseDate","11/12/12")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "14.1.438.0"
        {
        $hashDetail.Add("Rollup",8)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","12/10/12")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      #endregion

      #region Exchange 2010 SP2
      "14.2.247.5"
        {
        $hashDetail.Add("Rollup",0)
        $hashDetail.Add("VersionOfRollup",0)
        $hashDetail.Add("ReleaseDate","12/4/11")
        $hashDetail.Add("IsExchangePreview",$False)
        }
       
      "14.2.283.3"
        {
        $hashDetail.Add("Rollup",1)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","2/13/12")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.2.298.4"
        {
        $hashDetail.Add("Rollup",2)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","4/16/12")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.2.309.2"
        {
        $hashDetail.Add("Rollup",3)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","5/29/12")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.2.318.2"
        {
        $hashDetail.Add("Rollup",4)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","8/13/12")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.2.318.4"
        {
        $hashDetail.Add("Rollup",4)
        $hashDetail.Add("VersionOfRollup",2)
        $hashDetail.Add("ReleaseDate","10/9/12")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.2.328.5"
        {
        $hashDetail.Add("Rollup",5)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","11/13/12")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.2.328.10"
        {
        $hashDetail.Add("Rollup",5)
        $hashDetail.Add("VersionOfRollup",2)
        $hashDetail.Add("ReleaseDate","12/10/12")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.2.342.3"
        {
        $hashDetail.Add("Rollup",6)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","2/11/13")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.2.375.0"
        {
        $hashDetail.Add("Rollup",7)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","8/13/13")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      #endregion

      #region Exchange 2010 SP3
      "14.3.123.4"
        {
        $hashDetail.Add("Rollup",0)
        $hashDetail.Add("VersionOfRollup",0)
        $hashDetail.Add("ReleaseDate","02/12/13")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.3.146.0"
        {
        $hashDetail.Add("Rollup",1)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","05/29/13")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.3.158.1"
        {
        $hashDetail.Add("Rollup",2)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","08/13/13")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.3.169.1"
        {
        $hashDetail.Add("Rollup",3)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","11/25/13")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.3.174.1"
        {
        $hashDetail.Add("Rollup",4)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","12/10/13")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.3.181.6"
        {
        $hashDetail.Add("Rollup",5)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","02/25/14")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.3.195.1"
        {
        $hashDetail.Add("Rollup",6)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","05/23/14")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.3.210.2"
        {
        $hashDetail.Add("Rollup",7)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","08/26/14")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.3.224.2"
        {
        $hashDetail.Add("Rollup",8)
        $hashDetail.Add("VersionOfRollup",2)
        $hashDetail.Add("ReleaseDate","12/12/14")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.3.235.1"
        {
        $hashDetail.Add("Rollup",9)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","03/17/15")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.3.248.2"
        {
        $hashDetail.Add("Rollup",10)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","06/16/15")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.3.266.1"
        {
        $hashDetail.Add("Rollup",11)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","09/11/15")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.3.279.2"
        {
        $hashDetail.Add("Rollup",12)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","12/10/15")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.3.294.0"
        {
        $hashDetail.Add("Rollup",13)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","03/14/16")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.3.301.0"
        {
        $hashDetail.Add("Rollup",14)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","06/21/16")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.3.319.2"
        {
        $hashDetail.Add("Rollup",15)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","09/21/16")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.3.336.0"
        {
        $hashDetail.Add("Rollup",16)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","12/13/16")
        $hashDetail.Add("IsExchangePreview",$False)
        }

      "14.3.352.0"
        {
        $hashDetail.Add("Rollup",17)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","03/21/17")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      #endregion

      #region Exchange 2013 Preview and RTM
      "15.0.466.13"
        {
        $hashDetail.Add("Rollup",0)
        $hashDetail.Add("VersionOfRollup",0)
        $hashDetail.Add("ReleaseDate","07/16/12")
        $hashDetail.Add("IsExchangePreview",$True)
        }
      
      "15.0.516.32"
        {
        $hashDetail.Add("Rollup",0)
        $hashDetail.Add("VersionOfRollup",0)
        $hashDetail.Add("ReleaseDate","10/11/12")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "15.0.620.29"
        {
        $hashDetail.Add("Rollup",1)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","04/02/13")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "15.0.712.22"
        {
        $hashDetail.Add("Rollup",2)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","07/09/13")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "15.0.712.24"
        {
        $hashDetail.Add("Rollup",2)
        $hashDetail.Add("VersionOfRollup",2)
        $hashDetail.Add("ReleaseDate","07/29/13")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "15.0.712.31"
        {
        $hashDetail.Add("Rollup",2)
        $hashDetail.Add("VersionOfRollup",3)
        $hashDetail.Add("ReleaseDate","07/29/13")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      
      "15.0.775.38"
        {
        $hashDetail.Add("Rollup",3)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","11/25/13")
        $hashDetail.Add("IsExchangePreview",$False)
        }
      #endregion      

      #region Exchange 2013 SP1
      "15.0.847.32"
        {
        $hashDetail.Add("Rollup",4)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","2/24/14")
        $hashDetail.Add("IsExchangePreview",$false)
        }

      "15.0.913.22"
        {
        $hashDetail.Add("Rollup",5)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","5/23/14")
        $hashDetail.Add("IsExchangePreview",$false)
        }

      "15.0.995.29"
        {
        $hashDetail.Add("Rollup",6)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","8/26/14")
        $hashDetail.Add("IsExchangePreview",$false)
        }

      "15.0.1044.25"
        {
        $hashDetail.Add("Rollup",7)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","12/8/14")
        $hashDetail.Add("IsExchangePreview",$false)
        }

      "15.0.1044.29" #There are two versions for CU7, no idea why...
        {
        $hashDetail.Add("Rollup",7)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","12/8/14")
        $hashDetail.Add("IsExchangePreview",$false)
        }

      "15.0.1076.9"
        {
        $hashDetail.Add("Rollup",8)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","03/17/15")
        $hashDetail.Add("IsExchangePreview",$false)
        }

      "15.0.1104.5"
        {
        $hashDetail.Add("Rollup",9)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","06/16/15")
        $hashDetail.Add("IsExchangePreview",$false)
        }

      "15.0.1130.7"
        {
        $hashDetail.Add("Rollup",10)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","09/14/15")
        $hashDetail.Add("IsExchangePreview",$false)
        }

      "15.0.1130.10"
        {
        $hashDetail.Add("Rollup",10)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","09/14/15")
        $hashDetail.Add("IsExchangePreview",$false)
        }

      "15.0.1156.6"
        {
        $hashDetail.Add("Rollup",11)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","12/10/15")
        $hashDetail.Add("IsExchangePreview",$false)
        }

      "15.0.1178.4"
        {
        $hashDetail.Add("Rollup",12)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","03/14/16")
        $hashDetail.Add("IsExchangePreview",$false)
        }

      "15.0.1210.3"
        {
        $hashDetail.Add("Rollup",13)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","06/21/16")
        $hashDetail.Add("IsExchangePreview",$false)
        }

      "15.0.1236.3"
        {
        $hashDetail.Add("Rollup",14)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","09/20/16")
        $hashDetail.Add("IsExchangePreview",$false)
        }

      "15.0.1263.5"
        {
        $hashDetail.Add("Rollup",15)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","12/13/17")
        $hashDetail.Add("IsExchangePreview",$false)
        }

      "15.0.1293.2"
        {
        $hashDetail.Add("Rollup",16)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","03/21/17")
        $hashDetail.Add("IsExchangePreview",$false)
        }
      #endregion

      #region Exchange 2016
      "15.1.225.16"
        {
        $hashDetail.Add("Rollup",0)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","06/22/15")
        $hashDetail.Add("IsExchangePreview",$true)
        }

      "15.1.225.17"
        {
        $hashDetail.Add("Rollup",0)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","06/22/15")
        $hashDetail.Add("IsExchangePreview",$true)
        }

      "15.1.225.42"
        {
        $hashDetail.Add("Rollup",0)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","09/28/15")
        $hashDetail.Add("IsExchangePreview",$false)
        }

      "15.1.396.30"
        {
        $hashDetail.Add("Rollup",1)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","03/14/16")
        $hashDetail.Add("IsExchangePreview",$false)
        }

      "15.1.466.34"
        {
        $hashDetail.Add("Rollup",2)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","06/21/16")
        $hashDetail.Add("IsExchangePreview",$false)
        }

      "15.1.544.27"
        {
        $hashDetail.Add("Rollup",3)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","09/20/16")
        $hashDetail.Add("IsExchangePreview",$false)
        }

      "15.1.669.32"
        {
        $hashDetail.Add("Rollup",4)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","12/13/16")
        $hashDetail.Add("IsExchangePreview",$false)
        }

      "15.1.845.34"
        {
        $hashDetail.Add("Rollup",5)
        $hashDetail.Add("VersionOfRollup",1)
        $hashDetail.Add("ReleaseDate","03/21/17")
        $hashDetail.Add("IsExchangePreview",$false)
        }
      #endregion

      #region Default
      Default
        {
        $hashDetail.Add("Rollup","?")
        $hashDetail.Add("VersionOfRollup","?")
        $hashDetail.Add("ReleaseDate","??/??/??")
        $hashDetail.Add("IsExchangePreview","?")
        $strWarningMsg  = "Could not find details for version number $Version`n"
        $strWarningMsg += "Please check that you have the latest version of this script`n"
        $strWarningMsg += "Http://GetExVerRU.CodePlex.com`n"
        $strWarningMsg += "Script version you are running is: $ScriptVersion`n"
        $strWarningMsg += "If you are running the newest version of this script send an email to jashinba@Microsoft.com`n"
        $strWarningMsg += "Subject: GetExVerRU - Missing Version Number $Version"
        Write-Warning $strWarningMsg
        }
      #endregion
      }
    Return $hashDetail
    }
  }

Process
  {
  $arrExServer += $ExchangeServer
  }

End
 {
  ForEach ($ExSrvName in $arrExServer)
    {
    $ExSrv      = Get-ExchangeServer $ExSrvName
    $strFullVer = Get-FileVersion $ExSrv.Name (Get-ExchangeServerInstallFolder $ExSrv)
    $Props      = New-Object System.Collections.Specialized.OrderedDictionary
    $hVerDetail = Convert-VersionToDetails $strFullVer

    #Add Name
    $Props.Add("Name", $ExSrv.Name)

    #Add Version
    $Props.Add("Version", $strFullVer)

    #Add Major Version
    $Props.Add("MajorVersion", $ExSrv.AdminDisplayVersion.Major)

    #Add Service Pack
    $Props.Add("ServicePack", $ExSrv.AdminDisplayVersion.Minor)

    #Add Rollup
    $Props.Add("Rollup",$hVerDetail.RollUp)

    #Add Version Of Rollup
    $Props.Add("VersionOfRollup",$hVerDetail.VersionOfRollup)

    #Add Release Date
    $Props.Add("ReleaseDate",$hVerDetail.ReleaseDate)

    #Add Is Exchange Preview
    $Props.Add("IsExchangePreview",$hVerDetail.IsExchangePreview)

    New-Object PSObject -Property $Props
    }
  }


