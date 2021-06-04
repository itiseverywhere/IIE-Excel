Function New-CsvFromExcel {
    Param (
        [Parameter(Mandatory=$true)]$SourceFileFullName,
        [Parameter(Mandatory=$false)]$ProcessType,
        [Parameter(Mandatory=$false)]$DestinationDir,
        [Parameter(Mandatory=$false)]$IncludeColumn,
        [Parameter(Mandatory=$false)]$IncludeIf,
        [Parameter(Mandatory=$false)]$IncludeIfNot,
        [Parameter(Mandatory=$false)]$ExcludeColumn,
        [Parameter(Mandatory=$false)]$ExcludeIf,
        [Parameter(Mandatory=$false)]$ExcludeIfNot,
        [Parameter(Mandatory=$false)][switch]$KeepExcelFileOpenForUpdates,
        [Parameter(Mandatory=$false)][switch]$SaveOnlyFirstWorkSheet,
        [Parameter(Mandatory=$false)][switch]$SaveAlsoAsTabSeperated,
        [Parameter(Mandatory=$false)][string]$WorkSheetToSave
    )
    
    # v Clear $Global:GlobalScopeItemsALL
    If ($ProcessType -eq "Scope" -or -not$ProcessType) {$Global:GlobalScopeItemsALL = ""}
    ElseIf ($ProcessType -eq "Merge") {$Global:GlobalMergeALL = ""}
   
    Write-Host "`t======== New-CsvFromExcel ========" -ForegroundColor DarkGray
    If ($Details -or $DetailsAll) {Write-Host "`tSourceFileFullName: $SourceFileFullName" -ForegroundColor Green}

    # v  Check SourceFile    
    If (-not($SourceFileFullName -like "*.xls*" -or $SourceFileFullName -like "*.csv" -or $SourceFileFullName -like "*.tsv" -or $SourceFileFullName -like "*.xlt")) {Write-Warning "No Excel File: SourceFileFullName" ; exit}
    If ($SourceFileFullName -like "*:\*") {
        $UseFileFullName = $SourceFileFullName
    } else {
        $UseFileFullName = "$LogRoot\$SourceFileFullName"
    }
    Write-Host "`tUseFileFullName: ""$UseFileFullName""" -ForegroundColor DarkGray

    $UseFileItem = "" ; $UseFileItem = (Get-Item $UseFileFullName)
    If (-not$UseFileItem) {Write-Host "Please Close Excel, file not found (SourceFileFullName): ""$SourceFileFullName""" -ForegroundColor Yellow ; exit}
    # v Check File Age
    $UseFileItemFullName = $UseFileItem.FullName
    $UseFileItemDirectoryName = $UseFileItem.DirectoryName ; If ($Details -or $DetailsAll) {Write-Host "`tUseFileItemDirectoryName: ""$UseFileItemDirectoryName""" -ForegroundColor DarkGreen}
    $UseFileItemName = $UseFileItem.Name ; If ($Details -or $DetailsAll) {Write-Host "`tUseFileItemName:`t`t""$UseFileItemName""" -ForegroundColor DarkGreen}
    $UseFileItemBaseName = $UseFileItem.BaseName ; If ($Details -or $DetailsAll) {Write-Host "`tWFindFileToCheckBaseName:`t""$UseFileItemBaseName""" -ForegroundColor DarkGreen}

    $UseFileItemLastWriteTime = $UseFileItem.LastWriteTime ; If ($Details -or $DetailsAll) {Write-Host "`tFindFileToCheckLastWriteTime:`t""$UseFileItemLastWriteTime""" -ForegroundColor DarkCyan}
    $checkdaysold = 1 ; $timespan = new-timespan -days $checkdaysold ; if (((Get-Date) - $UseFileItemLastWriteTime) -gt $timespan) {Write-Host "`t""$UseFileItemName"" is older than $checkdaysold days" -ForegroundColor Yellow}
    Else {Write-Host "`t""$UseFileItemName"" is more recent than $checkdaysold days" -ForegroundColor Green}
    $checkhoursold = 4 ; $timespan = new-timespan -Hours $checkhoursold ; if (((Get-Date) - $UseFileItemLastWriteTime) -gt $timespan) {Write-Host "`t""$UseFileItemName"" is older than $checkhoursold hours" -ForegroundColor Yellow}
    Else {Write-Host "`t""$UseFileItemName"" is more recent than $checkhoursold hours" -ForegroundColor Green}
    $checkhoursold = 1 ; $timespan = new-timespan -Hours $checkhoursold ; if (((Get-Date) - $UseFileItemLastWriteTime) -gt $timespan) {Write-Host "`t""$UseFileItemName"" is older than $checkhoursold hours" -ForegroundColor Yellow}
    Else {Write-Host "`t""$UseFileItemName"" is more recent than $checkhoursold hours" -ForegroundColor Green}
    $checkminutessold = 15 ; $timespan = new-timespan -Minutes $checkminutessold ; if (((Get-Date) - $UseFileItemLastWriteTime) -gt $timespan) {Write-Host "`t""$UseFileItemName"" is older than $checkminutessold minutes" -ForegroundColor Yellow}
    Else {Write-Host "`t""$UseFileItemName"" is more recent than $checkminutessold minutes" -ForegroundColor Green}

    If (-not$DestinationDir) {$DestinationDir = $LogRoot}

    # v Process File
    If ($UseFileItemName -like "*.csv") {
        Write-Host "`tImporting CSV: ""$UseFileFullName"""
        $Global:GlobalScopeItemsALL = Import-Csv -Path $UseFileItemFullName -Delimiter ","
    } # If ($SourceFileFullName -like "*.csv")
    ElseIf ($UseFileItemName -like "*.xlt" -or $SourceFileFullName -like "*.tsv") {
        Write-Host "`tImporting TSV: ""$UseFileFullName"""
        $Global:GlobalScopeItemsALL = Import-Csv -Path $UseFileItemFullName -Delimiter "`t"
    } # ElseIf ($SourceFileFullName -like "*.xlt" -or $SourceFileFullName -like "*.tsv")
    ElseIf ($UseFileItemName -like "*.xls*") {

        ####
        # v  OPEN EXCEL FILE (TO SAVE AS CSV)
        $Excel = "" ; $WorkBook = ""
        If (-not$Excel) {
            $Excel = New-Object -ComObject Excel.Application
            If ($Silent) {$Excel.Visible = $False} else {$Excel.Visible = $True}
            If (-not$WorkBook) {
                Write-Host "`t""$UseFileItemName"" (Open XLS to save CSV files)" -ForegroundColor DarkGray
                Try{$WorkBook = $Excel.Workbooks.Open($UseFileItemFullName)} Catch {Write-Warning "Unable to open ""$UseFileItemFullName""" ; exit}
            } # Open Workbook to save as CSV and close afterwards
        } # Open Excel File

        # v Foreach $WorkSheet
        $ArrSavedCSVFileFullNames = @()
        $WorkSheetNames = $WorkBook.Worksheets | Select -Expandproperty name
        
        If ($WorkSheetToSave) {
            $ProcessWorkSheets = $WorkBook.Worksheets | ? Name -eq $WorkSheetToSave   
        } else {
            $ProcessWorkSheets = $WorkBook.Worksheets
        }

        $WorkSheetNames = $ProcessWorkSheets | Select -Expandproperty name
          
        Foreach ($WorkSheet in $ProcessWorkSheets) {

            $WorkSheetname = $WorkSheet.name
            $Iteration = [array]::IndexOf($WorkSheetNames, $WorkSheetname)

            If (-not$SaveOnlyFirstWorkSheet -or ($SaveOnlyFirstWorkSheet -and $Iteration -eq 0)) {

                Write-Host "`t""$UseFileItemName""`tWorksheet #: ""$Iteration""`t$WorkSheetName (processing)" -ForegroundColor DarkCyan

                Try {$WorkSheet.ShowAllData() ; Write-Host "$WorkSheet.ShowAllData()" -ForegroundColor Green} catch {}

                $Global:GlobalExcelWorkSheet = $WorkSheet
                $DestinationFileCSVName = $UseFileItemBaseName + " " + $WorkSheet.name + ".csv" ; $DestinationFileCSVFullName = "$DestinationDir\$DestinationFileCSVName"
                If (Test-Path $DestinationFileCSVFullName) {
                    Write-Host "`t""$UseFileItemName""`tWorksheet #: ""$Iteration""`t$WorkSheetName (overwrite CSV ""$DestinationFileCSVName"")" -ForegroundColor DarkCyan
                    Remove-Item $DestinationFileCSVFullName -ErrorAction SilentlyContinue -Force | Out-Null
                } else {
                    Write-Host "`t""$UseFileItemName""`tWorksheet #: ""$Iteration""`t$WorkSheetName (Save CSV ""$DestinationFileCSVName"")" -ForegroundColor DarkCyan
                } 
                $WorkSheet.SaveAs($DestinationFileCSVFullName, 6) # SAVE THE WORKSHEET TO CSV
            
                Write-Host "`t""$UseFileItemName""`tWorksheet #: ""$Iteration""`t$WorkSheetName (determine headers)" -ForegroundColor DarkCyan
                $Global:GlobalScopeHeaders = (Get-Content $DestinationFileCSVFullName)[0] -split ';' -split ","
            
                $ArrSavedCSVFileFullNames += $DestinationFileCSVFullName

                If ($SaveOnlyFirstWorkSheet) {
                    $WorkBook.Saved = $true
                    Write-Host "`t""$UseFileItemName""`tWorksheet #: ""$Iteration""`t$WorkSheetName (SaveOnlyFirstWorkSheet, close excel)"
                    #Try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)} catch {}
                    Try {$LastExcelFileStartedId = ((get-process excel | select MainWindowTitle, ID, StartTime | Sort StartTime)[-1]).Id ; Stop-Process -Id $LastExcelFileStartedId} Catch {}
                    If ($LastExcelFileStartedId) {
                        Write-Host "`t""$UseFileItemName"" (close excel) ID: ""$LastExcelFileStartedId""" -ForegroundColor Yellow
                        Stop-Process -Id $LastExcelFileStartedId
                    }
                    Write-Host "`t""$UseFileItemName"" (closed excel)" -ForegroundColor DarkCyan
                } # Do not save as Tab Seperated, return with the $DestinationFileCSVFullName
                Write-Host "`t""$UseFileItemName""`tWorksheet #: ""$Iteration""`t$WorkSheetName (import CSV)" -ForegroundColor DarkCyan
                $Global:GlobalScopeItemsALL = Import-Csv -Path $DestinationFileCSVFullName -Delimiter ","

            } else {
                Write-Host "`t""$UseFileItemName""`tWorksheet #: ""$Iteration""`t$WorkSheetName (skipped)" -ForegroundColor DarkGray
            } # Determine how many worksheets to process

        } # foreach ($WorkSheet in $WorkBook.Worksheets)
        # ^ Foreach $WorkSheet
        
        <#
        Write-Host "`t""$UseFileItemName"" (close excel)" -ForegroundColor DarkCyan
        $WorkBook.Saved = $true
        Try{$LastExcelFileStartedId = ((get-process excel | select MainWindowTitle, ID, StartTime | Sort StartTime)[-1]).Id} Catch {}
        If ($LastExcelFileStartedId) {
            Write-Host "`t""$UseFileItemName"" (close excel) ID: ""$LastExcelFileStartedId""" -ForegroundColor Yellow
            Stop-Process -Id $LastExcelFileStartedId
        }
        Write-Host "`t""$UseFileItemName"" (closed excel)" -ForegroundColor DarkCyan
        #>

        # ^ OPEN EXCEL FILE (TO SAVE AS CSV)
        ####

        If ($SaveAlsoAsTabSeperated) {
            Start-Sleep 1
            Foreach ($SavedCSVFileFullName in $ArrSavedCSVFileFullNames) {

                $CSVFileToCheck = "" ; $CSVFileToCheck = (Get-Item $SavedCSVFileFullName)
                
                If ($CSVFileToCheck) {
                    $CSVFileToCheckFullName = $CSVFileToCheck.FullName
                    $CSVFileToCheckBaseName = $CSVFileToCheck.BaseName
                    $CSVFileToCheckName = $CSVFileToCheck.Name
                    $DestinationFileXLTName = $CSVFileToCheckBaseName + ".txt" ; $DestinationFileXLTFullName = "$DestinationDir\$DestinationFileXLTName"
                    Write-Host "`tOpen CSV:`t`t`t""$CSVFileToCheckName""" -ForegroundColor DarkCyan
                    If (Test-Path $DestinationFileXLTFullName) {
                        Write-Host "`tSave XLT (overwrite):`t""$DestinationFileXLTName""" -ForegroundColor DarkCyan
                        Remove-Item $DestinationFileXLTFullName -ErrorAction SilentlyContinue -Force | Out-Null
                    } else {
                        Write-Host "`tSave XLT:`t`t`t""$DestinationFileXLTName""" -ForegroundColor DarkCyan
                    }
                    Import-Csv -Path $CSVFileToCheckFullName -Delimiter "," | Export-Csv -Path $DestinationFileXLTFullName -Delimiter "`t" -NoTypeInformation -ErrorAction Stop -WarningAction Stop
                } # $CSVFileToCheck
                else {
                    Write-Host "CSV file not found (CSVFileToCheckFullName): ""$CSVFileToCheckFullName""" -ForegroundColor Yellow
                } # -not$CSVFileToCheck

                } # Foreach SavedCSVFileFullName
        } # If ($SaveAlsoAsTabSeperated)

        
        ####
        # v KeepExcelFileOpenForUpdates
        If ($KeepExcelFileOpenForUpdates -and $Global:GlobalScopeItemsALL) {
            # v Open Excel File (to keep open for editing)
            $Excel = "" ; $WorkBook = ""
            If (-not$Excel) {
                Write-Host "`tNew-Object -ComObject Excel.Application (to keep open for editing)" -ForegroundColor Green
                $Excel = New-Object -ComObject Excel.Application
                If ($Silent) {$Excel.Visible = $False} else {$Excel.Visible = $True}
                If (-not$WorkBook) {
                    Write-Host "`tWorkbooks.Open (to keep open for editing)" -ForegroundColor Green
                    Try{$WorkBook = $Excel.Workbooks.Open($UseFileItemFullName)} Catch {Write-Warning "Unable to open ""$UseFileItemFullName""" ; exit}
                }
            }            
            # v Foreach $WorkSheet
            Foreach ($WorkSheet in $WorkBook.Worksheets) {
                
                Try {$WorkSheet.ShowAllData() ; Write-Host "$WorkSheet.ShowAllData()" -ForegroundColor Green} catch {}
    
                $Global:GlobalExcelWorkSheet = $WorkSheet
                $Global:GlobalExcelWorkSheet.ShowAllData()

                $WorkSheetname = $WorkSheet.name
                Write-Host "`tWorkSheet open for editing:`t$WorkSheetname" -ForegroundColor Green
                If ($SaveOnlyFirstWorkSheet) {
                    $WorkBook.Saved = $true
                    Write-Host "`tWorkSheet open for editing (saved):`t$WorkSheetname" -ForegroundColor Green
                    $Excel.ActiveWindow.WindowState = -4140 # https://docs.microsoft.com/en-us/office/vba/api/excel.application.windowstate
                    break
                }
            } # foreach ($WorkSheet in $WorkBook.Worksheets) 
            # ^ Foreach $WorkSheet

        } # KeepExcelFileOpenForUpdates
        # ^ KeepExcelFileOpenForUpdates
        ####

    } # ElseIf ($SourceFileFullName -like "*.xls*")
    Else {
        Write-Warning "Unknown File Type: ""$UseFileItemName"""
    } # Unknown SourceFile

    # $Global:GlobalScopeItemsALLColumns = $Global:GlobalScopeItemsALL[0].psobject.Properties | foreach { $_.Name } #This is no longer needed because we search for the right Column in Excel directly

    If (($ProcessType -eq "Scope" -or -not$ProcessType) -and $Global:GlobalScopeItemsALL) {

        $Global:GlobalScopeItemsALLUpdated = $Global:GlobalScopeItemsALL

        $Global:GlobalScopeItemsALLCount = ($Global:GlobalScopeItemsALL | Measure-Object).count
        
        #####
        # v LIMIT SCOPE
        $ReportColumns = @("Name","DisplayName",$EmailAddressColumn,$OldEmailAddressColumn)
        If ($MailForwardEnableColumn) {$ReportColumns += $MailForwardEnableColumn}
        If ($InboxRuleRedirectEnableColumn) {$ReportColumns += $InboxRuleRedirectEnableColumn}
        $Global:GlobalScopeItemsScope = $Global:GlobalScopeItemsALL ; $Global:GlobalScopeItemsScopeCount = ($Global:GlobalScopeItemsScope | Measure-Object).Count
        
        ####
        # v Scope Based on -Query parameter
        If ($Query) {

            $QueryLength = $Query.Length
            If ($Query.substring(0,1) -ne "{") {$Query = '{' + $Query}
            If ($Query.substring($QueryLength,1) -ne "}") {$Query = $Query + '}'}
            
            $ThisScopeCommand = '$Global:GlobalScopeItemsALL' + ' | ? ' + $Query
            Write-host $ThisScopeCommand
            $ThisScope = Invoke-Expression $ThisScopeCommand ; $ThisScopeCount = ($ThisScope | Measure-Object).Count
            $AllScopeCommand = '$Global:GlobalScopeItemsScope' + ' | ? ' + $Query
            Write-host $AllScopeCommand
            $Global:GlobalScopeItemsScope = Invoke-Expression $AllScopeCommand ; $Global:GlobalScopeItemsScopeCount = ($Global:GlobalScopeItemsScope | Measure-Object).Count
            If ($ThisScopeCount -eq 0 -or $Global:GlobalScopeItemsScopeCount -eq 0) {$ResultColor = "Yellow"} else {$ResultColor = "Cyan"}
            Write-Host "-Query`t`t$ThisScopeCount`t(most restrictive $Global:GlobalScopeItemsScopeCount)`t$Query" -ForegroundColor $ResultColor

        } # If ($Query)
        # ^ Scope Based on -Query parameter
        ####

        If ($IncludeColumn -or $ExcludeColumn) {$TotalQuery = '{'} # Prepare TotalQuery

        ####            
        # v Update scope based on $IncludeColumn
        If ($IncludeColumn -and ($IncludeIf -or $IncludeIfNot)) {
            $ReportColumns += $IncludeColumn
            If (-not($Global:GlobalScopeItemsALL | ? $IncludeColumn)) {Write-Host "`tNotFound column ""$IncludeColumn"" in $UseFileItemFullName" -ForegroundColor Yellow ; exit}
            $ScopeItemsIncludeColumnScope = $Global:GlobalScopeItemsALL
            If (-not($Global:GlobalScopeItemsALL | ? $IncludeColumn)) {Write-Host "No`tIncludeColumn`t(""$IncludeColumn"") in $ScopeFileToUseFullName" -ForegroundColor Yellow ; Stop-Transcript ; continue} #else {Write-Host "`r`nFound`tIncludeColumn:`t(""$IncludeColumn"")" -ForegroundColor Cyan}
            
            If ($IncludeIf) {
                $Operator = "-OR" ; $ThisQuery = '{'
                foreach ($Item in $IncludeIf) {
                    If ($Item.contains('*') -or $Item.contains('?')) {$Comparison = "-LIKE"} else {$Comparison = "-EQ"}
                    If ([array]::IndexOf($IncludeIf, $Item) -eq 0) {
                        $ThisQuery += '$_.' + $IncludeColumn + " " + $comparison + " ""$Item"""
                    } else {
                        $ThisQuery += ' ' + $Operator + ' $_.' + $IncludeColumn + " " + $Comparison + " ""$Item"""
                    }
                    If ($TotalQuery -eq '{') {$TotalQuery += '$_.' + $IncludeColumn + " " + $comparison + " ""$Item"""} else {$TotalQuery += ' ' + $Operator + ' $_.' + $IncludeColumn + " " + $Comparison + " ""$Item"""} # Append TotalQuery
                } # foreach $Item
                $ThisQuery += '}'
                $ThisScopeCommand = '$Global:GlobalScopeItemsALL' + ' | ? ' + $ThisQuery ; $ThisScope = Invoke-Expression $ThisScopeCommand ; $ThisScopeCOunt = ($ThisScope | Measure-Object).Count
                $AllScopeCommand = '$Global:GlobalScopeItemsScope' + ' | ? ' + $ThisQuery ; $Global:GlobalScopeItemsScope = Invoke-Expression $AllScopeCommand ; $Global:GlobalScopeItemsScopeCount = ($Global:GlobalScopeItemsScope | Measure-Object).Count
                If ($ThisScopeCount -eq 0 -or $Global:GlobalScopeItemsScopeCount -eq 0) {$ResultColor = "Yellow"} else {$ResultColor = "Cyan"}
                Write-Host "IncludeIf`t$ThisScopeCount`t(most restrictive $Global:GlobalScopeItemsScopeCount)`t$ThisQuery" -ForegroundColor $ResultColor
            } # If ($IncludeIfUSed)

            If ($IncludeIfNot) {
                $Operator = "-AND" ; $ThisQuery = '{'
                foreach ($Item in $IncludeIfNot) {
                    If ($Item.contains('*') -or $Item.contains('?')) {$Comparison = "-NOTLIKE"} else {$Comparison = "-NE"}
                    If ([array]::IndexOf($IncludeIfNot, $Item) -eq 0) {
                        $ThisQuery += '$_.' + $IncludeColumn + " " + $comparison + " ""$Item"""
                    } else {
                        $ThisQuery += ' ' + $Operator + ' $_.' + $IncludeColumn + " " + $Comparison + " ""$Item"""
                    }
                    If ($TotalQuery -eq '{') {$TotalQuery += '$_.' + $IncludeColumn + " " + $comparison + " ""$Item"""} else {$TotalQuery += ' ' + $Operator + ' $_.' + $IncludeColumn + " " + $Comparison + " ""$Item"""} # Append TotalQuery
                } # foreach $Item
                $ThisQuery += '}'
                $ThisScopeCommand = '$Global:GlobalScopeItemsALL' + ' | ? ' + $ThisQuery ; $ThisScope = Invoke-Expression $ThisScopeCommand ; $ThisScopeCOunt = ($ThisScope | Measure-Object).Count
                $AllScopeCommand = '$Global:GlobalScopeItemsScope' + ' | ? ' + $ThisQuery ; $Global:GlobalScopeItemsScope = Invoke-Expression $AllScopeCommand ; $Global:GlobalScopeItemsScopeCount = ($Global:GlobalScopeItemsScope | Measure-Object).Count
                If ($ThisScopeCount -eq 0 -or $Global:GlobalScopeItemsScopeCount -eq 0) {$ResultColor = "Yellow"} else {$ResultColor = "Cyan"}
                Write-Host "IncludeIfNot`t$ThisScopeCount`t(most restrictive $Global:GlobalScopeItemsScopeCount)`t$ThisQuery" -ForegroundColor $ResultColor
            } # If ($IncludeIfNotUsed)

            $ScopeItemsIncludeColumnCount = $ThisScopeCount
        } # Update scope based on $IncludeColumn
        # ^ Update scope based on $IncludeColumn
        ####

        ####
        # v Update scope based on $ExcludeColumn
        If ($ExcludeColumn -and ($ExcludeIf -or $ExcludeIfNot)) {
            $ReportColumns += $ExcludeColumn
            If (-not($Global:GlobalScopeItemsALL | ? $ExcludeColumn)) {Write-Host "`tNotFound column ""$ExcludeColumn"" in $UseFileItemFullName" -ForegroundColor Yellow ; exit}
            $ScopeItemsExcludeColumnScope = $Global:GlobalScopeItemsALL
            If (-not($Global:GlobalScopeItemsALL | ? $ExcludeColumn)) {Write-Host "No`tExcludeColumn`t(""$ExcludeColumn"") in $ScopeFileToUseFullName" -ForegroundColor Yellow ; Stop-Transcript ; continue} #else {Write-Host "`r`nFound`tExcludeColumn:`t(""$ExcludeColumn"")" -ForegroundColor Cyan}
            
            If ($ExcludeIf) {
                $Operator = "-AND" ; $ThisQuery = '{'
                foreach ($Item in $ExcludeIf) {
                    If ($Item.contains('*') -or $Item.contains('?')) {$Comparison = "-NOTLIKE"} else {$Comparison = "-NE"}
                    If ([array]::IndexOf($ExcludeIf, $Item) -eq 0) {
                        $ThisQuery += '$_.' + $ExcludeColumn + " " + $comparison + " ""$Item"""
                    } else {
                        $ThisQuery += ' ' + $Operator + ' $_.' + $ExcludeColumn + " " + $Comparison + " ""$Item"""
                    }
                    If ($TotalQuery -eq '{') {$TotalQuery += '$_.' + $IncludeColumn + " " + $comparison + " ""$Item"""} else {$TotalQuery += ' ' + $Operator + ' $_.' + $IncludeColumn + " " + $Comparison + " ""$Item"""} # Append TotalQuery
                } # foreach $Item
                $ThisQuery += '}'
                $ThisScopeCommand = '$Global:GlobalScopeItemsALL' + ' | ? ' + $ThisQuery ; $ThisScope = Invoke-Expression $ThisScopeCommand ; $ThisScopeCOunt = ($ThisScope | Measure-Object).Count
                $AllScopeCommand = '$Global:GlobalScopeItemsScope' + ' | ? ' + $ThisQuery ; $Global:GlobalScopeItemsScope = Invoke-Expression $AllScopeCommand ; $Global:GlobalScopeItemsScopeCount = ($Global:GlobalScopeItemsScope | Measure-Object).Count
                If ($ThisScopeCount -eq 0 -or $Global:GlobalScopeItemsScopeCount -eq 0) {$ResultColor = "Yellow"} else {$ResultColor = "Cyan"}
                Write-Host "ExcludeIf`t$ThisScopeCount`t(most restrictive $Global:GlobalScopeItemsScopeCount)`t$ThisQuery" -ForegroundColor $ResultColor      
            } # If ($ExcludeIfUsed)

            If ($ExcludeIfNot) {
                $Operator = "-OR" ; $ThisQuery = '{'
                foreach ($Item in $ExcludeIfNot) {
                    If ($Item.contains('*') -or $Item.contains('?')) {$Comparison = "-LIKE"} else {$Comparison = "-EQ"}
                    If ([array]::IndexOf($ExcludeIfNot, $Item) -eq 0) {
                        $ThisQuery += '$_.' + $ExcludeColumn + " " + $comparison + " ""$Item"""
                    } else {
                        $ThisQuery += ' ' + $Operator + ' $_.' + $ExcludeColumn + " " + $Comparison + " ""$Item"""
                    }
                    If ($TotalQuery -eq '{') {$TotalQuery += '$_.' + $IncludeColumn + " " + $comparison + " ""$Item"""} else {$TotalQuery += ' ' + $Operator + ' $_.' + $IncludeColumn + " " + $Comparison + " ""$Item"""} # Append TotalQuery
                } # foreach $Item
                $ThisQuery += '}'
                $ThisScopeCommand = '$Global:GlobalScopeItemsALL' + ' | ? ' + $ThisQuery ; $ThisScope = Invoke-Expression $ThisScopeCommand ; $ThisScopeCOunt = ($ThisScope | Measure-Object).Count
                $AllScopeCommand = '$Global:GlobalScopeItemsScope' + ' | ? ' + $ThisQuery ; $Global:GlobalScopeItemsScope = Invoke-Expression $AllScopeCommand ; $Global:GlobalScopeItemsScopeCount = ($Global:GlobalScopeItemsScope | Measure-Object).Count
                If ($ThisScopeCount -eq 0 -or $Global:GlobalScopeItemsScopeCount -eq 0) {$ResultColor = "Yellow"} else {$ResultColor = "Cyan"}
                Write-Host "ExcludeIfNot`t$ThisScopeCount`t(most restrictive $Global:GlobalScopeItemsScopeCount)`t$ThisQuery" -ForegroundColor $ResultColor      
            } # If ($ExcludeIfNotUsed)
            $ScopeItemsExcludeColumnCount = $ThisScopeCount
        } # Update scope based on $ExcludeColumn
        # ^ Update scope based on $ExcludeColumn
        
        If ($ExcludeColumn -or $ExcludeColumn) {$TotalQuery += '}'} # Close TotalQuery
    
        # v Show Scope Results
        $ReportColumns = $ReportColumns | Sort-Object -Unique
        If ($DetailsAll -or $ListScopeOnly) {
            Write-Host "-ListScope`t$TotalQuery" -ForegroundColor Yellow
            $Global:GlobalScopeItemsScope | Sort $EmailAddressColumn | Select-Object $ReportColumns | ft
        }

        If ($Global:GlobalScopeItemsALLCount -eq $Global:GlobalScopeItemsScopeCount) {
            If (($Commit -or $CommitWithoutAsking) -and -not$WhatIf) {
                Write-Host "`tScope ($Global:GlobalScopeItemsScopeCount) equals all $Global:GlobalScopeItemsALLCount items in ""$UseFileItemName"", -Commit switch is used!" -ForegroundColor Yellow
            } else {
                Write-Host "`tScope ($Global:GlobalScopeItemsScopeCount) equals all $Global:GlobalScopeItemsALLCount items in ""$UseFileItemName""" -ForegroundColor DarkCyan
            }
        } # If ($Global:GlobalScopeItemsScopeCount -eq $Global:GlobalScopeItemsALLCount)

        If ($IncludeColumn -and $ScopeItemsIncludeColumnCount -eq 0) {
            Write-Host "`r`n#########`r`n# v IncludeColumn`t""$IncludeColumn"" value ""$IncludeIf"" not found, available -IncludeColumn values:" -ForegroundColor Yellow
            #Write-Host "`t# Include`t""$IncludeColumn""`t[$IncludeIf]" -ForegroundColor Yellow
            $Global:GlobalScopeItemsALL | Select-Object $IncludeColumn | Sort-Object $IncludeColumn | Get-Unique -AsString | ft
            Write-Host "# ^ IncludeColumn`t""$IncludeColumn"" values`r`n#########`r`n" -ForegroundColor Yellow
        } # If ($ScopeItemsIncludeColumnCount -eq 0)
        If ($ExcludeColumn -and $ScopeItemsExcludeColumnCount -eq 0) {
            Write-Host "`r`n#########`r`n# v ExcludeColumn`t""$ExcludeColumn"" value ""$ExcludeIf"" not found, available -ExcludeColumn values:" -ForegroundColor Yellow
            #Write-Host "`t# Exclude`t""$ExcludeColumn""`t[$ExcludeIf]" -ForegroundColor Yellow
            $Global:GlobalScopeItemsALL | Select-Object $ExcludeColumn | Sort-Object $ExcludeColumn | Get-Unique -AsString | ft
            Write-Host "# ^ ExcludeColumn`t""$ExcludeColumn""`t values`r`n#########`r`n" -ForegroundColor Yellow
        } # If ($ScopeItemsExcludeColumnCount -eq 0) {
               
        #Write-Host "`tGlobal:GlobalScopeItemsALLCount: ""$Global:GlobalScopeItemsALLCount""" -ForegroundColor DarkGray
        #Write-Host "`t================`r`n" -ForegroundColor DarkGray
        If ($Global:GlobalScopeItemsALLCount -eq 0) {continue}
        # ^ Show Scope Results

        # ^ LIMIT SCOPE
        #####

    } # $ProcessType -eq "Scope"

} # New-CsvFromExcel