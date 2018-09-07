cd "P:\ANG_System_Files"
Add-Type -AssemblyName System.Windows.Forms

function Load-Dll
{
    param(
        [string]$assembly
    )
    Write-Host "Loading $assembly"

    $driver = $assembly
    $fileStream = ([System.IO.FileInfo] (Get-Item $driver)).OpenRead();
    $assemblyBytes = new-object byte[] $fileStream.Length
    $fileStream.Read($assemblyBytes, 0, $fileStream.Length) | Out-Null;
    $fileStream.Close();
    $assemblyLoaded = [System.Reflection.Assembly]::Load($assemblyBytes);
}

function Save-AttachmentToSheetRow
{
    param(
        [long]$sheetId,
        [long]$rowId,
        [System.IO.FileInfo]$file,
        [string]$mimeType
    )

    $result = $client.SheetResources.RowResources.AttachmentResources.AttachFile($sheetId, $rowId, $file.FullName, $mimeType)

    return $result
}

Load-Dll -assembly P:\ANG_System_Files\ItextDLLs\itextsharp.dll
Load-Dll -assembly P:\ANG_System_Files\ItextDLLs\itextsharp.pdfa.dll
Load-Dll -assembly P:\ANG_System_Files\ItextDLLs\itextsharp.xmlworker.dll
Load-Dll -assembly P:\ANG_System_Files\ItextDLLs\itextsharp.xtra.dll

Load-Dll ".\smartsheet-csharp-sdk.dll"                     
Load-Dll ".\RestSharp.dll"
Load-Dll ".\Newtonsoft.Json.dll"
Load-Dll ".\NLog.dll"

function Get-CellObjects
{
    param([Smartsheet.Api.Models.Sheet]$sheet)

    Write-Host "Getting Sheet $($sheet.Name) Comparison Objects"

    $data = $sheet.Rows | foreach {
        $checkVal1 = $false
        $checkVal2 = $false

        if($_.Cells[4].Value -eq $true)
        {
            $checkVal1 = $true
        } 

        if($_.Cells[46].Value -eq $true)
        {
            $checkVal2 = $true
        }

        [pscustomobject]@{
            Attachments = $_.Attachments;
            RowId = $_.Id;
            RowNumber = $_.RowNumber;
            JobNumCol = $_.Cells[0].ColumnId;
            JobNum = $_.Cells[0].Value;
            CompletedByCol = $_.Cells[1].ColumnId;
            CompletedBy = $_.Cells[1].Value;
            DateComplByCol = $_.Cells[2].ColumnId;
            DateComplBy = $_.Cells[2].Value;
            PercentCol = $_.Cells[3].ColumnId;
            Percent = $_.Cells[3].Value;
            PaperWorkOfficeCol = $_.Cells[4].ColumnId;
            PaperWorkOffice = $checkVal1;
            JobNameCol = $_.Cells[5].ColumnId;
            JobName = $_.Cells[5].Value;
            EstAmmountCol = $_.Cells[6].ColumnId;
            EstAmmount = $_.Cells[6].Value;
            EstExpenseCol = $_.Cells[7].ColumnId;
            EstExpense = $_.Cells[7].Value;
            ProjCityCol = $_.Cells[9].ColumnId;
            ProjCity = $_.Cells[9].Value;
            ProjStateCol = $_.Cells[10].ColumnId;
            ProjState = $_.Cells[10].Value;
            ProjCountryCol = $_.Cells[12].ColumnId;
            ProjCountry = $_.Cells[12].Value;
            EstStartCol = $_.Cells[13].ColumnId;
            EstStart = $_.Cells[13].Value;
            EstEndCol = $_.Cells[14].ColumnId;
            EstEnd = $_.Cells[14].Value;
            ProjManCol = $_.Cells[15].ColumnId;
            ProjMan = $_.Cells[15].Value;
            RequesterCol = $_.Cells[16].ColumnId;
            Requester = $_.Cells[16].Value;
            ReqDateCol = $_.Cells[17].ColumnId;
            ReqDate = $_.Cells[17].Value;
            GCCol = $_.Cells[18].ColumnId;
            GC = $_.Cells[18].Value;
            GCContactCol = $_.Cells[19].ColumnId;
            GCContact = $_.Cells[19].Value;
            GCContactPhoneCol = $_.Cells[20].ColumnId;
            GCContactPhone = $_.Cells[20].Value;
            GCContactEmailCol = $_.Cells[21].ColumnId;
            GCContactEmail = $_.Cells[21].Value;
            TypeWorkCol = $_.Cells[22].ColumnId;
            TypeWork = $_.Cells[22].Value;
            NameCharCol = $_.Cells[39].ColumnId;
            NameChar = $_.Cells[39].Value;
            ContactNameCol = $_.Cells[40].ColumnId;
            ContactName = $_.Cells[40].Value;
            CharAddressCol = $_.Cells[41].ColumnId;
            CharAddress = $_.Cells[41].Value;
            CharTaxIDCol = $_.Cells[42].ColumnId;
            CharTaxID = $_.Cells[42].Value;
            CharPhoneNumCol = $_.Cells[43].ColumnId;
            CharPhoneNum = $_.Cells[43].Value;
            ModifiedCol = $_.Cells[44].ColumnId;
            Modified = $_.Cells[44].Value;
            FileStructureCol = $_.Cells[46].ColumnId;
            FileStructure = $checkVal2;
        }                                                  
    } | where {![string]::IsNullOrWhiteSpace($_.JobName)} 

    Write-Host "$($data.Count) Returned"      
    return $data                                           
}   

function Send-Notification 
{
    param (
        [string]$emailTo,
        [string]$subject,
        [string]$body,
        [string]$emailFrom = "Alerts@allnewglass.com",
        [string]$attachment,
        [string]$password
    )

    $username = $emailfrom
    $smtpserver = "smtp.office365.com" 
    $smtpmessage = New-Object System.Net.Mail.MailMessage($emailfrom,$emailto,$subject,$body)

    if (![string]::IsNullOrWhiteSpace($attachment))
    {
        $smtpattachment = New-Object System.Net.Mail.Attachment($attachment)
        $smtpmessage.Attachments.Add($smtpattachment)
    }

    $smtpclient = New-Object Net.Mail.SmtpClient($SmtpServer, 587) 
    $smtpclient.EnableSsl = $true 
    $smtpclient.Credentials = New-Object System.Net.NetworkCredential($username, $password); 
    $smtpclient.Send($smtpmessage)

    Remove-Variable -Name smtpclient
    Remove-Variable -Name password
} 

function Get-ContactsFromDirectory
{
    $dirId = "6255005150799748"
    $dirSheet = $client.SheetResources.GetSheet($dirId, $includes, $null, $null, $null, $null, $null, $null);

    $contacts = @();

    foreach($row in $dirSheet.Rows | select -Skip 1)
    {
        $contact = [PSCustomObject]@{
            First = $row.Cells[0].Value;
            Last = $row.Cells[1].Value;
            Location = $row.Cells[2].Value;
            Position = $row.Cells[3].Value;
            Ext = $row.Cells[4].Value;
            Phone = $row.Cells[5].Value;
            Email = $row.Cells[6].Value;
        }

        $contacts += $contact
    }

    return $contacts
}

$token      = "e41266qmwuasa15w9rwe5321ob"
$smartsheet = [Smartsheet.Api.SmartSheetBuilder]::new()
$builder    = $smartsheet.SetAccessToken($token)
$client     = $builder.Build()
$includes   =  @([Smartsheet.Api.Models.SheetLevelInclusion]::ATTACHMENTS)
$includes   = [System.Collections.Generic.List[Smartsheet.Api.Models.SheetLevelInclusion]]$includes
$jobDataId = "1549079680444292"
$JobData  = $client.SheetResources.GetSheet($jobDataId, $includes, $null, $null, $null, $null, $null, $null);

$jobDataCOs = Get-CellObjects $JobData
$contacts = Get-ContactsFromDirectory

foreach ($jobDataCO in $jobDataCOs)
{
    if(![string]::IsNullOrWhiteSpace($jobDataCO.JobNum) -and ![string]::IsNullOrWhiteSpace($jobDataCO.JobName) -and ($jobDataCO.PaperWorkOffice -eq $false))
    {
        if([string]::IsNullOrWhiteSpace($jobDataCO.JobName) -or [string]::IsNullOrWhiteSpace($jobDataCO.EstAmmount) -or [string]::IsNullOrWhiteSpace($jobDataCO.EstExpense) -or 
           [string]::IsNullOrWhiteSpace($jobDataCO.ProjCity) -or [string]::IsNullOrWhiteSpace($jobDataCO.ProjState) -or [string]::IsNullOrWhiteSpace($jobDataCO.ProjCountry) -or 
           [string]::IsNullOrWhiteSpace($jobDataCO.EstStart) -or [string]::IsNullOrWhiteSpace($jobDataCO.EstEnd) -or [string]::IsNullOrWhiteSpace($jobDataCO.ProjMan) -or 
           [string]::IsNullOrWhiteSpace($jobDataCO.Requester) -or [string]::IsNullOrWhiteSpace($jobDataCO.ReqDate) -or [string]::IsNullOrWhiteSpace($jobDataCO.GC) -or 
           [string]::IsNullOrWhiteSpace($jobDataCO.GCContact) -or [string]::IsNullOrWhiteSpace($jobDataCO.GCContactPhone) -or [string]::IsNullOrWhiteSpace($jobDataCO.GCContactEmail) -or 
           [string]::IsNullOrWhiteSpace($jobDataCO.TypeWork))
        {
            $form1 = New-Object System.Windows.Forms.Form  
            $form1.Text = 'All New Glass'
            $form1.Size = [System.Drawing.Size]::new(650,125)
            $form1.StartPosition = 'CenterScreen'
            
            $OKButton = New-Object System.Windows.Forms.Button
            $OKButton.Location = [System.Drawing.Point]::new(275,60)
            $OKButton.Size = [System.Drawing.Size]::new(75,23)
            $OKButton.Text = 'OK'
            $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form1.AcceptButton = $OKButton
            $form1.Controls.Add($OKButton)
            
            $label1 = New-Object System.Windows.Forms.Label
            $label1.Location = [System.Drawing.Point]::new(10,20)
            $label1.Size = [System.Drawing.Size]::new(650,20)
            $label1.Text = "$($jobDataCO.Requester) has left required cells blank. They will now be notified via text and email about completing the line."
            $form1.Controls.Add($label1)
            
            $form1.Topmost = $true
            
            $result1 = $form1.ShowDialog()  

            $dataBaseLink = "https://app.smartsheet.com/b/home?lx=eNtoesd7A6eqcIAMsROIUA"

            $emailBody = "You have failed to completely fill out all of the required cells in the ANG Job Database and therefore your job number request has been denied at this time.`n`nGoing forward you must fill out ALL of the cells in the light blue section for approval.  All of the cells between the columns marked Job Name and Type of Work are necessary for the job number request paperwork. If it is a charity job then you must also fill out the tan colored charity section towards the end of the sheet.  `n`nPlease return to the database with the link below and fill out the remainder of the required cells.`n`n$dataBaseLink `n`nOnce you have completed filling out the required cells inform wheitmann@allnewglass.com as soon as possible."

            Send-Notification -emailTo $($jobDataCO.Requester) -subject "JOB NUMBER REQUEST DENIED"  -body $emailBody -emailFrom "Alerts@allnewglass.com" -password "Allnew123"

            foreach ($contact in $contacts)
            {
                if ($jobDataCO.Requester -eq $contact.Email)
                {
                    $carriers        = @{
                        Alltel       = "message.alltel.com"
                        ATT          = "txt.att.net"
                        BoostMobile  = "myboostmobile.com"
                        MetroPCS     = "mymetropcs.com"
                        Nextel       = "messaging.nextel.com"
                        SprintPCS    = "messaging.sprintpcs.com"
                        TMobile      = "tmomail.net"
                        Verizon      = "vtext.com"
                        VirginMobile = "vmobl.com"
                    }

                    foreach($carrier in $carriers.Values)
                    { 
                        $phoneNumber = $contact.Phone -replace "[^0-9]", ''
                        Send-Notification -emailTo "$phoneNumber@$carrier" -subject "JOB NUMBER REQUEST DENIED" -body "Job number request incomplete.  Check email for details." -emailFrom "Alerts@allnewglass.com" -password "Allnew123"
                    }
                }
            }
        }

        else
        {
            $JobDataPMCol = $JobData.Columns | where {$_.Title -eq ("Paperwork Filled Out PM")}
            $JobDataOfficeCol = $JobData.Columns | where {$_.Title -eq ("Paperwork Filled Out Office")}
            $PercentCol = $JobData.Columns | where {$_.Title -eq ("Estimated Margin (Office Use Only)")}
            $PathCol = $JobData.Columns | where {$_.Title -eq ("Po Log File Path")}
            $FileStructCol = $JobData.Columns | where {$_.Title -eq ("File Structure Created")}
            $x = $jobDataCO.EstAmmount
            $y = $jobDataCO.EstExpense 
            $percent = (($x - $y)/ $x).ToString("P")
            $ammountCurrency = ($jobDataCO.EstAmmount).ToString("C")
            $expenseCurrency = ($jobDataCO.EstExpense).ToString("C")
            $estimatedStart = ([Nullable[DateTime]]$jobDataCO.EstStart).ToString('MM/dd/yyyy') 
            $estimatedEnd = ([Nullable[DateTime]]$jobDataCO.EstEnd).ToString('MM/dd/yyyy')
            $dateCompleted = ([Nullable[DateTime]]$jobDataCO.DateComplBy).ToString('MM/dd/yyyy')
            $reqesterDate = ([Nullable[DateTime]]$jobDataCO.ReqDate).ToString('MM/dd/yyyy')

            $dictionary = @{
            Job_Name                  = $jobDataCO.JobName
            Contractor                = $jobDataCO.GC
            Contractor_Contact_Name   = $jobDataCO.GCContact
            Contractor_Phone          = $jobDataCO.GCContactPhone
            Contractor_Email          = $jobDataCO.GCContactEmail
            Estimated_Start_Date      = $estimatedStart 
            Estimated_End_Date        = $estimatedEnd 
            Estimated_Contract_Amount = $ammountCurrency
            Estimated_Expenses        = $expenseCurrency
            Project_Location_City     = $jobDataCO.ProjCity
            Project_Location_State    = $jobDataCO.ProjState
            Project_Location_Country  = $jobDataCO.ProjCountry
            Charity_Name              = $jobDataCO.NameChar
            Charity_Contact_Name      = $jobDataCO.ContactName
            Charity_Address           = $jobDataCO.CharAddress
            Charity_Tax_ID            = $jobDataCO.CharTaxID
            Charity_Phone             = $jobDataCO.CharPhoneNum
            ANG_PM                    = $jobDataCO.ProjMan
            Person_Requesting_Job_No  = $jobDataCO.Requester
            Date_Requested            = $reqesterDate
            Assigned_Job_No           = $jobDataCO.JobNum
            Completed_By              = $jobDataCO.CompletedBy
            Date_Completed            = $dateCompleted
            Estimated_Margin          = $percent
            }

            $currentPDF = "P:\ANG_System_Files\commonFormsUsedInScripts\ANG Job # Request Form.pdf"
            $outputPDF = "P:\A N G\ANG COMPANY FORMS\JobNumberRequests\ANG Job # Request Form_$($jobDataCO.JobNum)_$($jobDataCO.JobName).pdf"    

            if ([System.Io.File]::Exists($outputPDF))
            {
                [System.Io.File]::Delete($outputPDF)  
            }

            $output = [System.IO.File]::Create($outputPDF)
            $reader = [iTextSharp.text.pdf.PdfReader]::new($currentPDF)
            $stamper = [iTextSharp.text.pdf.PdfStamper]::new($reader, $output)
            $keys = $reader.AcroFields.Fields.Key    ####gets the names of the fields in the form####

            Write-Host "Creating Form for $($jobDataCO.JobName)"

            foreach ($d in $dictionary.Keys)
            {
               $stamper.AcroFields.SetField("$d", "$($dictionary.Item($d))") | Out-null
            }

            if($jobDataCO.EstAmmount -gt "15000")
            {
                $stamper.AcroFields.SetField("Contract", "YES") | Out-null
            }

            if($jobDataCO.EstAmmount -le "15000")
            {
                $stamper.AcroFields.SetField("Misc", "YES") | Out-null
            }

            if($jobDataCO.TypeWork -eq "Commercial")
            {

               $stamper.AcroFields.SetField("Commercial", "X") | Out-null 
            }

            if($jobDataCO.TypeWork -eq "Private")
            {

               $stamper.AcroFields.SetField("Private", "X") | Out-null 
            }

            if($jobDataCO.TypeWork -eq "Other")
            {

               $stamper.AcroFields.SetField("Other", "X") | Out-null 
            }

            if($jobDataCO.TypeWork -eq "Charity Work")
            {

               $stamper.AcroFields.SetField("Charity", "X") | Out-null 
            }

            $stamper.Close()

            if($jobDataCO.FileStructure -eq $false)
            {
                $folderName = "$($jobDataCO.JobNum) - $($jobDataCO.JobName)"
                $pDriveLocation = "P:\A N G\Projects\$folderName"
                New-Item -ItemType directory -Path $pDriveLocation
                Copy-Item "P:\ANG_System_Files\commonFormsUsedInScripts\PROJECT FOLDERS\000 Bid"                        -Destination $pDriveLocation -Recurse -Container
                Copy-Item "P:\ANG_System_Files\commonFormsUsedInScripts\PROJECT FOLDERS\100 GC name -GC job #-"         -Destination $pDriveLocation -Recurse -Container
                Copy-Item "P:\ANG_System_Files\commonFormsUsedInScripts\PROJECT FOLDERS\150 Submittals"                 -Destination $pDriveLocation -Recurse -Container
                Copy-Item "P:\ANG_System_Files\commonFormsUsedInScripts\PROJECT FOLDERS\200 Finance"                    -Destination $pDriveLocation -Recurse -Container
                Copy-Item "P:\ANG_System_Files\commonFormsUsedInScripts\PROJECT FOLDERS\300 ANG Key Processes"          -Destination $pDriveLocation -Recurse -Container
                Copy-Item "P:\ANG_System_Files\commonFormsUsedInScripts\PROJECT FOLDERS\400 PO"                         -Destination $pDriveLocation -Recurse -Container
                Copy-Item "P:\ANG_System_Files\commonFormsUsedInScripts\PROJECT FOLDERS\401 Vendors"                    -Destination $pDriveLocation -Recurse -Container
                Copy-Item "P:\ANG_System_Files\commonFormsUsedInScripts\PROJECT FOLDERS\500 Subcontractors"             -Destination $pDriveLocation -Recurse -Container
                Copy-Item "P:\ANG_System_Files\commonFormsUsedInScripts\PROJECT FOLDERS\600 ANG Engineering & Drafting" -Destination $pDriveLocation -Recurse -Container
                Copy-Item "P:\ANG_System_Files\commonFormsUsedInScripts\PROJECT FOLDERS\700 Field"                      -Destination $pDriveLocation -Recurse -Container
                Copy-Item "P:\ANG_System_Files\commonFormsUsedInScripts\PROJECT FOLDERS\800 Shop"                       -Destination $pDriveLocation -Recurse -Container
                Copy-Item "P:\ANG_System_Files\commonFormsUsedInScripts\PROJECT FOLDERS\900 Closeout & Warranty"        -Destination $pDriveLocation -Recurse -Container

                $newPoLogName = "$folderName PO Log.xlsx"
                $newPoLogPath = "$pDriveLocation\400 PO\$newPoLogName"

                Rename-Item -path "$pDriveLocation\400 PO\_ANG PO log.xlsx" -NewName "$newPoLogName"  

                $Excel = New-Object -ComObject excel.application 
                $Excel.visible = $false
                $Workbook = $Excel.Workbooks.open("$newPoLogPath")
                $Worksheet = $Workbook.Worksheets.item("po log") 
                $worksheet.activate() | Out-Null
                $worksheet.Cells.Item(1, 8) = if ($jobDataCO.JobName -ne $null){$jobDataCO.JobName} else {[string]::Empty} 
                $worksheet.Cells.Item(2, 8) = if ($jobDataCO.ProjMan -ne $null){$jobDataCO.ProjMan} else {[string]::Empty}
                $worksheet.Cells.Item(2, 12) = if ($jobDataCO.JobNum -ne $null){$jobDataCO.JobNum} else {[string]::Empty}
                $workbook.Save() | Out-Null
                $workbook.Close() | Out-Null  
                $Excel.Quit() | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Worksheet)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)

                $PathCell = [Smartsheet.Api.Models.Cell]::new()
                $PathCell.ColumnId = $PathCol.Id
                $PathCell.Value    = if ($newPoLogPath -ne $null){$newPoLogPath} else {[string]::Empty}

                $row = [Smartsheet.Api.Models.Row]::new()
                $row.Id = $jobDataCO.RowId
                $row.Cells = [Smartsheet.Api.Models.Cell[]]@($PathCell)

                $updateRow = $client.SheetResources.RowResources.UpdateRows($jobDataId, [Smartsheet.Api.Models.Row[]]@($row))

                start $pDriveLocation
            }

            $PaperworkOfficeCell = [Smartsheet.Api.Models.Cell]::new()
            $PaperworkOfficeCell.ColumnId = $JobDataOfficeCol.Id
            $PaperworkOfficeCell.Value    =  $true

            $PercentCell = [Smartsheet.Api.Models.Cell]::new()
            $PercentCell.ColumnId = $PercentCol.Id
            $PercentCell.Value    =  $percent

            $FileStructureCell = [Smartsheet.Api.Models.Cell]::new()
            $FileStructureCell.ColumnId = $FileStructCol.Id
            $FileStructureCell.Value    =  $true

            $row = [Smartsheet.Api.Models.Row]::new()
            $row.Id = $jobDataCO.RowId
            $row.Cells = [Smartsheet.Api.Models.Cell[]]@($PaperworkOfficeCell, $PercentCell, $FileStructureCell)
            try
            {
                $updateRow = $client.SheetResources.RowResources.UpdateRows($jobDataId, [Smartsheet.Api.Models.Row[]]@($row))

                $result = Save-AttachmentToSheetRow -sheetId $jobDataId -rowId $updateRow.Id -file $outputPDF -mimeType "application/pdf"
            }
            catch
            {
                Write-Error $_.Exception.Message
                Write-Host ""
            }

            start "P:\A N G\ANG COMPANY FORMS\JobNumberRequests"
       } 
    }
}