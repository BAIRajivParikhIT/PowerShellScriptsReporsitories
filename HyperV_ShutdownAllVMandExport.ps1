## Powershell Script to Shutdown and Export Hyper-V 2012 VMs, one at a time.  

#Visit my Blog for more scripts or to comment on this one at http://www.czerno.com/blog/

## Destination Folder of the Exports
## A subfolder for each VM will be created.
$dest = "\\ServerName\Sharename"
## You can also use "F:\Foldername"

## SMTP Server and Email Address Settings
$smtpServer = "smtp.domain.name" 
$From = "$env:computername@domain.name"
$To = "email@domain.name"

##Log File
$LogFile = "ExportVMs.htm"
$Node = $env:computername

"<img src=""http://www.microsoft.com/global/en-us/Server-Cloud/PublishingImages/Product%20Logos/Microsoft_HyperV_Server_R2_320x50.png"" alt=""Hyper-V Export Script Log"" width=""320"" height=""50""><br>" > $LogFile

"<p><b>Starting Export Script on Node <font color=""#0000FF"">" + $Node + "</font></b><br>" >> $LogFile
"<b>Script Start Time </b>" + (get-date).ToShortTimeString() + " <br>" >> $LogFile
"<hr />" >> $LogFile

## Get a list of all VMs on Node
$VMs = Get-VM

"<b>Starting Export of Virtual Machines on this Node: </b>" >> $LogFile
"<br>" >>$LogFile

## For each VM on Node, Shutdown, Export and Start 
foreach ($VM in $VMs) 
{
	$VMName = $VM.Name
	$VMName
    $summofvm = Get-VM -Name $VMName | Get-VMIntegrationService -Name Heartbeat
    $HBStatus = $summofvm.OperationalStatus
    $VMState = $VM.State
    $doexport = "no"
    if ($CVMState -eq $null) {} else {Remove-variable CVMState}

    Write-Host "Checking $VMName"
	"<b>Checking the current state of <font color=""#0000FF"">$VMName </font></b>" >> $LogFile

    ## Shutdown the VM if HeartBeat Service responds
    if ($HBStatus -eq "OK")
    {
    write-host Heartbeat Status is $HBStatus
    $doexport = "yes"
	write-host "HeartBeat Service for $VMName is responding $HBStatus, beginning shutdown sequence"
	"<dd>HeartBeat Service for <font color=""#0000FF"">$VMName </font>is responding $HBStatus, beginning shutdown sequence at " + (get-date).ToShortTimeString() + "</dd>" >> $LogFile
    
    ## Shutting Down Guest OS
	Stop-VM $VMName -force
    }

    ## Checks to see if the VM is already stopped
    elseif ($VMState -eq "Off")
    {
    $CVMState = $VMState
    $doexport = "yes"
	write-host "$VMName is already off, starting export"
	"<dd><font color=""#0000FF"">$VMName </font>is already stopped, starting export </dd>" >> $LogFile
    }

    ## If the HeartBeat service is not OK, aborting this VM
	elseif ($HBStatus -ne "OK" -and $VMState -ne "Off")
	{
	$doexport = "no"
	write-host "The HeartBeat Service for $VMName is not responding, shutdown and export aborted for this VM"
	"<dd><font color=""#FF0000"">******************************************************************************************** </dd>" >> $LogFile
	"<dd>The HeartBeat Service for <b>$VMName</b> is not responding, shutdown and export aborted for this VM </dd>" >> $LogFile
	"<dd>******************************************************************************************** </font></dd>" >> $LogFile
	}

    $i = 1

    ## Loop until the VMs State is off
	if ($doexport -eq "yes")
	{
		Do {
		$i++
		
		## Get the VMs Current State
        $HBStatus = $summofvm.OperationalStatus
        $VMState = $VM.State
		write-host "$VMName is $VMState"
		Start-Sleep -s 10
		}
			while ($VMState -ne "Off" -or $i -gt 300)
			
			## If a folder already exists for the current VM, delete it.
			if ([IO.Directory]::Exists("$dest\$VMName"))
			{
				[IO.Directory]::Delete("$dest\$VMName", $True)
			}

		"<dd><font color=""#0000FF"">$VMName </font>is $VMEnabledState at " + (get-date).ToShortTimeString() + "</dd>" >> $LogFile
		
		write-host "Exporting $VMName"
		"<dd>Export of <font color=""#0000FF"">$VMName</font> began at " + (get-date).ToShortTimeString() + "</dd>" >> $LogFile
		
		## Begin export of the VM
		
		export-vm $VMName -path $dest 
		
		"<dd>Export of <font color=""#0000FF"">$VMName</font> finished at " + (get-date).ToShortTimeString() + "</dd>" >> $LogFile
		
		## Start the VM if it was previously Running and wait for a Heartbeat with time-out
        if ($CVMState -ne "Off")
		{
        write-host "Starting $VMName and waiting for Heartbeat"
		"<dd>Startup of <font color=""#0000FF"">$VMName</font> started at " + (get-date).ToShortTimeString() + "</dd>" >> $LogFile
		Start-VM $VMName
		
		$j = 1

        Do {
		    $j++
		
		    ## Get the VMs Current State
            $HBStatus = $summofvm.OperationalStatus
            $VMState = $VM.State
		    Start-Sleep -s 10
		   }
			until ($j -gt 30 -or $HBStatus -eq "Ok")
		
		write-host $VMName "is" $VMState "and Heartbeat is" $HBStatus
        write-host
		"<dd><font color=""#0000FF"">" + $VMName + "</font> is " + $VMState + " and Heartbeat is " + $HBStatus + ". Startup completed at " + (get-date).ToShortTimeString() + "</dd>" >> $LogFile
        }
        else
        {
        write-host "Not starting $VMName since it was not Running when script started"
        write-host
		"<dd>Not starting <font color=""#0000FF"">$VMName</font> since it was not Running when script started </dd>" >> $LogFile
        }	
    }

    ## Listing Export Folder Details for Log File
    "<dd><table>" >> $LogFile
    #Query Folders
    $dirpath = $dest
    $subfolders = Get-ChildItem $dirpath | Where-Object {$_.PsIsContainer} 
    #Calculate sizes
    $folder = $dirpath + "\" + $VMName 
    $colItems = (Get-ChildItem $folder -Recurse | Measure-Object -property length -sum -ErrorAction SilentlyContinue) 
    $colItemsSum = ("{0:N2}" -f ($colItems.sum / 1GB) + " GB")
    $lastwrite = (get-item $folder).lastwritetime
    "<tr><td><b>Export Folder Details:</b></td><td>$folder</td><td><b>Size:</b></td><td>$colItemsSum</td><td><b>Last modified:</b></td><td>$lastwrite</td></tr>" >> $LogFile
    "</table></dd>" >> $LogFile
    "<br>" >> $LogFile    
}

"<hr />" >> $LogFile
"<b>Script Completion Time</b> " + (get-date).ToShortTimeString() + "<br>" >> $LogFile
"<br></p>"  >> $LogFile


$Subject = "$env:computername Hyper-V Exports"
$Body =  [string]::join([environment]::NewLine, (get-content $LogFile)) 

send-mailmessage -from $From -to $To -subject $Subject -body $Body -BodyAsHtml -priority High -smtpServer $smtpServer
