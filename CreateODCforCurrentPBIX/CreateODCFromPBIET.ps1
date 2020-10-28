Function MZ-PBIDesktopODC
{
# modified the https://github.com/DevScope/powerbi-powershell-modules/blob/master/Modules/PowerBIPS.Tools/PowerBIPS.Tools.psm1
# inspired by Erik Svensen blog: https://eriksvensen.com/2020/07/27/powerbi-external-tool-to-connect-excel-to-the-current-pbix-file/
# based on https://github.com/donsvensen/erikspbiexcelconnector

<#
.SYNOPSIS
Exports a PBIDesktop ODC connection file
.DESCRIPTION
Exports a PBIDesktop ODC connection file with the name of the current file (name PBI Desktop window) via External Tools
.PARAMETER server
Server address provided by Power BI Desktop External Tools, f.e. "localhost:65534"
.PARAMETER path
ODC file path to be created
.EXAMPLE
MZ-PBIDesktopODCSample -server "localhost:65534" -path "C:\Temp"
#>
    param
    (
        [Parameter(Mandatory = $false)]        
		[string]
        $server,
        [Parameter(Mandatory = $false)]        
		[string]
        $path	
    )
    $address = $server
    $port = [int]$address.Split(":")[1]
    $processes =  Get-NetTCPConnection |? State -eq "Established" |? RemotePort -eq $port | Select OwningProcess
    $title = get-process |? ProcessName -eq "PBIDesktop" |? Id -eq $processes.OwningProcess[0] | select MainWindowTitle
    $name = $($title.MainWindowTitle).Replace(" - Power BI Desktop","")

    # Write-Output @{WindowTitle=$name; Port = $address} 

    $odcXml = "<html xmlns:o=""urn:schemas-microsoft-com:office:office""xmlns=""http://www.w3.org/TR/REC-html40""><head><meta http-equiv=Content-Type content=""text/x-ms-odc; charset=utf-8""><meta name=ProgId content=ODC.Cube><meta name=SourceType content=OLEDB><meta name=Catalog content=164af183-2454-4f45-964a-c200f51bcd59><meta name=Table content=Model><title>$name</title><xml id=docprops><o:DocumentProperties  xmlns:o=""urn:schemas-microsoft-com:office:office""  xmlns=""http://www.w3.org/TR/REC-html40"">  <o:Name>$name</o:Name> </o:DocumentProperties></xml><xml id=msodc><odc:OfficeDataConnection  xmlns:odc=""urn:schemas-microsoft-com:office:odc""  xmlns=""http://www.w3.org/TR/REC-html40"">  <odc:Connection odc:Type=""OLEDB"">   
        <odc:ConnectionString>Provider=MSOLAP;Integrated Security=ClaimsToken;Data Source=$address;MDX Compatibility= 1; MDX Missing Member Mode= Error; Safety Options= 2; Update Isolation Level= 2; Locale Identifier= 1033</odc:ConnectionString>   
        <odc:CommandType>Cube</odc:CommandType>   <odc:CommandText>Model</odc:CommandText>  </odc:Connection> </odc:OfficeDataConnection></xml></head></html>"
    
    $odcFile = "$path\$name.odc"
    $odcXml | Out-File $odcFile -Force	
}

MZ-PBIDesktopODC -server $args[0] -path "C:\Temp"