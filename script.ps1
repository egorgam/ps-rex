param(
  [string]$Regex,
  [string]$Server,
  [int]$Port
)
function ConvertTo-Json20([object] $item){
    add-type -assembly system.web.extensions
    return $(new-object system.web.script.serialization.javascriptSerializer).Serialize($item)
}
function SendReport ($Status, $Path) {
    $Data = @{
        "ip" = [string](gwmi Win32_NetworkAdapterConfiguration | ? { $_.IPAddress -ne $null }).ipaddress
        "hostname" = $([System.Net.DNS]::GetHostByName('').HostName)
        "path" = $Path
        "timestamp" = Get-Date -Format o | ForEach-Object {$_ -replace ":", "."}
        "status" = $Status
        }
    $EndPoints = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Parse($Server), $Port)
    $Socket = New-Object System.Net.Sockets.UDPClient 
    $Socket.TTL = 128
    $EncodedText = [Text.Encoding]::UTF8.GetBytes($(ConvertTo-Json20 $Data)) 
    $SendMessage = $Socket.Send($EncodedText, $EncodedText.Length, $EndPoints) 
    $script:trigger = $false
    Write-Host $Status $Path -f "green"
}
function SearchIntoLabel ($Path) {
    Get-ChildItem -Path $Path -Recurse -ErrorVariable Errors -ErrorAction SilentlyContinue | Where-Object { $_.Attributes -ne [System.IO.FileAttributes]::Directory } | ForEach-Object {
        $Files = @($_)
        for ($i = 0; ($Files[$i]); $i++) {
            if ([regex]::IsMatch($_.Extension, ".doc.*|.rtf|.odt|.ott|.oth|.odm|.wps")) {
                $Word = New-Object -ComObject Word.application
                $Word.DisplayAlerts = $False
                $Document = $Word.Documents.Open($_.FullName, $false, $true)
                if ([regex]::IsMatch($Document.Range().Text, $Regex)){
                    SendReport -Status "MATCH" -Path $_.FullName 
                }
                $Word.Quit()
                [GC]::Collect()
                [GC]::WaitForPendingFinalizers()
            }
            if ([regex]::IsMatch($_.Extension, ".xl.*|.csv|.ods|.ots|.sxc|.stc|.wk.*")) {     
                $Excel = New-Object -ComObject Excel.Application
                $Excel.DisplayAlerts = $False
                $Workbook = $Excel.Workbooks.Open($_.FullName)
                if ([regex]::IsMatch($Workbook.ActiveSheet.UsedRange.Rows.Formula, $Regex)){
                    SendReport -Status "MATCH" -Path $_.FullName 
                }
                $Excel.Quit()
                [GC]::Collect()
                [GC]::WaitForPendingFinalizers()
            }   
            if ([regex]::IsMatch($_.Extension, ".ppt.*|.odf|.odg|.otp|.sxi|.sti")) {
                $PowerPoint = New-Object -ComObject Powerpoint.application
                $PowerPoint.DisplayAlerts = $False
                $Presentation = $PowerPoint.Presentations.Open($Files[$i].FullName, $true, $true, $false)
                ForEach ($Slide in $Presentation.Slides){
                    ForEach ($Shape in $Slide.Shapes){
                        if ([regex]::IsMatch($Shape.TextFrame.TextRange.Text, $Regex)){
                            SendReport -Status "MATCH" -Path $Files[$i].FullName 
                            $o = $true
                            break
                        }
                    }
                    if($o -eq $true) {
                        $o = $false
                        break
                    }
                }
                $PowerPoint.Quit()
                [GC]::Collect()
                [GC]::WaitForPendingFinalizers()
            }
            if ([regex]::IsMatch($_.Extension, ".txt|.log|.xml|.htm.*|.pdf")) {
                if (Write-Output $_ | Select-String $Regex) {
                    SendReport -Status "MATCH" -Path $_.FullName 
                }
            }
        }
    }
}
Get-WmiObject win32_logicaldisk| ? {$_.drivetype -ne 5, 4} | ForEach-Object {
    $Labels = @($_)
    for ($i = 0; $Labels[$i]; $i++) {
        if($_.name -eq "C:"){
            SearchIntoLabel -Path "C:\Users"
        }
        if ($_.name -ne "C:"){
            SearchIntoLabel -Path $_.name
        }
    }
}
if ($Errors) {
    SendReport -Status "ACCESS_DENIED" -Path $Error[0].CategoryInfo.TargetName
}
if ($script:trigger -ne $false){
    SendReport -Status "MATCH" -Path $Null
}
[GC]::Collect()
Write-Output 'end job'
