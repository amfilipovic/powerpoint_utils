function clear-screen {
    if ($env:OS -match 'Windows') {
        cls
    }
    else {
        clear
    }
}

function convert-file-size {
    param([double]$size)
    $power = [Math]::Pow(2, 10)
    $sizeLabels = @('B', 'KB', 'MB', 'GB', 'TB')
    switch($size) {
        {$_ -lt $power} { '{0:F2} {1}' -f $_, $sizeLabels[0]; break}
        default {
            $n = [Math]::Floor([Math]::Log($_, $power))
            '{0:F2} {1}' -f ($_ / [Math]::Pow($power, $n)), $sizeLabels[$n]
        }
    }
}

clear-screen

$currentFolder = Split-Path -Parent $MyInvocation.MyCommand.Definition
$presentations = Get-ChildItem -Path $currentFolder -Filter *.ppt* -File
$totalNumberOfSlides = 0
$totalFileSize = 0

Write-Host "Scan results of '$currentFolder':"
foreach ($presentation in $presentations) {
    $currentPresentation = New-Object -ComObject PowerPoint.Application
    $currentPresentationFile = $currentPresentation.Presentations.Open($presentation.FullName)
    $numberOfSlides = $currentPresentationFile.Slides.Count
    $fileSize = $presentation.Length
    $currentPresentationFile.Close()
    $currentPresentation.Quit()
    Write-Host "Number of slides in '$($presentation.Name)': $numberOfSlides ($((convert-file-size $fileSize)))"
    $totalNumberOfSlides += $numberOfSlides
    $totalFileSize += $fileSize
}

if ($presentations) {
    Write-Host "Total number of slides in $($presentations.Count) presentation(s): $totalNumberOfSlides ($(convert-file-size $totalFileSize))"
}
else {
    clear-screen
    Write-Host "Error: There are no presentations in this folder."
}