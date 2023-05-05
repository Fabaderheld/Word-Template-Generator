<#
.SYNOPSIS
    Short description
.DESCRIPTION
    Long description
.EXAMPLE
    PS C:\> <example usage>
    Explanation of what the example does
.PARAMETER Name
    The description of a parameter. Add a ".PARAMETER" keyword for each parameter in the function or script syntax.
.INPUTS
    Inputs (if any)
.OUTPUTS
    Output (if any)
.NOTES
    General notes
#>


[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [String]
    $WordTemplate,
    [String]
    $FromCSV,
    [String]
    $FromExcel,
    $Image,
    $AddressLine,
    $FirstName,
    $LastName,
    $DestinationPath,
    $NewTemplateName
)

if ($FromCSV) {
    $Data = Get-Content -Path $FromCSV
}
else {
    $Data = [array][PSCustomObject]@{
        Image       = $Image
        Addressline = $AddressLine
        FirstName   = $FirstName
        LastName    = $LastName
    }

}

foreach ($Item in $Data) {

    # Setup variables
    $tmpPath = "$env:Temp\wordtemplate"
    $ImagePath = Join-Path -Path $tmpPath -ChildPath "word\media\image1.jpeg"
    $DocumentPath = Join-Path -Path $tmpPath -ChildPath "word\document.xml"

    # extract template to tmp path
    New-Item -Path $tmpPath -ItemType Directory -ErrorAction SilentlyContinue
    Expand-Archive -Path $WordTemplate -DestinationPath $tmpPath

    # Parse document
    $Document = Get-Content -Path $DocumentPath

    # replace placeholders with values
    $Document.Replace("<w:t>$FirstNamePlaceHolder</w:t>", "<w:t>$($Item.FirstName)</w:t>") | Set-Content -Encoding utf8 -Path $DocumentPath -Force

    # replace Image
    Copy-Item -Path $Item.Image -Destination $ImagePath -Force

    # Generate new template file
    $DestinationFile = Join-Path -Path $DestinationPath -ChildPath $NewTemplateName -AdditionalChildPath ".dotx"
    Compress-Archive -Path "$tmpPath\*" -DestinationPath $DestinationFile
}


function Update-WordTemplate {
    param (
        [Parameter(Mandatory)]
        [String]
        $WordFile,
        $Image,
        $AddressLine,
        $FirstName,
        $FirstNamePlaceHolder,
        $LastName,
        $DestinationPath,
        $NewTemplateName
    )

}