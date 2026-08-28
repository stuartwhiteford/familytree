$ErrorActionPreference = 'Stop'

$sourceFiles = Get-ChildItem $PSScriptRoot -Filter '*.docx' | Where-Object { $_.Name -like 'Whitefords 1881 Census *.docx' }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
function Add-CurrentPerson {
    if ($null -ne $script:currentPerson) {
        $script:records.Add([pscustomobject]@{
            household_type = $script:currentPerson.household_type
            dwelling = $script:currentPerson.dwelling
            census_place = $script:currentPerson.census_place
            film = $script:currentPerson.film
            reference_type = $script:currentPerson.reference_type
            reference_volume_or_piece = $script:currentPerson.reference_volume_or_piece
            folio = $script:currentPerson.folio
            enumeration_district = $script:currentPerson.enumeration_district
            page = $script:currentPerson.page
            name = $script:currentPerson.name
            marital_status = $script:currentPerson.marital_status
            age = $script:currentPerson.age
            sex = $script:currentPerson.sex
            birthplace = $script:currentPerson.birthplace
            relationship = $script:currentPerson.relationship
            occupation = $script:currentPerson.occupation
        })
        $script:currentPerson = $null
    }
}

try {
    foreach ($sourceFile in $sourceFiles) {
        $document = $null
        $records = [System.Collections.Generic.List[object]]::new()
        $currentPerson = $null
        $householdType = 'Dwelling'
        $dwelling = ''
        $censusPlace = ''
        $film = ''
        $volume = ''
        $referenceType = ''
        $referenceVolumeOrPiece = ''
        $folio = ''
        $enumerationDistrict = ''
        $page = ''
        $outputPath = [IO.Path]::ChangeExtension($sourceFile.FullName, '.csv')
        $document = $word.Documents.Open($sourceFile.FullName, $false, $true)
        foreach ($paragraph in $document.Paragraphs) {
        $text = ($paragraph.Range.Text -replace '[\r\a]', '').Trim()
        if ([string]::IsNullOrWhiteSpace($text)) { continue }

        if ($text -match '^Dwelling:\s*(?<value>.*)$') {
            Add-CurrentPerson
            $householdType = 'Dwelling'
            $dwelling = $Matches.value.Trim()
            continue
        }
        if ($text -match '^(?<type>Institution|Vessel):\s*(?<value>.*)$') {
            Add-CurrentPerson
            $householdType = $Matches.type
            $dwelling = $Matches.value.Trim()
            continue
        }
        if ($text -match '^Census Place:\s*(?<value>.*)$') {
            Add-CurrentPerson
            $censusPlace = $Matches.value.Trim()
            continue
        }
        if ($text -match '^Source:\s*FHL Film\s+(?<film>\S+)\s+(?<referenceType>GRO|PRO) Ref\s+(?:(?:Volume\s+(?<volume>\S+))|(?:RG11\s+Piece\s+(?<piece>\S+)\s+Folio\s+(?<folio>\S+)))\s+(?:(?:EnumDist\s+(?<district>\S+)\s+)|)Page\s+(?<page>\S+)') {
            Add-CurrentPerson
            $film = $Matches.film
            $referenceType = $Matches.referenceType
            $referenceVolumeOrPiece = if ($Matches.volume) { $Matches.volume } else { $Matches.piece }
            $folio = $Matches.folio
            $enumerationDistrict = $Matches.district
            $page = $Matches.page
            continue
        }
        if ($text -match '^Rel:\s*(?<value>.*)$') {
            if ($null -ne $currentPerson) { $currentPerson.relationship = $Matches.value.Trim() }
            continue
        }
        if ($text -match '^Occ:\s*(?<value>.*)$') {
            if ($null -ne $currentPerson) { $currentPerson.occupation = $Matches.value.Trim() }
            continue
        }
        if ($text -match '^_+$' -or $text -match '^Extract of ' -or $text -match '^Marr\s+Age\s+Sex\s+Birthplace$') {
            continue
        }

        if ($text -match '^(?<name>.+?)\s+(?:(?<marital>[MUW])\s+)?(?<age>\d+|\.\.\.)\s+(?<sex>[MF])(?:\s+(?<birthplace>.*))?$') {
            Add-CurrentPerson
            $currentPerson = [pscustomobject]@{
                household_type = $householdType
                dwelling = $dwelling
                census_place = $censusPlace
                film = $film
                reference_type = $referenceType
                reference_volume_or_piece = $referenceVolumeOrPiece
                folio = $folio
                enumeration_district = $enumerationDistrict
                page = $page
                name = $Matches.name.Trim()
                marital_status = if ($null -eq $Matches.marital) { '' } else { $Matches.marital.Trim() }
                age = if ($Matches.age -eq '...') { '' } else { $Matches.age }
                sex = $Matches.sex
                birthplace = if ($null -eq $Matches.birthplace) { '' } else { $Matches.birthplace.Trim() }
                relationship = ''
                occupation = ''
            }
        }
        }
        Add-CurrentPerson
        $records | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
        "Wrote $($records.Count) records to $outputPath"
        $document.Close([ref]$false)
        $document = $null
    }
}
finally {
    if ($null -ne $document) { $document.Close([ref]$false) }
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
