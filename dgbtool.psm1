<#
.Synopsis
   This reads the Word dokument, that is used to gather information about new employes. 
.DESCRIPTION
   Denne Cmdlet kan trække information ud af word dokumentet "Ny bruger Skema" og Laver et custum Objekt
   der hedder DGBSkemaData, som bliver brugt af andre DGBTools Cmdlets. "Import-DGBSkemaData" har -Path som en Parameter,
   det er en påkrævet parameter. 
.EXAMPLE
   Import-DGBSkemaData -Path "F:\Chris Test.docx"
.EXAMPLE
   Get-ChildItem -Path f:\ | Import-DGBSkemaData
#>
function Import-DGBSkemaData{
    [CmdletBinding()]
    
Param(
    # Path to the document "Ny Medarbejder Skema"
    [Parameter(ValueFromPipelineByPropertyName=$true, Position=0, Mandatory=$true)]
    [string]$Path
)#Param

    Begin{
        # Laver et COM objekt for Microsoft Word
        $objWord = New-Object -Com Word.Application
        $objWord.Visible=$false
    }#Begin

    Process{
        # Open current document
        $objDocument = $objWord.Documents.Open("$Path")
         
        # Find the version
        $paras = $objDocument.Paragraphs
        $Version = $paras[3].range.text
        Write-Verbose "Dokumentet er $Version"

        # Set variable for tables
        $tables = $objDocument.Tables
  
        # Henter data fra Medarbejder delen af Ny ansat dokumentet
        Write-Verbose "Collecting user info."
        $Fornavne = $tables[1].cell(2,2).range.text
        $Efternavn = $tables[1].cell(2,3).range.text
        $Addresse = $tables[1].cell(4,2).range.text
        $Postnummer = $tables[1].cell(4,3).range.text
        $By = $tables[1].cell(4,4).range.text
        $Telefonnummer = $tables[1].cell(6,3).range.text
        $Email = $tables[1].cell(6,4).range.text

        # Henter data fra Afdelingslede delen af Ny ansat dokumentet
        Write-Verbose "Collecting Department info."
        $Beskrivelse = $tables[2].cell(3,2).range.text
        $Afdelingsnummer = $tables[2].cell(3,3).range.text
        $Afdeling = $tables[2].cell(3,4).range.text
        $Leder = $tables[2].cell(3,5).range.text
        $StartDato = $tables[2].cell(5,4).range.text

        # Gather data from the Administration/HR part of the "Ny ansat" document
        Write-Verbose "Collecting HR info."
        $Brugernavn = $tables[4].cell(2,2).range.text
        $Personalenummer = $tables[4].cell(2,3).range.text

        # Splatting the properties for the custom object
        $Property = [ordered]@{
            'Dokumentversion'="$Version";
            'GivenName'="$Fornavne";
            'Surname'="$Efternavn";
            'StreetAddress'="$Addresse";
            'PostalCode'="$Postnummer"
            'City'="$By"
            'HomePhone'="$Telefonnummer"
            'EmailAddress'="$Email"
            'Name'="$Brugernavn";
            'EmployeeNumber'="$Personalenummer";
            'Title'="$Beskrivelse";
            'Description'="$Beskrivelse"
            'DepartmentID'="$Afdelingsnummer";
            'Department'= "$Afdeling"
            'Manager'="$Leder";
            'Startdato'="$StartDato"
            'Company'= 'Den Gamle By'
            'Country'="DK"
        }# $Property

        $DGBSkemaData = New-Object -TypeName PSObject -Property $Property

        Write-Output $DGBSkemaData

        $objDocument.ActiveDocument.Close

    }# Process

    End{
        $objWord.Quit()
    }#End
}#Import-DGBSkemaData

<#
.Synopsis
   This funtion cleans up the DGBSkemaData object
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Optimize-DGBSkemaData{
    [CmdletBinding()]
    
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $DGBSkemaData  
    )

    Begin
    {
    }
    Process{
        if ($DGBSkemaData -eq 'Klik her.')
        {
            $DGBSkemaData = $false
        }
        Else{}
        Write-Output $test
    }
    End
    {
    }
}#Optimize-DGBSkemaData

<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
function Write-DGBSkemaDataLog{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  HelpUri = 'http://www.microsoft.com/',
                  ConfirmImpact='Medium')]
    [OutputType([String])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [ValidateCount(0,5)]
        [ValidateSet("sun", "moon", "earth")]
        [Alias("p1")] 
        $Param1,

        # Param2 help description
        [Parameter(ParameterSetName='Parameter Set 1')]
        [AllowNull()]
        [AllowEmptyCollection()]
        [AllowEmptyString()]
        [ValidateScript({$true})]
        [ValidateRange(0,5)]
        [int]
        $Param2,

        # Param3 help description
        [Parameter(ParameterSetName='Another Parameter Set')]
        [ValidatePattern("[a-z]*")]
        [ValidateLength(0,15)]
        [String]
        $Param3
    )

    Begin
    {
        $Dato = Get-Date
        $LogPath = "c:\Logs\NyeBrugere.txt"
    }
    Process
    {
        foreach ($Data in $DGBUserData)
        {
            $Dato | out-file 
            $Data.Fornavn $Data.Efternavn | Out-File $LogPath -Append
            "Manglende Informationer" | Out-File $LogPath -Append
        }
    }
    End
    {
    }
}


Send-MailMessage -From IT@dengamleby.dk -To IT@dengamleby.dk -Subject "Forkert version!" -SmtpServer dgbexch01.dengamleby.dk -Body "Medarbejder skemaet var en forkert version, dokumentets navn er TEST"
