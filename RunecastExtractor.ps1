<#
#Powershell Runecast API query script
#v1.0 vMan.ch, 13.06.2018 - Initial Version
 
    A lazy vMans module to hit the Runecast API and extract data.
 
    Script requires Powershell v3 and above.
 
    Make sure to install the Required Module --> https://github.com/dfinke/ImportExcel
 
    Usage
    .\RunecastExtractor.ps1 -Runecast runecast.vMan.ch -Token '7546e68b-96bc-406e-8d57-280e1de75670' -FileName 'RunecastExtract.xlsx' -OutputLocation D:\RunecastExtractor\
 
#>
 
param
(
    [String]$Runecast,
    [String]$Token,
    [String]$FileName,
    [String]$OutputLocation
)
 
 
 
#Take all certs.
add-type @"
    using System.Net;
    using System.Security.Cryptography.X509Certificates;
    public class TrustAllCertsPolicy : ICertificatePolicy {
        public bool CheckValidationResult(
            ServicePoint srvPoint, X509Certificate certificate,
            WebRequest request, int certificateProblem) {
            return true;
        }
    }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
 
#Stuff for Invoke-RestMethod
$ContentType = "application/json"
$header = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$header.Add("Accept", 'application/json')
$header.Add("Authorization", $Token)
$header.Add("User-Agent", 'vManRunecastExtractor/1.0')
 
 
#Get a list of every Issue
 
 
    $IssueList = @()
 
    $IssueUrl = 'https://'+$Runecast+'/rc2/api/v1/issues'
 
    $issues = Invoke-RestMethod -Uri $IssueUrl -Method Get -Headers $header -ContentType $ContentType
 
    ForEach ($issue in $issues.issues){
 
        $IssueList += New-Object PSObject -Property @{
 
                    id = $issue.id
                    affects = $issue.affects
                    appliesTo = $issue.appliesTo
                    severity = $issue.severity
                    type = $issue.type
                    title = $issue.title
                    url = $issue.url
                    annotation = $issue.annotation 
                    updatedDate = $issue.updatedDate
                    stigid = $issue.stigid
                    vulnid = $issue.vulnid
                    checkDescription = $issue.checkDescription
                    fixDescription = $issue.fixDescription
                    stigSection = $issue.stigSection
                
        }
 
    }
 
 
#Get a list of VC's
 
    $VCList = @()
 
    $VCUrl = 'https://'+$Runecast+'/rc2/api/v1/vcenters'
 
    $VCs = Invoke-RestMethod -Uri $VCUrl -Method Get -Headers $header -ContentType $ContentType
 
    ForEach ($VC in $VCs.vcenters){
 
        $VCList += New-Object PSObject -Property @{
 
                    vcUid = $VC.uid
                    address = $VC.address
        }
 
    }
 
 
#Get a list of results
 
    $ResultsList = @()
 
    $ResultsUrl = 'https://'+$Runecast+'/rc2/api/v1/results'
 
    $results = Invoke-RestMethod -Uri $ResultsUrl -Method Get -Headers $header -ContentType $ContentType
 
    ForEach ($Result in $Results.Results.issues){
 
        $id = $Result.id
        $status = $Result.Status
 
        ForEach ($affectedObject in $Result.affectedObjects){
 
            $ResultsList += New-Object PSObject -Property @{
 
                    id = $id
                    Name = $affectedObject.Name
                    vcUid = $affectedObject.vcUid
                    moid = $affectedObject.moid
                
        }
       }
    }
 
	
 
 
#Export it all to Excel baby!!
 
$File = $OutputLocation + $FileName
 
 
$IssueList | Select id,affects,appliesTo,severity,type,title,url,annotation,updatedDate,stigid,vulnid,checkDescription,fixDescription,stigSection | export-excel $File -WorkSheetname Issues
 
$VCList | Select vcUid,address | export-excel $File -WorkSheetname vCenters
 
$ResultsList | Select id,Name,vcUid,moid | export-excel $File -WorkSheetname Results