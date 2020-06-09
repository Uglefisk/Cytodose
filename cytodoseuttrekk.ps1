<#
.Synopsis
    20200518 - reinert.hansen@helse-vest-ikt.no
   Uttrekk fra Cytodose til Kreftregisteret
.DESCRIPTION
   Leser inn en liste med pasienter/dato fra-tidspunkt, og kjører en stored procedure for å hente alle 
    rader for alle pasienter i alle cytodosedatabaser som er definert.
.EXAMPLE
   Man definerer servernavn, brukernavn og passord (som er kryptert - se retningslinjer i koden), så kjøres scriptet i sin helhet.
.NOTES
   General notes
#>

$secureStringText = 'et hemmelig og sikkert passord til en sqllogin for alle baser på cytodose'
$secureStringPwd = $secureStringText | ConvertTo-SecureString -AsPlainText -Force
$secureStringText = $secureStringPwd | ConvertFrom-SecureString 
# lagrer kryptert passord for senere gjenbruk i automatisering
Set-Content "C:\Users\ninja\Downloads\passord.txt" $secureStringText
# henter kryptert passord fra fil
$pwdTxt = Get-Content "C:\Users\ninja\Downloads\passord.txt"
$securePwd = $pwdTxt | ConvertTo-SecureString
$pwdTxt # echo kryptert string
# lager binary string
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePwd)
# cast til string
$UnsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)


#henter liste med pasienter, levert av kreftregisteret
$pasientliste = Import-Csv "C:\Users\ninja\Downloads\Pasientliste_please_change.csv" -Delimiter ";"
#liste med databaser hvor man skal hente data
$databaser = @("Cytodose_HBE_Prod", "Cytodose_HST_Prod", "cytodose_hfd_prod", "Cytodose_hfo_prod")

# lager en datatable som holder for alle resultater
$resultatsett = New-Object System.Data.DataTable
$resultatsett.Columns.Add("Result", "System.String") |Out-Null

# kjører stored procedure for alle baser og alle pasienter
foreach ($database in $databaser)
{
    #for hver pasient i csvfilen
    Write-Host "Processing $database" -BackgroundColor DarkYellow -ForegroundColor Blue
    foreach ($pasient in $pasientliste)
    {
        #variabler som parametre til stored procedure
        $fnr = $pasient.FODSELSNR
        $dato = $pasient.diagnosedato
        #sqlstatement (stored procedure) som skal kjøres
        $sql = "exec pINTGetCompletedOrders @PIN = '" + $fnr + "', @MinDate = '" + $dato + "', @MaxDate = '2020-05-14'"
        # write-host  $database $sql
        #kjører sqlkommandoen og tar vare på resultatsettet i variabelen $r
        $r = Invoke-Sqlcmd -ServerInstance "vir-rdbms347" -Username AsimpleSQLLogin -Password $UnsecurePassword -Database $database -Query $sql -as DataSet # | Export-Csv "C:\Users\ninja\Downloads\tbProcExample.csv" -NoTypeInformation
      
        if ($r -ne $null)
        {
            # legger til alle rader som ble returnert inn i $resultatsett - skal skrives som csv til slutt.
            foreach ($record in $r.Tables[0].Rows)
            {           
                $resultatsett.Rows.Add($record.ItemArray)
            }
        }
    
    }
}

#sjekker antall rader totalt
$resultatsett.Rows.Count
#skriver resultatsettet til csv
$resultatsett|Export-Csv -Path C:\users\ninja\Downloads\liste.csv
# henter evt. resultatsettet ut i gridview
# $resultatsett|Out-GridView

$resultatsett.Rows.Item(0)|Out-GridView
$resultatsett|get-member
$resultatsett.Rows.Count

# legger til 7zip i powershellsesjonen
$7zipPath = "$env:ProgramFiles\7-Zip\7z.exe"

if (-not (Test-Path -Path $7zipPath -PathType Leaf)) {
    throw "7 zip file '$7zipPath' not found"
}

Set-Alias 7zip $7zipPath

$Source = "C:\users\ninja\Downloads\liste.csv"
$Target = "C:\users\ninja\Downloads\liste.zip"
$zippassword = "Secret with Capital S - please change :)"
# lagt zip-fil med passord
7zip a -mx=9 $Target $Source -p"$zippassword"
