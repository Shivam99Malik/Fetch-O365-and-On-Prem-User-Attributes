#Author: Shivam Malik
#Connect-MSOLService
#Connect Exchange online Powershell (Connect-EXOPSSession -UserPrincipalName malik.sh.adm@domine.com)
#Change CSV Input path
#Change CSV Output path
#Change input file CSV Header to UPN



$name=@()
$name= import-csv C:\\exp\GradeRefrigeration.csv
$report=@()

foreach ($names in $name)
{
    $a=Get-MSOLUser -userprincipalname $names.UPN
    $b=Get-Mailbox -Identity $names.UPN
    $c=Get-MailboxStatistics -Identity $names.UPN

    $LicensesName= $a.Licenses
    $LicenseArray = $LicensesName | foreach {$_.AccountSkuId}
    $licenseString = $licenseArray -join ";"

    $report += [PSCustomObject]@{
        UPN = $a.UserPrincipalName
        IsLicense= $a.islicensed
        ADCity= $a.City
        ADCountry= $a.Country
        Department= $a.Department
        DisplayName= $a.DisplayName
        ObjectId= $a.ObjectId
        UsageLocation= $a.UsageLocation
        WhenCreated= $a.WhenCreated
        LitigationHoldStatus= $b.litigationholdenabled
        LitigationHoldDuration= $b.LitigationHoldDuration
        LitigationHoldOwner= $b.LitigationHoldOwner
        LitigationHoldDate= $b.LitigationHoldDate
        O365UsageLocation= $b.UsageLocation
        CA13= $b.customattribute13
        PrimarySmtpAddress= $b.PrimarySmtpAddress
        O365WhenCreated= $b.WhenCreated
        RecipientType=$b.RecipientType
        MailboxSize= $c.TotalItemSize
        LastLogonTime= $c.lastlogontime
        License = $licenseString
    }
 }
 $report | Export-csv -Path C:\exp\outGradeRefrigerationO365.csv -notypeinformation

