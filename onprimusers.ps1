#Get-Mailbox -ResultSize unlimited | Export-Csv C:\Users\malik.sh\Desktop\OnPremUsers012321.csv
*******************************************************************************************************

#Author: Shivam Malik
#Connect Exchange On-Prem Powershell
#Change CSV Input path
#Change CSV Output path
#Change input file CSV Header to "UserPrincipalName"


$name=@()
$name=import-csv C:\Users\malik.sh\Desktop\M147.csv
$report=@()

foreach ($names in $name)
{
    $a=Get-User -Identity $names.UserPrincipalName
    $b=Get-Mailbox -Identity $names.UserPrincipalName
    $c=Get-MailboxStatistics -Identity $names.UserPrincipalName

    $report += [PSCustomObject]@{
        UPN = $a.UserPrincipalName
        City= $a.City
        Company= $a.Company
        Country= $a.CountryOrRegion
        Department= $a.Department
        DisplayName= $a.DisplayName
        WhenCreated= $a.WhenCreated
        Manager= $a.Manager
        WhenSoftDeleted= $a.WhenSoftDeleted
        RecipientType= $a.RecipientType
        RecipientTypeDetails= $a.RecipientTypeDetails
        Database= $b.Database          
        ArchiveDatabase= $b.ArchiveDatabase
        LitigationHoldStatus= $b.litigationholdenabled
        LitigationHoldDuration= $b.LitigationHoldDuration
        LitigationHoldOwner= $b.LitigationHoldOwner
        LitigationHoldDate= $b.LitigationHoldDate
        ExchangeGUID= $b.ExchangeGUID
        PrimarySmtpAddress= $b.PrimarySmtpAddress
        MailboxSize= $c.TotalItemSize
        OrganizationalUnit= $a.OrganizationalUnit

    }
}
$report | Export-csv -Path C:\Users\malik.sh\Desktop\On-PremUserDetailsM147.csv –notypeinformation

