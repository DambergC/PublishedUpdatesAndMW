$YearMonth = get-date

$YearMonthMin = $YearMonth.AddDays(-10)
$YearMonthMax = $YearMonth.AddDays(15)



$collectionmembers = @()
$result = @()


$Deployments = Import-Csv -Path "C:\Github\PublishedUpdatesAndMW\data.cfg" -Delimiter ";" -Encoding UTF7

cd PS1:

#$mw = Get-CMMaintenanceWindow -CollectionId SMS00001 | Where-Object {($_.starttime -ge $YearMonthMin) -and ($_.starttime -le $YearMonthMax)}


foreach ($Deployment in $Deployments)

{

   $collectionmembers =  Get-CMCollection -CollectionId $Deployment.collectionid | Get-CMCollectionMember 

   $mw = Get-CMMaintenanceWindow -CollectionId $Deployment.CollectionID | Where-Object {($_.starttime -ge $YearMonthMin) -and ($_.starttime -le $YearMonthMax)}

   $collection = Get-CMCollection -CollectionId $Deployment.CollectionID
   

    foreach ($member in $collectionmembers)
            
                    {
                        
                    
                    $myobj = New-Object -TypeName PSObject
                    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Name" -Value $member.name
                    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Collection" -Value $collection.name
                    add-member -InputObject $myobj -MemberType NoteProperty -Name "MW Name" -Value $mw.Name
                    add-member -InputObject $myobj -MemberType NoteProperty -Name "StartTime" -Value $mw.starttime
                    add-member -InputObject $myobj -MemberType NoteProperty -Name "Duration" -Value $mw.Duration
                    $result += $myobj
                    }
}





$result | Format-Table


cd e:
