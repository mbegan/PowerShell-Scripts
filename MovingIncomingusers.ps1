$UserBase = "OU=people,DC=ad,DC=domain,DC=tld"
$ProvBase = "OU=IncomingUsers,DC=ad,DC=domain,DC=tld"


########################################################################################################
# Grab all of the OU objects below the 'People' structure that are called 'New User'                   #
# Parse the level up OU Name from each 'New User' OU to derive the primary city value for that OU      #
# Evaluate the 'city' attribute for each 'New User' OU to derive the alternate city values for that OU #
########################################################################################################
$ous = Get-ADOrganizationalUnit -SearchBase $UserBase -Filter {name -eq 'New User'} -ResultSetSize $null
$destinations = New-Object System.Collections.Hashtable
$siteOUs = New-Object System.Collections.Hashtable
foreach ($ou in $ous)
{
    $parts = $ou.DistinguishedName.Split(',')
    $city = ($parts[1].Split("=")[1]).ToLower().Replace(" ","")
    if (!$destinations[$city])
    {
        $destinations[$city] = New-Object System.Collections.Hashtable
        $destinations[$city].OU = $ou
        $destinations[$city].Users = New-Object System.Collections.ArrayList
    }
    $parent = $parts[1..($parts.Count-1)] -join ","
    if (!$siteOUs[$parent])
    {
        $siteOUs[$parent] = New-Object System.Collections.Hashtable
        $siteOUs[$parent].ou = Get-ADOrganizationalUnit -Identity $parent
        $siteOUs[$parent].cities = New-Object System.Collections.ArrayList
        $siteOUs[$parent].CurrentUsers = New-Object System.Collections.ArrayList
        $siteOUs[$parent].notBelong = New-Object System.Collections.ArrayList
    }
    $_c = $siteOUs[$parent].cities.Add($city)
    if ($ou.City)
    {
        $cities = $ou.City.Split(";")
        foreach ($city in $cities)
        {
            $city = $city.ToLower().Replace(" ","")
            $_c = $siteOUs[$parent].cities.Add($city)
            if (!$destinations[$city])
            {
                $destinations[$city] = New-Object System.Collections.Hashtable
                $destinations[$city].OU = $ou
                $destinations[$city].Users = New-Object System.Collections.ArrayList
            }
        }
    }
}


########################################################################################################
# Grab all of the users from the Incoming Users OU                                                     #
# Find their destination OU and put them in bucket to be moved into that OU in the next step           #
# Place users that dont have a city or don't match a derived city into the nomatches bucket for alert  #
########################################################################################################
$nomatches = New-Object System.Collections.ArrayList
$users = Get-ADUser -Filter { ( ((EmployeeType -eq 'SAPHR') -and (Enabled -eq $true) ) -and ( (vmsHRStatus -eq 3) -or (vmsHRStatus -eq 1) ) ) } -SearchBase $ProvBase -ResultSetSize $null -Properties country,city,state,employeeType,vmshrstatus,whencreated,mail,vmsUserStartDate -SearchScope Subtree
foreach ($u in $users)
{
    if (!$u.City)
    {
        #User didn't have a city value, skip and alert
        $_c = $nomatches.Add($u)
        continue
    }

    #If there is a comma in the city take the first portion, remove spaces and cast to lowercase
    $dest = ($u.City.split(","))[0].ToLower().Replace(' ','')
    if ($destinations[$dest])
    {
        $_c = $destinations[$dest].Users.Add($u)
        continue
    }

    #If we don't find a match for the user, skip and alert
    $_c = $nomatches.Add($u)
}

########################################################################################################
# Parse through the destination buckets                                                                #
# Check to see if we are creating a conflict when we move the user into the dest OU if so rename CN    #
# Move users into their destination OU                                                                 #
########################################################################################################
$collisions = New-Object System.Collections.ArrayList
foreach ($dest in $destinations.Keys)
{
    if ($destinations[$dest].Users.Count -ge 1)
    {
        write-host $dest `t $destinations[$dest].Users.Count
        foreach ($u in $destinations[$dest].Users)
        {
            
            try
            {
                $x = Get-ADUser -Identity ( "CN=" + $u.Name + "," + $destinations[$dest].OU.DistinguishedName )
            }
            catch
            {
                $x = $false
            }
            if ( $x )
            {
                $collision = @{old = $x;new = $u}
                $_c = $collisions.Add($collision)
                $newName = $u.Name + " (" + $u.SamAccountName +")"
                try
                {
                    Rename-ADObject -Identity $u.ObjectGUID.Guid -NewName $newName
                }
                catch
                {
                    Write-Host $_.Exception
                    continue
                }
                sleep -Seconds 2
            }

            try
            {
                Move-ADObject -Identity $u.ObjectGUID.Guid -TargetPath $destinations[$dest].OU.DistinguishedName
            }
            catch
            {
                Write-Host $_.Exception.GetType().FullName -BackgroundColor White -ForegroundColor Red
                continue
            }
        }
    }
}

########################################################################################################
# Parse through the siteOUs buckets                                                                    #
# Check to see if the users that are in the given siteOUs have a city value that lines up with it      #
# Move users into their destination OU                                                                 #
########################################################################################################
foreach ($site in $siteOUs.Keys)
{
    Write-Host checking $siteOUs.$site.ou.DistinguishedName
    $siteOUs[$site].CurrentUsers = Get-ADUser -Filter {enabled -eq $true -and employeeType -eq 'SAPHR'} -Properties city -SearchBase $siteOUs[$site].ou.DistinguishedName
    foreach ($cu in $siteOUs[$site].CurrentUsers)
    {
        #If there is a comma in the city take the first portion, remove spaces and cast to lowercase
        if ( $cu.City -ne $null )
        {
            try
            {
                $city = ($cu.City.split(","))[0].ToLower().Replace(' ','')
            }
            catch
            {
                $city = 'blank'
            }
        } else {
            $city = 'blank'
        }
        
        if (! $siteOUs[$site].cities.Contains($city) )
        {
            Write-Host `t $cu.Name of $city does not belong in $site -BackgroundColor Red
            $_c = $siteOUs[$site].notBelong.Add($cu)
        }
    }
}

foreach ($site in $siteOUs.Keys)
{
    Write-Host $site $siteOUs[$site].notBelong.Count
}

$emailto = "persontonotify@domain.tld"
$emailFrom = "peoplemover@noreply.domain.tld"
$smptSrv = 'smtp.domain.tld'
if ($nomatches.Count -ge 1)
{
    $emailSubject = “People Mover has people without destinations " + $dateTime + " Running From " + $env:COMPUTERNAME
    $body = "No Destination for `r`n"
    foreach ($u in $nomatches)
    {
        $line = "`t " + $u.Name +" : " + $u.City + "`r`n"
        $body += $line
    }
    Send-MailMessage -To $emailto -From $emailFrom -Subject $emailSubject -SmtpServer $smptSrv -Body $body
}

if ($collisions.Count -ge 1)
{
    $emailSubject = “People Mover has detected a collision " + $dateTime + " Running From " + $env:COMPUTERNAME
    $body = "Collisions detected for `r`n"
    foreach ($u in $collisions)
    {
        $line = "`t" + $u.new.Name + " (" + $u.new.samAccountName + ") conflicted with " + $u.old.Name +" (" + $u.old.samAccountName + ") `r`n"
        $body += $line
    }
    Send-MailMessage -To $emailto -From $emailFrom -Subject $emailSubject -SmtpServer $smptSrv -Body $body
}