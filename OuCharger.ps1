#Fonction commune à plusieurs scripts

function OuCharger ($listOUSrv, $ExcludeSrvList)
{
    foreach ( $i in $listOUSrv )
    {
        switch ( $i )
        {
            'INTE'
            {
                $ServeurListe += @(Get-ADObject -LDAPFilter "(objectClass=Computer)" -SearchBase 'OU=INTE,OU=Serveur,DC=TONDOM,DC=lan' | Where-Object {$_.Name -notin $ExcludeSrvList} | Select-Object Name)
            }
            'PREPROD'
            {
                $ServeurListe += @(Get-ADObject -LDAPFilter "(objectClass=Computer)" -SearchBase 'OU=PREPROD,OU=Serveur,DC=TONDOM,DC=lan' | Where-Object {$_.Name -notin $ExcludeSrvList} | Select-Object Name)
            }
            'PROD'
            {
                $ServeurListe += @(Get-ADObject -LDAPFilter "(objectClass=Computer)" -SearchBase 'OU=PROD,OU=Serveur,DC=TONDOM,DC=lan' | Where-Object {$_.Name -notin $ExcludeSrvList} | Select-Object Name)
            }
            'TESTING'
            {
                $ServeurListe += @(Get-ADObject -LDAPFilter "(objectClass=Computer)" -SearchBase 'OU=TESTING,OU=Serveur,DC=TONDOM,DC=lan' | Where-Object {$_.Name -notin $ExcludeSrvList} | Select-Object Name)
            }
            default
            {
                Write-Information 'Veuillez remplir la variable$listOUSrv correctement'
                Exit
            }

        }

    }
    return $ServeurListe
}
