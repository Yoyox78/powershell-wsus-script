# Je fais appel à une fonction commune à plusieurs scripts
. ".\OuCharger.ps1"

# Je vide toutes les variables au cas ou elle sont deja utilisées, je met en silencieux les erreurs, car il risque d'indiquer que les varibles n'existe pas
clear-variable -Name "ServeurListe"-ErrorAction SilentlyContinue
clear-variable -Name "ListSRVError" -ErrorAction SilentlyContinue
clear-variable -Name "LogSrvRestart" -ErrorAction SilentlyContinue
clear-variable -Name "LogSrvInstall" -ErrorAction SilentlyContinue

# s'il a un argument, alors je crée la variable manuellement pour la tâche plannifier, le fonctionnement se fera par OU
if ($args[0] -ne $null)
{
    $readstart = "ou"
    $listConSrv = $args[0]
    $listConSrv = $listConSrv.Split(',')
}

#Je met le script en mode debug
$debug = 1

#J'initialise la variable dédiée aux erreurs
$ListSRVError = "Voici la liste des erreurs rencontrées, veuillez vérifier sur les serveurs si cela pourrait être corrigés `n"
$ListSRVError += "`n---------------------------------------------------------------------------------`n`n"

# Vérifie s'il a un argument
if ($args[0] -eq $null)
{
    $readstart = Read-Host "OU ou Poste (ou|poste)"
    if ($readstart -notmatch "ou|poste")
    {
        write-host "Veuillez indiquer OU ou POSTE lors de la demande, je quitte"
        exit
    }

    if ($readstart -eq "poste") 
    {
        $ServeurListe = @()
        Do{
            $StrSrvif = Read-Host "Veuillez indiquer le nom du serveur, si vous souhaitez vous arretez la appuyer sur entrée"
            if ($StrSrvif -ne "")
            {
                $ServeurListe += [pscustomobject]@{Name=$StrSrvif}
            }
        }Until($StrSrvif -eq "")
    }

    if ($readstart -eq "ou") 
    {
        Write-Host "OU dispo `n INTE PREPOD PROD TESTING"
        $listConSrv = Read-Host "Veuillez indiquer l'OU, si vous en indiquez plusieurs, veuillez mettre des **","** entre chaque OU (sans espace)"
        $listConSrv = $listConSrv.Split(',')
        if ($listConSrv -notmatch '^INTE$|^PREPOD$|^PROD$|^TESTING$')
        {
            write-host "Erreur, aucun élément correspondant pour le choix de l'OU (variable listConSrv) en debut de script, Veuillez entrer les bonnes valeurs. Je quitte"
            Exit
        }
    }
}
else
{
    if ($readstart -eq "ou") 
    {
        if ($listConSrv -notmatch '^INTE$|^PREPOD$|^PROD$|^TESTING$')
        {
            write-host "Erreur, L'argument donné lors de l'execution du script ne correspond pas au OU existante."
            write-host "Veuillez indiquer une ou plusieurs OU (séparées par une virgule) parmis le choix suivant : INTE PREPROD PROD TESTING"
            Exit
        }
    }
}

# fonction qui permet de concaténer la variable dédiée aux erreurs avec le nom de la machine et l'erreur rencontrée
function ErrReturn ($FuncErr, $FuncErrserveur)
{
    $FuncErrorLog +=  "Serveur : $FuncErrserveur`n"
    $FuncErrorLog +=  "`n---------------------------------------------------------------------------------`n"
    $FuncErrorLog +=  $FuncErr
    $FuncErrorLog +=  "`n---------------------------------------------------------------------------------`n`n"
    return $FuncErrorLog
}

# Je vide la variable dédiée aux erreurs pour ne pas avoir les erreurs précédentes
$Error.Clear()

# Je verifie si le module que je souhaite utiliser est bien installé
$ModInstallSrvRoot = Get-Module -ListAvailable -Name PSWindowsUpdate
#Si cela ne retourne pas d'erreur et que la variable est vide alors j'install le module
# Si la variable $ModInstallSrvRoot retourne une erreur alors je l'ajoute dans la variable des erreurs
if  (($Error.count -eq 0) -and ($ModInstallSrvRoot -eq $null))
{ 
    #j'installe les paquet requis et charge le module
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
    Install-Module PSWindowsUpdate -Confirm:$false -Force
    Import-module PSWindowsUpdate
}
elseif ($Error.count -ne 0)
{
    $RootSrvUse = $env:COMPUTERNAME
    $ListSRVError += ErrReturn $ErrorFuncSend $RootSrvUse
}

# Suivant les choix effectués, je crée une liste de serveur. J'exclu le nom des OU de la liste pour que ca ne soit pas pris pour un nom de serveur
if ($readstart -eq "ou")
{
    # j'utilise la fonction commune à plusieurs script (déclaré en debut de script)
    $ServeurListe = OuCharger $listConSrv, $null
}

# J'utilise la liste des serveurs pour faire li'nstallation du module, vérifier les mises à jour et les installer sans restart le serveur
# Le restart des serveurs se  fera avec le script PRTG
foreach ($serv in $ServeurListe)
{
    if (Test-Connection -ComputerName $serv.Name -Quiet)
    {
        if ($debug -eq 1)
        {
            $serv.name
        }

        $Error.Clear()

        $InvokeTest = Invoke-Command -ComputerName $serv.Name -ScriptBlock {Get-Module -ListAvailable -Name PSWindowsUpdate}
        if  (($Error.count -eq 0) -and ($InvokeTest -eq $null))
        { 
            $Error.Clear()
            Invoke-Command -ComputerName $serv.Name -ScriptBlock {Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force}
            if (Select-String -InputObject $Error[0].FullyQualifiedErrorId -Pattern 'NoMatchFoundForProvider' -Quiet) 
            {
                Invoke-Command -ComputerName $serv.Name -ScriptBlock {Set-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NetFramework\v4.0.30319" -Name "SchUseStrongCrypto" -Value "1" -Type DWord}
                Invoke-Command -ComputerName $serv.Name -ScriptBlock {Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\.NetFramework\v4.0.30319" -Name "SchUseStrongCrypto" -Value "1" -Type DWord}
                Invoke-Command -ComputerName $serv.Name -ScriptBlock {Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force}
                $Error.Clear()
            }
            Invoke-Command -ComputerName $serv.Name -ScriptBlock {Install-Module PSWindowsUpdate -Confirm:$false -Force}
            Invoke-Command -ComputerName $serv.Name -ScriptBlock {Import-module PSWindowsUpdate}
            sleep 5
            $VerifModuleInstall = Invoke-Command -ComputerName $serv.Name -ScriptBlock {Get-Module -ListAvailable -Name PSWindowsUpdate}
            if ($VerifModuleInstall -eq $null)
            {
                $ErrorFuncSend = $Error[0].Exception.message 
                $ListSRVError += ErrReturn $ErrorFuncSend $serv.Name
                Continue
            }
        }
        elseif ($Error.count -ne 0)
        {
            $ErrorFuncSend = $Error[0].Exception.message 
            $ListSRVError += ErrReturn $ErrorFuncSend $serv.Name
            Continue
        }

        $Error.Clear()
        # Je verifie la liste des mises à jour sur le serveur interrogé, dont les mise à jour exigent le redémarrage
        $LstUpdateReq = Get-WindowsUpdate -computername $serv.name

        # Si la variable $LstUpdateReq retourne une erreur, je la note et passe au serveur suivant
        if ($Error.count -ne 0)
        {
            $ErrorFuncSend = $Error[0].Exception.message 
            $ListSRVError += ErrReturn $ErrorFuncSend $serv.Name
            Continue
        }

        # On vérifie les MAJ en ignorant celle  qui ont besoin d'un redémarrage
        $LstUpdateNotReq = Get-WindowsUpdate -computername $serv.name -IgnoreRebootRequired
        if ($LstUpdateReq.count -ne 0)
        {
            # Je verifie que les mise à jour qui ont besoin d'un redémarrage ne sont pas les même que celle n'ayant pas besoin de redémarrage
            if ($LstUpdateReq.count -ne $LstUpdateNotReq.count) {
                $LogSrv += "Serveur : " + $serv.Name
                $LogSrv += "`n---------------------------------------------------------------------------------`n"
                $LogSrv += "############## Mises à jours ayant un besoin de redémarrage du serveur ##############"
                $LogSrv += "`n---------------------------------------------------------------------------------`n"
                $LogSrv += $LstUpdateReq | Where-Object {$_.KB -notin $LstUpdateNotReq.KB } | Out-String
                $LogSrv += "`n---------------------------------------------------------------------------------`n"
                $LogSrv += "########################### Installation mise(s) à jour ###########################" 
                $LogSrv += "`n---------------------------------------------------------------------------------`n"
                $LogSrv += Get-WindowsUpdate -computername $serv.name -Install -AcceptAll -Verbose -IgnoreReboot | Out-string
                $LogSrv += "`n---------------------------------------------------------------------------------`n"
            }
            else 
            {
                $LogSrv += "Serveur : " + $serv.Name
                $LogSrv += "`n---------------------------------------------------------------------------------`n"
                $LogSrv += "########################### Installation mise(s) à jour ###########################" 
                $LogSrv += "`n---------------------------------------------------------------------------------`n"
                $LogSrv += Get-WindowsUpdate -computername $serv.name -Install -AcceptAll -Verbose -IgnoreReboot | Out-string
                $LogSrv += "`n---------------------------------------------------------------------------------`n"
            }
        }
    }
    else
    { 
        $ErrorFuncSend = "La machine ne repond pas au ping"
        $ListSRVError += ErrReturn $ErrorFuncSend $serv.Name
        Continue
    }
}

# Si l'une des listes n'est pas vide, elle est convertit en caractère. On envoi le mail
if (($LogSrv -ne "") -or ($ListSRVError -ne ""))  {

    $LogSrv | out-string 
    $ListSRVError | out-string 

    if ($debug -eq 1)
    {
        Write-Host "------------------------------------------- Liste des mise à jours `n $LogSrv `n ------------------------------------------- Serveurs en erreurs `n $ListSRVError `n"   
    }
    else
    {
        #Encodage UTF8
        $encodingMail = [System.Text.Encoding]::UTF8 
        Send-MailMessage -From "EXPEDITEUR@TONDOM.FR" -To "DESTINATAIRE@TONDOM.FR" -Subject "Rapport des serveurs à restart" -SmtpServer "RELAISMTP(IP ou DNS)" -Body "------------------------------------------- Liste des mise à jour `n $LogSrv `n ------------------------------------------- Serveurs en erreurs `n $ListSRVError `n" -Encoding $encodingMail

    }

}
