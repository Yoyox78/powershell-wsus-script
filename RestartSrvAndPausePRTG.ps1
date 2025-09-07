# Je fais appel a une fonction commune aux deux script
. “.\OuCharger.ps1”

#Le mode debug sert à avoir plus d'information, il remplace l'envoi de mail par un affichage dans la console
$debug = 0

#On vide la variable
clear-variable -Name "ServeurListe" -ErrorAction SilentlyContinue

# s'il a un argument, alors je crée le variable manuellement pour la tâche plannifier, le fonctionnement se fera par OU
# l'argument en attente et une ou plusieurs OU séparé par des ,
if ($args[0] -ne $null)
{
    $readstart = "ou"
    $listConSrv = $args[0]
    $listConSrv = $listConSrv.Split(',')
}

if (($debug -ne 1) -and ($debug -ne 0))
{
    Write-Host "Veuillez mettre la variable debug à 1 ou 0, je quitte"
    exit
}

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
            $StrSrvif = Read-Host "Veuillez indiquer le nom du serveur, si vous souhaitez vous arreter la, appuyer sur entrée"
            if ($StrSrvif -ne "")
            {
                $ServeurListe += [pscustomobject]@{Name=$StrSrvif}
            }
        }Until($StrSrvif -eq "")
    }

    if ($readstart -eq "ou") 
    {
        Write-Host "OU dispo `n INTE/PREPROD/PROD/TESTING"
        $listConSrv = Read-Host "Veuillez indiquer l'OU, si vous en indiquez plusieurs, veuillez mettre des , sans espace"
        $listConSrv = $listConSrv.Split(',')
        $listConSrv
    }

    if ($readstart -eq "ou") 
    {
        if ($listConSrv -notmatch '^INTE$|^PREPROD$|^PROD$|^TESTING$')
        {
            write-host "Erreur, aucun élément correspondant pour le choix de l'OU (variable listConSrv en debut de script), Veuillez entrer les bonnes valeurs. Je quitte"
            Exit
        }
    }
}
else
{
    if ($readstart -eq "ou") 
    {
        if ($listConSrv -notmatch '^INTE$|^PREPROD$|^PROD$|^TESTING$')
        {
            write-host "Erreur, L'argument donné lors de l'execution du script ne correspond pas aux OU existante."
            write-host "Veuillez indiquer une ou plusieurs OU séparé par une virgule (si plusieurs) parmis le choix suivant : INTE PREPROD PROD TESTING"
            Exit
        }
    }
}

# Cette fonction permet de créer la tâche plannifier, attention les vieux serveur n'on pas cette fonction powershell
function Task ($SRVName, $TaskTime)
{
    $TaskExec= New-ScheduledTaskAction -Execute "shutdown.exe" -Argument "/r /t 0 /f"
    $TaskRun = New-ScheduledTaskTrigger -Once -At $TaskTime"pm"
    $TaskUser = "NT AUTHORITY\SYSTEM"
    Invoke-Command -ComputerName $SRVName -ScriptBlock {Register-ScheduledTask -TaskName $Using:NameTask -Trigger $Using:TaskRun -Action $Using:TaskExec -User $Using:TaskUser -RunLevel Highest –Force -Description "Restart le serveur, car en attente de redemarrage"}
}

# Fonction qui permet d'installer des modules, s'il sont manquant
function Module ($Module)
{
    if (-Not (Get-Module -ListAvailable -Name $Module)) 
    {
        Install-Module $Module -Confirm:$false -Force
    }
}

# Installation des modules dont on va avoir besoin
Module "PrtgAPI"
Module "PendingReboot"

# Fonction permettant de mettre en pause les sondes PRTG des serveurs aux heures prévues
function PRTGSched ($Srv, $h, $m, $PRTGopt)
{   
    $ReturnPRTGLog = $null 
    $PRTGParams = $null
    $PRTGId = $null

    # Identifiant de connexion PRTG
    $PRTGLogin = "USERPRTG"
    $PRTGHash = "HASHPASSWORD"
    $PRTGServeur = "https://URLSUPERVISION.LAN"

    # Connexion au serveur PRTG
    Connect-PrtgServer -Server $PRTGServeur (New-Credential $PRTGLogin $PRTGHash) -PassHash

    #Recupération de l'ID de la machine sur PRTG
    $PRTGId = Get-Device $Srv

    if ($PRTGId -ne $null)
    {
        # Suivant la valeur de $PRTGopt, on alimente les variables avec ce que l'on a besoin 
        switch ( $PRTGopt )
        {
            # Cette partie crée la planification, pour les serveurs 
            Schedule 
            {
                $PRTGStart = (get-date -Hour $h -Minute $m -Second 0)
                $PRTGEnd = $PRTGStart.AddMinutes(5)
                $PRTGEnd = $PRTGEnd.ToString("yyyy-MM-dd-HH-mm-ss")
                $PRTGStart = $PRTGStart.ToString("yyyy-MM-dd-HH-mm-ss")

                #api https://www.paessler.com/manuals/prtg/application_programming_interface_api_definition
                $PRTGParams = @{
                    "scheduledependency" = 0
                    "maintenable_" = 1
                    "maintstart_" = $PRTGStart
                    "maintend_" = $PRTGEnd
                }

            }
        }

        # l'appel est faite à PRTG avec les éléments dont il a besoin
        Get-Device -Id $PRTGId.id | Set-ObjectProperty -RawParameters $PRTGParams -Force
        # Je me déconnecte de PRTG
        Disconnect-PrtgServer
    }
    else {
        $ReturnPRTGLog = "True"
        # Je me déconnecte de PRTG
        Disconnect-PrtgServer
        return $ReturnPRTGLog
    }
}

#On recupère la liste des serveurs que l'on va exclure des tâches à effectuer (une ligne par serveur dans le fichier)
$ExcludeSrv = Get-Content ".\RestartSRVTachePlannifierPRTG-ListExcludeSrv.txt"

if ($readstart -eq "ou")
{
    # j'utilise la fonction commune aux deux scripts (déclarées en debut de script)
    $ServeurListe = OuCharger $listConSrv $ExcludeSrv
}

# On défini les variables de temps et du nom de la tâche qui sera crée
$Minute=55
#heure et hour sont crée car les serveurs sont en horaires anglaise alors que PRTG en francais
$Hour=7
$Heure=19
$GardeFou=0
$NameTask = "RebootSRV001"
$ListSRV = ""
$ListSRVError = ""
$ListSRVNotReboot = ""

# Gestion serveur par serveur
foreach ($serv in $ServeurListe)
{
    if (Test-Connection -ComputerName $serv.Name -Quiet)
    {
        # On verifie si le serveur est en attente de redémarrage est on met les warning dans la variable $WarningPending
        switch ((Test-PendingReboot -ComputerName $serv.Name -SkipConfigurationManagerClientCheck -WarningVariable WarningPending).IsRebootPending)
        {
            True
            {

                if ($debug -eq 1)
                {
                    Write-Host "--------------------------------------------------------------------------------------------"
                    Write-Host "Vous etes dans le switch True (en attente de redémarrage)"
                    Write-Host "Nom du serveur : " $serv.Name
                    Write-Host "debug : " $debug
                    Write-Host "warning : " $WarningPending
                    Write-Host "Erreur : " $Error[0].Exception.message
                    Write-Host "--------------------------------------------------------------------------------------------"
                }
                # On verifie si le garde-fou est activé. Le garde-fou sert à éviter de dépasser la plage de redémarrage 20h > 23h30
                if ($GardeFou -ne 1) 
                {

                    #On reinitialise la variable $Error
                    $Error.Clear()

                    # On verifie si la tache plannifier existe
                    $ReturnTask = Invoke-Command -ComputerName $serv.Name -ScriptBlock {Get-ScheduledTask | where-object {$_.TaskName -like $Using:NameTask}}

                    #si $ReturnTask retourne une erreur alors on quitte la boucle et on met le serveur en erreur en indiquant pourquoi
                    if ($Error.count -ne 0)
                    {
                        $ListSRVError +=  "Serveur : " + $serv.Name + "`n"
                        $ListSRVError +=  "`n-------------------------- Voici le message d'erreur ---------------------------------------`n"
                        $ListSRVError +=  $Error[0].Exception.message
                        $ListSRVError +=  "`n-------------------------- Fin du message d'erreur -----------------------------------------`n`n"
                        Continue
                    }

                    # le serveur doit être restart, j'implémente le nombre de minute et converti en heure si besoin
                    $Minute += 5
                    if (($Minute -eq 30) -and ($Hour -eq 11)) 
                    {
                        $GardeFou=1
                        $ListSRV += "############### RESTANT A RESTART PLUS TARD (HORS DELAI)`n"
                    }

                    if ($Minute -eq 60) 
                    {
                        $Minute = 0
                        $Hour += 1
                        $Heure += 1
                    }

                    # Si la variable n'est pas null 
                    if ($ReturnTask -ne $null)
                    {
                        # Alors on supprime la tâche, $using sert à utiliser des variables qui sont en local (dans le script)
                        Invoke-Command -ComputerName $serv.Name -ScriptBlock {Unregister-ScheduledTask $Using:NameTask -Confirm:$false}
                        # Lancement de la fonction avec l'heure défini actuellement
                        Task $serv.Name $Hour":"$Minute
                        # La sonde PRTG est mise en pause à l'heure prévu
                        $ReturnPRTGLog = PRTGSched $serv.Name $Heure $Minute "Schedule"
                    }
                    else
                    {
                        # Si aucune tâche n'est présente alors on lance la fonction avec l'heure défini actuellement
                        Task $serv.Name $Hour":"$Minute
                        # Je met la sonde PRTG en pause à l'heure prévu
                        $ReturnPRTGLog = PRTGSched $serv.Name $Heure $Minute "Schedule"
                    }

                    if ($ReturnPRTGLog -ne $null)
                    {
                       # Incrémentation de la liste des serveurs ayant besoin d'être redémarré
                        $ListSRV += $serv.Name + " restart à  " + $Heure + ":" + $Minute + "`n"     
                    }
                    else 
                    {
                        $ListSRV += $serv.Name + " restart à  " + $Heure + ":" + $Minute + " Pas de sonde PRTG trouvé, Suspension de la sonde non definie `n"
                    }
                }
                else 
                {
                    # Si le garde-fou est activé, alors on note les serveurs en attente de redémarrage. On ne plannifie pas de tâche car on n'a plus de temps disponible
                    $ListSRV += $serv.Name + " restart à  " + $Heure + ":" + $Minute + "`n"
                }
            }
            False
            {
                # On liste les serveurs qui n'ont pas besoin d'être redémarrer
                $ListSRVNotReboot +=  $serv.Name + "`n"
                if ($debug -eq 1)
                {
                    Write-Host "--------------------------------------------------------------------------------------------"
                    Write-Host "Vous etes dans le switch False (Pas besoin de restart)"
                    Write-Host "Nom du serveur : " $serv.Name
                    Write-Host "debug : " $debug
                    Write-Host "warning : " $WarningPending
                    Write-Host "--------------------------------------------------------------------------------------------"
                }

            }

            default
            {
                #Si le warning du switch Test-PendingReboot est egal a l'erreur RPC je l'indique
                if (Select-String -InputObject $WarningPending -Pattern '0x800706BA' -Quiet) 
                {
                    $ListSRVError +=  "Serveur : " + $serv.Name + "`n"
                    $ListSRVError +=  "`n-------------------------- Voici le message d'erreur ---------------------------------------`n"
                    $ListSRVError +=  "N'arrive pas à se connecter à la machine (RPC) `n"
                    $ListSRVError +=  "`n-------------------------- Fin du message d'erreur -----------------------------------------`n`n"                                                    
                }
                else 
                {
                    # On note toutes les autres erreurs et indique les serveurs comme étant en erreurs
                    $ListSRVError +=  "Serveur : " + $serv.Name + "`n"
                    $ListSRVError +=  "`n-------------------------- Voici le message d'erreur ---------------------------------------`n"
                    $ListSRVError +=  "$WarningPending `n"
                    $ListSRVError +=  "`n-------------------------- Fin du message d'erreur -----------------------------------------`n`n"

                }
                #Sert de test pour savoir ou resort le warning
                if ($debug -eq 1)
                {
                    Write-Host "--------------------------------------------------------------------------------------------"
                    Write-Host "Vous etes dans le switch default (autre cas que redémarrage ou pas de redémarrage)"
                    Write-Host "Nom du serveur : " $serv.Name
                    Write-Host "debug : " $debug
                    Write-Host "warning : " $WarningPending
                    Write-Host "error : " $Error[0].FullyQualifiedErrorId
                    Write-Host "--------------------------------------------------------------------------------------------"
                }

            }
        }
    }
    else
    {
        $ListSRVError +=  "Serveur : " + $serv.Name + "`n"
        $ListSRVError +=  "`n-------------------------- Voici le message d'erreur ---------------------------------------`n"
        $ListSRVError +=  "N'arrive pas à se connecter à la machine (PING)  `n"
        $ListSRVError +=  "`n-------------------------- Fin du message d'erreur -----------------------------------------`n`n"
    }
}

# Si l'une des listes n'est pas vide, elle est convertit en caractère. On envoi le mail
if (($ListSRV -ne "") -or ($ListSRVError -ne "") -or ($ListSRVNotReboot -ne ""))  {

    $ListSRV | out-string 
    $ListSRVError | out-string 
    $ListSRVNotReboot | out-string

    if ($debug -eq 1)
    {
        Write-Host "------------------------------------------- Serveurs qui n'ont pas besoin de redémarrage `n $ListSRVNotReboot `n ------------------------------------------- Serveurs en erreurs `n $ListSRVError `n ------------------------------------------- Serveurs qui vont redémarrer ce soir `n $ListSRV "   
    }
    else
    {
        #Encodage UTF8
        $encodingMail = [System.Text.Encoding]::UTF8 
        Send-MailMessage -From "EMAILEXPEDITEUR@TONDOM.FR" -To "DESTINATAIRE1@TONDOM.FR", "DESTINATAIRE2@TONDOM.FR" -Subject "Rapport des serveurs à restart" -SmtpServer "SERVEURSMTPRELAI" -Body "------------------------------------------- Serveurs qui n'ont pas besoin de redémarrage `n $ListSRVNotReboot `n ------------------------------------------- Serveurs en erreurs `n $ListSRVError `n ------------------------------------------- Serveurs qui vont redémarrer ce soir `n $ListSRV " -Encoding $encodingMail
    }

}
