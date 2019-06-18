###########################################################################################
# Script EAC Alca
# Auteur : Henrotte Alexandre
# 
# Options du script : 1) Créer un utilisateur
#
#                     2) Créer des utilisateurs à l'aide d'un fichier csv
#
#                     3) Changer un utilisateur de classe
#
#                     4) Supprimer un utilisateur (avec sauvegarde des documents)
#
#                     5) Effectuer une sauvegarde
#
#                     6) Ouvrir la page github du script
#
#
# NB : - Le script doit obligatoirement être exécuté à partir d'un Windows Serveur
#      - Le script à la capacité de gérer les erreurs qui peuvent arriver
#
# Réalisé dans le cadre du projet de find d'année (EAC5)
#
#                                                                  Année Scolaire 2018-2019
###########################################################################################
# 
# Adresse mail : alexandre.henrotte1@gmail.com
#
# Github : https://www.github.com/EAC-Alex
#
###########################################################################################

################################ VERIFICATION DE L'OS######################################

Clear-Host

Write-Progress -Activity "Vérification du système d'exploitation" -Status "Veuillez patienter..." # Superflue mais montre à l'utilisateur qu'une vérification du système d'exploitation se fait

Start-Sleep 3
$OS = Get-WmiObject -Class Win32_OperatingSystem | ForEach-Object -MemberName Caption

Write-Progress -Activity "Vérification du système d'exploitation" -Status "Terminé" -Completed

if ($OS -notlike "*Windows*Server*") 
 {
  Write-Host -ForegroundColor Red "Le script doit obligatoirement être exécuté à parti d'un Windows Serveur !"
  Start-Sleep 3
  Write-host -ForegroundColor Yellow "`n`nExtinction du script " -NoNewline
      Start-Sleep 1.2
      Write-host -ForegroundColor Yellow "." -NoNewline
      Start-Sleep 1.2
      Write-host -ForegroundColor Yellow "." -NoNewline
      Start-Sleep 1.2
      Write-host -ForegroundColor Yellow "." -NoNewline
      Start-Sleep 1.2
      Write-host -ForegroundColor Yellow "." -NoNewline
      exit
 }


###########################################################################################

############## TEST DE CONNEXION A INTERNET ET INSTALLATION DES MODULES ###################

$erroractionpreference = "SilentlyContinue"

$sortie = "non"
while ($sortie -ne "oui")
{

    Write-Progress -Activity "Test de connexion à internet" -Status "Veuillez patienter..."
    $testinternet = Test-NetConnection
    Write-Progress -Activity "Test de connexion à internet" -Status "Terminé" -Completed

    if ($testinternet.PingSucceeded -eq $False)
    {
     Clear-Host
     Write-host -ForegroundColor Red "`nLe test de connexion à internet a échoué !`n" -NoNewline

     for ($i = 1; $i -lt 100; $i++)
      {
       $userchoix = Read-Host "`nVoulez vous continuer sans internet (risque d'erreur accru) ? (oui/non)"
       if ($userchoix -eq "oui")
       {
        Clear-Host
        break
       }
       elseif ($userchoix -eq "non")
            {Write-host "`n`n`nMerci d'avoir utilisé le script officiel du domaine Alca.tqi, à bientôt !" -ForegroundColor Cyan
             Start-Sleep 5
             exit}
       else
        {
         Write-Host -ForegroundColor Red "`nVeuillez répondre par oui ou par non" # On continue de poser la question
        }

       } # For

     } # If


    else
    {
        Clear-Host
        Write-Progress -Activity "Installation des modules" -Status "1/3"
        Install-WindowsFeature AD-Domain-Services | Out-Null -ErrorAction SilentlyContinue
        Write-Progress -Activity "Installation des modules" -Status "2/3"
        Install-WindowsFeature -Name "DHCP" -IncludeManagementTools | Out-Null -ErrorAction SilentlyContinue
        Write-Progress -Activity "Installation des modules" -Status "3/3"
        Install-Module NtfSsecurity | Out-Null -ErrorAction SilentlyContinue
        Write-Progress -Activity "Installation des modules" -Status "Terminé" -Completed
    }

 $sortie = "oui"

} # While

$erroractionpreference = "Continue"

###########################################################################################

########################## VERIFICATION DES OU ET DES GROUPES #############################


    <#Si la condition est vrai, une erreur va apparaître, mais c'est justement ca qui fait la condition donc il n'y a pas besoin d'afficher les messages d'erreurs#>

    $erroractionpreference = "SilentlyContinue"

    if (-not (dsquery ou "OU=Eleves,DC=alca,DC=tqi"))
          {
          New-ADOrganizationalUnit "Eleves" -path "DC=alca, DC=tqi"
          }

    if (-not (dsquery ou "OU=Profs,DC=alca,DC=tqi"))
          {
          New-ADOrganizationalUnit "Profs" -path "DC=alca, DC=tqi"
          }

    if (-not (dsquery group "CN=Ensemble_Des_Eleves(NTFS),OU=Eleves,DC=alca,DC=tqi"))
          {
          New-ADGroup -Name "Ensemble_Des_Eleves(NTFS)" -SamAccountName "Ensemble_Des_Eleves(NTFS)"  -GroupCategory Security -GroupScope Global -path "OU=Eleves,DC=alca,DC=tqi"
          }

    if (-not (dsquery group "CN=Ensemble_Des_Professeurs(NTFS),OU=Profs,DC=alca,DC=tqi"))
          {
          New-ADGroup -Name "Ensemble_Des_Professeurs(NTFS)" -SamAccountName "Ensemble_Des_Professeurs(NTFS)"  -GroupCategory Security -GroupScope Global -path "OU=Profs,DC=alca,DC=tqi"
          }

    if (-not (dsquery group "CN=Internet_Eleves,OU=Eleves,DC=alca,DC=tqi"))
          {
          New-ADGroup -Name "Internet_Eleves" -SamAccountName "Internet_Eleves"  -GroupCategory Security -GroupScope Global -path "OU=Eleves,DC=alca,DC=tqi"
          }

    if (-not (dsquery group "CN=Internet_Profs,OU=Profs,DC=alca,DC=tqi"))
          {
          New-ADGroup -Name "Internet_Profs" -SamAccountName "Internet_Profs"  -GroupCategory Security -GroupScope Global -path "OU=Profs,DC=alca,DC=tqi"
          }

    if (-not (dsquery ou "OU=Classe,DC=alca,DC=tqi"))
         {
          New-ADOrganizationalUnit "Classe" -path "DC=alca, DC=tqi"

          $CsvData = Import-Csv -path "C:\Users\Administrateur\Documents\ListeEAC.csv" -Delimiter ";"

          foreach ($i in $CsvData)
          {
          $nomclasse = $i.Classe
          $parentou = "OU=Classe, DC=alca, DC=tqi"
          New-ADGroup -Name "$nomclasse" -SamAccountName "$nomclasse" -GroupCategory Security -GroupScope Global -path "$parentou"
          }
         }

    $erroractionpreference = "Continue"


#######################################################################################

########################## DECLARATION DES FONCTIONS ##################################

Function CreationDossier ($String, $name, $firstname)

    {

      if (-not (Test-path "\\SERVEUR\Dossier_Ecole\$String\Dossier_Document\$name`_$firstname"))
      {
      New-item -ItemType directory -Path "\\SERVEUR\Dossier_Ecole\$String\Dossier_Document\$name`_$firstname" -ErrorAction SilentlyContinue
      }

    }

Function DroitsNTFS ($String, $String2, $name, $firstname)

    {
      # Dossier_Ecole

      Get-item "\\SERVEUR\Dossier_Ecole" | Disable-NTFSAccessInheritance

      Clear-NTFSAccess -Path "\\SERVEUR\Dossier_Ecole"

      Add-NTFSAccess -Path "\\SERVEUR\Dossier_Ecole" -Account "Administrateur" -AccessRights FullControl

      Add-NTFSAccess -Path "\\SERVEUR\Dossier_Ecole" -Account "Ensemble_Des_Eleves(NTFS)" -AccessRights ReadAndExecute

      Add-NTFSAccess -Path "\\SERVEUR\Dossier_Ecole" -Account "Ensemble_Des_Professeurs(NTFS)" -AccessRights ReadAndExecute

      # Partage_Eleves (avec -ErrorAction SilentlyContinue pour éviter une erreur si encore aucun élève ou professeur n'a été créé)

      Get-item "\\SERVEUR\Dossier_Ecole\Partage_Eleves" -ErrorAction SilentlyContinue | Disable-NTFSAccessInheritance 

      Clear-NTFSAccess -Path "\\SERVEUR\Dossier_Ecole\Partage_Eleves" -ErrorAction SilentlyContinue

      Add-NTFSAccess -Path "\\SERVEUR\Dossier_Ecole\Partage_Eleves" -Account "Administrateur" -AccessRights FullControl -ErrorAction SilentlyContinue

      Add-NTFSAccess -Path "\\SERVEUR\Dossier_Ecole\Partage_Eleves" -Account "Ensemble_Des_Eleves(NTFS)" -AccessRights ReadAndExecute -ErrorAction SilentlyContinue

      Add-NTFSAccess -Path "\\SERVEUR\Dossier_Ecole\Partage_Eleves" -Account "Ensemble_Des_Professeurs(NTFS)" -AccessRights ReadAndExecute -ErrorAction SilentlyContinue

      # Partage_Profs (avec -ErrorAction SilentlyContinue pour éviter une erreur si encore aucun élève ou professeur n'a été créé)

      Get-item "\\SERVEUR\Dossier_Ecole\Partage_Profs" -ErrorAction SilentlyContinue | Disable-NTFSAccessInheritance

      Clear-NTFSAccess -Path "\\SERVEUR\Dossier_Ecole\Partage_Profs" -ErrorAction SilentlyContinue

      Add-NTFSAccess -Path "\\SERVEUR\Dossier_Ecole\Partage_Profs" -Account "Administrateur" -AccessRights FullControl -ErrorAction SilentlyContinue

      Add-NTFSAccess -Path "\\SERVEUR\Dossier_Ecole\Partage_Profs" -Account "Ensemble_Des_Eleves(NTFS)" -AccessRights None -ErrorAction SilentlyContinue

      Add-NTFSAccess -Path "\\SERVEUR\Dossier_Ecole\Partage_Profs" -Account "Ensemble_Des_Professeurs(NTFS)" -AccessRights ReadAndExecute -ErrorAction SilentlyContinue
      
      # Dossier_Document $String

      Get-item "\\SERVEUR\Dossier_Ecole\$String\Dossier_Document\$name`_$firstname" | Disable-NTFSAccessInheritance

      Clear-NTFSAccess -Path "\\SERVEUR\Dossier_Ecole\$String\Dossier_Document\$name`_$firstname" 

      Add-NTFSAccess -Path "\\SERVEUR\Dossier_Ecole\$String\Dossier_Document\$name`_$firstname" -Account "$name`_$firstname@alca.tqi" -AccessRights FullControl

      Add-NTFSAccess -Path "\\SERVEUR\Dossier_Ecole\$String\Dossier_Document\$name`_$firstname" -Account "Administrateur" -AccessRights FullControl

      Add-NTFSAccess -Path "\\SERVEUR\Dossier_Ecole\Partage_Eleves\Dossier_Document\$name`_$firstname" -Account "Ensemble_Des_Professeurs(NTFS)" -AccessRights ReadAndExecute -ErrorAction SilentlyContinue

      # Profil_Itinérant $String

      Get-item "\\SERVEUR\Dossier_Ecole\$String\Profil_Itinérant" | Disable-NTFSAccessInheritance

      Clear-NTFSAccess -Path "\\SERVEUR\Dossier_Ecole\$String\Profil_Itinérant" 

      Add-NTFSAccess -Path "\\SERVEUR\Dossier_Ecole\$String\Profil_Itinérant" -Account "$String2" -AccessRights Modify

      Add-NTFSAccess -Path "\\SERVEUR\Dossier_Ecole\$String\Profil_Itinérant" -Account "Administrateur" -AccessRights FullControl

    }

Function CréationDesCours($name, $firstname, $cours, $ligne)

    {
      
      if (-not (Test-path "\\SERVEUR\Dossier_Ecole\Partage_Eleves\Dossier_Document\$name`_$firstname\Cours\Structure"))
      {

          $cours = "Math","Français","FHG","FSE","AIP","OS","Sciences","Réseaux","Structure","Anglais","Religion"

          foreach ($ligne in $cours)
          {
           New-item -ItemType directory -Path "\\SERVEUR\Dossier_Ecole\Partage_Eleves\Dossier_Document\$name`_$firstname\Cours\$ligne" | Out-Null
          }

      }
           
    }

Function VerificationDossierEtPartage ($String)

    {

      # Partie Eleves

      if (-not (Test-path "C:\Dossier_Ecole\$String"))
      {
      New-item -ItemType directory -Path "C:\Dossier_Ecole\$String" | Out-Null
      }

      if (-not (Test-path "C:\Dossier_Ecole\$String\Dossier_Document"))
      {
      New-item -ItemType directory -Path "C:\Dossier_Ecole\$String\Dossier_Document" | Out-Null
      }

      if (-not (Test-path "C:\Dossier_Ecole\$String\Profil_Itinérant"))
      {
      New-item -ItemType directory -Path "C:\Dossier_Ecole\$String\Profil_Itinérant" | Out-Null
      }

      if(!(Get-SMBShare -Name "Dossier_Ecole" -ErrorAction SilentlyContinue)){
      New-SmbShare -Name "Dossier_Ecole" -Path "C:\Dossier_Ecole" -FullAccess "Tout le monde" | Out-Null
      }

    }

Function VerifUtilisateurPresent ($String, $name, $firstname)

    {

      $verif = dsquery user -samid $name`_$firstname
      if ($null -eq $verif )
      {
      Remove-item -Path "\\SERVEUR\Dossier_Ecole\$String\Dossier_Document\$name`_$firstname" -Force
      Remove-item -Path "\\SERVEUR\Dossier_Ecole\$String\Profil_Itinérant\$name`_$firstname" -Force
      }

    }

Function Internet

    {

      for ($i = 1; $i -lt 100; $i++)
      {
       $global:InternetAccess = Read-Host "`nVoulez vous que cet personne puisse accéder à Internet ? (oui/non)"

       if ($global:InternetAccess -eq "oui")
       {
       break
       }

       elseif ($global:InternetAccess -eq "non")
       {
       break
       }

       else
       {
        Write-Host -ForegroundColor Red "`nVeuillez répondre par oui ou par non" # On continue de poser la question
       }
      }

     }

Function IdentiteUtilisateur ($name, $firstname)

    {

      Write-Host "----------------------------------------"
      Write-Host " Informations sur le nouvel utilisateur"
      Write-Host "----------------------------------------"
      Get-AdUser -Identity "$name`_$firstname" | Select-Object DistinguishedName, GivenName, Name, SamAccountName, UserPrincipalName

    }

Function Reessayer

    {

      for ($i = 1; $i -lt 100; $i++)
      {
       $decision = Read-Host "`n`nVoulez vous réessayez ? (oui/non)"

       if ($decision -eq "oui")
       {
       Break
       }

       elseif ($decision -eq "non")
       {
       Final
       }

       else
       {
       Write-Host -ForegroundColor Red "`nVeuillez répondre par oui ou par non" # On continue de poser la question
       }
      }

     }

Function CatchMessage ($Message)

     {
      Write-Host -ForegroundColor Red "$Message"
      Reessayer
      Start-Sleep 1.3
      Attente
     }

Function Final ($i, $decision)

    {

      for ($i = 1; $i -lt 100; $i++)
      {
       $decision = Read-Host "`nVoulez vous retourner au menu principal ? (oui/non)"
       #$decision = $Host.UI.ReadLine()

       if ($decision -eq "oui")
       {
       Attente2
       break
       }

       elseif ($decision -eq "non")
       {
       Write-host "`n`nMerci d'avoir utilisé le script officiel du domaine Alca.tqi, à bientôt !" -ForegroundColor Cyan
       Start-Sleep 5
       exit
       }

       else
       {
        Write-Host -ForegroundColor Red "`nVeuillez répondre par oui ou par non" # On continue de poser la question
       }
      }

     }

Function Attente

      {

      Write-host -ForegroundColor Yellow "`n`nRelancement du processus de configuration " -NoNewline
      Start-Sleep 1.2
      Write-host -ForegroundColor Yellow "." -NoNewline
      Start-Sleep 1.2
      Write-host -ForegroundColor Yellow "." -NoNewline
      Start-Sleep 1.2
      Write-host -ForegroundColor Yellow "." -NoNewline
      Start-Sleep 1.2
      Write-host -ForegroundColor Yellow "." -NoNewline

      }

Function Attente2

      {

      Write-host -ForegroundColor Yellow "`n`nRetour vers le menu principal " -NoNewline
      Start-Sleep 1.2
      Write-host -ForegroundColor Yellow "-" -NoNewline
      Start-Sleep 1.2
      Write-host -ForegroundColor Yellow "-" -NoNewline
      Start-Sleep 1.2
      Write-host -ForegroundColor Yellow "-" -NoNewline
      Start-Sleep 1.2
      Write-host -ForegroundColor Yellow " ✔" -NoNewline
      Start-Sleep 0.75

      }

#######################################################################################

Clear-Host
Write-host -ForegroundColor Cyan "

  ____                        ____  _          _ _ 
 |  _ \ _____      _____ _ __/ ___|| |__   ___| | |
 | |_) / _ \ \ /\ / / _ \ '__\___ \| '_ \ / _ \ | |
 |  __/ (_) \ V  V /  __/ |   ___) | | | |  __/ | |
 |_|   \___/ \_/\_/ \___|_|  |____/|_| |_|\___|_|_|
                                                   
                                       
"

Write-host  -ForegroundColor Cyan "Bienvenue dans le script PowerShell du domaine Alca.tqi"
Write-host
Write-host
Write-host
Write-host
Write-host
Write-host
Pause
Clear-Host

################################# MENU PRINCIPAL ######################################

for ($i = 1; $i -lt 100; $i++)

{

Clear-Host

Write-Host -ForegroundColor Cyan "---------------------------------------"
Write-Host -ForegroundColor Cyan "      Interface de configuration"
Write-Host -ForegroundColor Cyan "---------------------------------------"
Write-Host
Write-Host "1) Créer un utilisateur"
Write-Host
Write-Host "2) Créer des utilisateurs à l'aide d'un fichier csv"
Write-Host
Write-Host "3) Changer un utilisateur de classe"
Write-Host
Write-Host "4) Supprimer un utilisateur (avec sauvegarde des documents)"
Write-Host 
Write-Host "5) Effectuer une sauvegarde manuelle ou une sauvegarde d'un utilisateur"
Write-Host
Write-Host "6) Ouvrir la page github du script"
Write-Host
Write-Host
Write-Host -ForegroundColor Cyan '---Encodez la lettre "q" pour quitter----'
Write-Host
Write-Host
Write-Host

[string]$choix = Read-Host "Que voulez vous faire ?"
switch ($choix)
  {
      1{ # Création d'utilisateur simple

        $test = dsquery user -samid "0"
        while($null -eq $test)
        {
        Clear-Host
        Write-Host -ForegroundColor Cyan "/// Création d'un utilisateur ///"
        
        $name = Read-Host "`nQuel est le nom de famille de l'utilisateur ?"
        $firstname = Read-Host "`nQuel est le prénom de l'utilisateur ?"
        $password = Read-Host -AsSecureString "`nVeuillez encoder le mot de passe de l'utilisateur"
        $statut = Read-Host "`nEst-ce un élève (1) ou un professeur (2) ?"

        ############################ SI ELEVE #######################################

        if($statut -eq 1)

          {
            $groupes = Get-ADGroup -Filter "*" -SearchBase "OU=Classe,DC=alca,DC=tqi" | Format-Table Name -HideTableHeaders
            Write-Host "`nClasses disponibles :"
            $groupes
            $classe = Read-Host "Quel est sa classe ?"
            Internet
            
            try
            {

              VerificationDossierEtPartage "Partage_Eleves"
          
              CreationDossier "Partage_Eleves" $name $firstname  | Out-Null
              
              if (Get-ADGroup $classe)
              {
              New-ADUser -SamAccountName "$name`_$firstname" -GivenName "$firstname" -Name "$name" -DisplayName "$name $firstname" -ProfilePath "\\SERVEUR\Dossier_Ecole\Partage_Eleves\Profil_Itinérant\$name`_$firstname" -HomeDir "\\SERVEUR\Dossier_Ecole\Partage_Eleves\Dossier_Document\$name`_$firstname" -PATH "OU=Eleves, DC=alca, DC=tqi" -AccountPassword $password -UserPrincipalName "$name`_$firstname@labo.be"  -Enable $true
              Add-ADGroupMember -Identity "$classe" -Members "CN=$name, OU=Eleves, DC=alca, DC=tqi"
              Add-ADGroupMember -Identity "Ensemble_Des_Eleves(NTFS)" -Members "CN=$name, OU=Eleves, DC=alca, DC=tqi"
              if ($global:InternetAccess -eq "oui") {Add-ADGroupMember -Identity "Internet_Eleves" -Members "CN=$name, OU=Eleves, DC=alca, DC=tqi"}
              }

              DroitsNTFS "Partage_Eleves" "Ensemble_Des_Eleves(NTFS)" $name $firstname

              CréationDesCours $name $firstname

              VerifUtilisateurPresent "Partage_Eleves" $name $firstname

            } # Try (élèves)

            catch [Microsoft.ActiveDirectory.Management.ADIdentityAlreadyExistsException]
            {
            CatchMessage "`nLe dossier personnel où le compte de $name $firstname existe déja !"
            }

            catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
            {
            CatchMessage "`nLa classe que vous avez encodé n'existe pas."
            }

            catch [Microsoft.ActiveDirectory.Management.ADPasswordComplexityException]
            {
            Get-ADUser -Filter 'Enabled -eq "False"' -SearchBase "OU=Eleves,DC=alca,DC=tqi" | Remove-ADUser -Confirm:$false
            CatchMessage "`nLe mot de passe que vous avez encodez n'est pas assez fort. Assurez vous qu'il possède au moins 6 lettres, un chiffre et un caractère spécial."
            }

            catch [Microsoft.ActiveDirectory.Management.ADException]
            {
            CatchMessage"`nL'utilisateur que vous essayer de créer existe déjà."
            } 

            catch
            {
            CatchMessage "`nUn problème est survenu, veuillez réessayer, tout en vérifiant que les informations que vous encodez sont correctes et en vous assurant que l'utilisateur n'a pas déja été créé."
            }


            $test = dsquery user -samid $name`_$firstname # Fait sortir de la boucle si l'utilisateur a bel et bien été créé
            
          } # If -eq 1 (si élèves)
          
          ##########################################################################

          ################################ SI PROFS ################################

          elseif ($statut -eq 2)
          {
            try
            {

              VerificationDossierEtPartage "Partage_Profs"
          
              CreationDossier "Partage_Profs" $name $firstname | Out-Null

              New-ADUser -SamAccountName "$name`_$firstname" -GivenName "$firstname" -Name "$name" -DisplayName "$name $firstname" -ProfilePath "\\SERVEUR\Dossier_Ecole\Partage_Profs\Profil_Itinérant\$name`_$firstname" -HomeDir "\\SERVEUR\Dossier_Ecole\Partage_Profs\Dossier_Document\$name`_$firstname" -PATH "OU=Profs, DC=alca, DC=tqi" -AccountPassword $password -UserPrincipalName "$name`_$firstname@alca.tqi"  -Enable $true
              Add-ADGroupMember -Identity "Ensemble_Des_Professeurs(NTFS)" -Members "CN=$name, OU=Profs, DC=alca, DC=tqi"
              if ($global:InternetAccess -eq "oui") {Add-ADGroupMember -Identity "Internet_Profs" -Members "CN=$name, OU=Profs, DC=alca, DC=tqi"}


              DroitsNTFS "Partage_Profs" "Ensemble_Des_Professeurs(NTFS)" $name $firstname

              VerifUtilisateurPresent "Partage_Eleves" $name $firstname

            } # Try (profs)

            catch [Microsoft.ActiveDirectory.Management.ADIdentityAlreadyExistsException]
            {
            CatchMessage "`nLe dossier personnel de $name $firstname existe déja !"
            }

            catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
            {
            CatchMessage "`nLa classe que vous avez encodé n'existe pas."
            }

            catch [Microsoft.ActiveDirectory.Management.ADPasswordComplexityException]
            {
            Get-ADUser -Filter 'Enabled -eq "False"' -SearchBase "OU=Profs,DC=alca,DC=tqi" | Remove-ADUser -Confirm:$false
            CatchMessage "`nLe mot de passe que vous avez encodez n'est pas assez fort. Assurez vous qu'il possède au moins 6 lettres, un chiffre et un caractère spécial."
            }

            catch [Microsoft.ActiveDirectory.Management.ADException]
            {
            CatchMessage "`nL'utilisateur que vous essayer de créer existe déjà."
            }

            catch
            {
            CatchMessage "`nUn problème est survenu, veuillez réessayer, tout en vérifiant que les informations que vous encodez sont correctes et en vous assurant que l'utilisateur n'a pas déja été créé.`n"
            }


            $test = dsquery user -samid $name`_$firstname # Fait sortir de la boucle si l'utilisateur a bel et bien été créé

          } # If -eq 2 (si profs)

          #########################################################################

          else
          {
          Write-Host -ForegroundColor Red "`nLe nombre que vous avez encodez est invalide" -NoNewline
          Start-Sleep 2
          Attente
          }

        } # While

        <#Clear-Host
         IdentiteUtilisateur#>
        Write-Host -ForegroundColor Green "`nL'utilisateur '$name $firstname' a bel et bien été créé !"
        Final

      } # Switch 1
    
    
    2 { # Création d'utilisateurs avec fichier.csv

       Clear-Host
       Write-Host -ForegroundColor Cyan "/// Création d'utilisateurs avec un fichier csv ///"

       Write-Host -ForegroundColor Yellow "`nAttention, ce type de configuration part du principe que votre liste contient au minimum ces 3 colonnes : Nom, Prenom, Password"

       $check ="a"
       while ($check -eq "a")
       {
        try
        {
         $path = Read-Host "`nEntrez l'emplacement de votre fichier csv [Par défaut : 'C:\Users\Administrateur\Documents\ListeEAC.csv']" # E:\Autres\PowerShell\AlcaEAC\ListeEAC.csv
         
         if ($path -eq "")
         {
         $path = "C:\Users\Administrateur\Documents\ListeEAC.csv"
         $check = "b"
         }
   
         elseif (-not (Test-Path $path))
         {
         Write-Host -ForegroundColor Red "`nL'emplacement que vous avez encodé est invalide, veuillez réessayer"
         }
         
         else
         {
         $check = "b"
         }
        }
        catch
        {
        Write-Host -ForegroundColor Red "`nL'emplacement que vous avez encodé est invalide, veuillez réessayer"
        }
       } # While

        $userlist = Import-Csv -path "$path" -Delimiter ";"

        $test = dsquery user -samid "0"

        while($null -eq $test)
        {

               try
               {

                  foreach ($a in $userlist)
                  {
                   $name = $a.Nom
                   $firstname = $a.Prenom
                   $classe = $a.Classe
                   $password = $a.Password
                   $securepassword = (ConvertTo-SecureString -AsPlainText "$password" -Force)
                   $global:InternetAccess = $a.Internet
                   $statut = $a.Statut

                   VerificationDossierEtPartage "Partage_$statut`s"
                   if (-not (Test-path "\\SERVEUR\Dossier_Ecole\$statut`s\Dossier_Document\$name`_$firstname"))
                   {
                    CreationDossier "Partage_$statut`s" $name $firstname | Out-Null
                   }
                   if (-not (dsquery user -samid $name`_$firstname))
                   {
                    New-ADUser -SamAccountName "$name`_$firstname" -GivenName "$firstname" -Name "$name" -DisplayName "$name $firstname" -ProfilePath "\\SERVEUR\Dossier_Ecole\Partage_$statut`s\Profil_Itinérant\$name`_$firstname" -HomeDir "\\SERVEUR\Dossier_Ecole\Partage_$statut`s\Dossier_Document\$name`_$firstname" -PATH "OU=$statut`s, DC=alca, DC=tqi" -AccountPassword $securepassword -UserPrincipalName "$name`_$firstname@alca.tqi"  -Enable $true
                   }
                   if ($statut -eq "Eleve") {Add-ADGroupMember -Identity "$classe" -Members "CN=$name, OU=Eleves, DC=alca, DC=tqi"}
                   if ($statut -eq "Eleve") {Add-ADGroupMember -Identity "Ensemble_Des_$statut`s(NTFS)" -Members "CN=$name, OU=$statut`s, DC=alca, DC=tqi"}
                   if ($statut -eq "Prof") {Add-ADGroupMember -Identity "Ensemble_Des_$statut`esseurs(NTFS)" -Members "CN=$name, OU=$statut`s, DC=alca, DC=tqi"}
                   if ($global:InternetAccess -eq "oui") {Add-ADGroupMember -Identity "Internet_$statut`s" -Members "CN=$name, OU=$statut`s, DC=alca, DC=tqi"}
                   if ($statut -eq "Eleve") {DroitsNTFS "Partage_$statut`s" "Ensemble_Des_$statut`s(NTFS)" $name $firstname}
                   if ($statut -eq "Prof") {DroitsNTFS "Partage_$statut`s" "Ensemble_Des_$statut`esseurs(NTFS)" $name $firstname}
                   if ($statut -eq "Eleve") {CréationDesCours $name $firstname}
                   VerifUtilisateurPresent "Partage_$statut`s" $name $firstname

                   if ($statut -eq "Eleve") {Write-Host -ForegroundColor Green "`n$statut '$name $firstname' créé !`n"}
                   if ($statut -eq "Prof") {Write-Host -ForegroundColor Green "`n$statut`esseur '$name $firstname' créé !`n"}

                   $test = dsquery user -samid $name`_$firstname # Fait sortir de la boucle si l'utilisateur a bel et bien été créé
                   }

                } # Try (Eleves)

                catch [System.UnauthorizedAccessException]
                {
                CatchMessage "`nVous n'avez pas le droit d'accéder à cette liste !"
                }

                catch [Microsoft.ActiveDirectory.Management.ADIdentityAlreadyExistsException]
                {
                Write-Host -ForegroundColor Red "`nLe dossier personnel de $name $firstname existe déja !"
                Continue
                }

                catch
                { 
                CatchMessage "`nUn problème est survenu, vérifiez votre liste, un ou plusieurs utilisateurs sont peut-être déjà existant."
                Clear-Host
                }

        } # While

        Write-Host -ForegroundColor Green "`nTout s'est correctement déroulé !`n"
        Final

     } # Switch 2

    3 {     
           $testuser = dsquery user -samid "0"
           while ($null -eq $testuser)
           {

            Clear-Host
            Write-Host -ForegroundColor Cyan  "/// Changement de classe d'un utilisateur ///"

            Write-Host "Liste des élèves :"
             $ListeEleves = Get-Aduser -Filter "*" -SearchBase "OU=Eleves,DC=alca,DC=tqi" | Format-Table SamAccountName -HideTableHeaders
             if ($null -eq $ListeEleves) { Write-Host -ForegroundColor Red "`nIl n'y a pas d'élèves`n" }
             $ListeEleves
             Write-Host "Liste des professeurs :"
             $ListeProfs = Get-Aduser -Filter "*" -SearchBase "OU=Profs,DC=alca,DC=tqi" | Format-Table SamAccountName -HideTableHeaders
             if ($null -eq $ListeProfs) { Write-Host -ForegroundColor Red "`nIl n'y a pas de professeurs" }
             $ListeProfs

             $erroractionpreference = "SilentlyContinue"
             
              $choixuser = Read-Host "`nEntrez le nom_prénom de l'élève ou du professeur que vous voulez supprimer"
              $testuser = dsquery user -samid $choixuser
              if ($null -eq $testuser)
               {
                Write-Host -ForegroundColor Red "`nL'élève ou le professeur que vous avez encoder n'existe pas, attention le symbole underscore est obligatoire entre le nom et le prénom`n"
                Start-Sleep 1.8
                Attente
                }

            } # While

             $erroractionpreference = "SilentlyContinue"

            $verifclasse = dsquery group "0"
            while ($null -eq $verifclasse)
            {
             Clear-Host

             $groupes = Get-ADGroup -Filter "*" -SearchBase "OU=Classe,DC=alca,DC=tqi" | Format-Table Name -HideTableHeaders
             Write-Host "`nClasses disponibles :"
             $groupes

             $erroractionpreference = "SilentlyContinue"

              $classe = Read-Host "`nDans quel classe voulez vous le placer ?"
              $verifclasse = dsquery group "CN=$classe,OU=Classe,DC=alca,DC=tqi"
              if ($null -eq $verifclasse) 
              {
               Write-Host -ForegroundColor Red "`nLa classe que vous avez encodé est incorrecte"
               Start-Sleep 1.8
               Attente
              }
            } #While

             $erroractionpreference = "Continue"

             $username = get-aduser -identity $choixuser
             $adgroup = get-adgroup -Filter "*" -SearchBase "OU=Classe,DC=alca,DC=tqi" -Properties members | Where-Object {$_.members -eq $username.distinguishedname}
             Remove-ADGroupMember -identity $adgroup.DistinguishedName -Members $username.DistinguishedName -Confirm:$false -ErrorAction SilentlyContinue

             Add-ADGroupMember -Identity $classe -Members $choixuser

             Write-Host -ForegroundColor Green "`nLa classe de l'utilisateur a bel et bien été modifié !"
             Final

      } # Switch 3    

    4 {
          $testuser = dsquery user -samid "0"
          while ($null -eq $testuser)
          {
           
           Clear-Host
           Write-Host -ForegroundColor Cyan  "/// Suppression d'utilisateur avec sauvgarde des données ///"

             Write-Host "Liste des élèves :"
             $ListeEleves = Get-Aduser -Filter "*" -SearchBase "OU=Eleves,DC=alca,DC=tqi" | Format-Table SamAccountName -HideTableHeaders
             if ($null -eq $ListeEleves) { Write-Host -ForegroundColor Red "`nIl n'y a pas d'élèves`n" }
             $ListeEleves
             Write-Host "Liste des professeurs :"
             $ListeProfs = Get-Aduser -Filter "*" -SearchBase "OU=Profs,DC=alca,DC=tqi" | Format-Table SamAccountName -HideTableHeaders
             if ($null -eq $ListeProfs) { Write-Host -ForegroundColor Red "`nIl n'y a pas de professeurs" }
             $ListeProfs

             $erroractionpreference = "SilentlyContinue"
             
             $testuser = dsquery user -samid "0"
             while ($null -eq $testuser)
             {
              $choixuser = Read-Host "`nEntrez le nom_prénom de l'élève ou du professeur que vous voulez supprimer"
              $testuser = dsquery user -samid $choixuser
              if ($null -eq $testuser)
              {
               Write-Host -ForegroundColor Red "`nL'élève ou le professeur que vous avez encoder n'existe pas, attention le symbole underscore est obligatoire entre le nom et le prénom`n"
               Start-Sleep 1.8
               Attente
              }

             }

            } # While

             $erroractionpreference = "Continue"

              $test2 = "False"
              while ($test2 -ne "True")
               {

                Clear-Host
                try
                {
                 $destination = Read-Host "`nOu voulez vous placer la sauvegarde des documents ? (Exemple :'D:\Sauvegarde_Repo')"
                 $test2 = Test-Path "$destination"
                 if ($test2 -ne "True") {Write-Host -ForegroundColor Red "`nCe dossier n'existe pas, veuillez encoder un dossier valide"} 
                }
                catch
                {
                 Write-Host -ForegroundColor Red "`nCe dossier n'existe pas, veuillez encoder un dossier valide"
                }
               } # While

               $date = Get-date -UFormat "%d_%m_%Y"
               New-item -ItemType directory -Path "$destination\Sauvegarde_$choixuser`_$date" | Out-Null
               Copy-item "\\SERVEUR\Dossier_Ecole\Partage_Eleves\Dossier_Document\$choixuser" "$destination\Sauvegarde_$choixuser`_$date" -Recurse -ErrorAction SilentlyContinue
               Copy-item "\\SERVEUR\Dossier_Ecole\Partage_Profs\Dossier_Document\$choixuser" "$destination\Sauvegarde_$choixuser`_$date" -Recurse -ErrorAction SilentlyContinue
               Remove-ADUser -Identity "$choixuser"
               Remove-item "\\SERVEUR\Dossier_Ecole\Partage_Eleves\Dossier_Document\$choixuser" -Recurse -ErrorAction SilentlyContinue -Force
               Remove-item "\\SERVEUR\Dossier_Ecole\Partage_Eleves\Profil_Itinérant\$choixuser" -Recurse -ErrorAction SilentlyContinue -Force               
               Remove-item "\\SERVEUR\Dossier_Ecole\Partage_Profs\Dossier_Document\$choixuser" -Recurse -ErrorAction SilentlyContinue -Force
               Remove-item "\\SERVEUR\Dossier_Ecole\Partage_Profs\Profil_Itinérant\$choixuser" -Recurse -ErrorAction SilentlyContinue -Force

    
              Write-Host -ForegroundColor Green "`nL'utilisateur a été supprimé et sa sauvegarde a bien été créée!"
              Final

      } # Switch 4

    5 {
           Clear-Host
           Write-Host -ForegroundColor Cyan "/// Sauvegarde de données ///"
           
           $test = Test-path "False"
           while ($test -ne "True")
           {

            $question = Read-Host "`nVoulez vous faire une sauvegarde manuelle (1) ou une sauvegarde d'un utilisateur (2) ?"

            if ($question -eq "1")
            {
              $test1 = "False"
              while ($test1 -ne "True")
               {
                 try
                 {
                 $backupdir = Read-Host "`nEntrez le chemin d'accès du dossier ou du fichier que vous voulez sauvegarder? (Exemple :'C:\DossierImportant\')"
                 $test1 = Test-path "$backupdir"
                 if ($test1 -ne "True") {Write-Host -ForegroundColor Red "`nCe dossier ou ce fichier n'existe pas, veuillez encoder un chemin d'accès valide"} 
                 }
                 catch
                 {
                    Write-Host -ForegroundColor Red "`nCe dossier ou ce fichier n'existe pas, veuillez encoder un chemin d'accès valide"
                 }
                }

                $test2 = "False"
                while ($test2 -ne "True")
                {

                 Clear-Host
                 try
                 {
                    $destination = Read-Host "`nOu voulez vous placer la sauvegarde ? (Exemple :'D:\Sauvegarde_Repo')"
                    $test2 = Test-Path "$destination"
                    if ($test2 -ne "True") {Write-Host -ForegroundColor Red "`nCe dossier n'existe pas, veuillez encoder un dossier valide"} 
                 }
                 catch
                 {
                    Write-Host -ForegroundColor Red "`nCe dossier n'existe pas, veuillez encoder un dossier valide" 

                 }
                }

                $date = Get-date -UFormat "%d_%m_%Y"
                New-item -ItemType directory -Path "$destination\Sauvegarde du $date" | Out-Null
                Copy-item $backupdir "$destination\Sauvegarde du $date" -Recurse

            $test = Test-path "$destination\Sauvegarde du $date"

             } # If Sauvegarde manuelle

             elseif ($question -eq "2")
             {

            $testuser = dsquery user -samid "0"
            while ($null -eq $testuser)
            {

             Write-Host "Liste des élèves :"
             $ListeEleves = Get-Aduser -Filter "*" -SearchBase "OU=Eleves,DC=alca,DC=tqi" | Format-Table SamAccountName -HideTableHeaders
             $ListeEleves
             Write-Host "Liste des professeurs :"
             $ListeProfs = Get-Aduser -Filter "*" -SearchBase "OU=Profs,DC=alca,DC=tqi" | Format-Table SamAccountName -HideTableHeaders
             $ListeProfs

             $erroractionpreference = "SilentlyContinue"
             
              $choixuser = Read-Host "`nEntrez le nom_prénom de l'élève ou du professeur concerné par la sauvegarde des documents"
              $testuser = dsquery user -samid $choixuser
              if ($null -eq $testuser) 
              {
               Write-Host -ForegroundColor Red "`nL'élève ou le professeur que vous avez encoder n'existe pas, attention le symbole underscore est obligatoire entre le nom et le prénom`n"
               Start-Sleep 1.8
               Attente
              }
             
            } # While

             $erroractionpreference = "Continue"

              $test2 = "False"
              while ($test2 -ne "True")
               {

                Clear-Host
                try
                {
                 $destination = Read-Host "`nOu voulez vous placer la sauvegarde ? (Exemple :'D:\Sauvegarde_Repo')"
                 $test2 = Test-Path "$destination"
                 if ($test2 -ne "True") {Write-Host -ForegroundColor Red "`nCe dossier n'existe pas, veuillez encoder un dossier valide"} 
                }
                catch
                {
                 Write-Host -ForegroundColor Red "`nCe dossier n'existe pas, veuillez encoder un dossier valide"
                }
               }

               $date = Get-date -UFormat "%d_%m_%Y"
               New-item -ItemType directory -Path "$destination\Sauvegarde_$choixuser`_$date" | Out-Null
               Copy-item "\\SERVEUR\Dossier_Ecole\Partage_Eleves\Dossier_Document\$choixuser" "$destination\Sauvegarde_$choixuser`_$date" -Recurse -ErrorAction SilentlyContinue
               Copy-item "\\SERVEUR\Dossier_Ecole\Partage_Profs\Dossier_Document\$choixuser" "$destination\Sauvegarde_$choixuser`_$date" -Recurse -ErrorAction SilentlyContinue
 
               $test = Test-path "$destination\Sauvegarde_$choixuser'_$date"

             } # If Sauvegarde élèves ou profs

             else
             {
              Write-Host -ForegroundColor Red "`nLe nombre que vous avez encodez est invalide"
             }


      } # While


    Write-Host -ForegroundColor Green "`nLa sauvegarde a bien été créée!"
    Final

      } # Switch 5
      
    6 {
       Start-Process "https://www.github.com/EAC-Alex"
      }

    q {
        Write-host "`n`n`nMerci d'avoir utilisé le script officiel du domaine Alca.tqi, à bientôt !" -ForegroundColor Cyan
        Start-Sleep 5
        exit
      } 

      Default {
              Write-Host -ForegroundColor Red "`nLe nombre que vous avez encodez est invalide"
              Start-Sleep 1.8
              Attente
              }

  } # Switch Global

} # Boucle For Global du menu principal