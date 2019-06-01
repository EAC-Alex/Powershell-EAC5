########################## INSTALLATION DES MODULES #######################################

cls
Write-Progress -Activity "Installation des modules" -Status "1/3"
Install-WindowsFeature AD-Domain-Services | Out-Null
Write-Progress -Activity "Installation des modules" -Status "2/3"
Install-WindowsFeature -Name "DHCP" -IncludeManagementTools | Out-Null
Write-Progress -Activity "Installation des modules" -Status "3/3"
Install-Module NtfSsecurity | Out-Null
Write-Progress -Activity "Installation des modules" -Status "Terminé" -Completed

###########################################################################################

########################## VERIFICATION DES OU #######################################

if (-not (dsquery ou "OU=Eleves,DC=alca,DC=tqi"))
      {
      New-ADOrganizationalUnit "Eleves" -path "dc=alca, dc=tqi"
      }

if (-not (dsquery ou "OU=Profs,DC=alca,DC=tqi"))
      {
      New-ADOrganizationalUnit "Profs" -path "dc=alca, dc=tqi"
      }

if (-not (dsquery ou "OU=Classe,DC=alca,DC=tqi"))
     {
      New-ADOrganizationalUnit "Classe" -path "dc=alca, dc=tqi"

      $CsvData = Import-Csv -path "C:\Users\Administrateur\Documents\ListeEAC.csv" -Delimiter ";"

      foreach ($i in $CsvData)
      {
      $nomclasse = $i.Classe
      $parentou = "OU=Classe, DC=alca, DC=tqi"
      New-ADGroup -Name "$nomclasse" -SamAccountName "$nomclasse" -GroupCategory Security -GroupScope Global -path "$parentou"
      }
     }

#######################################################################################

########################## DECLARATION DES FONCTIONS ##################################

Function CreationDossier ($name, $firstname)

    {

      if (-not (Test-path "\\127.0.0.1\Partage_Eleves\Dossier_Document\$name`_$firstname"))
      {
      New-item -ItemType directory -Path "\\127.0.0.1\Partage_Eleves\Dossier_Document\$name`_$firstname"
      }

      if (-not (Test-path "\\127.0.0.1\Partage_Eleves\Profil_Itinérant\$name`_$firstname"))
      {
      New-item -ItemType directory -Path "\\127.0.0.1\Partage_Eleves\Profil_Itinérant\$name`_$firstname"
      }

    }

Function DroitsNTFS ($name, $firstname)

    {

    # Desactiver l'héritage tout en copiant les autorisations NTFS héritées

    Get-item "\\127.0.0.1\Partage_Eleves\Dossier_Document\$name`_$firstname" | Disable-NTFSAccessInheritance
    Get-item "\\127.0.0.1\Partage_Eleves\Profil_Itinérant\$name`_$firstname" | Disable-NTFSAccessInheritance

    # Modfication du propriétaire des dossier

    Set-NTFSOwner -path "\\127.0.0.1\Partage_Eleves\Dossier_Document\$name`_$firstname" -Account "$firstname@labo.be"
    Set-NTFSOwner -path "\\127.0.0.1\Partage_Eleves\Profil_Itinérant\$name`_$firstname" -Account "$firstname@labo.be"

    # Ajout des autorisations NTFS

    Add-NTFSAccess -Path "\\127.0.0.1\Partage_Eleves\Dossier_Document\$name`_$firstname" -Account "$firstname@labo.be" -AccessRights FullControl
    Add-NTFSAccess -Path "\\127.0.0.1\Partage_Eleves\Profil_Itinérant\$name`_$firstname" -Account "$firstname@labo.be" -AccessRights FullControl
       
    }

Function VerificationDossierEtPartage

    {

      if (-not (Test-path "C:\Partage_Eleves\Dossier_Document"))
      {
      New-item -ItemType directory -Path "C:\Partage_Eleves\Dossier_Document" | Out-Null
      }

      if (-not (Test-path "C:\Partage_Eleves\Profil_Itinérant"))
      {
      New-item -ItemType directory -Path "C:\Partage_Eleves\Profil_Itinérant" | Out-Null
      }

      if(!(Get-SMBShare -Name "Partage_Eleves" -ErrorAction SilentlyContinue)){
      New-SmbShare -Name "Partage_Eleves" -Path "C:\Partage_Eleves" -FullAccess "Administrateur" | Out-Null
      }

    }

Function VerifUtilisateurPresent ($name, $firstname)

    {

      $verif = dsquery user -samid $firstname
      if ($verif -eq $null)
      {
      Remove-item -ItemType directory -Path "\\127.0.0.1\Partage_Eleves\Dossier_Document\$name`_$firstname"
      Remove-item -ItemType directory -Path "\\127.0.0.1\Partage_Eleves\Profil_Itinérant\$name`_$firstname"
      }

    }

Function Reessayer ($i, $decision)

    {

      for ($i = 1; $i -lt 100; $i++)
      {
       $decision = Read-Host "`n`nVoulez vous réessayez ? (oui/non)"
       #$decision = $Host.UI.ReadLine()

       if ($decision -eq "oui")
       {
       break
       }

       elseif ($decision -eq "non")
       {
       Write-host "`n`n`nDommage, merci d'avoir utilisé le script officiel du domaine Alca.tqi, à bientôt !" -ForegroundColor Cyan
       Sleep 5
       exit
       }

       else
       {
       # On continue de poser la question
       }
      }

     }


Function Final ($i, $decision)

    {

      for ($i = 1; $i -lt 100; $i++)
      {
       $decision = Read-Host "`nVoulez vous retourner au menu principal ? (oui/non)"
       #$decision = $Host.UI.ReadLine()

       if ($decision -eq "oui")
       {
       break
       }

       elseif ($decision -eq "non")
       {
       Write-host "`n`n`nMerci d'avoir utilisé le script officiel du domaine Alca.tqi, à bientôt !" -ForegroundColor Cyan
       Sleep 5
       exit
       }

       else
       {
        Write-Host -ForegroundColor Red "Répondez par oui ou par non"# On continue de poser la question
       }
      }

     }

Function Attente

      {

      Write-host "`n`nRelancement du processus de configuration " -NoNewline
      Sleep 1.2
      Write-host "." -NoNewline
      Sleep 1.2
      Write-host "." -NoNewline
      Sleep 1.2
      Write-host "." -NoNewline
      Sleep 1.2
      Write-host "." -NoNewline

      }

#######################################################################################

################################# MENU PRINCIPAL ######################################

for ($i = 1; $i -lt 100; $i++)

{

cls

Write-Host "---------------------------------------"
Write-Host "      Interface de configuration"
Write-Host "---------------------------------------"
Write-Host
Write-Host "/// Gestion des utilisateurs ///"
Write-Host
Write-Host "1) Créer un utilisateur"
Write-Host "2) Créer des utilisateurs à l'aide d'un fichier csv"
Write-Host 
Write-Host "/// Gestion des groupes ///"
Write-Host 
Write-Host "3) Créer un groupe"
Write-Host "4) Créer des groupes à l'aide d'un fichier csv"
Write-Host
Write-Host "/// Divers ///"
Write-Host
Write-Host
Write-Host
Write-Host
Write-Host '---Appuyez sur la lettre "q" pour quitter----'
Write-Host

[string]$choix = Read-Host "Que voulez vous faire ?"
switch ($choix)
  {
      1{ # Création d'utilisateur simple

        $test = dsquery user -samid "0"
        while($test -eq $null)
        {
        cls
        Write-Host "/// Création d'un utilisateur ///"
        
        $name = Read-Host "`nQuel est le nom de famille de l'utilisateur ?"
        $firstname = Read-Host "`nQuel est le prénom de l'utilisateur ?"
        $password = Read-Host -AsSecureString "`nVeuillez encoder le mot de passe de l'utilisateur"
        $statut = Read-Host "`nEst-ce un élève (1) ou un professeur (2) ?"

        ############################ SI ELEVE #######################################

        if($statut -eq 1)

          {
            $groupes = Get-ADGroup -Filter "*" -SearchBase "OU=Classe,DC=alca,DC=tqi" | Ft Name
            Write-Host "`nClasses disponibles :"
            $groupes
            $classe = Read-Host "Quel est sa classe ?"
            
            try
            {

              VerificationDossierEtPartage
          
              CreationDossier | Out-Null
              
              if(Get-ADGroup $classe)
              {
              New-ADUser -SamAccountName "$firstname" -GivenName "$firstname" -Name "$name" -DisplayName "$name $firstname" -ProfilePath "\\127.0.0.1\Partage_Eleves\Profil_Itinérant\$name" -HomeDir "\\127.0.0.1\Partage_Eleves\Dossier_Document\$name" -PATH "OU=Eleves, DC=alca, DC=tqi" -AccountPassword $password -UserPrincipalName "$firstname@labo.be"  -Enable $true
              Add-ADGroupMember -Identity "$classe" -Members "CN=$name, OU=Eleves, DC=alca, DC=tqi"
              }

              DroitsNTFS

              VerifUtilisateurPresent

            } # Try (élèves)

            catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
            {
            Write-Host -ForegroundColor Red "`nLa classe que vous avez encodez n'existe pas." -NoNewline
            Reessayer
            Sleep 2
            Attente
            }

            catch [Microsoft.ActiveDirectory.Management.ADPasswordComplexityException]
            {
            Get-ADUser -Filter 'Enabled -eq "False"' -SearchBase "OU=Eleves,DC=alca,DC=tqi" | Remove-ADUser -Confirm:$false
            Write-Host -ForegroundColor Red "`nLe mot de passe que vous avez encodez n'est pas assez fort. Assurez vous qu'il possède au moins 6 lettres, un chiffre et un caractère spécial." -NoNewline
            Reessayer
            Sleep 2
            Attente
            }

            catch [Microsoft.ActiveDirectory.Management.ADException]
            {
            Write-Host -ForegroundColor Red "`nL'utilisateur que vous essayer de créer existe déjà." -NoNewline
            Reessayer
            Sleep 2
            Attente

            }
            catch
            {
            Write-Host -ForegroundColor Red "`nUn problème est survenu, veuillez réessayer, tout en vérifiant que les informations que vous encodez sont correctes et en vous assurant que l'utilisateur n'a pas déja été créé." -NoNewline
            Reessayer
            Sleep 3
            Attente
            }


            $test = dsquery user -samid $firstname # Fait sortir de la boucle si l'utilisateur a bel et bien été créé
            
          } # If -eq 1 (si élèves)
          
          ##########################################################################

          ################################ SI PROFS ################################

          elseif ($statut -eq 2)
          {
            try
            {

              VerificationDossierEtPartage
          
              CreationDossier | Out-Null

              New-ADUser -SamAccountName "$firstname" -GivenName "$firstname" -Name "$name" -DisplayName "$name $firstname" -ProfilePath "\\127.0.0.1\Partage_Eleves\Profil_Itinérant\$name" -HomeDir "\\127.0.0.1\Partage_Eleves\Dossier_Document\$name" -PATH "OU=Profs, DC=alca, DC=tqi" -AccountPassword $password -UserPrincipalName "$firstname@labo.be"  -Enable $true

              DroitsNTFS

              VerifUtilisateurPresent

            } # Try (profs)

            catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
            {
            Write-Host -ForegroundColor Red "`nLa classe que vous avez encodez n'existe pas." -NoNewline
            Reessayer
            Sleep 2
            Attente
            }

            catch [Microsoft.ActiveDirectory.Management.ADPasswordComplexityException]
            {
            Get-ADUser -Filter 'Enabled -eq "False"' -SearchBase "OU=Eleves,DC=alca,DC=tqi" | Remove-ADUser -Confirm:$false
            Write-Host -ForegroundColor Red "`nLe mot de passe que vous avez encodez n'est pas assez fort. Assurez vous qu'il possède au moins 6 lettres, un chiffre et un caractère spécial." -NoNewline
            Reessayer
            Sleep 3
            Attente
            }

            catch [Microsoft.ActiveDirectory.Management.ADException]
            {
            Write-Host -ForegroundColor Red "`nL'utilisateur que vous essayer de créer existe déjà." -NoNewline
            Reessayer
            Sleep 2
            Attente
            }

            catch
            {
            Write-Host -ForegroundColor Red "`nUn problème est survenu, veuillez réessayer, tout en vérifiant que les informations que vous encodez sont correctes et en vous assurant que l'utilisateur n'a pas déja été créé." -NoNewline
            Reessayer
            Sleep 3
            Attente
            }

        
            $test = dsquery user -samid $firstname # Fait sortir de la boucle si l'utilisateur a bel et bien été créé

          } # If -eq 2 (si profs)

          #########################################################################

          else
          {
          Write-Host -ForegroundColor Red "`nLe nombre que vous avez encodez est invalide" -NoNewline
          Sleep 2
          Attente
          }

        } # While

        Write-Host -ForegroundColor Green "`nL'utilisateur $name $firstname a bel et bien été créé !`n"
        Final

      } # Switch 1
    
    
    2 { # Création d'utilisateurs avec liste.csv

       Write-Host -ForegroundColor Red "Attention, ce type de configuration part du principe que votre liste contient au minimum ces 3 colonnes : Nom, Prenom, Password"
      
       while ($testpath -eq $null)
       {
        try
        {
         $path = Read-Host "`nEntrez l'emplacement de votre liste"
         $testpath = Test-path "$path"
        }
        catch
        {
         Write-Host -ForegroundColor Red "L'emplacement que vous avez encodez est incorrect, veuillez réessayer."
        }
       } 

       $userlist = Import-Csv -path "$path" -Delimiter ";"
      
       foreach ($i in $userlist)
       {
        $name = root
        $i.Nom
        $firstname = $i.Prenom
        $classe = $i.Classe
        $password = $i.Password
        $securepassword = (ConvertTo-SecureString -AsPlainText "$password" -Force)

        VerificationDossierEtPartage
            
        CreationDossier | Out-Null
        
        for ($i = 1; $i -lt 100; $i++)

        {
          $statut = Read-Host "`nCette liste contient-elle des élèves (1) ou des professeurs (2) ?"
          
          if ($statut -eq 1)
          {
            New-ADUser -SamAccountName "$firstname" -GivenName "$firstname" -Name "$name" -DisplayName "$name $firstname" -ProfilePath "\\127.0.0.1\Partage_Eleves\Profil_Itinérant\$name" -HomeDir "\\127.0.0.1\Partage_Eleves\Dossier_Document\$name" -PATH "OU=Eleves, DC=alca, DC=tqi" -AccountPassword $password -UserPrincipalName "$firstname@labo.be"  -Enable $true
            Add-ADGroupMember -Identity "$classe" -Members "CN=$name, OU=Eleves, DC=alca, DC=tqi"

            DroitsNTFS

            VerifUtilisateurPresent

            Write-Host -ForegroundColor Green "`nLes élèves ont bien été créés !`n"
            Final
    
          } # If (Eleves)

          elseif ($statut -eq 2)
          {
            New-ADUser -SamAccountName "$firstname" -GivenName "$firstname" -Name "$name" -DisplayName "$name $firstname" -ProfilePath "\\127.0.0.1\Partage_Eleves\Profil_Itinérant\$name" -HomeDir "\\127.0.0.1\Partage_Eleves\Dossier_Document\$name" -PATH "OU=Profs, DC=alca, DC=tqi" -AccountPassword $password -UserPrincipalName "$firstname@labo.be"  -Enable $true
            
            DroitsNTFS

            VerifUtilisateurPresent

            Write-Host -ForegroundColor Green "`nLes professeurs ont bien été créés !`n"
            Final

          } # If (Profs)

          else
          {
            Write-Host -ForegroundColor Red "`nLe nombre que vous avez encodez est invalide" -NoNewline
          }

        } # Boucle For

        break # Sortie vers menu principal

       } # Foreach csv

     } # Switch 2
      

    q {
        Write-host "`n`n`nMerci d'avoir utilisé le script officiel du domaine Alca.tqi, à bientôt !" -ForegroundColor Cyan
        Sleep 5
        exit
      }
        

      Default {
              Write-Host -ForegroundColor Red "`nLe nombre que vous avez encodez est invalide"
              Sleep 1.8
              Attente
              }

  } # Switch Global

} # Boucle For Global du menu principal



<#

      $Error[4].Exception.GetType().fullname


      for ($i = 0; $i -lt 200; $i++)
      {
      $Error[$i].Exception.GetType().fullname
      }
      














      Do {
    $strTitle = Read-host "Enter the book title"
} until ($strTitle -notMatch "[^[:alpha:]]")


$name = Read-Host "Quel est le nom de famille de l'utilisateur ?"
while ($name -contains 0..9)
      {
      [string]$name = Read-Host "Quel est le nom de famille de l'utilisateur ?"
      }



#$ssn4 = ("a" -or "b" -or "c" -or "d" -or "e" -or "f" -or "g" -or "h" -or "i" -or "j" -or "k" -or "l" -or "m" -or "n" -or "o" -or "p" -or "q" -or "r" -or "s" -or "t" -or "u" -or "v" -or "w" -or "x" -or "y" -or "z")
#do
#{
#	$name = read-host "Please enter a 4 digit number"
#	
#} while ($name -notmatch $ssn4)
#{
#Write-host "non"
#}

$firstname = Read-Host
$test = dsquery user -samid $firstname

if($test -eq $null){"shit"} else {"nice"}


$pass = Read-Host -AsSecureString "entre stp"



cls
write-host "

  ____                        ____  _          _ _ 
 |  _ \ _____      _____ _ __/ ___|| |__   ___| | |
 | |_) / _ \ \ /\ / / _ \ '__\___ \| '_ \ / _ \ | |
 |  __/ (_) \ V  V /  __/ |   ___) | | | |  __/ | |
 |_|   \___/ \_/\_/ \___|_|  |____/|_| |_|\___|_|_|
                                                   
                                       
" -ForegroundColor Cyan

Write-host "Bienvenue dans le script PowerShell du domaine Alca.tqi" -ForegroundColor Cyan
Write-host
Write-host
Write-host
Write-host
Write-host
Write-host
Pause
cls


$a = Read-Host "a"
$b = Read-Host "b"

try
{
 $a / $b 
}

catch [System.DivideByZeroException]
{
Write-host "Zero"
}

catch 
{
Write-host "Problème"
}

Search-ADAccount -AccountDisabled

Get-ADUser -Filter 'Enabled -eq "False"' -SearchBase "OU=Eleves,DC=alca,DC=tqi" | Remove-ADUser

try{
$password = Read-Host -AsSecureString "`nVeuillez encoder le mot de passe de l'utilisateur"
New-ADUser -SamAccountName "$firstname" -GivenName "$firstname" -Name "$name" -DisplayName "$name $firstname" -ProfilePath "\\127.0.0.1\Partage_Eleves\Profil_Itinérant\$name" -HomeDir "\\127.0.0.1\Partage_Eleves\Dossier_Document\$name" -PATH "OU=Eleves, DC=alca, DC=tqi" -AccountPassword $password -UserPrincipalName "$firstname@labo.be"  -Enable $true
}
catch [Microsoft.ActiveDirectory.Management.ADPasswordComplexityException]
      {
      Write-Host -ForegroundColor Red "`nLe mot de passe que vous avez encodez n'est pas assez fort. Assurez vous qu'il possède au moins 6 lettres, un chiffre et un caractère spécial"
      }



Write-Progress -Activity "Installation des modules" -Status "1/3"
Install-WindowsFeature AD-Domain-Services
Write-Progress -Activity "Installation des modules" -Status "2/3"
Install-WindowsFeature -Name "DHCP" -IncludeManagementTools
Write-Progress -Activity "Installation des modules" -Status "3/3"
Install-Module NtfSsecurity

#>