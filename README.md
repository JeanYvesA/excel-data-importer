# Importation de Données Excel vers MySQL en C#
C'est une application console écrit en C# (.NET 8) qui permet d'importer des données depuis un fichier Excel vers une base de données MySQL.
# Fonctionnalités
Lecture et récupération des données de fichier Excel (.xlsx)

Importation des données dans une base de données relationnelle MySQL

Gestion des erreurs

Tests unitaires avec NUnit
# Prérequis
.NET 8

MySQL Server

# Installation
1. Cloner le dépôt

  ``` 
  git clone https://github.com/JeanYvesA/excel-data-importer.git
```
  
2. Restaurer les dépendances
   
``` 
dotnet restore
```

3. Configurer la connexion à la base de données

   Modifier le fichier appsettings.json situé à la racine du projet avec les informations de votre base de données MySQL.

   Exemple de appsettings.json :

   ```
   {
        "ConnectionStrings": {
          "DefaultConnection": "server=localhost;database=ImportDB;user=root;password=1234"
        }
   }
   ```

4. Installer les outils de l'ORM Entity Framework

   ```
   dotnet tool install --global dotnet-ef
   ```

5. Exécuter la migration pour créer les tables

   ```
   dotnet ef database update
   ```

6. Exécuter de l'application

   ```
   dotnet run
   ```

    L'utilisateur devra choisir entre :
    
    1. Utiliser le fichier de test : c'est un fichier excel qui contient les informations pour tester l'application
    
    2. Entrer le chemin d'un fichier Excel

7. Exécuter les tests

   ```
   dotnet test
   ```

# Guide d'utilisation

1. Exécuter l'application

   ```
   dotnet run
   ```

2. Saisir "1" pour utiliser un fichier de test ou "2" pour entrer un chemin de fichier excel

3. L'application traite le fichier et insère les données dans la base MySQL. Le temps d'exécution est affiché à la fin de l'importation.
