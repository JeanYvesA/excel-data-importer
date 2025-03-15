
using System;
using System.Diagnostics;
using ExcelDataImporterProject.Models;
using ExcelDataImporterProject.Services;
using Microsoft.EntityFrameworkCore;

//Variables
int choice;
string option, filePath, projectDirectory, filesDirectory;
List<Person> persons;
Stopwatch stopwatch;

Console.WriteLine("Bonjour, bienvenue !\n");

//Test de connexion à la base de données
using (var context = new MyDbContext())
{
    try
    {
        context.Database.OpenConnection(); // Ouvre une connexion à la base de données
        Console.WriteLine("Connexion à la base de données réussie !\n");
        context.Database.CloseConnection(); // Ferme la connexion
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Erreur de connexion à la base de données : {ex.Message}");
        return;
    }
}
do
{
    Console.WriteLine("Veuillez choisir une option :  \n");
    Console.WriteLine("1 - Utiliser le fichier de test");
    Console.WriteLine("2 - Entrer le chemin du fichier manuellement \n");
    Console.WriteLine("Veuillez entrer 1 ou 2");

    //Récupération du chemin saisi par l'utilisateur
    option = Console.ReadLine();
}
while (!int.TryParse(option, out choice) || (choice != 1 && choice != 2));

if (choice==1)
{
    //Récupération du fichier de test
    projectDirectory = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.Parent.FullName;
    filesDirectory = Path.Combine(projectDirectory, "Files");
    filePath = Path.Combine(filesDirectory, "people sample.xlsx");
}
else
{
    Console.WriteLine("Entrez le chemin du fichier Excel puis valider");

    //Récupération du chemin saisi par l'utilisateur
    filePath = Console.ReadLine();
}

// Vérification de l'existence du fichier
if (!File.Exists(filePath))
{
    Console.WriteLine("Le fichier est introuvable !\n");
    return;
}


Console.WriteLine("\nDébut de l'importation !\n");
stopwatch = Stopwatch.StartNew(); // Démarrage du chronomètre afin d'évaluer le temps d'exécution

try
{
    //Récupération de la liste des utilisateurs (Person)
    persons= ExcelReader.ReadExcelFile(filePath);

    //Insertion dans la base de données relationnelle MySQL
    DatabaseService.InsertPersons(persons);
}
finally
{
    stopwatch.Stop(); // Arrête le chronomètre

    Console.WriteLine("Fin de l'importation !\n");

    //Affichage du temps d'exécution
    if (stopwatch.Elapsed.Minutes > 0)
    {
        Console.WriteLine($"Temps d'exécution : {stopwatch.Elapsed.Minutes} minutes et {stopwatch.Elapsed.Seconds} secondes\n");
    }
    else
    {
        Console.WriteLine($"Temps d'exécution : {stopwatch.Elapsed.Seconds} secondes\n");
    }
}
