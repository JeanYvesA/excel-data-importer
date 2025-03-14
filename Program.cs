
using System.Diagnostics;
using ExcelDataImporterProject.Models;
using ExcelDataImporterProject.Services;

Console.WriteLine("Bonjour, bienvenue !\n");
Console.WriteLine("Entrez le chemin du fichier Excel puis valider");

//Récupération du chemin saisi par l'utilisateur
string filePath = Console.ReadLine();

// Vérifier si le fichier existe
if (!File.Exists(filePath))
{
    Console.WriteLine("Fichier introuvable !\n");
    return;
}

Stopwatch stopwatch = Stopwatch.StartNew(); // Démarre le chronomètre afin d'évaluer le temps d'exécution
Console.WriteLine("\nDébut de l'importation !\n");

try
{
    //Récupération de la liste des utilisateurs (Person)
    List<Person>persons= ExcelReader.ReadExcelFile(filePath);

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
