using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataImporterProject.Models;
using ExcelDataImporterProject.Services;
using NUnit.Framework;
using OfficeOpenXml;

namespace ExcelDataImporterProject.TestClass
{
    [TestFixture]
    class ExcelReaderTests
    {
        

        [SetUp]
        public void SetUp()
        {

        }
        /// <summary>
        /// Teste la lecture du fichier excel et la récupération des données
        /// </summary>
        [Test]
        public void TestReadExcelFile()
        {
            
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "test_data.xlsx");

            // Création d'un fichier Excel pour simuler des données
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                //Ajout l'en-tete
                worksheet.Cells[1, 1].Value = "matricule";
                worksheet.Cells[1, 2].Value = "nom";
                worksheet.Cells[1, 3].Value = "prenom";
                worksheet.Cells[1, 4].Value = "datedenaissance";
                worksheet.Cells[1, 5].Value = "status";
                //Ajout de deux lignes au fichier
                worksheet.Cells[2, 1].Value = "GDFEREUE";  
                worksheet.Cells[2, 2].Value = "Doe";  
                worksheet.Cells[2, 3].Value = "John";  
                worksheet.Cells[2, 4].Value = "1990-01-01";  
                worksheet.Cells[2, 5].Value = "Actif"; 
                worksheet.Cells[3, 1].Value = "POIUYTRB";  
                worksheet.Cells[3, 2].Value = "Smith";  
                worksheet.Cells[3, 3].Value = "Jane";  
                worksheet.Cells[3, 4].Value = "1985-06-15";  
                worksheet.Cells[3, 5].Value = "Inactif";  
                package.Save();
            }

            //Récupération des données
            List<Person> persons = ExcelReader.ReadExcelFile(filePath);

            // Les tests
            Assert.That(persons.Count, Is.EqualTo(2), "Le nombre de personnes lues dans le fichier Excel est incorrect.");
            Assert.That(persons[0].Number, Is.EqualTo("GDFEREUE"), "La référence de la première personne est incorrect.");
            Assert.That(persons[0].Surname, Is.EqualTo("Doe"), "Le nom de la première personne est incorrect.");
            Assert.That(persons[0].Firstname, Is.EqualTo("John"), "Le prénom de la première personne est incorrect.");
            Assert.That(persons[1].Date, Is.EqualTo(new DateOnly(1985, 06, 15)), "La date de naissance de la deuxième personne est incorrecte.");
            Assert.That(persons[1].Firstname, Is.EqualTo("Jane"),  "Le ^prénom de la deuxième personne est incorrect.");
            Assert.That(persons[1].Status, Is.EqualTo("Inactif"),  "La status de la deuxième personne est incorrecte.");

            
            File.Delete(filePath);  // Supprime le fichier après le test
        }

        /// <summary>
        /// Teste la conversion de différents format de date
        /// </summary>
        [Test] 
        public void TestDateFormat()
        {
            // Arrange
            string validDate1 = "05/11/1998"; // Format YYYY/mm/dd
            string validDate2 = "18-07-2001"; // Format dd-mm-YYYY
            string validDate3 = "12/23/2017"; // Format mm/dd/YYYY
            string validDate4 = "10-15-2010"; // Format mm-dd-YYYY

            //Conversion des dates
            bool result1 = DateOnly.TryParse(validDate1, out DateOnly parsedDate1);
            bool result2 = DateOnly.TryParse(validDate2, out DateOnly parsedDate2);
            bool result3 = DateOnly.TryParse(validDate3, CultureInfo.InvariantCulture, out DateOnly parsedDate3);
            bool result4 = DateOnly.TryParse(validDate4, CultureInfo.InvariantCulture, out DateOnly parsedDate4);

            // Les Tests
            Assert.That(result1, Is.True, "La conversion a echoué"); 
            Assert.That(result2, Is.True, "La conversion a echoué"); 
            Assert.That(result3, Is.True, "La conversion a echoué"); 
            Assert.That(result4, Is.True, "La conversion a echoué"); 
            Assert.That(parsedDate1, Is.EqualTo(new DateOnly(1998, 11, 05)), "La première date convertie est incorrecte");
            Assert.That(parsedDate2, Is.EqualTo(new DateOnly(2001, 07, 18)), "La deuxième date convertie est incorrecte");
            Assert.That(parsedDate3, Is.EqualTo(new DateOnly(2017, 12, 23)), "La troisième date convertie est incorrecte");
            Assert.That(parsedDate4, Is.EqualTo(new DateOnly(2010, 10, 15)), "La quatrième date convertie est incorrecte");
        }

    }
}
