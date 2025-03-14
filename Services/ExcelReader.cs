using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataImporterProject.Models;
using OfficeOpenXml;

namespace ExcelDataImporterProject.Services
{
    class ExcelReader
    {
        /// <summary>
        /// Lit un fichier Excel (.xlsx) à partir du chemin spécifié, extrait les données des personnes et retourne une liste de Person
        /// </summary>
        /// <param name="filePath">chemin du fichier excel</param>
        /// <returns></returns>
        public static List<Person> ReadExcelFile(string filePath)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            var persons = new List<Person>();
            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    //Récupération de la première feuille de calcul du fichier Excel
                    var worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;

                    //Parcourt toutes les lignes de la feuille de calcul, en ignorant l'en-tête
                    for (int row = 2; row <= rowCount; row++)  
                    {
                        try
                        {
                            string dateText = worksheet.Cells[row, 4].Text;
                            DateOnly? birthDate = null;

                            //Vérification de la validité de la date de naissance
                            //Traite les formats YYYY-mm-dd , dd-mm-YYYY , YYYY/mm/dd et dd/mm/YYYY
                            if (DateOnly.TryParse(dateText, out DateOnly parsedDate))
                            {
                                birthDate = parsedDate;
                            }
                            //Traite les formats YYYY-dd-mm , mm-dd-YYYY , YYYY/dd/mm et mm/dd/YYYY
                            else if (DateOnly.TryParse(dateText, CultureInfo.InvariantCulture, out DateOnly parsedbirthDate))
                            {
                                birthDate = parsedbirthDate;
                            }
                            else
                            {
                                Console.WriteLine($"Ligne {row} : Date invalide ('{dateText}'), valeur ignorée.");
                            }

                            // Crée une nouvelle instance de Person à partir des données de la ligne courante
                            // et l'ajoute à la liste des personnes
                            var person = new Person
                            {
                                Number = worksheet.Cells[row, 1].Text,
                                Surname = worksheet.Cells[row, 2].Text,
                                Firstname = worksheet.Cells[row, 3].Text,
                                Date = birthDate,
                                Status = worksheet.Cells[row, 5].Text,
                            };
                            persons.Add(person);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Erreur de lecture à la ligne {row}: {ex.Message}");
                        }
                        
                    }
                }
                
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"Erreur : {fnfEx.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Une erreur est survenue lors de la lecture du fichier : {ex.Message}");
            }
            return persons;
        }
    }
}
