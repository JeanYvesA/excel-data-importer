using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataImporterProject.Models;

namespace ExcelDataImporterProject.Services
{
    /// <summary>
    /// Classe responsable de l'interaction avec la base de données
    /// </summary>
    class DatabaseService
    {
        /// <summary>
        /// Insère une liste de personnes dans la base de données
        /// </summary>
        /// <param name="persons">Liste des personnes à insérer</param>
        public static void InsertPersons(List<Person> persons)
        {
            using (var context = new MyDbContext())
            {
                context.Persons.AddRange(persons);
                context.SaveChanges();
            }
        }
    }
}
