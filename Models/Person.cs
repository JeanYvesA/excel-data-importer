using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataImporterProject.Models
{
    /// <summary>
    /// Représente une personne avec ses informations personnelles
    /// </summary>
    class Person
    {
        public int Id { get; set; }
        public string? Number { get; set; }
        public string? Surname { get; set; }
        public string? Firstname { get; set; }
        public DateOnly? Date { get; set; }
        public string? Status { get; set; }
    }
}
