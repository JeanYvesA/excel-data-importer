using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;

namespace ExcelDataImporterProject.Models
{
    /// <summary>
    /// Contexte de données pour l'interaction avec la base de données MySQL
    /// </summary>
    class MyDbContext : DbContext
    {
        const string connectionString = "server=localhost;database=excel;user=root;";
        public DbSet<Person> Persons { get; set; }
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            var connectionString = "server=localhost;database=excel;user=root;";
            var serverVersion = ServerVersion.AutoDetect(connectionString);
            optionsBuilder.UseMySql(connectionString, serverVersion);
        }
    }
}
