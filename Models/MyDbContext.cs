using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;

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
            // Charge la configuration depuis appsettings.json
            var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory()) // Récupère le dossier de base
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true) // Charge le fichier JSON
                .Build();

            // Lire la chaîne de connexion
            string? connectionString = configuration.GetConnectionString("DefaultConnection");
            var serverVersion = ServerVersion.AutoDetect(connectionString);
            optionsBuilder.UseMySql(connectionString, serverVersion);
        }
    }
}
