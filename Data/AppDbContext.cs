using Microsoft.EntityFrameworkCore;
using DataLink.Models;

namespace DataLink.Data
{
    public class AppDbContext : DbContext
    {
        // Before starting, you need to enter the name of your server and database
        private const string ConnectionString = "Server=(localdb)\\MSSQLLocalDB;Database=DataLinkDb;Trusted_Connection=True;TrustServerCertificate=True;";

        public DbSet<Record> Records { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer(ConnectionString);
        }
    }

}
