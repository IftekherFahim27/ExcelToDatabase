using ExcelToDatabase.Models;
using Microsoft.EntityFrameworkCore;

namespace ExcelToDatabase.Services
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions options) : base(options)
        {

        }

        public DbSet<Category> Categories { get; set; }
        public DbSet<Item> Items { get; set; }


    }
}
