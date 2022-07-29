using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.SqlServer;

namespace App
{
    public class AppLogicContext : DbContext
    {
        private const string connectionString = @"Data Source=localhost\sqlexpress; Initial Catalog=WTW-WorkFlow-Example; Integrated Security=True";

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer(connectionString);
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {//indico que LibroAutor tiene dos claves primarias en una entidad
            modelBuilder.Entity<Users>().HasKey(xi => xi.UserId);
            modelBuilder.Entity<ValidatedUsers>().HasKey(xi => xi.UserId);
            modelBuilder.Entity<IncorrectUsers>().HasKey(xi => xi.UserId);
            modelBuilder.Entity<FinalReport>().HasKey(xi => xi.UserId);
        }


        public DbSet<Users> Users { get; set; }

        public DbSet<ValidatedUsers> ValidatedUsers { get; set; }
        public DbSet<IncorrectUsers> IncorrectUsers { get; set; }
        public DbSet<FinalReport> FinalReport { get; set; }


    }
}