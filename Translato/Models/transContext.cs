using System;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata;

namespace Translato.Models
{
    public partial class transContext : DbContext
    {
        public transContext()
        {
        }

        public transContext(DbContextOptions<transContext> options)
            : base(options)
        {
        }

        public virtual DbSet<Trans> Trans { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
#warning To protect potentially sensitive information in your connection string, you should move it out of source code. See http://go.microsoft.com/fwlink/?LinkId=723263 for guidance on storing connection strings.
                optionsBuilder.UseSqlServer("Data Source=localhost\\SQLEXPRESS;Initial Catalog=Translate;Integrated Security=True");
            }
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Trans>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("trans");

                entity.Property(e => e.Code)
                    .HasColumnName("CODE")
                    .HasMaxLength(255);

                entity.Property(e => e.Group).HasMaxLength(255);

                entity.Property(e => e.It)
                    .HasColumnName("IT")
                    .HasMaxLength(255);

                entity.Property(e => e.Ua)
                    .HasColumnName("UA")
                    .HasMaxLength(255);
            });

            OnModelCreatingPartial(modelBuilder);
        }

        partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
    }
}
