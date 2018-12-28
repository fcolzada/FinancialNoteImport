namespace FinancialNoteImport
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class FinanceDB : DbContext
    {
        public FinanceDB()
            : base("name=FinanceDB")
        {
        }

        public virtual DbSet<AppConfiguration> AppConfiguration { get; set; }
        public virtual DbSet<FeeType> FeeType { get; set; }
        public virtual DbSet<Raw_Finance> Raw_Finance { get; set; }
        public virtual DbSet<sysdiagrams> sysdiagrams { get; set; }
        public virtual DbSet<Transfer> Transfer { get; set; }
        public virtual DbSet<TransferType> TransferType { get; set; }
        public virtual DbSet<TransferWalletAllocation> TransferWalletAllocation { get; set; }
        public virtual DbSet<Wallet> Wallet { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<AppConfiguration>()
                .Property(e => e._namespace)
                .IsUnicode(false);

            modelBuilder.Entity<AppConfiguration>()
                .Property(e => e.name)
                .IsUnicode(false);

            modelBuilder.Entity<AppConfiguration>()
                .Property(e => e.value)
                .IsUnicode(false);

            modelBuilder.Entity<FeeType>()
                .Property(e => e.pkFee)
                .IsUnicode(false);

            modelBuilder.Entity<FeeType>()
                .Property(e => e.description)
                .IsUnicode(false);

            modelBuilder.Entity<FeeType>()
                .HasMany(e => e.Transfer)
                .WithOptional(e => e.FeeType)
                .HasForeignKey(e => e.idFee);

            modelBuilder.Entity<Raw_Finance>()
                .Property(e => e.amount)
                .HasPrecision(19, 4);

            modelBuilder.Entity<Raw_Finance>()
                .Property(e => e.description)
                .IsUnicode(false);

            modelBuilder.Entity<Raw_Finance>()
                .Property(e => e.notes)
                .IsUnicode(false);

            modelBuilder.Entity<Raw_Finance>()
                .Property(e => e.appliedTaxes)
                .HasPrecision(19, 4);

            modelBuilder.Entity<Raw_Finance>()
                .Property(e => e.appliedWelfare)
                .HasPrecision(19, 4);

            modelBuilder.Entity<Raw_Finance>()
                .Property(e => e.appliedSavings)
                .HasPrecision(19, 4);

            modelBuilder.Entity<Raw_Finance>()
                .Property(e => e.appliedAi)
                .HasPrecision(19, 4);

            modelBuilder.Entity<Raw_Finance>()
                .Property(e => e.appliedFree)
                .HasPrecision(19, 4);

            modelBuilder.Entity<Raw_Finance>()
                .Property(e => e.feeType)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<Raw_Finance>()
                .Property(e => e.hash)
                .IsUnicode(false);

            modelBuilder.Entity<Transfer>()
                .Property(e => e.m_amount)
                .HasPrecision(19, 4);

            modelBuilder.Entity<Transfer>()
                .Property(e => e.description)
                .IsUnicode(false);

            modelBuilder.Entity<Transfer>()
                .Property(e => e.note)
                .IsUnicode(false);

            modelBuilder.Entity<Transfer>()
                .Property(e => e.idFee)
                .IsUnicode(false);

            modelBuilder.Entity<Transfer>()
                .Property(e => e.idTransferType)
                .IsUnicode(false);

            modelBuilder.Entity<Transfer>()
                .HasMany(e => e.TransferWalletAllocation)
                .WithRequired(e => e.Transfer)
                .HasForeignKey(e => new { e.d_transfer, e.m_amount, e.n_sequence })
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<TransferType>()
                .Property(e => e.pkTransferType)
                .IsUnicode(false);

            modelBuilder.Entity<TransferType>()
                .Property(e => e.description)
                .IsUnicode(false);

            modelBuilder.Entity<TransferType>()
                .HasMany(e => e.Transfer)
                .WithRequired(e => e.TransferType)
                .HasForeignKey(e => e.idTransferType)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<TransferWalletAllocation>()
                .Property(e => e.m_amount)
                .HasPrecision(19, 4);

            modelBuilder.Entity<TransferWalletAllocation>()
                .Property(e => e.m_allocated_amount)
                .HasPrecision(19, 4);

            modelBuilder.Entity<Wallet>()
                .Property(e => e.p_rate)
                .IsFixedLength();

            modelBuilder.Entity<Wallet>()
                .Property(e => e.name)
                .IsUnicode(false);

            modelBuilder.Entity<Wallet>()
                .Property(e => e.description)
                .IsUnicode(false);

            modelBuilder.Entity<Wallet>()
                .HasMany(e => e.TransferWalletAllocation)
                .WithRequired(e => e.Wallet)
                .HasForeignKey(e => e.idWallet)
                .WillCascadeOnDelete(false);
        }
    }
}
