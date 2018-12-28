namespace FinancialNoteImport
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("TransferWalletAllocation")]
    public partial class TransferWalletAllocation
    {
        [Key]
        [Column(Order = 0, TypeName = "date")]
        public DateTime d_transfer { get; set; }

        [Key]
        [Column(Order = 1, TypeName = "money")]
        public decimal m_amount { get; set; }

        [Key]
        [Column(Order = 2)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int n_sequence { get; set; }

        [Key]
        [Column(Order = 3)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public long idWallet { get; set; }

        [Column(TypeName = "money")]
        public decimal m_allocated_amount { get; set; }

        public virtual Transfer Transfer { get; set; }

        public virtual Wallet Wallet { get; set; }
    }
}
