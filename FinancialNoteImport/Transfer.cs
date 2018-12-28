namespace FinancialNoteImport
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Transfer")]
    public partial class Transfer
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Transfer()
        {
            TransferWalletAllocation = new HashSet<TransferWalletAllocation>();
        }

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

        [Required]
        [StringLength(1000)]
        public string description { get; set; }

        [StringLength(200)]
        public string note { get; set; }

        [StringLength(2)]
        public string idFee { get; set; }

        [Required]
        [StringLength(2)]
        public string idTransferType { get; set; }

        public virtual FeeType FeeType { get; set; }

        public virtual TransferType TransferType { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<TransferWalletAllocation> TransferWalletAllocation { get; set; }
    }
}
