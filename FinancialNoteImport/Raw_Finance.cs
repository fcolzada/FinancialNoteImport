namespace FinancialNoteImport
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class Raw_Finance
    {
        [Key]
        public long pkRaw_HouseFinance { get; set; }

        [Column(TypeName = "money")]
        public decimal amount { get; set; }

        [Column(TypeName = "date")]
        public DateTime valueDate { get; set; }

        [StringLength(1000)]
        public string description { get; set; }

        [StringLength(200)]
        public string notes { get; set; }

        [Column(TypeName = "money")]
        public decimal? appliedTaxes { get; set; }

        [Column(TypeName = "money")]
        public decimal? appliedWelfare { get; set; }

        [Column(TypeName = "money")]
        public decimal? appliedSavings { get; set; }

        [Column(TypeName = "money")]
        public decimal? appliedAi { get; set; }

        [Column(TypeName = "money")]
        public decimal? appliedFree { get; set; }

        [StringLength(2)]
        public string feeType { get; set; }

        public int? n_sequence { get; set; }

        [Column(TypeName = "smalldatetime")]
        public DateTime? loadDate { get; set; }

        [StringLength(32)]
        public string hash { get; set; }
    }
}
