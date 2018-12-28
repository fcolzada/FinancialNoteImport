namespace FinancialNoteImport
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("conf.AppConfiguration")]
    public partial class AppConfiguration
    {
        [Key]
        [Column("namespace", Order = 0)]
        [StringLength(200)]
        public string _namespace { get; set; }

        [Key]
        [Column(Order = 1)]
        [StringLength(50)]
        public string name { get; set; }

        [Required]
        [StringLength(200)]
        public string value { get; set; }
    }
}
