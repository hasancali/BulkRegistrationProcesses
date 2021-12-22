using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text;

namespace BulkRegistrationProcesses.Models
{
    /// <summary>
    /// Stok
    /// </summary>
    [TableAttribute("Stock")]
    public class Stock
    {
        /// <summary>
        /// Id
        /// </summary>
        /// 
        [Key]
        public int Id { get; set; }
        /// <summary>
        /// Name
        /// </summary>
        [Required]
        [StringLength(50)]
        public string Name { get; set; }
    }
}
