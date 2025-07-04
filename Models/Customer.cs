using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;

namespace DemoMVCProject1.Models
{
    public class Customer
    {
        [Key]
        public int Id { get; set; }
        [Required,MaxLength(100)]
        public string Name { get; set; }
        [Required, MaxLength(50)]
        public string Email { get; set; }
        [Required, MaxLength(50)]
        public string Phone { get; set; }
        [MaxLength(250)]
        public string Address { get; set; }
    }

}