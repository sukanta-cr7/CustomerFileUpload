using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.Entity;
using DemoMVCProject1.Models;

namespace DemoMVCProject1.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext() : base("DefaultConnection") { }

        public DbSet<Customer> Customers { get; set; }
    }

}