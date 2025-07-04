using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DemoMVCProject1.Data;
using DemoMVCProject1.Models;

namespace DemoMVCProject1.Services
{
    public static class ImportService
    {
        public static void SaveCustomers(List<Customer> customers)
        {
            var context = new ApplicationDbContext();
            context.Customers.AddRange(customers);
            context.SaveChanges();
        }
    }

}