using BulkRegistrationProcesses.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Text;

namespace BulkRegistrationProcesses
{
    public class Context : DbContext
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="options"></param>
        public Context(DbContextOptions<DbContext> options) : base(options)
        {
        }

        public DbSet<Stock> Stocks { get; set; }
    }
}
