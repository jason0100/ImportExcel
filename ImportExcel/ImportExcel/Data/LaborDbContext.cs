using ImportExcel.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ImportExcel.Data
{
	public class LaborDbContext:DbContext
	{
		public LaborDbContext(DbContextOptions<LaborDbContext> options) : base(options)
		{ }
		public DbSet<Labor> Labors { get; set; }
	}
}
