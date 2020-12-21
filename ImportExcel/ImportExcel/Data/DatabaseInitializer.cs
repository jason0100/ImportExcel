using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ImportExcel.Data
{
	public class DatabaseInitializer
	{
		public static void Initialize(LaborDbContext context)
		{
			context.Database.EnsureCreated();


		}
	}
}
