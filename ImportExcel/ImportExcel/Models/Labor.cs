using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace ImportExcel.Models
{
	public class Labor
	{
		[DatabaseGenerated(DatabaseGeneratedOption.Identity)]
		[Key]
		public int sn { get; set; }
		public string Name { get; set; }
		public string Id { get; set; }
		public DateTime Birthdate { get; set; }
		public decimal Salary { get; set; }
		public int aa { get; set; }
		public DateTime aaDate { get; set; }
		

	}
}
