using ImportExcel.Attributes;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ImportExcel.Models.Upload
{
    public class uploadFile
    {
		[Required]
		[DataType(DataType.Upload)]
		[MaxFileSize(5 * 1024 * 1024)]
		[AllowedExtensions(new string[] { ".xls", ".xlsx", ".exe" })]
		public IFormFile file { get; set; }
        
        public string folder { get; set; }
    }
  
}
