
using ImportExcel.Attributes;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ImportExcel.Models.Upload
{
    public class uploadPhoto:IValidatableObject
    {
        [Required]
        //[AllowedExtensions(new string[] { ".jpg", ".png" })]
        public List<IFormFile> files { get; set; }
        [Required]
        public string folder { get; set; }

        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
        {
            string[] fileFormats = new string[] {".jpg",".png" };
            bool isValidFormat = true;
            foreach (var i in files)
            {
                var extension = Path.GetExtension(i.FileName);
                if (!Array.Exists(fileFormats, a => a == extension))
                {
                    isValidFormat = false;
                    break;
                }
               
            }
            if (!isValidFormat)
                yield return new ValidationResult("檔案格式不符 .jpg or .png", new string[] { "files" });
        }
    }
}
