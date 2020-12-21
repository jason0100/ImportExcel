
using ImportExcel.Models;
using ImportExcel.Models.Upload;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ImportExcel.Helpers
{

    public class UploadHelper : IUploadHelper
    {

        private readonly IConfiguration _config;
        private readonly IWebHostEnvironment _env;

        public UploadHelper(IWebHostEnvironment env, IConfiguration config)
        {
            _env = env;
            _config = config;

        }


        public ResultModel UploadData(uploadFile uploadData)
        {
            var result = new ResultModel();
            List<string> fileNames = new List<string>();

            //if (uploadData.files.Count() <= 0) return null;
            string path = System.IO.Path.Combine(_env.WebRootPath, uploadData.folder);
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
         /*   foreach (var file in uploadData.files)
            {*/
                var fileName = DateTime.Now.ToString("yyyyMMdd") + "_" + DateTime.Now.Ticks + Path.GetExtension(uploadData.file.FileName).ToLower();

                var filePath = Path.Combine(path, fileName);
                try
                {
                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        uploadData.file.CopyTo(stream);
                        fileNames.Add(fileName);
                    }

                }
                catch (Exception e)
                {
                    result.IsSuccess = false;
                    result.Message = "上傳失敗";
                    return result;
                }
            //}
            result.IsSuccess = true;
            result.Message = "上傳成功";
            result.Data = fileNames;
            return result;


        }

    }
}
