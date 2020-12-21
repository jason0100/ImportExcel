using ImportExcel.Models;
using ImportExcel.Models.Upload;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ImportExcel.Helpers
{
    public interface IUploadHelper
    {

        ResultModel UploadData(uploadFile uploadData);
    }
}
