using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using ImportExcel.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using System.IO;
using ImportExcel.Helpers;
using ImportExcel.Models.Upload;
using System.Text;
using ImportExcel.Fillters;
using Microsoft.Extensions.Hosting;
using OfficeOpenXml;
using ImportExcel.Data;
using System.Globalization;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Dapper;
using OfficeOpenXml.Style;
using System.Drawing;

namespace ImportExcel.Controllers
{
	public class HomeController : Controller
	{
		private readonly ILogger<HomeController> _logger;
		private readonly IUploadHelper _uploadHelper;
		private readonly IWebHostEnvironment _env;
		private readonly LaborDbContext _DBContext;
		private readonly IConfiguration _config;

		private readonly static Dictionary<string, string> _contentTypes = new Dictionary<string, string>
		{
			{".png", "image/png"},
			{".jpg", "image/jpeg"},
			{".jpeg", "image/jpeg"},
			{".gif", "image/gif"}
		};
		private readonly string _folder;

		public HomeController(ILogger<HomeController> logger, IWebHostEnvironment env, IUploadHelper uploadHelper, LaborDbContext DBContext, IConfiguration config)
		{
			_logger = logger;
			// 把上傳目錄設為：wwwroot\UploadFolder
			_folder = $@"{env.WebRootPath}\UploadFolder";
			_uploadHelper = uploadHelper;
			_env = env;
			_DBContext = DBContext;
			_config = config;
		}

		public IActionResult Index()
		{
			return View();
		}

		public IActionResult Privacy()
		{
			return View();
		}

		[ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
		public IActionResult Error()
		{
			return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
		}

		[HttpGet]
		public IActionResult ImportExcel()
		{

			DirectoryInfo dir = new DirectoryInfo(_folder);
			var files = dir.GetFiles();
			var sb = new StringBuilder();
			foreach (var i in files)
			{
				if (_contentTypes.ContainsKey(i.Extension))
				{
					sb = sb.Append("<a target=\"_blank\" href=Download?fileName=" + i.Name + " >" + i.Name + "</a>" + "\t");
					sb = sb.Append(i.Length / 1024 + " KB\t");
					sb = sb.Append(@"<br>");
				}
			}
			ViewBag.fileList = sb;
			return View();
		}



		//適合小檔案上傳
		[HttpPost]
		public async Task<IActionResult> Upload(uploadFile data)
		{
			var result = new ResultModel();


			if (!ModelState.IsValid)
			{
				Console.WriteLine("error");
				result.IsSuccess = false;
				//result.Message = ModelState.Values();
				//return result;
				var error = ModelState.Values.Where(a => a.Errors.Count() > 0).SelectMany(a => a.Errors).Select(a => a.ErrorMessage).ToList();
				var errorStr = error[0].ToString();
				var sb = new StringBuilder();
				TempData["result"] = data.file.FileName + "<a href=上傳失敗>" + errorStr;
				return RedirectToAction("ImportExcel");
			}

			data.folder = "uploadFile";

			result = _uploadHelper.UploadData(data);
			if (result.IsSuccess)
				TempData["result"] = data.file.FileName + "<a href=上傳成功>";

			else
				TempData["result"] = data.file.FileName + "上傳失敗";
			return RedirectToAction("ImportExcel");
		}

		//大檔案上傳
		[Route("album")]
		[HttpPost]
		[DisableFormValueModelBindingFilter]
		public async Task<IActionResult> Album()
		{
			var photoCount = 0;
			var formValueProvider = await Request.StreamFile((file) =>
			{
				photoCount++;
				return System.IO.File.Create($"{_folder}\\{file.FileName}");
			});

			var model = new AlbumModel
			{
				Title = formValueProvider.GetValue("title").ToString(),
				Date = Convert.ToDateTime(formValueProvider.GetValue("date").ToString())
			};

			// ...

			return Ok(new
			{
				title = model.Title,
				date = model.Date.ToString("yyyy/MM/dd"),
				photoCount = photoCount
			});
		}

		//匯入Excel
		[HttpPost]
		public async Task<IActionResult> Import(IFormFile excelfile)
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial; //EEplus 關閉新許可模式通知
			string sWebRootFolder = _env.ContentRootPath;
			string sFileName = $"{Guid.NewGuid()}.xlsx";
			FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, _folder, sFileName));
			try
			{
				using (FileStream fs = new FileStream(file.ToString(), FileMode.Create))
				{
					excelfile.CopyTo(fs);
					fs.Flush();
				}
				using (ExcelPackage package = new ExcelPackage(file))
				{
					//轉換日期使用
					CultureInfo culture = new CultureInfo("zh-TW");
					culture.DateTimeFormat.Calendar = new TaiwanCalendar();
					StringBuilder sb = new StringBuilder();
					ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
					int skipRow = 6;
					int rowCount = worksheet.Dimension.Rows;
					int ColCount = worksheet.Dimension.Columns;

					bool stop = false;
					for (int row = 1; row <= rowCount; row++)
					{
						if (row <= skipRow)
							continue;
						if (stop)
							break;

						var labor = new Labor();
						for (int col = 1; col <= ColCount; col++)
						{
							//當序號沒數字時不再讀取
							if (col == 1 && worksheet.Cells[row, col].Value == null)
							{
								stop = true;
								break;
							}


							string text = "";
							if (worksheet.Cells[row, col].Value != null)
							{
								text = worksheet.Cells[row, col].Value.ToString();
							}

							switch (col)
							{
								case 2:
									labor.Name = text.Trim();
									break;
								case 3:
									labor.Id = text.ToUpper().Trim();
									break;
								case 4:
									text = text.PadLeft(8, '0');
									labor.Birthdate = DateTime.ParseExact(text.Trim(), "yyyMMdd", culture);
									break;
								case 5:
									labor.Salary = Convert.ToInt32(text.Trim());
									break;
								case 6:
									labor.aa = Convert.ToInt32(text.Trim());
									break;
								case 7:
									text = text.PadLeft(8, '0');
									labor.aaDate = DateTime.ParseExact(text.Trim(), "yyyMMdd", culture);
									break;

							}
							//印出畫面使用
							//	sb.Append(worksheet.Cells[row, col].Value==null?"": worksheet.Cells[row, col].Value.ToString() +  "\t");


						}
						//印出畫面使用
						//sb.Append(Environment.NewLine);
						if (!stop)
							await _DBContext.Labors.AddAsync(labor);
					}
					await _DBContext.SaveChangesAsync();
					return Content(sb.ToString());
				}
			}
			catch (Exception ex)
			{
				return Content(ex.Message);
			}
		}

		[HttpGet]
		public ActionResult ShowDB()
		{

			string sqlCommand = @"select * from Labors";
			List<Labor> labors = null;
			//query from sql command
			/*	using (var conn = new SqlConnection(_config.GetConnectionString("serverConnection")))
				{
					conn.Open();
					labors = conn.Query<Labor>(sqlCommand).ToList();
				}*/
			
			//Query from store procedure
			using (var conn = new SqlConnection(_config.GetConnectionString("serverConnection")))
			{
				
				labors = conn.Query<Labor>("dbo.列出不重複員工名冊",commandType:System.Data.CommandType.StoredProcedure).ToList();
			}


			ViewBag.labors = labors;
			return View();

		}

		[HttpGet]
		public async Task<IActionResult> Download(string fileName)
		{
			if (string.IsNullOrEmpty(fileName))
			{
				return NotFound();
			}

			var path = $@"{_folder}\{fileName}";
			var memoryStream = new MemoryStream();
			using (var stream = new FileStream(path, FileMode.Open))
			{
				await stream.CopyToAsync(memoryStream);
			}
			memoryStream.Seek(0, SeekOrigin.Begin);

			// 回傳檔案到 Client 需要附上 Content Type，否則瀏覽器會解析失敗。
			return new FileStreamResult(memoryStream, _contentTypes[Path.GetExtension(path).ToLowerInvariant()]);
		}



		[HttpGet]
		public async Task<IActionResult> ExportExcel()
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial; //EEplus 關閉新許可模式通知
			var labors = new List<Labor>();
			using (var conn = new SqlConnection(_config.GetConnectionString("serverConnection")))
			{

				labors = conn.Query<Labor>("dbo.列出不重複員工名冊", commandType: System.Data.CommandType.StoredProcedure).ToList();
			}
			var memoryStream = new MemoryStream();

			var output = new FileInfo(Path.Combine(_env.WebRootPath, _folder, "ExportExcelTest-" + DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss") + ".xlsx"));
		
			using (var excel = new ExcelPackage(memoryStream))
			{
				var ws = excel.Workbook.Worksheets.Add("sheet1");
				var properties = typeof(Labor).GetProperties();
				var rows = labors.Count() + 1;// 直：資料筆數（記得加標題列）
				var cols = properties.Count();// 橫：類別中有別名的屬性數量
				if (rows > 0 && cols > 0)
				{
					ws.Cells[1, 1].LoadFromCollection(labors, true);

					// 儲存格格式
					var colNumber = 1;
					foreach (var prop in properties)
					{
						if (prop.PropertyType.Equals(typeof(DateTime)))
						{
							ws.Cells[2, colNumber, rows, colNumber].Style.Numberformat.Format = "mm-dd-yy";
						}
						colNumber++;
					}

					// 樣式準備
					using (var range = ws.Cells[1, 1, rows, cols])
					{
						ws.Cells.Style.Font.Name = "新細明體";
						ws.Cells.Style.Font.Size = 12;
						ws.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // 置中
						ws.Cells.AutoFitColumns(); // 欄寬

						// 框線
						range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
						range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
						range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
						range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

						// 標題列
						var title = ws.Cells[1, 1, 1, cols];
						title.Style.Fill.PatternType = ExcelFillStyle.Solid; // 設定背景填色方法
						title.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
					}
				}

				else
				{
					Debug.WriteLine("未列印資料，請檢查是否傳入資料為空，或指定類別未具有公開且加上 DisplayName 的屬性。");
				}
				excel.Save(); // 儲存 Excel
			}
			//string fileName = "ExportExcelTest-" + DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss") + ".xlsx";
			var path = Path.Combine(_folder, output.Name);


			//法1:從fileContext
			//var fileContents = System.IO.File.ReadAllBytes(path);

			//return File(memoryStream, "application/vnd.ms-excel", "勞保員工清冊.xlsx");
			//return File(fileContents, "application/vnd.ms-excel","勞保員工清冊.xlsx");

			//法2:從memoryStream

			memoryStream.Seek(0, SeekOrigin.Begin);
			return new FileStreamResult(memoryStream, "application/vnd.ms-excel")
			{
				FileDownloadName = "勞保員工清冊.xlsx"
			};
				
		}


	}
}
