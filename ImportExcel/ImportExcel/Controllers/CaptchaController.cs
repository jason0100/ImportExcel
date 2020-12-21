using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Hei.Captcha;
using Microsoft.AspNetCore.Mvc;

namespace ImportExcel.Controllers
{
	public class CaptchaController : Controller
	{
		private readonly SecurityCodeHelper _securityCode;

		public CaptchaController(SecurityCodeHelper securityCode)
		{
			_securityCode = securityCode;
		}
		public IActionResult Test()
		{

			return View();
		}

		public ActionResult GetValidateCode(string guid)
		{
			byte[] data = null;
			string code = RandomCode(5);

			TempData["code"] = code;
			//TempData.Add(guid, 0);
			//定義一個畫板
			MemoryStream ms = new MemoryStream();
			using (Bitmap map = new Bitmap(100, 40))
			{
				//畫筆,在指定畫板畫板上畫圖
				//g.Dispose();
				using (Graphics g = Graphics.FromImage(map))
				{
					g.Clear(Color.White);
					g.DrawString(code, new Font("黑體", 18.0F), Brushes.Blue, new Point(10, 8));
					//繪製干擾線(數字代表幾條)
					PaintInterLine(g, 10, map.Width, map.Height);
				}
				map.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
			}
			data = ms.GetBuffer();
			return File(data, "image/jpeg");
		}



		[HttpPost]
		public ActionResult Verify(string guid)
		{
			if (!TempData.ContainsKey("code"))
				return RedirectToAction("Test");
			string code = Request.Form["code"].ToString();
			if (code == TempData["code"].ToString())
			{
				ViewBag.code = code;
				ViewBag.Ans = TempData["code"];
				ViewBag.Result = "驗證正確";
				return View();
			}
			else
			{
				ViewBag.code = code;
				ViewBag.Ans = TempData["code"];
				ViewBag.Result = "驗證錯誤";
				return View();
			}


		}


		//隨機生成指定長度的驗證碼字符串
		private string RandomCode(int length)
		{
			string s = "0123456789zxcvbnmasdfghjklqwertyuiop";
			StringBuilder sb = new StringBuilder();
			Random rand = new Random();
			int index;
			for (int i = 0; i < length; i++)
			{
				index = rand.Next(0, s.Length);
				sb.Append(s[index]);
			}
			return sb.ToString();


		}

		//產生刪除線 num 代表幾條
		private void PaintInterLine(Graphics g, int num, int width, int height)
		{
			Random r = new Random();
			int startX, startY, endX, endY;
			for (int i = 0; i < num; i++)
			{
				startX = r.Next(0, width);
				startY = r.Next(0, height);
				endX = r.Next(0, width);
				endY = r.Next(0, height);
				g.DrawLine(new Pen(Brushes.Red), startX, startY, endX, endY);
			}
		}



		/// <summary>
		/// 泡泡中文验证码 
		/// </summary>
		/// <returns></returns>
		public IActionResult BubbleCode()
		{
			var code = _securityCode.GetRandomCnText(2);
			var imgbyte = _securityCode.GetBubbleCodeByte(code);

			return File(imgbyte, "image/png");
		}

		/// <summary>
		/// 数字字母组合验证码
		/// </summary>
		/// <returns></returns>
		public IActionResult HybridCode()
		{
			var code = _securityCode.GetRandomEnDigitalText(4);
			var imgbyte = _securityCode.GetEnDigitalCodeByte(code);

			return File(imgbyte, "image/png");
		}

		/// <summary>
		/// gif泡泡中文验证码 
		/// </summary>
		/// <returns></returns>
		public IActionResult GifBubbleCode()
		{
			var code = _securityCode.GetRandomCnText(2);
			var imgbyte = _securityCode.GetGifBubbleCodeByte(code);

			return File(imgbyte, "image/gif");
		}

		/// <summary>
		/// gif数字字母组合验证码
		/// </summary>
		/// <returns></returns>
		public IActionResult GifHybridCode()
		{
			var code = _securityCode.GetRandomEnDigitalText(4);
			var imgbyte = _securityCode.GetGifEnDigitalCodeByte(code);

			return File(imgbyte, "image/gif");
		}

	}
}
