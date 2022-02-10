using System;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using System.Xml;
using log4net;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.StaticFiles;
using Microsoft.Net.Http.Headers;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using ToPdf.Common;

namespace ToPdf.Controllers
{
	[Route("api/[controller]")]
	[ApiController]
	public class WordPreviewController : ControllerBase
	{
		// 默认存储位置
		private readonly string ROOT_PATH;

		private readonly IHostingEnvironment _hostingEnvironment;
		private readonly ILog _log;

		public WordPreviewController(IHostingEnvironment hostingEnvironment)
		{
			_hostingEnvironment = hostingEnvironment;
			ROOT_PATH = _hostingEnvironment.ContentRootPath;
			_log = Log4NetHelper.LogInit<WordPreviewController>();
		}

		/// <summary>
		/// 获取文件存储路径
		/// </summary>
		/// <param name="nId"></param>
		/// <returns></returns>
		private string GetFolderPath(string WordTemplateName)
		{
			//var dir = nId % 100 != 0 ? (int)(nId / 100) + 1 : (nId / 100);
			//var path = Path.Combine(ROOT_PATH, "WordFiles", WordTemplateName);
			var path = Path.Combine(ROOT_PATH, "App_Data", WordTemplateName);
			if (!Directory.Exists(path))
				Directory.CreateDirectory(path);
			return path;
		}

		/// <summary>
		/// 获取文件存储路径
		/// </summary>
		/// <param name="nId"></param>
		/// <returns></returns>
		private string GetWordPath(string WordTemplateName)
		{
			var rootPath = GetFolderPath(WordTemplateName);
			var path = Path.Combine(rootPath, "Word");
			if (!Directory.Exists(path))
				Directory.CreateDirectory(path);
			return path;
		}

		/// <summary>
		/// 通过Office生成并返回Pdf文件
		/// </summary>
		/// <param name="input"></param>
		/// <returns>File</returns>
		[HttpPost]
		[Route("GetPdf")]
		public async Task<IActionResult> GetPdf()
		{
			var fileName = Guid.NewGuid().ToString();
			var wordFilePath = Path.Combine(ROOT_PATH, "App_Data", "TemplateFiles");
			if (!Directory.Exists(wordFilePath))
			{
				Directory.CreateDirectory(wordFilePath);
			}
			try
			{
				var request = this.HttpContext.Request;
				// 确保可以多次读取
				request.EnableBuffering();
				var requestReader = new StreamReader(request.Body);
				var requestContent = await requestReader.ReadToEndAsync();
				request.Body.Position = 0;
				if (requestContent.Length > 0)
				{
					if (!Directory.Exists(wordFilePath))
					{
						Directory.CreateDirectory(wordFilePath);
					}
					// 将收到的流存成文件
					using (var fileStream = System.IO.File.Create(Path.Combine(wordFilePath, string.Format("{0}.docx", fileName))))
					{
						await requestReader.BaseStream.CopyToAsync(fileStream);
					}
					// Word文件转Pdf
					var pdfFilePath = Path.Combine(wordFilePath, string.Format("{0}.pdf", fileName));
					wordFilePath = Path.Combine(wordFilePath, string.Format("{0}.docx", fileName));
					var res = false;
					if (System.IO.File.Exists(wordFilePath))
					{
						res = WordToPDF(wordFilePath, pdfFilePath);
					}
					else
					{
						_log.Error("原始Word文件未保存成功！");
					}
					if (res)
					{
						// 获取文件的ContentType
						var provider = new FileExtensionContentTypeProvider();
						var memi = provider.Mappings[".pdf"];
						var contentDisposition = new ContentDispositionHeaderValue("inline");
						contentDisposition.SetHttpFileName(string.Format("{0}.pdf", "fileName"));
						Response.Headers[HeaderNames.ContentDisposition] = contentDisposition.ToString();
						FileStream stream = new FileStream(pdfFilePath, FileMode.Open);
						return new FileStreamResult(stream, memi);
					}
					else
					{
						return new ContentResult
						{
							Content = "Word文件转PDF报告文件失败！",
							ContentType = "text/html",
							StatusCode = 500
						};
					}
				}
				else
				{
					_log.Error("未接收到Word文件流！");
					return new ContentResult
					{
						Content = "未接收到Word文件流！",
						ContentType = "text/html",
						StatusCode = 404
					};
				}
			}
			catch (Exception e)
			{
				_log.Error(e.Message);
				return Content(e.Message);
			}
			/*
			finally
			{
				// 删除缓存
				if (Directory.Exists(wordFilePath))
				{
					var pdfPath = Path.Combine(wordFilePath, string.Format("{0}.pdf", fileName));
					if (System.IO.File.Exists(pdfPath))
					{
						System.IO.File.Delete(pdfPath);
					}
					var wordPath = Path.Combine(wordFilePath, string.Format("{0}.docx", fileName));
					if (System.IO.File.Exists(wordPath))
					{
						System.IO.File.Delete(wordPath);
					}
				}
			}
			*/
		}

		/// <summary>
		/// 通过Wps生成并返回Pdf文件
		/// </summary>
		/// <param name="length"></param>
		/// <returns></returns>
		[HttpPost]
		[Route("GetPdfByWps")]
		public async Task<IActionResult> GetPdfByWps()
		{
			var fileName = Guid.NewGuid().ToString();
			var wordFilePath = Path.Combine(ROOT_PATH, "App_Data", "TemplateFiles");
			if (!Directory.Exists(wordFilePath))
			{
				Directory.CreateDirectory(wordFilePath);
			}
			try
			{
				var request = this.HttpContext.Request;
				request.EnableBuffering();//确保可以多次读取
				var requestReader = new StreamReader(request.Body);
				//var requestContent = requestReader.ReadToEnd();
				var requestContent = requestReader.ReadToEndAsync().Result;
				request.Body.Position = 0;
				if (requestContent.Length > 0)
				{
					if (!Directory.Exists(wordFilePath))
					{
						Directory.CreateDirectory(wordFilePath);
					}
					// 将收到的流存成文件
					using (var fileStream = System.IO.File.Create(Path.Combine(wordFilePath, string.Format("{0}.wps", fileName))))
					{
						await requestReader.BaseStream.CopyToAsync(fileStream);
					}
					// Word文件转Pdf
					var pdfFilePath = Path.Combine(wordFilePath, string.Format("{0}.pdf", fileName));
					wordFilePath = Path.Combine(wordFilePath, string.Format("{0}.wps", fileName));
					bool res;
					res = WordToPdfByWPS(wordFilePath, pdfFilePath);
					if (res)
					{
						// 获取文件的ContentType
						var provider = new FileExtensionContentTypeProvider();
						var memi = provider.Mappings[".pdf"];
						var contentDisposition = new ContentDispositionHeaderValue("inline");
						contentDisposition.SetHttpFileName(string.Format("{0}.pdf", "fileName"));
						Response.Headers[HeaderNames.ContentDisposition] = contentDisposition.ToString();
						FileStream stream = new FileStream(pdfFilePath, FileMode.Open);
						return new FileStreamResult(stream, memi);
					}
					else
					{
						return new ContentResult
						{
							Content = "Word文件转Pdf文件失败！",
							ContentType = "text/html",
							StatusCode = 500
						};
					}
				}
				else
				{
					_log.Error("未接收到Word文件流！");
					return new ContentResult
					{
						Content = "获取上传Word文件流失败！",
						ContentType = "text/html",
						StatusCode = 404
					};
				}
			}
			catch (Exception e)
			{
				_log.Error(e.Message);
				return new ContentResult
				{
					Content = e.Message,
					ContentType = "text/html",
					StatusCode = 500
				};
			}
			/*
			finally
			{
				// 删除缓存
				if (Directory.Exists(wordFilePath))
				{
					var pdfPath = Path.Combine(wordFilePath, string.Format("{0}.pdf", fileName));
					if (System.IO.File.Exists(pdfPath))
					{
						System.IO.File.Delete(pdfPath);
					}
					var wordPath = Path.Combine(wordFilePath, string.Format("{0}.docx", fileName));
					if (System.IO.File.Exists(wordPath))
					{
						System.IO.File.Delete(wordPath);
					}
				}
			}
			*/
		}

		/// <summary>
		/// 使用Office将Word转Pdf
		/// </summary>
		/// <param name="sourcePath"></param>
		/// <param name="targetPath"></param>
		/// <returns></returns>
		private bool WordToPDF(string sourcePath, string targetPath)
		{
			bool result = false;
			Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
			Document document = null;
			try
			{
				application.Visible = false;
				document = application.Documents.Open(sourcePath);
				if (document != null)
				{
					document.ExportAsFixedFormat(targetPath, WdExportFormat.wdExportFormatPDF);
					result = true;
				}
				else
				{
					_log.Error("获取原始word文件Document失败");
					result = false;
				}
			}
			catch (Exception e)
			{
				_log.Error(e.Message);
				result = false;
			}
			finally
			{
				if (document != null)
				{
					document.Close();
				}
				if (application != null)
				{
					application.Quit();
				}
				foreach (System.Diagnostics.Process pro in System.Diagnostics.Process.GetProcessesByName("WINWORD"))
				{
					pro.Kill();
				}
			}
			return result;
		}

		/// <summary>
		/// 使用WPS将本地Word转PDF
		/// </summary>
		/// <param name="sourcePath"></param>
		/// <param name="targetPath"></param>
		/// <returns></returns>
		private bool WordToPdfByWPS(string sourcePath, string targetPath)
		{
			if (!FileHelper.Exists(sourcePath))
			{
				_log.Info("原始Word文件不存在！");
				return false;
			}
			Type type = null;
			dynamic doc = null;
			dynamic wpsDocs = null;
			object fileDoc = null;
			try
			{
				type = Type.GetTypeFromProgID("KWps.Application");
				if (type == null)
				{
					type = Type.GetTypeFromProgID("wps.Application");
				}
				if (type != null)
				{
					// 创建wps实例，需提前安装wps
					dynamic wps = Activator.CreateInstance(type);
					if (wps != null)
					{
						/*
						// 方式一：
						// 用wps打开word不显示界面
						doc = wps.Documents.Open(sourcePath, Visible: false);
						// 关闭拼写检查、关闭显示拼写错误提示框
						doc.SpellingChecked = false;
						doc.ShowSpellingErrors = false;
						// doc转pdf
						doc.ExportAsFixedFormat(targetPath, WdExportFormat.wdExportFormatPDF);
						*/

						// 方式二：
						// 设置为可见
						type.InvokeMember("Visible", BindingFlags.SetProperty, null, wps, new object[1] { false });
						// 得到Documents对象
						wpsDocs = type.InvokeMember("Documents", BindingFlags.GetProperty, null, wps, null);
						// 设置关键参数即可，例如: 在打开的方法中，只要指定打开的文件名与是否可见
						object[] args = new object[15];
						args[0] = sourcePath;
						args[11] = false;
						// 打开原始文件
						fileDoc = type.InvokeMember("Open", BindingFlags.InvokeMethod, null, wpsDocs, new object[1] { sourcePath });
						object[] args2 = new object[3];
						// 生成PDF
						args[0] = targetPath;
						type.InvokeMember("ExportPdf", BindingFlags.InvokeMethod, null, fileDoc, args2);
						//关闭文档工具。
						type.InvokeMember("Close", BindingFlags.InvokeMethod, null, fileDoc, null);
					}
					else
					{
						return false;
					}
				}
				else
				{
					return false;
				}
			}
			catch (Exception ex)
			{
				if (type != null && wpsDocs != null)
				{
					// 关闭文档工具。
					type.InvokeMember("Close", BindingFlags.InvokeMethod, null, wpsDocs, null);
				}
				return false;
			}
			finally
			{
				// 强制关闭所有wps的功能慎用,尤其是带并发的
				System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("wps");
				foreach (System.Diagnostics.Process prtemp in process)
				{
					prtemp.Kill();
				}

				// 关闭文档工具。
				if (type != null && fileDoc != null)
				{
					type.InvokeMember("Close", BindingFlags.InvokeMethod, null, fileDoc, null);
				}
				if (doc != null)
				{
					doc.Close(false);
				}
			}
			return true;
		}
	}
}