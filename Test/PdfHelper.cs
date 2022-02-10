using Aspose.Words;
using Microsoft.Office.Interop.Word;
using Spire.Doc;
using System;
using System.IO;
using System.Reflection;

namespace Test
{
	public static class PdfHelper
	{
		/// <summary>
		/// 使用Office将本地Word转PDF
		/// </summary>
		/// <param name="sourcePath"></param>
		/// <param name="targetPath"></param>
		/// <returns></returns>
		public static bool Word2PDF(string sourcePath, string targetPath)
		{
			bool result = false;
			Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
			Microsoft.Office.Interop.Word.Document document = null;
			try
			{
				application.Visible = false;
				document = application.Documents.Open(sourcePath);
				document.ExportAsFixedFormat(targetPath, WdExportFormat.wdExportFormatPDF);
				result = true;
			}
			catch (Exception e)
			{
				Console.WriteLine(e.Message);
				result = false;
			}
			finally
			{
				document.Close();
				application.Quit();
			}

			return result;
		}

		/// <summary>
		/// 使用WPS将本地Word转PDF
		/// </summary>
		/// <param name="sourcePath"></param>
		/// <param name="targetPath"></param>
		/// <returns></returns>
		public static bool Word2PdfByWPS(string sourcePath, string targetPath)
		{
			if (!File.Exists(sourcePath))
			{
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
				//if (type != null && wpsDocs != null)
				//{
				//	// 关闭文档工具。
				//	type.InvokeMember("Close", BindingFlags.InvokeMethod, null, wpsDocs, null);
				//}
				//return false;
				throw ex;
			}
			finally
			{
				// 强制关闭所有wps的功能慎用,尤其是带并发的
				//System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("wps");
				//foreach (System.Diagnostics.Process prtemp in process)
				//{
				//	prtemp.Kill();
				//}
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

		/// <summary>
		/// 使用FreeSpire.Doc将本地Word转PDF，收费版，免费版只支持文档前3页转换
		/// </summary>
		/// <param name="sourcePath">原始文件路径</param>
		/// <param name="targetPath">目标文件路径</param>
		/// <returns></returns>
		public static bool Word2PdfByFreeSpire(string sourcePath, string targetPath)
		{
			if (!File.Exists(sourcePath))
			{
				return false;
			}
			Spire.Doc.Document document = new Spire.Doc.Document();
			document.LoadFromFile(sourcePath);
			//保存为PDF格式
			document.SaveToFile(targetPath, FileFormat.PDF);
			return true;
		}

		/// <summary>
		/// 使用Aspose.Word将本地Word转PDF，收费版，带水印，有页数限制
		/// </summary>
		/// <param name="sourcePath">原始文件路径</param>
		/// <param name="targetPath">目标文件路径</param>
		public static bool Word2PdfByAspose(string sourcePath, string targetPath)
		{
			if (!File.Exists(sourcePath))
			{
				return false;
			}
			// 打开word文档，将doc文档转为pdf文档
			Aspose.Words.Document doc = new Aspose.Words.Document(sourcePath);
			if (doc == null)
			{
				return false;
			}
			doc.Save(targetPath, SaveFormat.Pdf);
			return true;
		}

		/// <summary>
		/// 使用Aspose.Word将本地Word转Png图片，收费版，带水印，有页数限制
		/// </summary>
		/// <param name="docFile"></param>
		/// <param name="pngDir"></param>
		/// <param name="pngCount"></param>
		/// <returns></returns>
		public static bool Word2Png(string docFile, string pngDir, out int pngCount)
		{
			Aspose.Words.Saving.ImageSaveOptions options = new Aspose.Words.Saving.ImageSaveOptions(SaveFormat.Png);
			options.Resolution = 300;
			options.PrettyFormat = true;
			options.UseAntiAliasing = true;

			pngCount = 0;
			try
			{
				Aspose.Words.Document doc = new Aspose.Words.Document(docFile);
				for (int i = 0; i < doc.PageCount; i++)
				{
					options.PageIndex = i;
					doc.Save(Path.Combine(pngDir, i + ".png"), options);

					pngCount++;
				}
				return true;
			}
			catch
			{
				return false;
			}
		}

		/// <summary>
		/// byte[]数组保存为文件
		/// </summary>
		/// <param name="path"></param>
		/// <param name="byteArr"></param>
		/// <returns></returns>
		public static bool ByteToFile(string path, byte[] byteArr)
		{
			var result = false;
			try
			{
				using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write))
				{
					fs.Write(byteArr, 0, byteArr.Length);
					result = true;
				}
			}
			catch
			{
				result = false;
			}
			return result;
		}
	}
}