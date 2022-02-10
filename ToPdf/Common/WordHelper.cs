using Newtonsoft.Json.Linq;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace ToPdf.Common
{
	/// <summary>
	/// Word模版
	/// </summary>
	public class WordHelper
	{
		/// <summary>
		///
		/// </summary>
		private WmlDocument _WordDoc;

		/// <summary>
		/// Word模版路径
		/// </summary>
		private string _TemplateFilePath;

		/// <summary>
		/// 实体类型
		/// </summary>
		private Type _ModelType;

		/// <summary>
		/// 实体数据
		/// </summary>
		private XmlDocument _ModelData;

		/// <summary>
		/// Word文件输出文件夹
		/// </summary>
		private string _OutputFolderPath;

		/// <summary>
		/// Word文件输出路径
		/// </summary>
		private string _OutputFilePath;

		/// <summary>
		/// 构造函数
		/// </summary>
		/// <param name="WordTemplate">Word模版</param>
		/// <param name="ModelType">数据类型</param>
		/// <param name="ModelData">数据</param>
		/// <param name="OutputFilePath">Word输出路径</param>
		public WordHelper(string TemplateFilePath, Type ModelType, XmlDocument ModelData, string OutputFilePath)
		{
			_ModelData = ModelData ?? throw new ArgumentNullException("缺少实体数据");
			_TemplateFilePath = TemplateFilePath ?? throw new ArgumentNullException("缺少模版路径");
			_ModelType = ModelType;
			_OutputFilePath = OutputFilePath;
			_OutputFolderPath = Path.GetDirectoryName(OutputFilePath);
			// 获取word模版文件
			var templateDoc = new FileInfo(_TemplateFilePath);
			_WordDoc = new WmlDocument(templateDoc.FullName);
		}

		/// <summary>
		/// 构造函数
		/// </summary>
		/// <param name="fileName">文件路径</param>
		public WordHelper(string fileName)
		{
			var templateDoc = new FileInfo(fileName);
			_WordDoc = new WmlDocument(templateDoc.FullName);
		}

		/// <summary>
		/// 生成Word
		/// </summary>
		/// <returns></returns>
		public void MakeWord()
		{
			var templateError = false;
			// 根据模版和数据生成文档
			var wmlAssembledDoc = DocumentAssembler.AssembleDocument(_WordDoc, _ModelData, out templateError);
			if (templateError)
			{
				//throw new Exception("模版中有错误");
			}
			//生成新文档的存储目录
			if (!Directory.Exists(_OutputFolderPath))
			{
				Directory.CreateDirectory(_OutputFolderPath);
			}
			// 保存新文档
			var assembledDoc = new FileInfo(Path.Combine(_OutputFolderPath, Path.GetFileName(_OutputFilePath)));
			wmlAssembledDoc.SaveAs(assembledDoc.FullName);
		}

		/// <summary>
		/// Xml序列化
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="obj"></param>
		/// <returns></returns>
		public static string XmlSerialize<T>(T obj)
		{
			using (var ms = new System.IO.MemoryStream())
			using (var writer = new StreamWriter(ms, Encoding.UTF8))
			{
				var serializer = new System.Xml.Serialization.XmlSerializer(typeof(T));
				serializer.Serialize(writer, obj);
				writer.Close();
				var xml = Encoding.UTF8.GetString(ms.ToArray());
				var byteOrderMarkUtf8 = Encoding.UTF8.GetString(Encoding.UTF8.GetPreamble());
				if (xml.StartsWith(byteOrderMarkUtf8))
				{
					xml = xml.Remove(0, byteOrderMarkUtf8.Length);
				}
				return xml;
			}
		}
	}
}