using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Test
{
	public partial class Form1 : Form
	{
		/// <summary>
		/// 默认存储位置
		/// </summary>
		private string ROOT_PATH;

		/// <summary>
		/// 是否调用本地服务，默认远程服务
		/// </summary>
		private bool isLocal;

		/// <summary>
		/// 转换工具，默认Word
		/// </summary>
		private enum ToolType : byte
		{
			/// <summary>
			/// Office版
			/// </summary>
			Office = 0,

			/// <summary>
			/// Wps版
			/// </summary>
			Wps = 1,

			/// <summary>
			/// FreeSpire收费版
			/// </summary>
			FreeSpire = 2,

			/// <summary>
			/// Aspose.Word版
			/// </summary>
			AsposeWord = 3
		}

		/// <summary>
		/// 转换工具，默认Word
		/// </summary>
		private ToolType convertTool;

		/// <summary>
		/// 是否转换成功
		/// </summary>
		private bool isExist;

		/// <summary>
		/// pdf文件保存路径
		/// </summary>
		private string wordPath = "";

		/// <summary>
		/// pdf文件保存路径
		/// </summary>
		private string pdfPath = "";

		/// <summary>
		///
		/// </summary>
		public Form1()
		{
			InitializeComponent();
			ROOT_PATH = System.IO.Directory.GetParent(System.Environment.CurrentDirectory).Parent.FullName;
			ROOT_PATH = Path.Combine(System.IO.Directory.GetParent(ROOT_PATH).FullName, "Files");
			isExist = false;
			// 默认本地调试
			isLocal = true;
			comboBox1.SelectedIndex = 0;
			// 默认Office工具装换
			comboBox2.SelectedIndex = 0;
			convertTool = (ToolType)comboBox2.SelectedIndex;
		}

		/// <summary>
		/// 打开文件
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void OpenFile_Click(object sender, EventArgs e)
		{
			OpenFileDialog openFileDialog = new OpenFileDialog();
			openFileDialog.FilterIndex = 1;
			openFileDialog.RestoreDirectory = true;
			if (openFileDialog.ShowDialog() == DialogResult.OK)
			{
				var byteArr = File.ReadAllBytes(openFileDialog.FileName);
				textBox1.Text = openFileDialog.FileName;
				//ROOT_PATH = Path.Combine(Path.GetDirectoryName(openFileDialog.FileName), "Files");
				ROOT_PATH = Path.GetDirectoryName(openFileDialog.FileName);
				if (!Directory.Exists(ROOT_PATH))
				{
					Directory.CreateDirectory(ROOT_PATH);
				}
				var fileName = Path.GetFileNameWithoutExtension(openFileDialog.FileName);
				textBox2.Text = Path.Combine(ROOT_PATH, string.Format("{0}.docx", fileName));
				wordPath = Path.Combine(ROOT_PATH, string.Format("{0}.docx", fileName));
				textBox3.Text = Path.Combine(ROOT_PATH, string.Format("{0}.pdf", fileName));
				pdfPath = Path.Combine(ROOT_PATH, string.Format("{0}.pdf", fileName));
				// 创建一个目标Word文件流
				FileStream fs = new FileStream(textBox2.Text, FileMode.Create);
				fs.Write(byteArr, 0, byteArr.Length);
				fs.Close();
			}
		}

		/// <summary>
		/// 更改服务类型
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
		{
			if (comboBox1.SelectedIndex == 0)
			{
				isLocal = true;
			}
			else
			{
				isLocal = false;
			}
		}

		/// <summary>
		/// 更改工具
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
		{
			convertTool = (ToolType)comboBox2.SelectedIndex;
		}

		/// <summary>
		///  Word文件转Pdf文件
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void ToPdf_Click(object sender, EventArgs e)
		{
			if (string.IsNullOrEmpty(textBox2.Text))
			{
				MessageBox.Show("请先选择需要转换的Word文件！");
				return;
			}

			if (isLocal)
			{
				// 本地
				try
				{
					var message = "Office";
					if (convertTool == ToolType.Office)
					{
						isExist = PdfHelper.Word2PDF(wordPath, pdfPath);
						//var count = WordToPDFHelper.WordsToPDFs(new string[1] { wordPath }, pdfPath);
						//isExist = false;
						//if (count > 0)
						//{
						//	isExist = true;
						//}
					}
					else if (convertTool == ToolType.Wps)
					{
						isExist = PdfHelper.Word2PdfByWPS(wordPath, pdfPath);
						message = "Wps";
					}
					else if (convertTool == ToolType.FreeSpire)
					{
						isExist = PdfHelper.Word2PdfByFreeSpire(wordPath, pdfPath);
						message = "FreeSpire";
					}
					else if (convertTool == ToolType.AsposeWord)
					{
						PdfHelper.Word2Png(wordPath, ROOT_PATH, out int pngCount);
						isExist = PdfHelper.Word2PdfByAspose(wordPath, pdfPath);
						message = "AsposeWord";
					}
					if (!isExist)
					{
						MessageBox.Show("本地" + message + "转换失败！");
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
			else
			{
				// 远程
				try
				{
					var bytes = File.ReadAllBytes(wordPath);
					// 同步事件中调用异步方法
					ToAsync(bytes, (byte)convertTool);
					//await GetPdfAsync(bytes);
					isExist = true;
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		/// <summary>
		/// 同步事件中调用异步方法
		/// </summary>
		/// <param name="bytes"></param>
		/// <param name="toolType"></param>
		private async void ToAsync(byte[] bytes, byte toolType)
		{
			if (!await GetPdfAsync(bytes, toolType))
			{
				MessageBox.Show("Word文件转Pdf文件成功");
			};
		}

		/// <summary>
		/// 远程服务器转换
		/// </summary>
		/// <param name="bytes"></param>
		/// <returns></returns>
		private async Task<bool> GetPdfAsync(byte[] bytes, byte toolType)
		{
			try
			{
				var res = false;
				var baseUrl = "http://192.168.20.98:8056";
				baseUrl = "http://192.168.20.14:10000";
				var url = baseUrl + "/api/WordPreview/GetPdf";
				if (toolType == 1)
				{
					url = baseUrl + "/api/WordPreview/GetPdfByWps";
				}
				url = baseUrl + "/api/Word2Pdf/GetPdf";
				System.Diagnostics.Stopwatch watch = new System.Diagnostics.Stopwatch();
				watch.Start();
				bytes = File.ReadAllBytes(wordPath);
				Stream stream = new MemoryStream(bytes);
				var content = new StreamContent(stream);
				using (var client = new HttpClient())
				{
					var result = await client.PostAsync(url, content);
					if (result.StatusCode != System.Net.HttpStatusCode.OK)
					{
						result.EnsureSuccessStatusCode();
						stream.Dispose();
						MessageBox.Show($"Word文件转Pdf文件失败,{result.Content.ReadAsStringAsync().Result}");
					}
					bytes = result.Content.ReadAsByteArrayAsync().Result;
				}
				watch.Stop();
				MessageBox.Show("Word文件成功转换为Pdf文件，耗时：" + watch.Elapsed);
				if (bytes != null && bytes.Length > 0)
				{
					File.WriteAllBytes(pdfPath, bytes);
					res = true;
				}
				return res;
			}
			catch (Exception ex)
			{
				throw;
			}
		}
	}
}