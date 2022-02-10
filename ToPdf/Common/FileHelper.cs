using System;
using System.IO;
using System.Threading.Tasks;

namespace ToPdf.Common
{
	/// <summary>
	/// 文件操作
	/// </summary>
	public class FileHelper
	{
		/// <summary>
		/// 文件是否存在
		/// </summary>
		/// <param name="filePath"></param>
		public static bool Exists(string filePath)
		{
			return File.Exists(filePath);
		}

		/// <summary>
		/// 写入文件
		/// </summary>
		/// <param name="filePath"></param>
		/// <param name="value"></param>
		public static async Task WriteFile(string filePath, string value)
		{
			try
			{
				//var fi = new FileInfo(filePath);
				//var path = fi.DirectoryName;
				var nIndex = filePath.LastIndexOf("\\");
				var path = filePath.Substring(0, nIndex);
				if (!Directory.Exists(path))
					Directory.CreateDirectory(path);

				await File.WriteAllTextAsync(filePath, value);
			}
			catch (Exception ex)
			{
				throw new Exception($"写入文件失败，{ex.Message}");
			}
		}

		/// <summary>
		/// 递归拷贝文件
		/// </summary>
		/// <param name="srcPath"></param>
		/// <param name="destPath"></param>
		public static void CopyDir(string srcPath, string destPath)
		{
			try
			{
				if (!Directory.Exists(srcPath))
				{
					return;
				}
				//如果不存在目标路径，则创建之
				if (!Directory.Exists(destPath))
				{
					Directory.CreateDirectory(destPath);
				}
				DirectoryInfo dir = new DirectoryInfo(srcPath);
				FileSystemInfo[] fileinfo = dir.GetFileSystemInfos();  //获取目录下（不包含子目录）的文件和子目录
				foreach (FileSystemInfo i in fileinfo)
				{
					if (i is DirectoryInfo)     //判断是否文件夹
					{
						if (!Directory.Exists(destPath + "\\" + i.Name))
						{
							Directory.CreateDirectory(destPath + "\\" + i.Name);   //目标目录下不存在此文件夹即创建子文件夹
						}
						CopyDir(i.FullName, destPath + "\\" + i.Name);    //递归调用复制子文件夹
					}
					else
					{
						File.Copy(i.FullName, destPath + "\\" + i.Name, true);      //不是文件夹即复制文件，true表示可以覆盖同名文件
					}
				}
			}
			catch (Exception e)
			{
			}
		}

		/// <summary>
		/// 递归拷贝文件
		/// </summary>
		/// <param name="fromDirectory">源路径</param>
		/// <param name="toDirectory">目标路径</param>
		public static void CopyFiles(string fromDirectory, string toDirectory)
		{
			if (!Directory.Exists(fromDirectory))
			{
				return;
			}
			//如果不存在目标路径，则创建之
			if (!Directory.Exists(toDirectory))
			{
				Directory.CreateDirectory(toDirectory);
			}
			string[] directories = Directory.GetDirectories(fromDirectory);

			if (directories.Length > 0)
			{
				foreach (string d in directories)
				{
					CopyFiles(d, toDirectory + d.Substring(d.LastIndexOf(@"\")));
				}
			}

			if (!Directory.Exists(toDirectory))
			{
				Directory.CreateDirectory(toDirectory);
			}

			string[] files = Directory.GetFiles(fromDirectory);

			if (files.Length > 0)
			{
				foreach (string s in files)
				{
					File.Copy(s, toDirectory + s.Substring(s.LastIndexOf(@"\")), true);
				}
			}
		}

		/// <summary>
		/// 从一个目录将其内容移动到另一目录
		/// </summary>
		/// <param name="directorySource">源目录</param>
		/// <param name="directoryTarget">目标目录</param>
		public static void MoveFolderTo(string directorySource, string directoryTarget)
		{
			//检查是否存在目的目录
			if (!Directory.Exists(directoryTarget))
			{
				Directory.CreateDirectory(directoryTarget);
			}
			//先来移动文件
			DirectoryInfo directoryInfo = new DirectoryInfo(directorySource);
			FileInfo[] files = directoryInfo.GetFiles();
			//移动所有文件
			foreach (FileInfo file in files)
			{
				//如果自身文件在运行，不能直接覆盖，需要重命名之后再移动
				if (File.Exists(Path.Combine(directoryTarget, file.Name)))
				{
					if (File.Exists(Path.Combine(directoryTarget, file.Name + ".bak")))
					{
						File.Delete(Path.Combine(directoryTarget, file.Name + ".bak"));
					}
					File.Move(Path.Combine(directoryTarget, file.Name), Path.Combine(directoryTarget, file.Name + ".bak"));
				}
				file.MoveTo(Path.Combine(directoryTarget, file.Name));
			}
			//最后移动目录
			DirectoryInfo[] directoryInfoArray = directoryInfo.GetDirectories();
			foreach (DirectoryInfo dir in directoryInfoArray)
			{
				MoveFolderTo(Path.Combine(directorySource, dir.Name), Path.Combine(directoryTarget, dir.Name));
			}
		}

		/// <summary>
		/// 判断文件是否存在
		/// </summary>
		/// <param name="path"></param>
		public static bool CheckExists(string path)
		{
			var isExists = false;
			try
			{
				var folderPath = Path.GetDirectoryName(path);
				if (!Directory.Exists(folderPath))
				{
					Directory.CreateDirectory(folderPath);
					return isExists;
				}
				if (File.Exists(path)) isExists = true;
			}
			catch (Exception)
			{
			}
			return isExists;
		}

		public static Stream FileToStream(string fileName)
		{
			// 打开文件
			FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
			// 读取文件的 byte[]
			byte[] bytes = new byte[fileStream.Length];
			fileStream.Read(bytes, 0, bytes.Length);
			fileStream.Close();
			// 把 byte[] 转换成 Stream
			Stream stream = new MemoryStream(bytes);
			return stream;
		}
	}
}