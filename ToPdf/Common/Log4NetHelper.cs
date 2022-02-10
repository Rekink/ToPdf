using log4net;
using log4net.Repository;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace ToPdf.Common
{
	/// <summary>
	/// Log4Net帮助类
	/// </summary>
	public class Log4NetHelper
	{
		private static ILog _log;
		private static ILoggerRepository _loggerRepository;

		public static void LogInit(Stream stream)
		{
			if (_log == null)
			{
				_loggerRepository = LogManager.CreateRepository("LoggerRepository");
				log4net.Config.XmlConfigurator.Configure(_loggerRepository, stream);
				_log = LogManager.GetLogger("LoggerRepository", "loginInfo");
			}
		}

		public static ILog LogInit<T>() where T : ControllerBase
		{
			ILog log = LogManager.GetLogger(Startup.repository.Name, typeof(T));
			return log;
		}
	}
}