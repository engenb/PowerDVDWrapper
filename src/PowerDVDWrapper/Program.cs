using System;
using System.Diagnostics;
using System.Management;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NLog;
using NLog.Config;
using NLog.Targets;

namespace PowerDVDWrapper
{
	public class Program
	{
		private static Logger _log;

		static Program()
		{
			ConfigureNLog();
			_log = LogManager.GetCurrentClassLogger();
		}

		static void Main(string[] args)
		{
			var exe = args[0];
			var arguments = string.Join(" ", args, 1, args.Length - 1);

			_log.Debug($"Starting application {exe} with args {arguments}");
			using (var process = Process.Start(new ProcessStartInfo
			{
				FileName = exe,
				Arguments = arguments
			}))
			{
				_log.Debug($"Process started: {process.ProcessName} [{process.Id}]");
				process.EnableRaisingEvents = true;

				var hookId = SetWindowsHookEx(
					13,
					(nCode, wParam, lParam) =>
					{
						if (nCode >= 0 && wParam == (IntPtr)0x0100)
						{
							var vkCode = Marshal.ReadInt32(lParam);
							if (Keys.Back == (Keys)vkCode)
							{
								_log.Debug("Backspace pressed.  Exiting.");
								KillProcessAndChildren(process);
							}
						}
						return CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);
					},
					GetModuleHandle(process.MainModule.ModuleName), 0);

				process.Exited += (sender, eventArgs) =>
				{
					_log.Debug($"Process exited: {process.ProcessName} [{process.Id}]");
					UnhookWindowsHookEx(hookId);
					Application.Exit();
				};

				Application.Run();
			}
		}

		private static void ConfigureNLog()
		{
			var config = new LoggingConfiguration();

			var consoleTarget = new ColoredConsoleTarget
			{
				Layout = @"${date:format=HH\:mm\:ss} ${message}"
			};
			config.AddTarget("console", consoleTarget);

			var fileTarget = new FileTarget
			{
				FileName = "${specialfolder:folder=ApplicationData}/PowerDVDWrapper/Logs/log.txt",
				Layout = @"${date:format=HH\:mm\:ss} ${message}",
				ArchiveFileName = "${specialfolder:folder=ApplicationData}/PowerDVDWrapper/Logs/Archives/log.{#}.txt",
				ArchiveNumbering = ArchiveNumberingMode.Date,
				ArchiveEvery = FileArchivePeriod.Day,
				ArchiveDateFormat = "yyyy-MM-dd",
				MaxArchiveFiles = 14,
				KeepFileOpen = false
			};
			config.AddTarget("file", fileTarget);

			config.LoggingRules.Add(new LoggingRule("*", LogLevel.Debug, consoleTarget));
			config.LoggingRules.Add(new LoggingRule("*", LogLevel.Debug, fileTarget));

			LogManager.Configuration = config;
		}

		private static void KillProcessAndChildren(Process proc)
		{
			_log.Debug($"Killing process {proc.ProcessName} [{proc.Id}]");

			var mgmtObjCollection = new ManagementObjectSearcher($"Select * From Win32_Process Where ParentProcessID={proc.Id}").Get();
			foreach (var mgmtObj in mgmtObjCollection)
			{
				try
				{
					var childProc = Process.GetProcessById(Convert.ToInt32(mgmtObj["ProcessId"]));
					_log.Debug($"Found child process {childProc.ProcessName} [{childProc.Id}]");
					KillProcessAndChildren(childProc);
				}
				catch
				{
					//eat it
				}
			}
			if (!proc.HasExited)
			{
				proc.Kill();
			}
		}

		private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
		
		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		[return: MarshalAs(UnmanagedType.Bool)]
		private static extern bool UnhookWindowsHookEx(IntPtr hhk);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

		[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern IntPtr GetModuleHandle(string lpModuleName);
	}
}
