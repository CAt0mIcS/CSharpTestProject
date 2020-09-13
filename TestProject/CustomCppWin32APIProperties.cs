using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace TestProject
{
	/// <summary>
	/// Command handler
	/// </summary>
	internal sealed class CustomCppWin32APIProperties
	{
		/// <summary>
		/// Command ID.
		/// </summary>
		public const int CommandId = 4129;

		/// <summary>
		/// Command menu group (command set GUID).
		/// </summary>
		public static readonly Guid CommandSet = new Guid("8f968e43-f7b0-4efc-8d59-143fb14bdf9b");

		/// <summary>
		/// VS Package that provides this command, not null.
		/// </summary>
		private readonly AsyncPackage package;

		/// <summary>
		/// Initializes a new instance of the <see cref="CustomCppWin32APIProperties"/> class.
		/// Adds our command handlers for menu (commands must exist in the command table file)
		/// </summary>
		/// <param name="package">Owner package, not null.</param>
		/// <param name="commandService">Command service to add command to, not null.</param>
		private CustomCppWin32APIProperties(AsyncPackage package, OleMenuCommandService commandService)
		{
			this.package = package ?? throw new ArgumentNullException(nameof(package));
			commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

			var menuCommandID = new CommandID(CommandSet, CommandId);
			var menuItem = new MenuCommand(this.Execute, menuCommandID);
			commandService.AddCommand(menuItem);
		}

		/// <summary>
		/// Gets the instance of the command.
		/// </summary>
		public static CustomCppWin32APIProperties Instance
		{
			get;
			private set;
		}

		/// <summary>
		/// Gets the service provider from the owner package.
		/// </summary>
		private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
		{
			get
			{
				return this.package;
			}
		}

		/// <summary>
		/// Initializes the singleton instance of the command.
		/// </summary>
		/// <param name="package">Owner package, not null.</param>
		public static async Task InitializeAsync(AsyncPackage package)
		{
			// Switch to the main thread - the call to AddCommand in CustomCppWin32APIProperties's constructor requires
			// the UI thread.
			await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

			OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
			Instance = new CustomCppWin32APIProperties(package, commandService);
		}

		/// <summary>
		/// This function is the callback used to execute the command when the menu item is clicked.
		/// See the constructor to see how the menu item is associated with this function using
		/// OleMenuCommandService service and MenuCommand class.
		/// </summary>
		/// <param name="sender">Event sender.</param>
		/// <param name="e">Event args.</param>
		private void Execute(object sender, EventArgs e)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			EnvDTE.DTE dte = (EnvDTE.DTE)this.ServiceProvider.GetServiceAsync(typeof(EnvDTE.DTE)).Result;
			EnvDTE.Projects projects = dte.Solution.Projects;

			string vcxfilepath = "";
			string projFilePath = "";
			string props = "";

			foreach (EnvDTE.Project project in projects)
			{
				foreach (EnvDTE.Property property in project.Properties)
				{
					try
					{
						props += property.Name + ":\t" + property.Value.ToString() + '\n';
					}
					catch (Exception)
					{
					}

					if (property.Name == "Kind" && property.Value.ToString() != "VCProject")
					{
						throw new Exception("Wrong project type, needs to be a C++ project");
					}

					if (property.Name == "ShowAllFiles")
					{
						property.Value = "True";
					}
					else if (property.Name == "ProjectFile")
					{
						vcxfilepath = property.Value.ToString();
					}
					else if (property.Name == "ProjectDirectory")
					{
						projFilePath = property.Value.ToString();
					}
				}
				break;
			}

			//VsShellUtilities.ShowMessageBox(this.package, props, "", 0, 0, 0);

			if (!vcxfilepath.Equals(""))
				VCXProjFileHandler.ModifyVCXProjWin32(vcxfilepath);

			Directory.CreateDirectory(projFilePath + "\\src");

			using (FileStream writer = new FileStream(projFilePath + "\\src\\pch.cpp", FileMode.Create))
			{
				byte[] arr = Encoding.ASCII.GetBytes("#include \"pch.h\"\n\n\n");
				writer.Write(arr, 0, 19);
			}

			using (FileStream writer = new FileStream(projFilePath + "\\src\\pch.h", FileMode.Create))
			{
				byte[] arr = Encoding.ASCII.GetBytes("#pragma once\n\n\n#include <Windows.h>\n\n");
				writer.Write(arr, 0, 37);
			}

			using (FileStream writer = new FileStream(projFilePath + "\\src\\main.cpp", FileMode.Create))
			{
				byte[] arr = Encoding.ASCII.GetBytes("#include \"pch.h\"\n\n\n\nint WINAPI wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ PWSTR pCmdLine, _In_ int nCmdShow)\n{\n\t\n}\n\n");
				writer.Write(arr, 0, 147);
			}
		}
	}
}
