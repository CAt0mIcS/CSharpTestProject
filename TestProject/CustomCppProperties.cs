using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace TestProject
{

	static class VCXProjFileHandler
	{
		/// <summary>
		/// 
		/// </summary>
		/// <param name="prfilePath">the file path to the .vcxproj file</param>
		public static void ModifyVCXProj(string prfilePath)
		{

			string line = "";
			int propCount = 0;
			string text = "";
			System.IO.StreamReader file =
				new System.IO.StreamReader(prfilePath);
			while ((line = file.ReadLine()) != null)
			{
				if (line.Contains("<PropertyGroup Condition="))
				{
					++propCount;
					if (propCount > 4)
					{
						text += line + '\n';
						line = file.ReadLine();

						text += "    <OutDir>$(SolutionDir)\\bin\\$(Configuration)-$(Platform)\\</OutDir>\n    <IntDir>$(SolutionDir)\\bin-int\\$(Configuration)-$(Platform)\\$(ProjectName)\\</IntDir>\n";
					}
				}

				text += line + '\n';
			}
			file.Close();

			#region insertString
			text = text.Insert(text.IndexOf("</Project>") - 1, "\n<ItemDefinitionGroup Condition=\"'$(Configuration)|$(Platform)'=='Debug|Win32'\">\n"
  + "  <ClCompile>"
  + "\n    <WarningLevel>Level3</WarningLevel>"
  + "\n    <SDLCheck>true</SDLCheck>"
  + "\n    <PreprocessorDefinitions>_DEBUG;_CONSOLE;%(PreprocessorDefinitions);</PreprocessorDefinitions>"
  + "\n    <ConformanceMode>true</ConformanceMode>"
  + "\n    <LanguageStandard>stdcpp17</LanguageStandard>"
  + "\n    <PrecompiledHeader>Use</PrecompiledHeader>"
  + "\n    <PrecompiledHeaderFile>pch.h</PrecompiledHeaderFile>"
  + "\n    <AdditionalIncludeDirectories>src;</AdditionalIncludeDirectories>"
  + "\n  </ClCompile>"
  + "\n  <Link>"
  + "\n    <SubSystem>Windows</SubSystem>"
  + "\n    <GenerateDebugInformation>true</GenerateDebugInformation>"
  + "\n  </Link>"
  + "\n  <PostBuildEvent>"
  + "\n    <Command>"
  + "\n    </Command>"
  + "\n  </PostBuildEvent>"
  + "\n</ItemDefinitionGroup>"
  + "\n<ItemDefinitionGroup Condition=\"'$(Configuration)|$(Platform)'=='Debug|x64'\">"
  + "\n  <ClCompile>"
  + "\n    <WarningLevel>Level3</WarningLevel>"
  + "\n    <SDLCheck>true</SDLCheck>"
  + "\n    <PreprocessorDefinitions>_DEBUG;_CONSOLE;%(PreprocessorDefinitions);</PreprocessorDefinitions>"
  + "\n    <ConformanceMode>true</ConformanceMode>"
  + "\n    <PrecompiledHeader>Use</PrecompiledHeader>"
  + "\n    <PrecompiledHeaderFile>pch.h</PrecompiledHeaderFile>"
  + "\n    <LanguageStandard>stdcpp17</LanguageStandard>"
  + "\n    <AdditionalIncludeDirectories>src;</AdditionalIncludeDirectories>"
  + "\n  </ClCompile>"
  + "\n  <Link>"
  + "\n    <SubSystem>Windows</SubSystem>"
  + "\n    <GenerateDebugInformation>true</GenerateDebugInformation>"
  + "\n  </Link>"
  + "\n  <PostBuildEvent>"
  + "\n    <Command>"
  + "\n    </Command>"
  + "\n  </PostBuildEvent>"
  + "\n</ItemDefinitionGroup>"
  + "\n<ItemDefinitionGroup Condition=\"'$(Configuration)|$(Platform)'=='Release|Win32'\">"
  + "\n  <ClCompile>"
  + "\n    <WarningLevel>Level3</WarningLevel>"
  + "\n    <FunctionLevelLinking>true</FunctionLevelLinking>"
  + "\n    <IntrinsicFunctions>true</IntrinsicFunctions>"
  + "\n    <SDLCheck>true</SDLCheck>"
  + "\n    <PreprocessorDefinitions>NDEBUG;_CONSOLE;%(PreprocessorDefinitions);</PreprocessorDefinitions>"
  + "\n    <ConformanceMode>true</ConformanceMode>"
  + "\n    <LanguageStandard>stdcpp17</LanguageStandard>"
  + "\n    <PrecompiledHeader>Use</PrecompiledHeader>"
  + "\n    <PrecompiledHeaderFile>pch.h</PrecompiledHeaderFile>"
  + "\n    <AdditionalIncludeDirectories>src;</AdditionalIncludeDirectories>"
  + "\n  </ClCompile>"
  + "\n  <Link>"
  + "\n    <SubSystem>Windows</SubSystem>"
  + "\n    <EnableCOMDATFolding>true</EnableCOMDATFolding>"
  + "\n    <OptimizeReferences>true</OptimizeReferences>"
  + "\n    <GenerateDebugInformation>true</GenerateDebugInformation>"
  + "\n  </Link>"
  + "\n  <PostBuildEvent>"
  + "\n    <Command>"
  + "\n    </Command>"
  + "\n  </PostBuildEvent>"
  + "\n</ItemDefinitionGroup>"
  + "\n<ItemDefinitionGroup Condition=\"'$(Configuration)|$(Platform)'=='Release|x64'\">"
  + "\n  <ClCompile>"
  + "\n    <WarningLevel>Level3</WarningLevel>"
  + "\n    <FunctionLevelLinking>true</FunctionLevelLinking>"
  + "\n    <IntrinsicFunctions>true</IntrinsicFunctions>"
  + "\n    <SDLCheck>true</SDLCheck>"
  + "\n    <PreprocessorDefinitions>NDEBUG;_CONSOLE;%(PreprocessorDefinitions);</PreprocessorDefinitions>"
  + "\n    <ConformanceMode>true</ConformanceMode>"
  + "\n    <PrecompiledHeader>Use</PrecompiledHeader>"
  + "\n    <PrecompiledHeaderFile>pch.h</PrecompiledHeaderFile>"
  + "\n    <LanguageStandard>stdcpp17</LanguageStandard>"
  + "\n    <AdditionalIncludeDirectories>src;</AdditionalIncludeDirectories>"
  + "\n  </ClCompile>"
  + "\n  <Link>"
  + "\n    <SubSystem>Windows</SubSystem>"
  + "\n    <EnableCOMDATFolding>true</EnableCOMDATFolding>"
  + "\n    <OptimizeReferences>true</OptimizeReferences>"
  + "\n    <GenerateDebugInformation>true</GenerateDebugInformation>"
  + "\n  </Link>"
  + "\n  <PostBuildEvent>"
  + "\n    <Command>"
  + "\n    </Command>"
  + "\n  </PostBuildEvent>"
  + "\n</ItemDefinitionGroup>"
  + "\n  <ItemGroup>"
  + "\n  <ClCompile Include=\"src\\main.cpp\" />"
  + "\n  <ClCompile Include=\"src\\pch.cpp\">"
  + "\n    <PrecompiledHeader Condition=\"'$(Configuration)|$(Platform)'=='Debug|Win32'\">Create</PrecompiledHeader>"
  + "\n    <PrecompiledHeader Condition=\"'$(Configuration)|$(Platform)'=='Release|Win32'\">Create</PrecompiledHeader>"
  + "\n    <PrecompiledHeader Condition=\"'$(Configuration)|$(Platform)'=='Debug|x64'\">Create</PrecompiledHeader>"
  + "\n    <PrecompiledHeader Condition=\"'$(Configuration)|$(Platform)'=='Release|x64'\">Create</PrecompiledHeader>"
  + "\n  </ClCompile>"
  + "\n</ItemGroup>"
  + "\n<ItemGroup>"
  + "\n  <ClInclude Include=\"src\\pch.h\" />"
  + "\n</ItemGroup>\n");
				#endregion

			System.IO.StreamWriter writer = new System.IO.StreamWriter(prfilePath);
			writer.WriteLine(text);
			writer.Close();

		}

		public static void ModifyVCXProjWin32(string prfilePath)
		{
			string line = "";
			int propCount = 0;
			string text = "";
			System.IO.StreamReader file =
				new System.IO.StreamReader(prfilePath);
			while ((line = file.ReadLine()) != null)
			{
				if (line.Contains("<PropertyGroup Condition="))
				{
					++propCount;
					if (propCount > 4)
					{
						text += line + '\n';
						line = file.ReadLine();

						text += "    <OutDir>$(SolutionDir)\\bin\\$(Configuration)-$(Platform)\\</OutDir>\n    <IntDir>$(SolutionDir)\\bin-int\\$(Configuration)-$(Platform)\\$(ProjectName)\\</IntDir>\n";
					}
				}

				text += line + '\n';
			}
			file.Close();

			#region insertString
			text = text.Insert(text.IndexOf("</Project>") - 1, "\n<ItemDefinitionGroup Condition=\"'$(Configuration)|$(Platform)'=='Debug|Win32'\">\n"
  + "  <ClCompile>"
  + "\n    <WarningLevel>Level3</WarningLevel>"
  + "\n    <SDLCheck>true</SDLCheck>"
  + "\n    <PreprocessorDefinitions>_DEBUG;_CONSOLE;%(PreprocessorDefinitions);</PreprocessorDefinitions>"
  + "\n    <ConformanceMode>true</ConformanceMode>"
  + "\n    <LanguageStandard>stdcpp17</LanguageStandard>"
  + "\n    <PrecompiledHeader>Use</PrecompiledHeader>"
  + "\n    <PrecompiledHeaderFile>pch.h</PrecompiledHeaderFile>"
  + "\n    <AdditionalIncludeDirectories>src;</AdditionalIncludeDirectories>"
  + "\n  </ClCompile>"
  + "\n  <Link>"
  + "\n    <SubSystem>Windows</SubSystem>"
  + "\n    <GenerateDebugInformation>true</GenerateDebugInformation>"
  + "\n  </Link>"
  + "\n  <PostBuildEvent>"
  + "\n    <Command>"
  + "\n    </Command>"
  + "\n  </PostBuildEvent>"
  + "\n</ItemDefinitionGroup>"
  + "\n<ItemDefinitionGroup Condition=\"'$(Configuration)|$(Platform)'=='Debug|x64'\">"
  + "\n  <ClCompile>"
  + "\n    <WarningLevel>Level3</WarningLevel>"
  + "\n    <SDLCheck>true</SDLCheck>"
  + "\n    <PreprocessorDefinitions>_DEBUG;_CONSOLE;%(PreprocessorDefinitions);</PreprocessorDefinitions>"
  + "\n    <ConformanceMode>true</ConformanceMode>"
  + "\n    <PrecompiledHeader>Use</PrecompiledHeader>"
  + "\n    <PrecompiledHeaderFile>pch.h</PrecompiledHeaderFile>"
  + "\n    <LanguageStandard>stdcpp17</LanguageStandard>"
  + "\n    <AdditionalIncludeDirectories>src;</AdditionalIncludeDirectories>"
  + "\n  </ClCompile>"
  + "\n  <Link>"
  + "\n    <SubSystem>Windows</SubSystem>"
  + "\n    <GenerateDebugInformation>true</GenerateDebugInformation>"
  + "\n  </Link>"
  + "\n  <PostBuildEvent>"
  + "\n    <Command>"
  + "\n    </Command>"
  + "\n  </PostBuildEvent>"
  + "\n</ItemDefinitionGroup>"
  + "\n<ItemDefinitionGroup Condition=\"'$(Configuration)|$(Platform)'=='Release|Win32'\">"
  + "\n  <ClCompile>"
  + "\n    <WarningLevel>Level3</WarningLevel>"
  + "\n    <FunctionLevelLinking>true</FunctionLevelLinking>"
  + "\n    <IntrinsicFunctions>true</IntrinsicFunctions>"
  + "\n    <SDLCheck>true</SDLCheck>"
  + "\n    <PreprocessorDefinitions>NDEBUG;_CONSOLE;%(PreprocessorDefinitions);</PreprocessorDefinitions>"
  + "\n    <ConformanceMode>true</ConformanceMode>"
  + "\n    <LanguageStandard>stdcpp17</LanguageStandard>"
  + "\n    <PrecompiledHeader>Use</PrecompiledHeader>"
  + "\n    <PrecompiledHeaderFile>pch.h</PrecompiledHeaderFile>"
  + "\n    <AdditionalIncludeDirectories>src;</AdditionalIncludeDirectories>"
  + "\n  </ClCompile>"
  + "\n  <Link>"
  + "\n    <SubSystem>Windows</SubSystem>"
  + "\n    <EnableCOMDATFolding>true</EnableCOMDATFolding>"
  + "\n    <OptimizeReferences>true</OptimizeReferences>"
  + "\n    <GenerateDebugInformation>true</GenerateDebugInformation>"
  + "\n  </Link>"
  + "\n  <PostBuildEvent>"
  + "\n    <Command>"
  + "\n    </Command>"
  + "\n  </PostBuildEvent>"
  + "\n</ItemDefinitionGroup>"
  + "\n<ItemDefinitionGroup Condition=\"'$(Configuration)|$(Platform)'=='Release|x64'\">"
  + "\n  <ClCompile>"
  + "\n    <WarningLevel>Level3</WarningLevel>"
  + "\n    <FunctionLevelLinking>true</FunctionLevelLinking>"
  + "\n    <IntrinsicFunctions>true</IntrinsicFunctions>"
  + "\n    <SDLCheck>true</SDLCheck>"
  + "\n    <PreprocessorDefinitions>NDEBUG;_CONSOLE;%(PreprocessorDefinitions);</PreprocessorDefinitions>"
  + "\n    <ConformanceMode>true</ConformanceMode>"
  + "\n    <PrecompiledHeader>Use</PrecompiledHeader>"
  + "\n    <PrecompiledHeaderFile>pch.h</PrecompiledHeaderFile>"
  + "\n    <LanguageStandard>stdcpp17</LanguageStandard>"
  + "\n    <AdditionalIncludeDirectories>src;</AdditionalIncludeDirectories>"
  + "\n  </ClCompile>"
  + "\n  <Link>"
  + "\n    <SubSystem>Windows</SubSystem>"
  + "\n    <EnableCOMDATFolding>true</EnableCOMDATFolding>"
  + "\n    <OptimizeReferences>true</OptimizeReferences>"
  + "\n    <GenerateDebugInformation>true</GenerateDebugInformation>"
  + "\n  </Link>"
  + "\n  <PostBuildEvent>"
  + "\n    <Command>"
  + "\n    </Command>"
  + "\n  </PostBuildEvent>"
  + "\n</ItemDefinitionGroup>"
  + "\n  <ItemGroup>"
  + "\n  <ClCompile Include=\"src\\main.cpp\" />"
  + "\n  <ClCompile Include=\"src\\pch.cpp\">"
  + "\n    <PrecompiledHeader Condition=\"'$(Configuration)|$(Platform)'=='Debug|Win32'\">Create</PrecompiledHeader>"
  + "\n    <PrecompiledHeader Condition=\"'$(Configuration)|$(Platform)'=='Release|Win32'\">Create</PrecompiledHeader>"
  + "\n    <PrecompiledHeader Condition=\"'$(Configuration)|$(Platform)'=='Debug|x64'\">Create</PrecompiledHeader>"
  + "\n    <PrecompiledHeader Condition=\"'$(Configuration)|$(Platform)'=='Release|x64'\">Create</PrecompiledHeader>"
  + "\n  </ClCompile>"
  + "\n</ItemGroup>"
  + "\n<ItemGroup>"
  + "\n  <ClInclude Include=\"src\\pch.h\" />"
  + "\n</ItemGroup>\n");
			#endregion

			System.IO.StreamWriter writer = new System.IO.StreamWriter(prfilePath);
			writer.WriteLine(text);
			writer.Close();
		}
	}



	/// <summary>
	/// Command handler
	/// </summary>
	internal sealed class CustomCppProperties
	{
		/// <summary>
		/// Command ID.
		/// </summary>
		public const int CommandId = 0x0100;

		/// <summary>
		/// Command menu group (command set GUID).
		/// </summary>
		public static readonly Guid CommandSet = new Guid("8f968e43-f7b0-4efc-8d59-143fb14bdf9b");

		/// <summary>
		/// VS Package that provides this command, not null.
		/// </summary>
		private readonly AsyncPackage package;

		/// <summary>
		/// Initializes a new instance of the <see cref="CustomCppProperties"/> class.
		/// Adds our command handlers for menu (commands must exist in the command table file)
		/// </summary>
		/// <param name="package">Owner package, not null.</param>
		/// <param name="commandService">Command service to add command to, not null.</param>
		private CustomCppProperties(AsyncPackage package, OleMenuCommandService commandService)
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
		public static CustomCppProperties Instance
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
			// Switch to the main thread - the call to AddCommand in CustomCppProperties's constructor requires
			// the UI thread.
			await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

			OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
			Instance = new CustomCppProperties(package, commandService);
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
					else if(property.Name == "ProjectFile")
					{
						vcxfilepath = property.Value.ToString();
					}
					else if(property.Name == "ProjectDirectory")
					{
						projFilePath = property.Value.ToString();
					}
				}
				break;
			}

			//VsShellUtilities.ShowMessageBox(this.package, props, "", 0, 0, 0);

			if(!vcxfilepath.Equals(""))
				VCXProjFileHandler.ModifyVCXProj(vcxfilepath);
			
			Directory.CreateDirectory(projFilePath + "\\src");

			using (FileStream writer = new FileStream(projFilePath + "\\src\\pch.cpp", FileMode.Create))
			{
				byte[] arr = Encoding.ASCII.GetBytes("#include \"pch.h\"\n\n\n");
				writer.Write(arr, 0, 19);
			}

			using (FileStream writer = new FileStream(projFilePath + "\\src\\pch.h", FileMode.Create))
			{
				byte[] arr = Encoding.ASCII.GetBytes("#pragma once\n\n\n");
				writer.Write(arr, 0, 15);
			}

			using (FileStream writer = new FileStream(projFilePath + "\\src\\main.cpp", FileMode.Create))
			{
				byte[] arr = Encoding.ASCII.GetBytes("#include \"pch.h\"\n\n\n\nint main(int argc, char** argv)\n{\n\t\n}\n\n");
				writer.Write(arr, 0, 59);
			}

		}
	}
}
