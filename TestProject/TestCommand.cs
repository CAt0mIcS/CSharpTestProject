using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Text.Editor;

namespace TestProject
{
	/// <summary>
	/// Command handler
	/// </summary>
	internal sealed class TestCommand
	{
		/// <summary>
		/// Command ID.
		/// </summary>
		public const int CommandId = 0x0100;

		/// <summary>
		/// Command menu group (command set GUID).
		/// </summary>
		public static readonly Guid CommandSet = new Guid("644ecc36-53c1-4f33-b690-cae7b2e459fa");

		/// <summary>
		/// VS Package that provides this command, not null.
		/// </summary>
		private readonly AsyncPackage package;

		/// <summary>
		/// Initializes a new instance of the <see cref="TestCommand"/> class.
		/// Adds our command handlers for menu (commands must exist in the command table file)
		/// </summary>
		/// <param name="package">Owner package, not null.</param>
		/// <param name="commandService">Command service to add command to, not null.</param>
		private TestCommand(AsyncPackage package, OleMenuCommandService commandService)
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
		public static TestCommand Instance
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
			// Switch to the main thread - the call to AddCommand in TestCommand's constructor requires
			// the UI thread.
			await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

			OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
			Instance = new TestCommand(package, commandService);
		}

		private IWpfTextView GetTextView()
		{
			var textManager = (IVsTextManager)ServiceProvider.GetServiceAsync(typeof(SVsTextManager)).Result;
			var componentModel = (IComponentModel)this.ServiceProvider.GetServiceAsync(typeof(SComponentModel)).Result;
			var editor = componentModel.GetService<IVsEditorAdaptersFactoryService>();

			textManager.GetActiveView(1, null, out IVsTextView textViewCurrent);
			return editor.GetWpfTextView(textViewCurrent);
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

			EnvDTE.TextSelection ts = dte.ActiveWindow.Selection as EnvDTE.TextSelection;
			if (ts == null)
				return;
			EnvDTE.CodeFunction func = ts.ActivePoint.CodeElement[EnvDTE.vsCMElement.vsCMElementFunction]
						as EnvDTE.CodeFunction;
			if (func == null)
				return;

			int lineNr = ts.CurrentLine;
			var wpfTextView = GetTextView();

			if (dte.ActiveDocument != null)
			{
				var selection = (EnvDTE.TextSelection)dte.ActiveDocument.Selection;
				string text = selection.Text;

				string text2 = "";
				// Modify the text, for example:
				bool lineSet = false;
				foreach(char c in text)
				{
					if (!lineSet)
					{
						lineSet = true;
						text2 += "//";
					}

					if (c == '\n')
						lineSet = false;

					text2 += c;
				}

				// Replace the selection with the modified text.
				selection.Text = text2;
			}

		}
	}
}
