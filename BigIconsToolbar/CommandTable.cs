using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Threading.Tasks;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace BigIconsToolbar
{
internal sealed class CommandTable
{
private static readonly Guid				CommandSet = new Guid( "99ea3419-d3e9-44f5-915d-932c95dbb020" );
private readonly DTE2						_dte;
private readonly AsyncPackage				_package;

private readonly Dictionary<int, string>	_commands = new Dictionary<int, string>()
{
	{ 0x0101, "File.StartWindow" },
	{ 0x0102, "Project.AddNewItem" },
	{ 0x0103, "Project.NewFolder" },
	{ 0x0104, "File.SaveAll" },
	{ 0x0105, "Edit.Undo" },
	{ 0x0106, "Edit.Redo" },
	{ 0x0107, "View.PropertiesWindow" },
	{ 0x0108, "Tools.Options" },
	{ 0x0109, "Project.ManageNuGetPackages" },
	{ 0x0110, "Build.CleanSolution" },
	{ 0x0111, "Build.RebuildSolution" },
	{ 0x0112, "Build.Cancel" },
	{ 0x0113, "Build.BuildSolution" },
	{ 0x0114, "Debug.Start" },
	{ 0x0115, "Debug.StartWithoutDebugging" },
};


public CommandTable( AsyncPackage package, OleMenuCommandService commandService, DTE2 dte )
{
	_package		= package ?? throw new ArgumentNullException( nameof( package ) );
	commandService	= commandService ?? throw new ArgumentNullException( nameof( commandService ) );

	_dte = dte;

	foreach( var command in _commands )
	{
		var menuCommandID	= new CommandID( CommandSet, command.Key );
		var menuItem		= new MenuCommand( Execute, menuCommandID );

		commandService.AddCommand( menuItem );
	}
}




private void Execute( object sender, EventArgs e )
{
	ThreadHelper.ThrowIfNotOnUIThread();

	if( sender is MenuCommand menuCommand )
	{
		var cmdId = menuCommand.CommandID;

		if( _commands.ContainsKey( cmdId.ID ) == true )
		{
			try
			{
				_dte.ExecuteCommand( _commands[ cmdId.ID ] );
			}
			catch( Exception ex )
			{
#if DEBUG
				VsShellUtilities.ShowMessageBox( _package,
												 ex.Message,
												 "The command is not available in this context.",
												 OLEMSGICON.OLEMSGICON_INFO,
												 OLEMSGBUTTON.OLEMSGBUTTON_OK,
												 OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST );
#endif
			}
		}
	}
}




internal static async Task InitializeAsync( AsyncPackage package, DTE2 dte )
{
	await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync( package.DisposalToken );

	var commandService = await package.GetServiceAsync( typeof( IMenuCommandService ) ) as OleMenuCommandService;

	_ = new CommandTable( package, commandService, dte );
}
}
}
