using PowerArgs;

namespace PublicFolderPollExtractFiles
{
	[ArgExample ( "PFP.exe -p \"Folder/subfolder/SubSubfolder\" -x \"c:\\temp\"", "Public folder and the destination for the extracted files." )]
	public class CLIArgs
	{
		[ArgDescription ( "Path to public folder that will be polled" )]
		[ArgShortcut ( "p" )]
		public string PublicFolderPath { get; set; }

		[ArgDescription ( "Path to where files should be extracted" )]
		[ArgShortcut ( "x" )]
		public string ExtractPath { get; set; }

		[ArgDescription ( "Shows the help documentation" )]
		//[ArgPosition(0)]
		public bool Help { get; set; }
	}
}