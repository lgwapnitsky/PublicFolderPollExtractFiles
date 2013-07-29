using System;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using Ionic.Zip;
using Microsoft.Exchange.WebServices.Data;

namespace PublicFolderPollExtractFiles
{
	internal class Program
	{
		private static ExchangeService service;
		private static CLIArgs parsed;

		private static void Main ( string[] args )
		{
			try
			{
				parsed = PowerArgs.Args.Parse<CLIArgs> ( args );

				if (parsed.Help)
				{
					PowerArgs.ArgUsage.GetStyledUsage<CLIArgs> ().Write ();
				}
				else if (parsed.ExtractPath == null || parsed.PublicFolderPath == null)
				{
					throw new ArgumentException ( "You must specify command line arguments\n" );
				}
				else
				{
					if (!(Directory.Exists ( parsed.ExtractPath )))
					{
						throw new DirectoryNotFoundException ();
					}
					ReadPublicFolders ();
				}
			}

			catch (Exception ex)
			{
				Console.WriteLine ( ex.Message );

				if (ex is ArgumentException || ex is DirectoryNotFoundException)
				{
					Console.WriteLine ( PowerArgs.ArgUsage.GetUsage<CLIArgs> () );
				}
			}
			//Console.ReadLine ();
		}

		private static void ReadPublicFolders ()
		{
			service = new ExchangeService ();
			service.Credentials = CredentialCache.DefaultNetworkCredentials;
			service.AutodiscoverUrl ( UserPrincipal.Current.EmailAddress );

			FolderId fid = GetPublicFolderID ( parsed.PublicFolderPath, true );
			ListMessagesFromSubFolder ( Folder.Bind ( service, fid ) );
		}

		private static void ListMessagesFromSubFolder ( Folder folder )
		{
			ItemView iv = new ItemView ( int.MaxValue );
			iv.Traversal = ItemTraversal.Shallow;

			FindItemsResults<Item> findItemsResults = folder.FindItems ( iv );
			foreach (Item i in findItemsResults)
				switch (i.GetType ().Name)
				{
					case "EmailMessage":
						EmailMessage msg = i as EmailMessage;
						Console.WriteLine ( msg.Subject );
						switch (msg.IsRead)
						{
							case true:
								Console.WriteLine ( "Message has been read" );
								break;

							case false:
								Console.WriteLine ( "Message is unread" );

								EmailMessage msg_att = EmailMessage.Bind ( service, new ItemId ( i.Id.UniqueId.ToString () ),
											 new PropertySet ( BasePropertySet.IdOnly, ItemSchema.Attachments ) );

								if (msg_att.Attachments.Count > 0)
									foreach (FileAttachment fa in msg_att.Attachments)
									{
										string filepath = parsed.ExtractPath + "\\";
										string filename = filepath + fa.Name;

										using (FileStream fs = new FileStream ( filename, FileMode.OpenOrCreate, FileAccess.ReadWrite ))
										{
											fa.Load ( fs );

											if (Regex.IsMatch ( Path.GetExtension ( filename ), @"\.zip$", RegexOptions.IgnoreCase ))
											{
												using (ZipFile zipfile = Ionic.Zip.ZipFile.Read ( filename ))
												{
													foreach (ZipEntry zEntry in zipfile.Entries)
													{
														if (!(File.Exists ( filepath + zEntry.FileName )) || (zEntry.ModifiedTime > File.GetLastWriteTimeUtc ( filepath + zEntry.FileName )))
															zEntry.Extract ( filepath, ExtractExistingFileAction.OverwriteSilently );
													}
												}
												fs.Close ();
												File.Delete ( filename );
											}
											fs.Dispose ();
										}
									}
								msg.IsRead = true;
								msg.Update ( ConflictResolutionMode.AutoResolve );

								break;
						}
						break;
				}
		}

		private static FolderId GetPublicFolderID ( string path, bool isPublicFolderRoot )
		{
			string[] folderNames = path.Split ( new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries );

			Folder parent = Folder.Bind ( service,
										isPublicFolderRoot
											? WellKnownFolderName.PublicFoldersRoot
											: WellKnownFolderName.MsgFolderRoot );

			foreach (string folder in folderNames)
			{
				SearchFilter searchFilter = new SearchFilter.SearchFilterCollection ( LogicalOperator.And,
																					new SearchFilter.IsEqualTo (
																						FolderSchema.DisplayName, folder ) );
				FindFoldersResults results = parent.FindFolders ( searchFilter, new FolderView ( 1 ) );

				if (results != null && results.TotalCount == 1)
				{
					parent = results.Folders[0];
				}
				else
				{
					parent = null; // Not Found
					break;
				}
			}

			return parent != null ? parent.Id : null;
		}
	}
}