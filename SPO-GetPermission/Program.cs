using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Security;
using SPO_GetPermission;
using Microsoft.Office.Interop.Excel;
namespace SPO_GetPermission
{
	/// <summary>  
	///  
	/// </summary>  
	public class Program
	{
		static void Main(string[] args)
		{
			
			#region Site Details - Read the details from config file  
			string siteURL = ConfigurationManager.AppSettings["siteURL"];
			string ListName = ConfigurationManager.AppSettings["listName"];
			string userName = ConfigurationManager.AppSettings["userName"];
			string password = ConfigurationManager.AppSettings["password"];
			string sitrUrl = siteURL;
			#endregion
			
			SecureString securePassword = new SecureString();
			foreach (char c in password.ToCharArray()) securePassword.AppendChar(c);
			/// Add the Folder object to List collection 
			#region Folder
			using (var ctx = new ClientContext(sitrUrl))
			{
				List<permissionItem> lstPermissionItem = new List<permissionItem>();
				ctx.Credentials = new SharePointOnlineCredentials(userName, securePassword);
				ctx.Load(ctx.Web, a => a.Lists);
				ctx.ExecuteQuery();

				List list = ctx.Web.Lists.GetByTitle(ListName);
				var Folderitems = list.GetItems(CamlQuery.CreateAllFoldersQuery());
				ctx.Load(Folderitems, icol => icol.Include(i => i.RoleAssignments.Include(ra => ra.Member), i => i.DisplayName),
					a => a.IncludeWithDefaultProperties(b => b.HasUniqueRoleAssignments),
					permsn => permsn.Include(a => a.RoleAssignments.Include(roleAsg => roleAsg.Member.LoginName,
					roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name,
					roleDef => roleDef.Description))));
				ctx.Load(Folderitems);
				ctx.ExecuteQuery();

				string sPermissionLevel = string.Empty;

				foreach (var item in Folderitems)
				{
					//Console.WriteLine("{0} folder permissions", item.DisplayName);
					//permissionItem oFolderPermissionItem = null;
					//if (item.HasUniqueRoleAssignments)
					//{
					foreach (var assignment in item.RoleAssignments)
					{
						var oFolderPermissionItem = new permissionItem();
						Console.WriteLine(assignment.Member.Title);
						oFolderPermissionItem.Title = item["FileRef"].ToString();
						oFolderPermissionItem.Type = "Folder";
						oFolderPermissionItem.UserGroup = assignment.Member.Title;
						List<string> roles = new List<string>();
						foreach (var role in assignment.RoleDefinitionBindings)
						{
							roles.Add(role.Description);
							switch (role.Name)
							{
								case "Full Control":
									oFolderPermissionItem.FullControl = assignment.Member.Title;
									break;
								case "Contribute":
									oFolderPermissionItem.Contribute = assignment.Member.Title;
									break;
								case "Read":
									oFolderPermissionItem.Read = assignment.Member.Title;
									break;
							}
							sPermissionLevel += role.Name + ",";
						}
						if (item.HasUniqueRoleAssignments)
						{
							oFolderPermissionItem.Inherited = "No";
						}
						else
						{
							oFolderPermissionItem.Inherited = "Yes";
						}
						oFolderPermissionItem.PermissionLevel = sPermissionLevel.TrimEnd(',');
						oFolderPermissionItem.PrincipalType = assignment.Member.PrincipalType.ToString();

						if (oFolderPermissionItem.FullControl == string.Empty && oFolderPermissionItem.Contribute == string.Empty && oFolderPermissionItem.Read == string.Empty)
						{
							
						}
						else
						{
							lstPermissionItem.Add(oFolderPermissionItem);
						}
						sPermissionLevel = string.Empty;
					}

				}
				#region Sample
				//ctx.Load(Folderitems, icol => icol.Include(i => i.RoleAssignments.Include(ra => ra.Member), i => i.DisplayName));
				//ctx.Load(Folderitems, a => a.IncludeWithDefaultProperties(b => b.HasUniqueRoleAssignments), a => a.di);
				//	ctx.Load(Folderitems, a => a.IncludeWithDefaultProperties(b => b.HasUniqueRoleAssignments),
				//permsn => permsn.Include(a => a.RoleAssignments.Include(roleAsg => roleAsg.Member.LoginName,
				//roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name,
				//roleDef => roleDef.Description))));

				//else
				/*{
					foreach (var assignment in item.RoleAssignments)
					{
						oFolderPermissionItem.Title = item.DisplayName;
						oFolderPermissionItem.Type = "Folder";
						oFolderPermissionItem.UserGroup = assignment.Member.Title;

						Console.WriteLine(assignment.Member.Title);
						List<string> roles = new List<string>();
						foreach (var role in assignment.RoleDefinitionBindings)
						{
							roles.Add(role.Description);
							sPermissionLevel += role.Description + ",";
						}
						oFolderPermissionItem.PermissionLevel = sPermissionLevel;
						{
							oFolderPermissionItem.Inherited = "No";
						}
						oFolderPermissionItem.PrincipalType = assignment.Member.PrincipalType.ToString();
						lstPermissionItem.Add(oFolderPermissionItem);
						sPermissionLevel = string.Empty;
					}
				}

				//lstPermissionItem.Add(oFolderPermissionItem);
			}
			*/
				#endregion

				/// <remarks>
				/// Add the item object to List collection 
				/// </remarks>
				#endregion
			#region File
				CamlQuery camlQuery = new CamlQuery();
				camlQuery.ViewXml =
					@"<View Scope='Recursive' />";

				var listItems = list.GetItems(camlQuery);
				//load all list items with default properties and HasUniqueRoleAssignments property
				ctx.Load(listItems, icol => icol.Include(i => i.RoleAssignments.Include(ra => ra.Member), i => i.File.Name),
					a => a.IncludeWithDefaultProperties(b => b.HasUniqueRoleAssignments),
					permsn => permsn.Include(a => a.RoleAssignments.Include(roleAsg => roleAsg.Member.LoginName,
					roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name,
					roleDef => roleDef.Description))));
				ctx.ExecuteQuery();

				foreach (var item in listItems)
				{
					//Console.WriteLine("List item: " + item["FileRef"].ToString());
					//load permissions if item has unique permission
					foreach (var roleAsg in item.RoleAssignments)
					{
						var oItemPermissionItem = new permissionItem();
						Console.WriteLine(roleAsg.Member.Title);
						oItemPermissionItem.Title = item["FileRef"].ToString(); 
						oItemPermissionItem.Type = "File";
						oItemPermissionItem.UserGroup = roleAsg.Member.Title;

						Console.WriteLine("User/Group: " + roleAsg.Member.LoginName);
						List<string> roles = new List<string>();
						foreach (var role in roleAsg.RoleDefinitionBindings)
						{
							roles.Add(role.Description);
							switch (role.Name)
							{
								case "Full Control":
									oItemPermissionItem.FullControl = roleAsg.Member.Title;
									break;
								case "Contribute":
									oItemPermissionItem.Contribute = roleAsg.Member.Title;
									break;
								case "Read":
									oItemPermissionItem.Read = roleAsg.Member.Title;
									break;
							}
							sPermissionLevel += role.Name + ",";
						}
						if (item.HasUniqueRoleAssignments)
						{
							oItemPermissionItem.Inherited = "No";
						}
						else
						{
							oItemPermissionItem.Inherited = "Yes";
						}
						oItemPermissionItem.PermissionLevel = sPermissionLevel.TrimEnd(',');
						oItemPermissionItem.PrincipalType = roleAsg.Member.PrincipalType.ToString();
						if (oItemPermissionItem.FullControl == string.Empty && oItemPermissionItem.Contribute == string.Empty && oItemPermissionItem.Read == string.Empty)
						{

						}
						else
						{
							lstPermissionItem.Add(oItemPermissionItem);
						}
						sPermissionLevel = string.Empty;
					}
				}
				#endregion

			ExportToExcel(lstPermissionItem,ListName);
			Console.WriteLine("Report generated");
			Console.ReadLine();
			}
		} //Main  

		public static void ExportToExcel(List<permissionItem> olstPermissionItem, string pListName)
		{
			// Load Excel application
			Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

			// Create empty workbook
			excel.Workbooks.Add();

			// Create Worksheet from active sheet
			Microsoft.Office.Interop.Excel._Worksheet workSheet = excel.ActiveSheet;

			// I created Application and Worksheet objects before try/catch,
			// so that i can close them in finnaly block.
			// It's IMPORTANT to release these COM objects!!
			try
			{
				// ------------------------------------------------
				// Creation of header cells
				// ------------------------------------------------
				workSheet.Cells[1, "A"] = "Title";
				workSheet.Cells[1, "B"] = "Type";
				//workSheet.Cells[1, "C"] = "User / Group";
				//workSheet.Cells[1, "D"] = "Permission Levels";
				//workSheet.Cells[1, "E"] = "Inherited";
				workSheet.Cells[1, "C"] = "Principal Type";
				workSheet.Cells[1, "D"] = "Full Control";
				workSheet.Cells[1, "E"] = "Contribute";
				workSheet.Cells[1, "F"] = "Read";

				// ------------------------------------------------
				// Populate sheet with some real data from "cars" list
				// ------------------------------------------------
				int row = 2; // start row (in row 1 are header cells)
				foreach (permissionItem pitem in olstPermissionItem)
				{
					workSheet.Cells[row, "A"] = pitem.Title;
					workSheet.Cells[row, "B"] = pitem.Type;
					//workSheet.Cells[row, "C"] = pitem.UserGroup;
					workSheet.Cells[row, "C"] = pitem.PrincipalType;
					workSheet.Cells[row, "D"] = pitem.FullControl;
					workSheet.Cells[row, "E"] = pitem.Contribute;
					workSheet.Cells[row, "F"] = pitem.Read;
					row++;
				}

				// Apply some predefined styles for data to look nicely :)
				//workSheet.Range["A1"].AutoFormat(Microsoft.Office.Interop.Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic1);

				// Define filename
				string fileName = string.Format(@"{0}\SPO_"+pListName+"_Permission.xlsx", Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

				// Save this data as a file
				workSheet.SaveAs(fileName);

				// Display SUCCESS message
				//MessageBox.Show(string.Format("The file '{0}' is saved successfully!", fileName));
			}
			catch (Exception exception)
			{
				//MessageBox.Show("Exception",
				//	"There was a PROBLEM saving Excel file!\n" + exception.Message,
				//	MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally
			{
				// Quit Excel application
				excel.Quit();

				// Release COM objects (very important!)
				if (excel != null)
					System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

				if (workSheet != null)
					System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheet);

				// Empty variables
				excel = null;
				workSheet = null;

				// Force garbage collector cleaning
				GC.Collect();
			}
		}

		

	} //cs 
} //ns  