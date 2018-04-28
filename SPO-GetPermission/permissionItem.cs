using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPO_GetPermission
{
	public class permissionItem
	{
		public permissionItem()
		{
		}
		public string Title { get; set; }
			public string Type { get; set; }
			public string UserGroup { get; set; }
			public string PermissionLevel { get; set; }
			public string Inherited { get; set; }

			//Security Group / User
			public string PrincipalType { get; set; }

			public string FullControl { get; set; } = string.Empty;
			public string Contribute { get; set; } = string.Empty;
			public string Read { get; set; } = string.Empty;




	}
}
