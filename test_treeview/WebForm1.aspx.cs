using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Web;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using System.Data.SqlClient;

namespace test_treeview
{
	/// <summary>
	/// Summary description for WebForm1.
	/// </summary>
	public class WebForm1 : System.Web.UI.Page
	{
		protected System.Web.UI.WebControls.Literal TreeView;
	
		#region Web Form Designer generated code
		override protected void OnInit(EventArgs e)
		{
			//
			// CODEGEN: This call is required by the ASP.NET Web Form Designer.
			//
			InitializeComponent();
			base.OnInit(e);
		}
		
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{    
			this.Load += new System.EventHandler(this.Page_Load);

		}
		#endregion
		
		DataSet ds ;
		public obout_ASPTreeView_2_NET.Tree oTree ;

		public DataSet SelectSql(string query) 
		{
			string connection = "server=(local);database=PBBS;Trusted_Connection=yes;";
			//string connection = "Initial Catalog=PBBS;Data Source=localhost;Trusted_Connection=true;";

			SqlConnection conn = new SqlConnection(connection);
			conn.Open();

			SqlDataAdapter adapter = new SqlDataAdapter();
			
			adapter.SelectCommand = new SqlCommand(query, conn);
			
			DataSet dataset = new DataSet();
			adapter.Fill(dataset);
			
			conn.Close();
			return dataset;
		}

		private void Page_Load(object sender, System.EventArgs e)
		{
			getdata();
			write_tree();
		}
		protected void getdata()
		{
			string sql = "Select * from VW_Tree_Account_Code";
			ds = SelectSql(sql);
			
			oTree = new obout_ASPTreeView_2_NET.Tree();
			foreach(DataRow dr in ds.Tables[0].Rows)
			{

				if ( !dr["ParentID"].Equals(DBNull.Value) )
				{	
					oTree.Add(dr["ParentID"], dr["Account_Code"], dr["Account_Name"] , dr["Expanded"], dr["Icon"], null);
				}
				else
				{	
					oTree.AddRootNode("", true, null);
					oTree.Add("root", dr["Account_Code"], dr["Account_Name"] , dr["Expanded"], dr["Icon"], null);
				}
			}//end Count rows
		}
		protected void write_tree()
		{
			
			// change this to your local TreeIcons folder
			oTree.FolderIcons = "/TreeIcons/Icons";
			oTree.FolderStyle = "/TreeIcons/Styles/Classic";
			oTree.FolderScript = "/TreeIcons/Tree_2028/Script";
			
			oTree.ShowIcons = false;

			oTree.Width = "100%";
			oTree.ShowIcons = false;
			oTree.EditNodeEnable = false;
			oTree.DragAndDropEnable = true;
			oTree.SelectedEnable = true;
			oTree.KeyNavigationEnable = true;
			oTree.EventList = "OnAddNode,OnNodeEdit,OnNodeDrop,OnRemoveNode";
			TreeView.Text = oTree.HTML();
			
		}
	}

}
