try
{
     //設定OLEDB，電腦需安裝2010可轉發套件64位元
     OleDbConnection myconnection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Form1.queueTableName + ";Extended Properties='Excel 12.0;HDR=YES';");
     //查詢Excel分頁
     OleDbDataAdapter oda = new OleDbDataAdapter("Select * from [sheet1$]", myconnection);
     DataSet ds = new DataSet();
     oda.Fill(ds);
     DataTable dt = ds.Tables[0];
	 
     dataGridView1.Rows.Clear();
     dataGridView1.Columns.Clear();
	 
     //設定欄位型態
     DataGridViewImageColumn imageColNum = new DataGridViewImageColumn();
     DataGridViewTextBoxColumn textColone = new DataGridViewTextBoxColumn();
     DataGridViewTextBoxColumn textColtwo = new DataGridViewTextBoxColumn();
     DataGridViewTextBoxColumn textColthree = new DataGridViewTextBoxColumn();
     DataGridViewTextBoxColumn textColfourth = new DataGridViewTextBoxColumn();
     DataGridViewTextBoxColumn textColfifth = new DataGridViewTextBoxColumn();
     DataGridViewTextBoxColumn textColsixth = new DataGridViewTextBoxColumn();
     DataGridViewTextBoxColumn textColseventh = new DataGridViewTextBoxColumn();
     //型態加入dataGridView
     dataGridView1.Columns.Add(imageColNum);
     dataGridView1.Columns.Add(textColone);
     dataGridView1.Columns.Add(textColtwo);
     dataGridView1.Columns.Add(textColthree);
     dataGridView1.Columns.Add(textColfourth);
     dataGridView1.Columns.Add(textColfifth);
     dataGridView1.Columns.Add(textColsixth);
     dataGridView1.Columns.Add(textColseventh);
     //設定欄位標題名稱給dataGridView
     imageColNum.HeaderText = "圖片";
     imageColNum.Name = "Img";
     textColone.HeaderText = "one";
     textColone.Name = "One";
     textColtwo.HeaderText = "two";
     textColtwo.Name = "Two";
     textColthree.HeaderText = "three";
     textColthree.Name = "Three";
     textColfourth.HeaderText = "four";
     textColfourth.Name = "Four";
     textColfifth.HeaderText = "fifth";
     textColfifth.Name = "Fifth";
     textColsixth.HeaderText = "sixth";
     textColsixth.Name = "Sixth";
     textColseventh.HeaderText = "seventh";
     textColseventh.Name = "Seventh";
     //宣告變數
     Image img;
     string countPath = "";
     string countOne = "";
     string countTwo = "";
     string countThree = "";
     string countFour = "";
     string countFive = "";
     string countSix = "";
     string countSeventh = "";
	 
     for (int i = 0; i < dt.Rows.Count; i++)
     {
        //圖片路徑
        countPath = Application.StartupPath + @"\img\" + dt.Rows[i]["name"] + ".png";
        countOne = "" + dt.Rows[i]["one"];   //["Excel欄位標題名稱"] 標題名稱當索引，搜尋此欄位資料，存到變數
        countTwo = "" + dt.Rows[i]["two"];
        countThree = "" + dt.Rows[i]["three"];
        countFour = "" + dt.Rows[i]["Four"];
        countFive = "" + dt.Rows[i]["Five"];
        countSix = "" + dt.Rows[i]["Six"];
        countSeventh = "" + dt.Rows[i]["Seven"];

        img = Image.FromFile(countPath);
        //資料塞回dataGridView1
        dataGridView1.Rows.Add(img,
                               countOne,
                               countTwo,
                               countThree,
                               countFour,
                               countFive,
                               countSix, 
			       countSeventh);
	 }
}catch(Exception ex)
{
	MessageBox.Show(ex.ToString());
}
