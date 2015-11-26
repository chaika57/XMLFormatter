using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;

using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace XMLFormater
{
    public partial class Form1 : Form
    {
        DataSet ds = new DataSet();

        public Form1()
        {
            InitializeComponent();
        }

        private void CreateXML()
        {
            int index = 0; //DataTable row index
            int tableNumberRows = ds.Tables[0].Rows.Count; //counts number of rows in DataTable
                        
            CreditTransferTransactionInformation10[] crdtTransTransacInfoTransactions = new CreditTransferTransactionInformation10[tableNumberRows];            

            foreach (DataRow row in ds.Tables[0].Rows)
            {                
                CreditTransferTransactionInformation10 crdtTransTransacInfo = new CreditTransferTransactionInformation10();
                
                PaymentIdentification1 paymentIndNum = new PaymentIdentification1();
                paymentIndNum.EndToEndId = index.ToString();

                EquivalentAmount2 eqvAmount = new EquivalentAmount2();
                ActiveOrHistoricCurrencyAndAmount curAndAmnt = new ActiveOrHistoricCurrencyAndAmount();
                curAndAmnt.Ccy = row[10].ToString();//Currency
                curAndAmnt.Value = decimal.Parse(row[9].ToString());//Transaction ammount

                eqvAmount.Amt = curAndAmnt;

                AmountType3Choice amntType = new AmountType3Choice();
                amntType.Item = eqvAmount;

                crdtTransTransacInfo.PmtId = paymentIndNum;
                crdtTransTransacInfo.Amt = amntType;

                crdtTransTransacInfoTransactions[index] = crdtTransTransacInfo;              

                index++;
            }

            PaymentInstructionInformation3[] pmntInstrInf = new PaymentInstructionInformation3[1];
            PaymentInstructionInformation3 pmntInstrInfIteam = new PaymentInstructionInformation3();
            pmntInstrInfIteam.CdtTrfTxInf = crdtTransTransacInfoTransactions;
            pmntInstrInf[0] = pmntInstrInfIteam;

            Authorisation1Choice[] authChoiceArray = new Authorisation1Choice[1];
            Authorisation1Choice authChoice = new Authorisation1Choice();
            authChoice.Item = "Sergei";
            authChoiceArray[0] = authChoice;

            GroupHeader32 grpHdr = new GroupHeader32();
            grpHdr.Authstn = authChoiceArray;

            CustomerCreditTransferInitiationV03 cctiv = new CustomerCreditTransferInitiationV03();
            cctiv.GrpHdr = grpHdr;
            cctiv.PmtInf = pmntInstrInf;

            Document document = new Document();
         
            document.CstmrCdtTrfInitn = cctiv;          

            var data = document;
            var serializer = new XmlSerializer(typeof(Document));
            using (var stream = new StreamWriter("C:\\Users\\yulya\\Desktop\\test.xml"))               
               
           serializer.Serialize(stream, data);

        }


        private void ctreateXMLButton_Click(object sender, EventArgs e)
        {
            CreateXML();
        }

        private void UploadExcel(string fileName)
        {           
            try
            {
                OleDbConnection MyConnection;                
                OleDbDataAdapter MyCommand;

                try
                {
                    MyConnection = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties=Excel 8.0;");
                }
                catch
                {
                    MyConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=Excel 12.0;");
                }

                MyConnection.Open();
                DataTable Sheets = MyConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
               
                string worksheets = Sheets.Rows[0]["TABLE_NAME"].ToString();

                MyCommand = new OleDbDataAdapter("select * from [" + worksheets + "]", MyConnection);
                MyCommand.TableMappings.Add("Table", "TestTable");
                ds = new System.Data.DataSet();
                MyCommand.Fill(ds);                
                MyConnection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }                

        }

        private void uploadButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            //dialog.Filter = "Excel Files | *.xls"; 
            dialog.Multiselect = false;
            try
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    string path = dialog.FileName;
                    UploadExcel(path);

                    fileNameLabel.Text = Path.GetFileName(path);//display file name when file is uploaded
                    createXMLButton.Enabled = true;//activate upload button
                    
                }
            }
            catch { }
        }
    }
}
