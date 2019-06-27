using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using System.Xml;
using System.Runtime.InteropServices;
using System.IO;
using System.Threading;
using System.Text.RegularExpressions;
using System.Collections; 




using System.Data.OleDb;


using System.Globalization;


namespace PricingEngine
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            intiliazePricer();
        }

        #region constants
        //cumnorm
        const double GAMMA = 0.2316419;
        const double A1 = 0.31938153;
        const double A2 = -0.356563782;
        const double A3 = 1.781477937;
        const double A4 = -1.821255978;
        const double A5 = 1.330274429;
        const double PI = 3.14159265358979;
        private DataTable pricer;
        private DataTable ccyDets;
        private DataTable crosses;




        private DataTable defaultBbgSource; //holds defaul bbg code for swap source
        private List<string> rowNames;
        private string systemFiles;
        private string xmlFilePath;
        private string xmlFileName;
        private string volPath;
        private string user;
        int endIni = 0;
        int scenRun = 0;



        private DataSet holidaySet; //contains holidays for each currency from bbg
        private DataSet fwdsSet;    //contains fwds from bbg
        private DataSet deposSet;   //contains ois rates from bbg
        private DataSet marketData; //contains all rateTiles
        private DataSet pricingData; //contains all updated ratetiles 2016
        private DataSet spreadSet; //contains spread info from excel sheet
        private DataSet smileMult; //contains rr and fly multipliers by tenor for each ccyPair
        private DataSet brokerRun; //contains bid/offer runs
        private DataSet skewSheet; //info for skewSheet program
        private DataSet allSurfs; //hold surfaces for wieghted vol shift on skewsheet
        private DataSet skewBps; //hold surfaces for wieghted vol shift on skewsheet
        private DataSet tradeSim; //hold surfaces for deal simulation
        private DataSet regaSega; //hold surfaces for deal simulation

        DateTime today = DateTime.Today;
        string homeCcy = "USD";


        //bbg api class
        Bloomberglp.Blpapi.Examples.Form1 fwds = new Bloomberglp.Blpapi.Examples.Form1();
        Bloomberglp.Blpapi.Examples.Form1 depos = new Bloomberglp.Blpapi.Examples.Form1();
        Bloomberglp.Blpapi.Examples.Form1 singleSpot = new Bloomberglp.Blpapi.Examples.Form1();
        Bloomberglp.Blpapi.Examples.Form1 strikeVol = new Bloomberglp.Blpapi.Examples.Form1();
        Bloomberglp.Blpapi.Examples.Form1 holidays = new Bloomberglp.Blpapi.Examples.Form1();
        Bloomberglp.Blpapi.Examples.Form1 mxSurface = new Bloomberglp.Blpapi.Examples.Form1();



        #region rowInt



        int CcyPairR;
        int SpotR;
        int ExpiryR;
        int StrikeR;
        int Put_CallR;
        int NotionalR;
        int Expiry_DaysR;
        int VolR;
        int SystemVolR;
        int BpsFromMidR;
        int PremiumFromMidR;
        int BloombergVolR;
        int Bid_OfferR;
        int BreakEvenR;
        int Vol_Spread_to_AtmR;
        int Bps_to_AtmR;
        int Spot_DeltaR;
        int AutoStrikeR;
        int ATM_VOLR;
        int RRR;
        int FLYR;
        int Swap_PtsR;
        int FwdR;
        int BbgSourceR;
        int ExpiryDateR;
        int DeliveryDateR;
        int PriceR;
        int Fwd_PriceR;
        int Fwd_DeltaR;
        int VegaR;
        int sVegaR;
        int sDeltaR;
        int VannaR;
        int VolgaR;
        int GammaR;
        int ThetaR;

        int Depo_BaseR;
        int Depo_TermsR;
        int Basis_BaseR;
        int Basis_TermsR;
        int TodayR;
        int SpotdateR;
        int DaycountR;
        int Deliver_DaysR;
        int DeltaAR;
        int GammaAR;
        int sGammaR;
        int VegaAR;
        int ThetaAR;
        int DV01_BaseR;
        int DV01_TermsR;
        int PremiumAR;
        int Premium_TermsR;
        int Premium_IncludedR;
        int Premium_TypeR;
        int Rega25R;
        int Rega10R;
        int Sega25R;
        int Sega10R;
        
        int VannaAR;
        int VolgaAR;
        int DV01_BaseAR;
        int DV01_TermsAR;


        #endregion

        #endregion

        #region pricer
        //There are 3 components to this program. 

        //1) MarketData - This combines fwds and depos from datasets(fwdSet and depoSet). Each one of these gets data from seperate call functions from bbg. The data is put together in datatables and will solve for either the base depo or terms depo depending on what is saved in the currency setup. 

        //2) PricingData - atmVols and smile come from text files saved from the dvi. The atm is lookuped by daycount as the vol file interpolates an individual vol for each day. The rr and flies are taken from a seperate file then linearly interpolated for the target option day. From here all necessary market data is available to price any option. This is all saved in dataset pricingdata with each table named for the ccyPair. Recently added long date capabilities -any option longer than 1y the atm vol is interpoloted via cubic spline - the inputs is the vol curve passed from the dvi. 

        //3) for each option the atmvol is taken for the dvi text file. The rr and fly are interpolated from pricingdata - the fly multiplier is also interopolated from pricing data which is then used to iterate a wingcontrol used in the vol interp. fwds and depos come from marketdata dataset. The user can update rr and flies on the displayed pricing data datagrid. the updated data will be used in the pricing interp. 

        //other - bid/offer spreads are taken from seperate excel file. smile multipliers as well as holidays and all individual curreny setup details are saved into xml files which are loaded on program initialize. 

        private void path()
        {

            string filename = @"C:\priceEngineUser.txt";
            string[] ReadFile = File.ReadAllLines(filename).ToArray();
            systemFiles = ReadFile[0];
            user = ReadFile[1];

            xmlFilePath = systemFiles + user + @"\xmlFiles\";


            filename = @"C:\DVIUSER.txt";
            string[] ReadFileVol = File.ReadAllLines(filename).ToArray();
            string directory = ReadFileVol[0];

            volPath = directory + user + @"\systemFiles\";

        }

        private void intiliazePricer()
        {

            if (pricingData == null)
                pricingData = new DataSet();

            if (brokerRun == null)
                brokerRun = new DataSet();

            if (tradeSim == null)
                tradeSim = new DataSet();
            
            if (regaSega == null)
                regaSega = new DataSet();

            path();

            //creates the dataTables.
            crossesDt(); //stores cross settings
            ccyDetsDt(); //stores individual currency settings
            pricerDt();//sets pricer interFace            
            LoadHols();//loads holidays from xml file
            loadSmileMult(); //loads smile data from xml file
            fill_Cross_Box(); //fill pulldown combobox on marketdata tab with each cross. 
            fill_tool_bar();
            spreadsFromExcel();//loadspreads from excel sheet


            //adds an instance of bbg class to new tab for forwards. 
            setBbgInterfacefwds();
            setBbgInterfaceDepos();

            rowNames = pricer.AsEnumerable().Select(x => x[0].ToString()).ToList();
            setRows();//rows

            endIni = 1;


            refreshAllToolStripMenuItem.BackColor = Color.Red;

        }

        private void refreshDataButton()
        {
            bool startUp = false;

            if (refreshAllToolStripMenuItem.BackColor == Color.Red)
            { startUp = true; }

            refreshAllToolStripMenuItem.Text = "Updating...";
            refreshAllToolStripMenuItem.ForeColor = Color.BlueViolet;
            refreshAllToolStripMenuItem.BackColor = Color.LightBlue;

            fwds.sendRequest();
            depos.sendRequest();

            DataRow lastFwds = fwds.d_data.Rows[fwds.d_data.Rows.Count - 1];
            DataRow lastDepos = depos.d_data.Rows[depos.d_data.Rows.Count - 1];

            bool endload = false;

            do
            {
                if (lastFwds["FWD_CURVE"] != DBNull.Value && lastDepos["PAR_CURVE"] != DBNull.Value)
                {
                    endload = true;
                }
            } while (endload == false);

            //once bbgData is loaded rateTiles are set then saved into MarketData
            dataSetFwds();
            dataSetDepos();
            setMarketData();

            refreshAllToolStripMenuItem.ForeColor = Color.Black;
            refreshAllToolStripMenuItem.Text = "RefreshMarketData";


            //on startup dataTables, formats and spreads from excel are set.
            if (startUp == true)
            {

                bindDataTables();
                addCombobox();
                formatDataTables();

                //set default surface view to first currency in crosses dt. 

                string ccyPair = crosses.Rows[0]["Cross"].ToString();
                displaySurface(ccyPair);

            }

        }

        private DataTable crossesDt()
        {
            if (crosses == null)
                crosses = new DataTable();

            if (defaultBbgSource == null)
                defaultBbgSource = new DataTable();

            crosses.Columns.Add("Cross");
            crosses.Columns.Add("Depo");
            crosses.Columns.Add("SolveBase");
            crosses.Columns.Add("Source");
            crosses.Columns.Add("Factor");
            crosses.Columns.Add("DeltaType");
            crosses.AcceptChanges();
            crosses.TableName = "crosses";
            LoadXmldt(crosses, xmlFilePath, "currSetup");

            defaultBbgSource.Columns.Add("security");
            defaultBbgSource.Columns.Add("bbgSource");

            //sets default bbgSource in case user changes it. Default source is saved to xml file and loaded on intialize
            foreach (DataRow row in crosses.Rows)
                defaultBbgSource.Rows.Add(row["Cross"].ToString(), row["Source"].ToString());

            return crosses;





        }

        private DataTable mxVolsDt()
        {

            DataTable mxVols = new DataTable();

            mxVols.Columns.Add("Term");
            mxVols.Columns.Add("10DFLYm");
            mxVols.Columns.Add("M25DFLYm");
            mxVols.Columns.Add("ATMm");
            mxVols.Columns.Add("25DRRm");
            mxVols.Columns.Add("10DRRm");
            mxVols.Columns.Add("Term1");
            mxVols.Columns.Add("10DFLYp");
            mxVols.Columns.Add("25DFLYp");
            mxVols.Columns.Add("ATMp");
            mxVols.Columns.Add("25DRRp");
            mxVols.Columns.Add("10DRRp");
            mxVols.Columns.Add("Term2");
            mxVols.Columns.Add("10DFLYc");
            mxVols.Columns.Add("25DFLYc");
            mxVols.Columns.Add("ATMc");
            mxVols.Columns.Add("25DRRc");
            mxVols.Columns.Add("10DRRc");
            mxVols.AcceptChanges();
            //  mxVols.TableName = "ccyDets";


            return mxVols;

        }

        private DataTable ccyDetsDt()
        {
            if (ccyDets == null)
                ccyDets = new DataTable();

            ccyDets.Columns.Add("Ccy");
            ccyDets.Columns.Add("CdrCode");
            ccyDets.Columns.Add("DayCountBasis");
            ccyDets.Columns.Add("PointsFactor");
            ccyDets.Columns.Add("CcyTerms");
            ccyDets.Columns.Add("CcySpotDays");
            ccyDets.Columns.Add("YieldCurveCode");
            ccyDets.AcceptChanges();
            ccyDets.TableName = "ccyDets";
            LoadXmldt(ccyDets, xmlFilePath, "ccySetupDets");

            return ccyDets;

        }

        private DataTable pricerDt()
        {
            if (pricer == null)
                pricer = new DataTable();

            pricer.Columns.Add("OptionNumber");


            for (int i = 1; i < 50; i++)
            {
                string col = (i).ToString();
                pricer.Columns.Add(col);
            }

            pricer.Rows.Add("CcyPair");
            pricer.Rows.Add("Spot");
            pricer.Rows.Add("Expiry");
            pricer.Rows.Add("Strike");
            pricer.Rows.Add("Put_Call");
            pricer.Rows.Add("Notional");
            pricer.Rows.Add();

            pricer.Rows.Add("Vol");
            pricer.Rows.Add("Bid_Offer");
            pricer.Rows.Add();

         

            
            pricer.Rows.Add("Spot_Delta");
            pricer.Rows.Add("Fwd_Delta");
            pricer.Rows.Add("Vega");
            pricer.Rows.Add("AutoStrike");
           
            pricer.Rows.Add();
            pricer.Rows.Add("Price");
            pricer.Rows.Add("Fwd_Price");       
            pricer.Rows.Add("Premium_Type");
            pricer.Rows.Add("BloombergVol");
            pricer.Rows.Add();

            pricer.Rows.Add("SystemVol");
            pricer.Rows.Add("BpsFromMid");
            pricer.Rows.Add("PremiumFromMid");
            pricer.Rows.Add("Vol_Spread_to_Atm");
            pricer.Rows.Add("Bps_to_Atm");
            pricer.Rows.Add();

            pricer.Rows.Add("ExpiryDate");
            pricer.Rows.Add("DeliveryDate");
            pricer.Rows.Add("Expiry_Days");
            pricer.Rows.Add("Deliver_Days");
            pricer.Rows.Add();

            pricer.Rows.Add("ATM_VOL");
            pricer.Rows.Add("RR");
            pricer.Rows.Add("FLY");
            pricer.Rows.Add("BreakEven");
            pricer.Rows.Add();

            pricer.Rows.Add("Swap_Pts");
            pricer.Rows.Add("Fwd");
            pricer.Rows.Add("Depo_Base");
            pricer.Rows.Add("Depo_Terms");
            pricer.Rows.Add("BBG Source");  
            pricer.Rows.Add();

         
            pricer.Rows.Add("sVega");
            pricer.Rows.Add("sDelta");
            pricer.Rows.Add("sGamma");
            pricer.Rows.Add("PremiumA");
            pricer.Rows.Add("VegaA");            
            pricer.Rows.Add("DeltaA");
            pricer.Rows.Add("GammaA");    
            pricer.Rows.Add("Rega25");
            pricer.Rows.Add("Rega10");
            pricer.Rows.Add("Sega25");           
            pricer.Rows.Add("Sega10");
            pricer.Rows.Add("VannaA");
            pricer.Rows.Add("VolgaA");
            pricer.Rows.Add("ThetaA");
            pricer.Rows.Add("DV01_BaseA");
            pricer.Rows.Add("DV01_TermsA");    
            pricer.Rows.Add();

           
            //pricer.Rows.Add("Gamma");
            //pricer.Rows.Add("Vanna");
            //pricer.Rows.Add("Volga");
            //pricer.Rows.Add("Theta");
            //pricer.Rows.Add("DV01_Base");
            //pricer.Rows.Add("DV01_Terms");

            pricer.Rows.Add("Today");
            pricer.Rows.Add("Spotdate");
            pricer.Rows.Add("Basis_Base");
            pricer.Rows.Add("Basis_Terms");
            pricer.Rows.Add("Premium_Included");

            pricer.AcceptChanges();



            return pricer;

        }

        private void setRows()
        {



            CcyPairR = rowNames.IndexOf("CcyPair");
            SpotR = rowNames.IndexOf("Spot");
            ExpiryR = rowNames.IndexOf("Expiry");
            StrikeR = rowNames.IndexOf("Strike");
            Put_CallR = rowNames.IndexOf("Put_Call");
            NotionalR = rowNames.IndexOf("Notional");
            Expiry_DaysR = rowNames.IndexOf("Expiry_Days");
            VolR = rowNames.IndexOf("Vol");
            SystemVolR = rowNames.IndexOf("SystemVol");
            BpsFromMidR = rowNames.IndexOf("BpsFromMid");
            PremiumFromMidR = rowNames.IndexOf("PremiumFromMid");
            BloombergVolR = rowNames.IndexOf("BloombergVol");
            Bid_OfferR = rowNames.IndexOf("Bid_Offer");
            BreakEvenR = rowNames.IndexOf("BreakEven");
            Vol_Spread_to_AtmR = rowNames.IndexOf("Vol_Spread_to_Atm");
            Bps_to_AtmR = rowNames.IndexOf("Bps_to_Atm");
            Spot_DeltaR = rowNames.IndexOf("Spot_Delta");
            Fwd_DeltaR = rowNames.IndexOf("Fwd_Delta");
            AutoStrikeR = rowNames.IndexOf("AutoStrike");
            ATM_VOLR = rowNames.IndexOf("ATM_VOL");
            RRR = rowNames.IndexOf("RR");
            FLYR = rowNames.IndexOf("FLY");
            Swap_PtsR = rowNames.IndexOf("Swap_Pts");
            FwdR = rowNames.IndexOf("Fwd");
            BbgSourceR = rowNames.IndexOf("BBG Source");
            ExpiryDateR = rowNames.IndexOf("ExpiryDate");
            DeliveryDateR = rowNames.IndexOf("DeliveryDate");
            PriceR = rowNames.IndexOf("Price");
            Fwd_PriceR = rowNames.IndexOf("Fwd_Price");
            VegaR = rowNames.IndexOf("Vega");
            VannaR = rowNames.IndexOf("Vanna");
            VolgaR = rowNames.IndexOf("Volga");
            GammaR = rowNames.IndexOf("Gamma");
            ThetaR = rowNames.IndexOf("Theta");

            Depo_BaseR = rowNames.IndexOf("Depo_Base");
            Depo_TermsR = rowNames.IndexOf("Depo_Terms");
            Basis_BaseR = rowNames.IndexOf("Basis_Base");
            Basis_TermsR = rowNames.IndexOf("Basis_Terms");
            TodayR = rowNames.IndexOf("Today");
            SpotdateR = rowNames.IndexOf("Spotdate");
            DaycountR = rowNames.IndexOf("Daycount");
            Deliver_DaysR = rowNames.IndexOf("Deliver_Days");
            DeltaAR = rowNames.IndexOf("DeltaA");
            GammaAR = rowNames.IndexOf("GammaA");
            sGammaR = rowNames.IndexOf("sGamma");
            VegaAR = rowNames.IndexOf("VegaA");
            ThetaAR = rowNames.IndexOf("ThetaA");
            DV01_BaseR = rowNames.IndexOf("DV01_Base");
            DV01_TermsR = rowNames.IndexOf("DV01_Terms");
            PremiumAR = rowNames.IndexOf("PremiumA");
            Premium_TermsR = rowNames.IndexOf("Premium_Terms");
            Premium_IncludedR = rowNames.IndexOf("Premium_Included");
            Premium_TypeR = rowNames.IndexOf("Premium_Type");
            //SegaR = rowNames.IndexOf("Sega");
            //RegaR = rowNames.IndexOf("Rega");
            Rega25R = rowNames.IndexOf("Rega25");
            Sega25R = rowNames.IndexOf("Sega25");
            Rega10R = rowNames.IndexOf("Rega10");
            Sega10R = rowNames.IndexOf("Sega10");
            sVegaR = rowNames.IndexOf("sVega");
            sDeltaR = rowNames.IndexOf("sDelta");
            VannaAR = rowNames.IndexOf("VannaA");
            VolgaAR = rowNames.IndexOf("VolgaA");
            DV01_BaseAR = rowNames.IndexOf("DV01_BaseA");
            DV01_TermsAR = rowNames.IndexOf("DV01_TermsA");

        }

        private DataTable surfaceDt()
        {
            DataTable surface = new DataTable();

            surface.Columns.Add("DayCount");
            surface.Columns.Add("Maturity");
            surface.Columns.Add("ExpiryDate");
            surface.Columns.Add("DeliveryDate");
            surface.Columns.Add("ATM");
            surface.Columns.Add("10DR");
            surface.Columns.Add("25DR");
            surface.Columns.Add("25D_BrokerFly");
            surface.Columns.Add("10D_BrokerFly");
            surface.Columns.Add("RR_Multiplier");
            surface.Columns.Add("SmileFly_Multiplier");
            surface.Columns.Add("BrokerFly_Multiplier");
         //   surface.Columns.Add("WingCtrlDn");
           // surface.Columns.Add("WingCtrlUp");
            surface.Columns.Add("Forward");
            surface.Columns.Add("Points");
            surface.Columns.Add("DepoDom");
            surface.Columns.Add("DepoFor");
            surface.Columns.Add("DFDom");
            surface.Columns.Add("DFFor");
            surface.Columns.Add("25D_SmileFly");
            surface.Columns.Add("10D_SmileFly");
            surface.Columns.Add("10dPutVol");
            surface.Columns.Add("25dPutVol");
            surface.Columns.Add("ATMVol");
            surface.Columns.Add("25dCallVol");
            surface.Columns.Add("10dCallVol");
            surface.Columns.Add("10dPutStrike");
            surface.Columns.Add("25dPutStrike");
            surface.Columns.Add("ATMStrike");
            surface.Columns.Add("25dCallStrike");
            surface.Columns.Add("10dCallStrike");

            return surface;
        }

        private void hide_rows()
        {

           
              bool t = dataGridView1.Rows[SystemVolR].Visible;
                
                {
                    for (int i = SystemVolR; i < sVegaR; i++)
                    {
                        if (t == false)
                        {
                            dataGridView1.Rows[i].Visible = true;
                        }
                        else
                        {
                            dataGridView1.Rows[i].Visible = false;
                        }
   
                    }

                    for (int i = TodayR; i < Premium_IncludedR+1; i++)
                    {
                        if (t == false)
                        {
                            dataGridView1.Rows[i].Visible = true;
                        }
                        else
                        {
                            dataGridView1.Rows[i].Visible = false;
                        }

                    }
                    
                }
                
        }

        private void bindDataTables()
        {
            dataGridView1.DataSource = pricer;
           
            dataGridView2.DataSource = ccyDets;
            dataGridView3.DataSource = crosses;

            DataTable surf = surfaceDt();
            dataGridView8.DataSource = surf;
        }

        private void formatDataTables()
        {
            formatPricerDt();
            formatSurfaceView();
        }

        private double smileVolGreeks(double atmVol, double rr, double fly, double spot, double fwdPts, double autoSt, DateTime dayStart, DateTime autoExp, double dfDomDel, double dfForDel, double dfForExp, double dfDomExp, int premoInc, double smileFlyMult, double rrMult)
        {
          
            double wingControl = 0;
          
            
            DateTime tod = today;

            double outRight = spot + fwdPts;
  
            if (autoSt < outRight)
            {
                wingControl = reCalibrateWingControl(atmVol, rr, fly, spot, dayStart, autoExp, dfDomDel, dfForExp, premoInc, smileFlyMult, rrMult, 0);
            }
            else
            {
                wingControl = reCalibrateWingControl(atmVol, rr, fly, spot, dayStart, autoExp, dfDomDel, dfForExp, premoInc, smileFlyMult, rrMult, 1);
            }


            double smileFly = equivalentfly(spot, tod, autoExp, atmVol, atmVol, rr, fly, dfDomExp, dfForExp, premoInc);
            double putVol = atmVol + smileFly - rr / 2;
            double callVol = atmVol + smileFly + rr / 2;

            double putStrike = FXStrikeVol(spot, dayStart, autoExp, 0.25, putVol, dfDomDel, dfForDel, "p", premoInc);
            double callStrike = FXStrikeVol(spot, dayStart, autoExp, 0.25, callVol, dfDomDel, dfForDel, "c", premoInc);
            double atmStrike = FXATMStrike(spot, dayStart, autoExp, atmVol, dfDomDel, dfForDel, premoInc);

            double vol = smileInterp(spot, dayStart, autoExp, wingControl * atmVol, autoSt, putStrike, putVol, atmStrike, atmVol, callStrike, callVol, dfDomExp, dfForExp, dfDomDel, dfForDel);

            return vol;


        }

        private object [] surfPts(double atmVol, double rr, double fly, double rrMult, double smileFlyMult,  double spot, DateTime dayStart, DateTime autoExp, double dfDomDel, double dfForDel, int premoInc)
        {
            object[] retVal = null;
            double wingControl = 1;
            //need to calc smilefly then get 25d and atm strikes and vols
            double smileFly = equivalentfly(spot, dayStart, autoExp, atmVol * wingControl, atmVol, rr, fly, dfDomDel, dfForDel, premoInc);
            double v25p = atmVol + smileFly - rr / 2;
            double v25c = atmVol + smileFly + rr / 2;

            double k25p = FXStrikeVol(spot, dayStart, autoExp, 0.25, v25p, dfDomDel, dfForDel, "p", premoInc);
            double k25c = FXStrikeVol(spot, dayStart, autoExp, 0.25, v25c, dfDomDel, dfForDel, "c", premoInc);
            double katm = FXATMStrike(spot, dayStart, autoExp, atmVol, dfDomDel, dfForDel, premoInc);
            double v10c= atmVol + 0.5 * (rr * rrMult) + (smileFly * smileFlyMult);
            double k10c = FXStrikeVol(spot, dayStart, autoExp, 0.1, v10c, dfDomDel, dfForDel, "c", premoInc);

            double v10p = atmVol - 0.5 * (rr * rrMult) + (smileFly * smileFlyMult);
            double k10p = FXStrikeVol(spot, dayStart, autoExp, 0.1, v10p, dfDomDel, dfForDel, "p", premoInc);

            retVal = new object[] { k10p, k25p, katm, k25c, k10c,v10p,v25p,atmVol,v25c,v10c };
            return retVal;
        }

        private double fxVol(double k, double atmVol, double rr, double fly, double rrMult, double smileFlyMult, double spot, DateTime dayStart, DateTime autoExp, double dfDomExp, double dfForExp,double dfDomDel, double dfForDel, int premoInc,  double splineFactor)
        {
            double retVal = 0;
            double wingControl = 1;
            double fwd = spot * dfForDel / dfDomDel;
            //need to calc smilefly then get 25d and atm strikes and vols
            double smileFly = equivalentfly(spot, dayStart, autoExp, atmVol * wingControl, atmVol, rr, fly, dfDomDel, dfForDel, premoInc);
            double v25p = atmVol + smileFly - rr / 2;
            double v25c = atmVol + smileFly + rr / 2;

            double k25p = FXStrikeVol(spot, dayStart, autoExp, 0.25, v25p, dfDomDel, dfForDel, "p", premoInc);
            double k25c = FXStrikeVol(spot, dayStart, autoExp, 0.25, v25c, dfDomDel, dfForDel, "c", premoInc);
            double katm = FXATMStrike(spot, dayStart, autoExp, atmVol, dfDomDel, dfForDel, premoInc);
            double v10c = atmVol + 0.5 * (rr * rrMult) + (smileFly * smileFlyMult);
            double k10c = FXStrikeVol(spot, dayStart, autoExp, 0.1, v10c, dfDomDel, dfForDel, "c", premoInc);

            double v10p = atmVol - 0.5 * (rr * rrMult) + (smileFly * smileFlyMult);
            double k10p = FXStrikeVol(spot, dayStart, autoExp, 0.1, v10p, dfDomDel, dfForDel, "p", premoInc);

          //  double volSimDeltaU = combinedInterp(spotSim * 1.005, dayStart, autoExp, wingControl * vatm, autoSt, k10p, v10p, k25p, v25p, katm, vatm, k25c, v25c, k10c, v10c, dfDomExp, dfForExp, dfDomDel, dfForDel, fwd, smileFactor);

            retVal = combinedInterp(spot, dayStart, autoExp, atmVol, k, k10p, v10p, k25p, v25p, katm, atmVol, k25c, v25c, k10c, v10c, dfDomExp, dfForExp, dfDomDel, dfForDel, fwd, splineFactor);

            if (retVal < 0) { retVal = 0.00000001; }
            return retVal;
        }


        private object[] smileGreeks(double k, double atmVol, double rr, double fly, double rrMult, double smileFlyMult, double spot, DateTime dayStart, DateTime autoExp, double dfDomExp, double dfForExp, double dfDomDel, double dfForDel, int premoInc, double splineFactor, string pC, string premoString, double notional)
        {
            object[] retVal = null;

           
         

            double vUnShocked = fxVol( k, atmVol, rr, fly, rrMult, smileFlyMult, spot, dayStart, autoExp, dfDomExp, dfForExp, dfDomDel, dfForDel, premoInc, splineFactor);
            double[] pUnShocked = FXOpts(spot, dayStart, autoExp, k, vUnShocked, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);

            double vDeltaUp = fxVol(k, atmVol, rr, fly, rrMult, smileFlyMult, spot * 1.0050, dayStart, autoExp, dfDomExp, dfForExp, dfDomDel, dfForDel, premoInc, splineFactor);
            double[] pDeltaUp = FXOpts(spot * 1.0050, dayStart, autoExp, k, vDeltaUp, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);

            double vDeltaDn = fxVol(k, atmVol, rr, fly, rrMult, smileFlyMult, spot * 0.9950, dayStart, autoExp, dfDomExp, dfForExp, dfDomDel, dfForDel, premoInc,  splineFactor);
            double[] pDeltaDn = FXOpts(spot * 0.9950, dayStart, autoExp, k, vDeltaDn, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);

            double vGammaUp = fxVol(k, atmVol, rr, fly, rrMult, smileFlyMult, spot * 1.01, dayStart, autoExp, dfDomExp, dfForExp, dfDomDel, dfForDel, premoInc, splineFactor);
            double[] pGammaUp = FXOpts(spot * 1.01, dayStart, autoExp, k, vGammaUp, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);
                      
            double vGamma1Up = fxVol(k, atmVol, rr, fly, rrMult, smileFlyMult, spot *1.01 * 1.0050, dayStart, autoExp, dfDomExp, dfForExp, dfDomDel, dfForDel, premoInc, splineFactor);
            double[] pGamma1Up = FXOpts(spot * 1.01 * 1.0050, dayStart, autoExp, k, vGamma1Up, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);

            double vGamma2Up = fxVol(k, atmVol, rr, fly, rrMult, smileFlyMult, spot * 1.01 * 0.9950, dayStart, autoExp, dfDomExp, dfForExp, dfDomDel, dfForDel, premoInc, splineFactor);
            double[] pGamma2Up = FXOpts(spot * 1.01 * 0.9950, dayStart, autoExp, k, vGamma2Up, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);

            double vGammaDn = fxVol(k, atmVol, rr, fly, rrMult, smileFlyMult, spot * .99, dayStart, autoExp, dfDomExp, dfForExp, dfDomDel, dfForDel, premoInc, splineFactor);
            double[] pGammaDn = FXOpts(spot * .99, dayStart, autoExp, k, vGammaDn, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);

            double vGamma1Dn = fxVol(k, atmVol, rr, fly, rrMult, smileFlyMult, spot * .99 * 1.0050, dayStart, autoExp, dfDomExp, dfForExp, dfDomDel, dfForDel, premoInc, splineFactor);
            double[] pGamma1Dn = FXOpts(spot * .99 * 1.0050, dayStart, autoExp, k, vGamma1Dn, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);

            double vGamma2Dn = fxVol(k, atmVol, rr, fly, rrMult, smileFlyMult, spot * .99 * 0.9950, dayStart, autoExp, dfDomExp, dfForExp, dfDomDel, dfForDel, premoInc, splineFactor);
            double[] pGamma2Dn = FXOpts(spot * .99 * 0.9950, dayStart, autoExp, k, vGamma2Dn, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);

            //smileVega
            double vVega = fxVol( k, atmVol + .01, rr, fly, rrMult, smileFlyMult, spot, dayStart, autoExp, dfDomExp, dfForExp, dfDomDel, dfForDel, premoInc,  splineFactor);
            //calc prmeiums 
            double[] pVega = FXOpts(spot, dayStart, autoExp, k, vVega, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);

            //Rega25
            double rrMult25 = (rr * rrMult) / (rr + .01);
            double vRega25 = fxVol( k, atmVol, rr + .01, fly, rrMult25, smileFlyMult, spot, dayStart, autoExp, dfDomExp, dfForExp, dfDomDel, dfForDel, premoInc,  splineFactor);
            double[] pRega25 = FXOpts(spot, dayStart, autoExp, k, vRega25, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);

            
            double rrMult10 = (rr*rrMult+.01)/rr;
            double vRega10 = fxVol( k, atmVol, rr, fly, rrMult10, smileFlyMult, spot, dayStart, autoExp, dfDomExp, dfForExp, dfDomDel, dfForDel, premoInc,  splineFactor);
            double[] pRega10 = FXOpts(spot, dayStart, autoExp, k, vRega10, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);

            double flyMult25 = (fly * smileFlyMult) / (fly + .01);
            double vSega25 = fxVol( k, atmVol, rr, fly + .01, rrMult, flyMult25, spot, dayStart, autoExp, dfDomExp, dfForExp, dfDomDel, dfForDel, premoInc,  splineFactor);
            double[] pSega25 = FXOpts(spot, dayStart, autoExp, k, vSega25, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);

            double flyMult10 = (fly * smileFlyMult + .01) / fly;
            double vSega10 = fxVol( k, atmVol, rr, fly, rrMult, flyMult10, spot, dayStart, autoExp, dfDomExp, dfForExp, dfDomDel, dfForDel, premoInc,  splineFactor);
            double[] pSega10 = FXOpts(spot, dayStart, autoExp, k, vSega10, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);

            //get premoConversions factors (price for fxopts comes in base pips. 

            double[] premoTypeInfoS = premoConventions(premoString, spot, k);
            double premoConversionS = premoTypeInfoS[0];
            double premoFactorS = premoTypeInfoS[1];

            double[] premoTypeInfoDU = premoConventions(premoString, spot * 1.005, k);
            double premoConversionDU = premoTypeInfoDU[0];

            double[] premoTypeInfoDD = premoConventions(premoString, spot * 0.9950, k);
            double premoConversionDD = premoTypeInfoDD[0];

            double[] premoTypeInfoGU = premoConventions(premoString, spot * 1.01, k);
            double premoConversionGU = premoTypeInfoGU[0];

            double[] premoTypeInfoGU1 = premoConventions(premoString, spot * 1.01 * 1.005, k);
            double premoConversionGU1 = premoTypeInfoGU1[0];

            double[] premoTypeInfoGU2 = premoConventions(premoString, spot * 1.01 * 0.9950, k);
            double premoConversionGU2 = premoTypeInfoGU2[0];

            double[] premoTypeInfoGD = premoConventions(premoString, spot * 0.99, k);
            double premoConversionGD = premoTypeInfoGD[0];  

            double[] premoTypeInfoGD1 = premoConventions(premoString, spot * 0.99 * 1.0050, k);
            double premoConversionGD1 = premoTypeInfoGD1[0];  
            
            double[] premoTypeInfoGD2 = premoConventions(premoString, spot * 0.99 * 0.9950, k);
            double premoConversionGD2 = premoTypeInfoGD2[0];  

            //calc premiums is correct units
            double premoUnShocked = pUnShocked[0] * premoConversionS;
            double premoVega = pVega[0] * premoConversionS;
            double premoRega25 = pRega25[0] * premoConversionS;
            double premoRega10 = pRega10[0] * premoConversionS;
            double premoSega25 = pSega25[0] * premoConversionS;
            double premoSega10 = pSega10[0] * premoConversionS;
            double premoDeltaUp = pDeltaUp[0] * premoConversionDU;
            double premoDeltaDn = pDeltaDn[0] * premoConversionDD;

            double premoGammaUp = pGammaUp[0] * premoConversionGU * premoFactorS * notional;
            double premoGamma1Up = pGamma1Up[0] * premoConversionGU1;
            double premoGamma2Up = pGamma2Up[0] * premoConversionGU2;

            double premoGammaDn = pGammaDn[0] * premoConversionGD * premoFactorS * notional;
            double premoGamma1Dn = pGamma1Dn[0] * premoConversionGD1;
            double premoGamma2Dn = pGamma2Dn[0] * premoConversionGD2;

             //calc smileGreeks
            double smileVega = (premoVega - premoUnShocked)*premoFactorS * notional;
            double smileRega25 = (premoRega25 - premoUnShocked)*premoFactorS * notional;
            double smileRega10 = (premoRega10 - premoUnShocked)*premoFactorS * notional;
            double smileSega25 = (premoSega25 - premoUnShocked)*premoFactorS * notional;
            double smileSega10 = (premoSega10 - premoUnShocked)*premoFactorS * notional;
            double smileDelta = (premoDeltaUp - premoDeltaDn) * premoFactorS * notional;
            double smileDeltaUp = (premoGamma1Up - premoGamma2Up) * premoFactorS * notional;
            double smileDeltaDn = (premoGamma1Dn - premoGamma2Dn) * premoFactorS * notional;

            

            double cashPremo = premoUnShocked * premoFactorS * notional;

            double chgSpot = Math.Abs(0.9950 / 1.0050 - 1);
            smileDelta = smileDelta / chgSpot;
            smileDeltaUp = smileDeltaUp / chgSpot;
            smileDeltaDn = smileDeltaDn / chgSpot;

            double smileGamma =0;

            smileGamma = (Math.Abs((smileDeltaUp + premoGammaUp) - (smileDelta+ cashPremo)) + Math.Abs((smileDeltaDn +premoGammaDn) - (smileDelta+cashPremo))) / 2;
            if (notional < 0) { smileGamma = smileGamma * -1; }

            if (premoInc != 1)
            {
                smileDelta = smileDelta / spot;
                smileDeltaUp = smileDeltaUp / spot;
                smileDeltaDn = smileDeltaDn / spot;

                smileGamma = (Math.Abs(smileDeltaUp - smileDelta) + Math.Abs(smileDeltaDn - smileDelta)) / 2;
                if (notional < 0) { smileGamma = smileGamma * -1; }
            }


            retVal = new object[] { smileVega, smileRega25, smileRega10, smileSega25, smileSega10, smileDelta, cashPremo,smileGamma };
            return retVal;
        }

        private void optPricer(string expiryText, string strText, string pC)
        {
            int curCol = dataGridView1.CurrentCell.ColumnIndex;


            string conditionCheck = pricer.Rows[ExpiryR][curCol].ToString();
            string conditionCheck1 = pricer.Rows[StrikeR][curCol].ToString();
            string conditionCheck2 = pricer.Rows[Put_CallR][curCol].ToString();

            //will exit the method if at least the 3 needed inputs arent there

            if (conditionCheck == "" || conditionCheck1 == "" || conditionCheck2 == "") { return; }


            //set strings for user updates. will be set to "red" if user updates. 
            // currency, spot, expiry, strike, smileVol, atmVol, rr, brokerfly, wingcontrol, fwd, depoF,depo, premotype
            string userOutRight = dataGridView1.Rows[FwdR].Cells[curCol].Style.BackColor.ToKnownColor().ToString();
            string userFwdPts = dataGridView1.Rows[Swap_PtsR].Cells[curCol].Style.BackColor.ToKnownColor().ToString();
            string userSpot = dataGridView1.Rows[SpotR].Cells[curCol].Style.BackColor.ToKnownColor().ToString();
            string userVol = dataGridView1.Rows[VolR].Cells[curCol].Style.BackColor.ToKnownColor().ToString();
            string userAtmVol = dataGridView1.Rows[ATM_VOLR].Cells[curCol].Style.BackColor.ToKnownColor().ToString();
            string userRR = dataGridView1.Rows[RRR].Cells[curCol].Style.BackColor.ToKnownColor().ToString();
            string userFly = dataGridView1.Rows[FLYR].Cells[curCol].Style.BackColor.ToKnownColor().ToString();
            string userSysVol = dataGridView1.Rows[SystemVolR].Cells[curCol].Style.BackColor.ToKnownColor().ToString();
            string userAmt = dataGridView1.Rows[NotionalR].Cells[curCol].Style.BackColor.ToKnownColor().ToString();
            string userDepoB = dataGridView1.Rows[Depo_BaseR].Cells[curCol].Style.BackColor.ToKnownColor().ToString();
            string userDepoT = dataGridView1.Rows[Depo_TermsR].Cells[curCol].Style.BackColor.ToKnownColor().ToString();

            //  setMarketData();
            double spot = Convert.ToDouble(pricer.Rows[SpotR][curCol]);
            string ccyPair = pricer.Rows[CcyPairR][curCol].ToString();


            string baseCcy = ccyPair.Substring(0, 3);
            string termsCcy = ccyPair.Substring(3, 3);


            DateTime dayStart = today;
            DateTime sptDate = SpotDate(dayStart, baseCcy, termsCcy);
            DateTime autoExp = AutoExpiryDate(expiryText, dayStart, homeCcy, baseCcy, termsCcy);
            DateTime delDate = SpotDate(autoExp, baseCcy, termsCcy);
            double dayCount = (autoExp - dayStart).TotalDays; //expiry to trade date
            double delDayCount = (delDate - sptDate).TotalDays; // delivery to spot date


            //gets daycount basis from ccydets dataTable

            int[] arr = dayCountBasis(ccyPair);
            int basisB = arr[0];
            int basisT = arr[1];

            // calls method to get cross info 
            object[] crossInfo = crossDtData(ccyPair);
            int volType = Convert.ToInt16(crossInfo[0]);
            double factor = Convert.ToDouble(crossInfo[1]);
            string bbgSource = crossInfo[2].ToString();

            //need to change smile factor for usdrub to control number of strikes are calcuted with cubic spline function
            double smileFactor = factor;
            if (ccyPair == "USDRUB" || ccyPair == "EURRUB" || ccyPair == "USDTRY") { smileFactor = 100; }


            //is needed to convert old delatype to new premo included. old was 1 = premo 2 = no, now 1 = premo 0 = no
            int premoInc = 0;


            if (volType == 1)
            {
                premoInc = 1;
            }
            else
            {
                premoInc = 0;
            }

            double[] rateComponents = rateBuilder(ccyPair, dayCount, delDayCount, spot, factor, basisB, basisT);
            //fwdPts, outRight, forDepo, domDepo

            double fwdPts = rateComponents[0];
            double outRight = rateComponents[1];
            double forDepo = rateComponents[2];
            double domDepo = rateComponents[3];



            //need to amend to  change either base depo or terms depo when fwd changes depending on premium currency 
            if (userOutRight == "Red")
            {
                outRight = Convert.ToDouble(pricer.Rows[FwdR][curCol]);
                fwdPts = (outRight - spot) * factor;
                domDepo = SolveTermsDepo(outRight, spot, delDayCount, forDepo, basisB, basisT);
            }

            if (userFwdPts == "Red")
            {
                fwdPts = Convert.ToDouble(pricer.Rows[Swap_PtsR][curCol]);
                outRight = fwdPts / factor + spot;
                domDepo = SolveTermsDepo(outRight, spot, delDayCount, forDepo, basisB, basisT);
            }

            if (userDepoB == "Red")
            {
                forDepo = convertPercent(pricer.Rows[Depo_BaseR][curCol].ToString());
                domDepo = SolveTermsDepo(outRight, spot, delDayCount, forDepo, basisB, basisT);
            }

            if (userDepoT == "Red")
            {
                domDepo = convertPercent(pricer.Rows[Depo_TermsR][curCol].ToString());
                forDepo = SolveBaseDepo(outRight, spot, delDayCount, domDepo, basisB, basisT);
            }
            double dfForExp = DiscountFactor(forDepo, Convert.ToInt16(dayCount), Convert.ToInt16(basisB));
            double dfForDel = DiscountFactor(forDepo, Convert.ToInt16(delDayCount), Convert.ToInt16(basisB));

            double dfDomExp = DiscountFactor(domDepo, Convert.ToInt16(dayCount), Convert.ToInt16(basisT));
            double dfDomDel = DiscountFactor(domDepo, Convert.ToInt16(delDayCount), Convert.ToInt16(basisT));


            //get fly, rr from pricerData displayed on main pricer screen 
            double[] volComponents = null;

            volComponents = volBuilder(ccyPair, dayCount);
            double atmVol = volComponents[0];
            double rr = volComponents[1];
            double fly = volComponents[2];
            double wingControl = volComponents[3];
            double targetFlyMult = volComponents[4];
            double smileFlyMult = volComponents[5];
            double rrMult = volComponents[6];

            if (userAtmVol == "Red") { atmVol = convertPercent(pricer.Rows[ATM_VOLR][curCol].ToString()); }
            if (userRR == "Red") { rr = convertPercent(pricer.Rows[RRR][curCol].ToString()); }
            if (userFly == "Red") { fly = convertPercent(pricer.Rows[FLYR][curCol].ToString()); }




            wingControl = 1;
            //need to calc smilefly then get 25d and atm strikes and vols
            double smileFly = equivalentfly(spot, dayStart, autoExp, atmVol * wingControl, atmVol, rr, fly, dfDomDel, dfForDel, premoInc);
            double putVol = atmVol + smileFly - rr / 2;
            double callVol = atmVol + smileFly + rr / 2;

            double putStrike = FXStrikeVol(spot, dayStart, autoExp, 0.25, putVol, dfDomDel, dfForDel, "p", premoInc);
            double callStrike = FXStrikeVol(spot, dayStart, autoExp, 0.25, callVol, dfDomDel, dfForDel, "c", premoInc);
            double atmStrike = FXATMStrike(spot, dayStart, autoExp, atmVol, dfDomDel, dfForDel, premoInc);
            double callVol10 = atmVol + 0.5 * (rr * rrMult) + (smileFly * smileFlyMult);
            double callStrike10 = FXStrikeVol(spot, dayStart, autoExp, 0.1, callVol10, dfDomDel, dfForDel, "c", premoInc);

            double putVol10 = atmVol - 0.5 * (rr * rrMult) + (smileFly * smileFlyMult);
            double putStrike10 = FXStrikeVol(spot, dayStart, autoExp, 0.1, putVol10, dfDomDel, dfForDel, "p", premoInc);

            string premoString = pricer.Rows[Premium_TypeR][curCol].ToString(); //checks datagridview for value

            //defaults premoInfo to base% for premo included and terms pips if not
            if (premoString == "" && premoInc == 1)
            {
                premoString = "Base %";
            }

            else if (premoString == "" && premoInc == 0)
            {
                premoString = "Terms Pips";
            }

            //get premoInc for  the delta solve. Keep in mind that the surface is built with premo included or not from the setup menu. This will ensure that strikes will have the correct vols despite what premo convention is used for an individual option. 

            double[] premoTypeInfo = premoConventions(premoString, spot, 1);

            int premoIncDeltSolve = Convert.ToInt16(premoTypeInfo[3]);

            double autoSt = 0;

            if (strText == "a") { autoSt = atmStrike; }

            if (IsNumeric(strText) == true)
            {
                autoSt = Convert.ToDouble(strText);
            }
            else
            {
                int TextLength = 0;
                TextLength = strText.Length;
                if (strText == "atmf")
                {
                    autoSt = outRight;
                }
                else if (strText == "atms")
                {
                    autoSt = spot;
                }
                else if (strText.Substring(TextLength - 1) == "d")
                {
                    double delt = Convert.ToDouble(strText.Substring(0, TextLength - 1)) / 100;

                    autoSt = FXStrikeDelta(delt, pC, outRight, atmVol, rr, fly, rrMult, smileFlyMult, spot, dayStart, autoExp, dfDomExp, dfForExp, dfDomDel, dfForDel, premoInc, smileFactor);

                    if (userVol == "Red")
                    {
                        double volU = convertPercent(pricer.Rows[VolR][curCol].ToString());
                        autoSt = FXStrikeVol(spot, dayStart, autoExp, delt, volU, dfDomDel, dfForDel, pC, premoIncDeltSolve);
                    }

                }
            }



            double vol = fxVol(autoSt, atmVol, rr, fly, rrMult, smileFlyMult, spot, dayStart, autoExp, dfDomExp, dfForExp, dfDomDel, dfForDel, premoInc, smileFactor);




           
           // double vol = combinedInterp(spot, dayStart, autoExp, wingControl * atmVol, autoSt, putStrike10, putVol10, putStrike, putVol, atmStrike, atmVol, callStrike, callVol, callStrike10, callVol10, dfDomExp, dfForExp, dfDomDel, dfForDel, outRight, smileFactor);



            double systemVol = vol;

            if (userSysVol == "Red")
            { systemVol = convertPercent(pricer.Rows[SystemVolR][curCol].ToString()); }

            if (userVol == "Red")
            { vol = convertPercent(pricer.Rows[VolR][curCol].ToString()); }

            premoTypeInfo = premoConventions(premoString, spot, autoSt);
            double premoConversion = premoTypeInfo[0];//applies this factor greeks to convert in proper units
            double premoFactor = premoTypeInfo[1];// will convert greeks to correct units in nominal amounts
            double notionalFactor = premoTypeInfo[2]; //notional factor is equal to spot or 1 - will convert notional to terms currency if % terms or base pips is selected


            double[] greeks = FXOpts(spot, dayStart, autoExp, autoSt, vol, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);
            double premium = greeks[0];
            double fpremium = greeks[6];
            double delta = greeks[1];
            double fdelta = greeks[5];
            delta = delta - premium / spot * premoIncDeltSolve;
            fdelta = fdelta - fpremium / outRight * premoIncDeltSolve;
            premium = premium * premoConversion;

            List<string> returnType = new List<string>(new string[] { "Base %", "Terms %", "Base Pips", "Terms Pips" });
            double fwd_premoConversion = 0;

            if (premoString == returnType[0]) { fwd_premoConversion = 1 / outRight * 100; }
            if (premoString == returnType[1]) { fwd_premoConversion = 1 / autoSt * 100; }
            if (premoString == returnType[2]) { fwd_premoConversion = 1 / (outRight * autoSt); }
            if (premoString == returnType[3]) { fwd_premoConversion = 1; }


            fpremium = fpremium * fwd_premoConversion;
            double gamma = greeks[2];
            gamma = gamma * spot;
            double vega = greeks[3];
            vega = vega / 100 * premoConversion;

            //theta - just rolls the day 1 day foward - this wont be real theta  as doesnt roll the vol or depo curves. Can work on this later. 
            DateTime dayTheta = AddWorkdays(dayStart, 1);
            double[] theta = FXOpts(spot, dayTheta, autoExp, autoSt, vol, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);
            double sTheta = theta[0] * premoConversion - premium;
           
            double[] smileVolga = FXOpts(spot, dayStart, autoExp, autoSt, vol + .01, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);
            double[] smileVanna = FXOpts(spot * 1.01, dayStart, autoExp, autoSt, vol, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);

            double sVolga = 0;
            double sVanna = 0;

            sVolga = (smileVolga[3] / 100 * premoConversion) - vega; //dvega/dvol
            sVanna = (smileVanna[3] / 100 * premoConversion) - vega; //dvega/dspot

           

            double dvForDepo;
            double dvDomDepo;

            //dv01
            if (fwdPts == 0)
            {
                dvForDepo = 0;
                dvDomDepo = 0;

            }
            else
            {
                dvForDepo = forDepo + .0001;
                dvDomDepo = domDepo + .0001;
            }

            double dvForExp = DiscountFactor(dvForDepo, Convert.ToInt16(dayCount), Convert.ToInt16(basisB));
            double dvForDel = DiscountFactor(dvForDepo, Convert.ToInt16(delDayCount), Convert.ToInt16(basisB));

            double dvDomExp = DiscountFactor(dvDomDepo, Convert.ToInt16(dayCount), Convert.ToInt16(basisT));
            double dvDomDel = DiscountFactor(dvDomDepo, Convert.ToInt16(delDayCount), Convert.ToInt16(basisT));

            double[] dv01For = FXOpts(spot, dayStart, autoExp, autoSt, vol, dfDomExp, dvForExp, dfDomDel, dvForDel, pC);
            double[] dv01Dom = FXOpts(spot, dayStart, autoExp, autoSt, vol, dvDomExp, dfForExp, dvDomDel, dfForDel, pC);

            double sDv01For = dv01For[0] * premoConversion - premium;
            double sDv01Dom = dv01Dom[0] * premoConversion - premium;

            double smileVolSpread = vol - atmVol;
            double[] sVSf = FXOpts(spot, dayStart, autoExp, autoSt, atmVol, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);
            double premoSmileVSflat = premium - sVSf[0] * premoConversion;

            double[] sysVolPremo = FXOpts(spot, dayStart, autoExp, autoSt, systemVol, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);
            double priceFromMid = premium - sysVolPremo[0] * premoConversion;

            double breakEven = atmVol / 24 * Math.Sqrt(dayCount) * spot;

            //greeks in nominal amounts 
            double notional = 100.0;
            if (userAmt == "Red")
            { notional = Convert.ToDouble(pricer.Rows[NotionalR][curCol]); }

            double premiumA = notional * premium * premoFactor * notionalFactor;
            double DeltaA = notional * delta * 1000000;
            double GammaA = notional * gamma * 10000;
            double VegaA = notional * vega * premoFactor * notionalFactor;
            double VannaA = notional * sVanna * premoFactor * notionalFactor;
            double VolgaA = notional * sVolga * premoFactor * notionalFactor;
            double ThetaA = notional * sTheta * premoFactor * notionalFactor;
            double Dv01_BaseA = notional * sDv01For * premoFactor * notionalFactor;
            double Dv01_TermsA = notional * sDv01Dom * premoFactor * notionalFactor;
            double premoFromMid = notional * priceFromMid * premoFactor * -1 * notionalFactor;

            //call smile greeks {smileVega,smileRega25,smileRega10,smileSega25,smileSega10,smileDelta }; function returns amounts times notional 
            object [] smileGreeksOuput = smileGreeks(autoSt, atmVol, rr, fly, rrMult, smileFlyMult, spot, dayStart, autoExp, dfDomExp, dfForExp, dfDomDel, dfForDel, premoInc, smileFactor, pC, premoString, notional);

            double sVega = Convert.ToDouble(smileGreeksOuput[0]);
            double rega25 = Convert.ToDouble(smileGreeksOuput[1]);
            double rega10 = Convert.ToDouble(smileGreeksOuput[2]);
            double sega25 = Convert.ToDouble(smileGreeksOuput[3]);
            double sega10 = Convert.ToDouble(smileGreeksOuput[4]);
            double sDelta = Convert.ToDouble(smileGreeksOuput[5]);
            double sGamma = Convert.ToDouble(smileGreeksOuput[7]);


            //set spreads
            //spreads are loaded into datatables on market data update . They are saved into dataset spreadset
            //get row of currency in datatable

            List<string> crossList = spreadSet.Tables[0].AsEnumerable().Select(x => x[0].ToString()).ToList();
            int rowInt = crossList.IndexOf(ccyPair);

            //load pillar daycounts from datatable columns then lookup daycount of current option
            var dayArr = spreadSet.Tables[0].Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray();
            int dayCol = 0;
            double volSpread = 0;

            for (int q = 2; q < dayArr.Count(); q++)
            {
                if (dayCount >= Convert.ToDouble(dayArr[q]))
                {
                    dayCol = q;
                }
            }

            //finaly  check delta to get spread from correct datatable. There are 3 (atm, 25delta, 10delta spreads)
            double deltaAdj = Math.Abs(delta);

            if (autoSt >= outRight)
            {
                double[] x = FXOpts(spot, dayStart, autoExp, autoSt, vol, dfDomExp, dfForExp, dfDomDel, dfForDel, "c");
                deltaAdj = x[1];

            }
            else
            {
                double[] x = FXOpts(spot, dayStart, autoExp, autoSt, vol, dfDomExp, dfForExp, dfDomDel, dfForDel, "p");
                deltaAdj = x[1];
            }


            deltaAdj = Math.Abs(deltaAdj);


            if (deltaAdj <= 0.60)
            {
                volSpread = Convert.ToDouble(spreadSet.Tables[0].Rows[rowInt][dayCol]);
            }

            if (deltaAdj <= 0.26)
            {
                volSpread = Convert.ToDouble(spreadSet.Tables[1].Rows[rowInt][dayCol]);
            }

            if (deltaAdj <= 0.11)
            {
                volSpread = Convert.ToDouble(spreadSet.Tables[2].Rows[rowInt][dayCol]);
            }

            volSpread = volSpread / 2;
            double bidVol = Math.Round((vol * 100 - volSpread) / .05, 0) * .05;
            double askVol = Math.Round((vol * 100 + volSpread) / .05, 0) * .05;
            string bidOffer = bidVol.ToString("0.00") + " / " + askVol.ToString("0.00");




            pricer.Rows[Expiry_DaysR][curCol] = dayCount.ToString("0.00");
            pricer.Rows[Deliver_DaysR][curCol] = delDayCount.ToString("0.00");
            pricer.Rows[Basis_BaseR][curCol] = basisB;
            pricer.Rows[Basis_TermsR][curCol] = basisT;
            pricer.Rows[Premium_IncludedR][curCol] = volType.ToString("0.00");
            pricer.Rows[NotionalR][curCol] = notional;
            pricer.Rows[TodayR][curCol] = dayStart.ToString("ddd-dd-MMM-yy");
            pricer.Rows[SpotdateR][curCol] = sptDate.ToString("ddd-dd-MMM-yy");
            pricer.Rows[BbgSourceR][curCol] = bbgSource;
            pricer.Rows[VegaR][curCol] = vega.ToString("0.0000");
            pricer.Rows[sVegaR][curCol] = sVega.ToString("#,##0"); 
            pricer.Rows[sDeltaR][curCol] = sDelta.ToString("#,##0");
            pricer.Rows[Sega25R][curCol] = sega25.ToString("#,##0");
            pricer.Rows[Rega25R][curCol] = rega25.ToString("#,##0");
            pricer.Rows[Sega10R][curCol] = sega10.ToString("#,##0");
            pricer.Rows[Rega10R][curCol] = rega10.ToString("#,##0");
            pricer.Rows[PriceR][curCol] = premium.ToString("0.0000");
            pricer.Rows[Fwd_PriceR][curCol] = fpremium.ToString("0.0000");
            pricer.Rows[BpsFromMidR][curCol] = priceFromMid.ToString("0.0000");
            pricer.Rows[PremiumFromMidR][curCol] = premoFromMid.ToString("#,##0");
            pricer.Rows[Bps_to_AtmR][curCol] = premoSmileVSflat.ToString("0.0000");
            pricer.Rows[Vol_Spread_to_AtmR][curCol] = smileVolSpread.ToString("0.00%");
            pricer.Rows[Spot_DeltaR][curCol] = delta.ToString("0.00%");
            pricer.Rows[Fwd_DeltaR][curCol] = fdelta.ToString("0.00%");
            pricer.Rows[Swap_PtsR][curCol] = fwdPts.ToString("0.00");
            pricer.Rows[FwdR][curCol] = outRight.ToString("0.0000");
            pricer.Rows[ATM_VOLR][curCol] = atmVol.ToString("0.00%");
            pricer.Rows[RRR][curCol] = rr.ToString("0.00%");
            pricer.Rows[FLYR][curCol] = fly.ToString("0.00%");
            pricer.Rows[VolR][curCol] = vol.ToString("0.00%");
            pricer.Rows[AutoStrikeR][curCol] = autoSt.ToString("0.0000");
            pricer.Rows[ExpiryDateR][curCol] = autoExp.ToString("ddd-dd-MMM-yy");
            pricer.Rows[DeliveryDateR][curCol] = delDate.ToString("ddd-dd-MMM-yy");
            pricer.Rows[Depo_BaseR][curCol] = forDepo.ToString("0.0000%");
            pricer.Rows[Depo_TermsR][curCol] = domDepo.ToString("0.0000%");
            pricer.Rows[SystemVolR][curCol] = systemVol.ToString("0.00%");
            pricer.Rows[Bid_OfferR][curCol] = bidOffer;
            pricer.Rows[Premium_TypeR][curCol] = premoString;
            pricer.Rows[BreakEvenR][curCol] = breakEven.ToString("0.0000");
            pricer.Rows[PremiumAR][curCol] = premiumA.ToString("#,##0");
            pricer.Rows[DeltaAR][curCol] = DeltaA.ToString("#,##0");
            pricer.Rows[sGammaR][curCol] = sGamma.ToString("#,##0");
            pricer.Rows[GammaAR][curCol] = GammaA.ToString("#,##0");
            pricer.Rows[VegaAR][curCol] = VegaA.ToString("#,##0");
            pricer.Rows[VannaAR][curCol] = VannaA.ToString("#,##0");
            pricer.Rows[VolgaAR][curCol] = VolgaA.ToString("#,##0");
            pricer.Rows[ThetaAR][curCol] = ThetaA.ToString("#,##0");
            pricer.Rows[DV01_BaseAR][curCol] = Dv01_BaseA.ToString("#,##0");
            pricer.Rows[DV01_TermsAR][curCol] = Dv01_TermsA.ToString("#,##0");

            //creates datatable for option simulation
            // add scenarios

            List<double> spotScen = new List<double> { -.2, -.18, -.16, -.12, -.1, -.08, -.06, -.04, -.03, -.01, 0, .01, .02, .03, .04, .06, .08, .1, .12, .16, .18, .2 };

            DataTable tSim = new DataTable();

            tSim.Columns.Add("Output");

            foreach (double d in spotScen)
            {
                tSim.Columns.Add(d.ToString("0.00%"));
            }

            int i = 0;
            tSim.Rows.Add("Spot");
            tSim.Rows.Add("Vol");
            tSim.Rows.Add("sDelta");
            tSim.Rows.Add("SDeltaCCent");
            tSim.Rows.Add("sDeltaCCol");
            tSim.Rows.Add("sVega");
            tSim.Rows.Add("Premium");
            tSim.Rows.Add("ChangePremium");
            tSim.Rows.Add("DetlaPnl");
            tSim.Rows.Add("TotalPnl");
            //col 0 is for row headers 

            List<string> rName = tSim.AsEnumerable().Select(x => x[0].ToString()).ToList();
            int rSpot = rName.IndexOf("Spot");
            int rVol = rName.IndexOf("Vol");
            int rDelta = rName.IndexOf("sDelta");
            int rDeltaC = rName.IndexOf("SDeltaCCent");
            int rDeltaCol = rName.IndexOf("sDeltaCCol");
            int rVega = rName.IndexOf("sVega");
            int rPremo = rName.IndexOf("Premium");
            int rPremoC = rName.IndexOf("ChangePremium");
            int rDeltaPnl = rName.IndexOf("DetlaPnl");
            int rTotalPnl = rName.IndexOf("TotalPnl");


            foreach (DataColumn col in tSim.Columns)
            {

                if (col.Ordinal != 0)
                {
                    //gets vols for smile each spot and vega shift 

                    double spotSim = spot * (1 + spotScen[i]);
                    object [] smileGreeksSc = smileGreeks(autoSt, atmVol, rr, fly, rrMult, smileFlyMult, spotSim, dayStart, autoExp, dfDomExp, dfForExp, dfDomDel, dfForDel, premoInc, smileFactor, pC, premoString, notional);
                  //   {smileVega,smileRega25,smileRega10,smileSega25,smileSega10,smileDelta };
                   

                    //calc smilevega
                    double smileVega = Convert.ToDouble(smileGreeksSc[0]);
                    double smileDelta = Convert.ToDouble(smileGreeksSc[5]);
                    double cashPremo = Convert.ToDouble(smileGreeksSc[6]);
                    //add to datatable
                    tSim.Rows[rSpot][col] = spotSim.ToString("0.0000");
                 //   tSim.Rows[rVol][col] = volSim.ToString("0.00%");
                    tSim.Rows[rDelta][col] = smileDelta.ToString("#,##0");
                    tSim.Rows[rVega][col] = (smileVega).ToString("#,##0");
                    tSim.Rows[rPremo][col] = (cashPremo).ToString("#,##0");
                    i++;
                }
            }


            // calc some derives - detla from center  col in row 3 and change in detla at each col (gamma) in row 4
            foreach (DataColumn col in tSim.Columns)
            {
                if (col.Ordinal != 0)
                {

                    tSim.Rows[rDeltaC][col] = (Convert.ToDouble(tSim.Rows[rDelta][col]) - Convert.ToDouble(tSim.Rows[rDelta][11])).ToString("#,##0");

                    double premoChange = Convert.ToDouble(tSim.Rows[rPremo][col]) - Convert.ToDouble(tSim.Rows[rPremo][11]);

                    tSim.Rows[rPremoC][col] = premoChange.ToString("#,##0");

                    double currentSpot = Convert.ToDouble(tSim.Rows[rSpot][11]);
                    double colSpot = Convert.ToDouble(tSim.Rows[rSpot][col]);
                    double spotMove = 0;

                    if (volType == 1)
                    {
                        spotMove = (colSpot - currentSpot) / colSpot;
                    }
                    else
                    {
                        spotMove = colSpot - currentSpot;
                    }


                    double currentDelta = Convert.ToDouble(tSim.Rows[rDelta][11]);
                    double deltaPnl = currentDelta * spotMove * -1;

                    tSim.Rows[rDeltaPnl][col] = deltaPnl.ToString("#,##0"); ;
                    tSim.Rows[rTotalPnl][col] = (deltaPnl + premoChange).ToString("#,##0");

                    if (col.Ordinal > 11)
                    {

                        tSim.Rows[rDeltaCol][col] = (Convert.ToDouble(tSim.Rows[rDelta][col]) - Convert.ToDouble(tSim.Rows[rDelta][col.Ordinal - 1])).ToString("#,##0");

                    }

                    if (col.Ordinal < 11)
                    {

                        tSim.Rows[rDeltaCol][col] = (Convert.ToDouble(tSim.Rows[rDelta][col]) - Convert.ToDouble(tSim.Rows[rDelta][col.Ordinal + 1])).ToString("#,##0");
                    }

                    if (col.Ordinal == 11)
                    {
                        tSim.Rows[rDeltaC][col] = 0;
                        tSim.Rows[rDeltaPnl][col] = 0;
                    }
                }

            }


            tSim.Rows[rPremoC].Delete();
            rName = tSim.AsEnumerable().Select(x => x[0].ToString()).ToList();
            rDeltaPnl = rName.IndexOf("DetlaPnl");

            tSim.Rows[rDeltaPnl].Delete();


            //add to dataset 
            tSim.TableName = curCol.ToString();

            if (tradeSim.Tables.Contains(curCol.ToString()))
            {
                tradeSim.Tables.Remove(curCol.ToString());
            }

            tradeSim.Tables.Add(tSim);
        }

        private int[] dayCountBasis(string ccyPair)
        {
            int[] retVal = null;
            string baseCcy = ccyPair.Substring(0, 3);
            string termsCcy = ccyPair.Substring(3, 3);
            List<string> ccyName = ccyDets.AsEnumerable().Select(x => x[0].ToString().ToUpper().Trim()).ToList();
            int c1 = ccyName.IndexOf(baseCcy);
            int c2 = ccyName.IndexOf(termsCcy);
            int basisB = Convert.ToInt32(ccyDets.Rows[c1]["DayCountBasis"]);
            int basisT = Convert.ToInt32(ccyDets.Rows[c2]["DayCountBasis"]);
            retVal = new int[] { basisB, basisT };
            return retVal;

        }

        private object[] crossDtData(string ccyPair)
        {
            object[] retVal = null;

            List<string> tempList = crosses.AsEnumerable().Select(x => x[0].ToString()).ToList();
            int i = tempList.IndexOf(ccyPair);
            int volSurfType = Convert.ToInt32(crosses.Rows[i].ItemArray[5]);
            double factor = Convert.ToDouble(crosses.Rows[i].ItemArray[4]);
            string bbgSource = crosses.Rows[i].ItemArray[3].ToString();
            retVal = new object[] { volSurfType, factor, bbgSource };
            return retVal;
        }

        private double[] premoConventions(string premoString, double spot, double strike)
        {
            double[] retval = null;
            double premoConversion = 0;
            double premoFactor = 0;
            double notionalFactor = 0;
            double premoIncDeltSolve = 0;

            List<string> returnType = new List<string>(new string[] { "Base %", "Terms %", "Base Pips", "Terms Pips" });

            if (premoString == returnType[0])
            {
                premoConversion = 1 / spot * 100;
                premoFactor = 10000;
                notionalFactor = 1;
                premoIncDeltSolve = 1;
                retval = new double[] { premoConversion, premoFactor, notionalFactor, premoIncDeltSolve };
                return retval;
            }

            if (premoString == returnType[1])
            {
                premoConversion = 1 / strike * 100;
                premoFactor = 10000;
                notionalFactor = spot;
                premoIncDeltSolve = 0;
                retval = new double[] { premoConversion, premoFactor, notionalFactor, premoIncDeltSolve };
                return retval;
            }

            if (premoString == returnType[2])
            {
                premoConversion = 1 / (spot * strike);
                premoFactor = 1000000;
                notionalFactor = spot;
                premoIncDeltSolve = 1;
                retval = new double[] { premoConversion, premoFactor, notionalFactor, premoIncDeltSolve };
                return retval;
            }

            if (premoString == returnType[3])
            {
                premoConversion = 1;
                premoFactor = 1000000;
                notionalFactor = 1;
                premoIncDeltSolve = 0;
                retval = new double[] { premoConversion, premoFactor, notionalFactor, premoIncDeltSolve };
                return retval;
            }

            return retval;


        }

        private double[] volBuilder(string ccyPair, double dayCount)
        {
            double[] retVal = null;

            //get fly, rr from pricerData displayed on main pricer screen 
            DataTable dt = pricingData.Tables[ccyPair];
            int iD = dt.Columns["DayCount"].Ordinal;
            int iR = dt.Columns["25DR"].Ordinal;
            int iF = dt.Columns["25D_BrokerFly"].Ordinal;
            int iW = dt.Columns["Forward"].Ordinal;//used to be wingcontrol - dont need it here but keeping the place 
            int iB = dt.Columns["BrokerFly_Multiplier"].Ordinal;
            int iA = dt.Columns["ATM"].Ordinal;
            int iB1 = dt.Columns["smileFly_Multiplier"].Ordinal;
            int iR1 = dt.Columns["RR_Multiplier"].Ordinal;


            List<double> iDays = dt.AsEnumerable().Select(x => Convert.ToDouble(x[iD])).ToList();
            List<double> iRr = dt.AsEnumerable().Select(x => Convert.ToDouble(convertPercent(x[iR].ToString()))).ToList();
            List<double> iFly = dt.AsEnumerable().Select(x => Convert.ToDouble(convertPercent(x[iF].ToString()))).ToList();
            List<double> iWing = dt.AsEnumerable().Select(x => Convert.ToDouble(x[iW])).ToList();
            List<double> iBfly = dt.AsEnumerable().Select(x => Convert.ToDouble(x[iB])).ToList();
            List<double> iATM = dt.AsEnumerable().Select(x => Convert.ToDouble(convertPercent(x[iA].ToString()))).ToList();
            List<double> ifly = dt.AsEnumerable().Select(x => Convert.ToDouble(x[iB1])).ToList();
            List<double> iRR = dt.AsEnumerable().Select(x => Convert.ToDouble(x[iR1])).ToList();

            //List<double> iDays = dt.AsEnumerable().Select(x => Convert.ToDouble(x[iD])).ToList();
            //List<double> iRr = dt.AsEnumerable().Select(x => Convert.ToDouble(convertPercent(x[iR].ToString()))).ToList();
            //List<double> iFly = dt.AsEnumerable().Select(x => Convert.ToDouble(x[iF])).ToList();
            //List<double> iWing = dt.AsEnumerable().Select(x => Convert.ToDouble(x[iW])).ToList();
            //List<double> iBfly = dt.AsEnumerable().Select(x => Convert.ToDouble(x[iB])).ToList();


            // for exiries longer than  atm vols are interpolated via cubic spline
            double atmVol = 0;

            if (dayCount > 370)
            {
                int n = iATM.Count;
                float[] xx = new float[n];
                float[] y = new float[n];

                for (int i = 0; i < n; i++)
                {
                    y[i] = Convert.ToSingle(iATM[i]);
                    xx[i] = Convert.ToSingle(iDays[i]);
                }


                int u = 1;

                float[] xs = new float[8000];

                for (int q = 0; q < 8000; q++)
                {
                    if (DateTime.Now.DayOfWeek == DayOfWeek.Friday)
                    {
                        u = 3;
                    }

                    if (DateTime.Now.DayOfWeek == DayOfWeek.Saturday)
                    {
                        u = 2;
                    }

                    xs[q] = u;
                    u++;
                }


                TestMySpline.CubicSpline spline = new TestMySpline.CubicSpline();
                float[] ys = spline.FitAndEval(xx, y, xs);

                atmVol = ys[Convert.ToInt16(dayCount)];

            }
            else
            {
                atmVol = Vol(ccyPair, dayCount); //comes from volfile which is updated from dvi
            }

            double rr = LinearInterp(iDays, iRr, dayCount);
            double fly = LinearInterp(iDays, iFly, dayCount); //brokerfly
            double wingControl = LinearInterp(iDays, iWing, dayCount);
            double brokerFlyMult = LinearInterp(iDays, iBfly, dayCount);
            double smileFlyMult = LinearInterp(iDays, ifly, dayCount);
            double rrMult = LinearInterp(iDays, iRR, dayCount);

            retVal = new double[] { atmVol, rr, fly, wingControl, brokerFlyMult, smileFlyMult, rrMult };
            return retVal;

        }

        private double[] rateBuilder(string ccyPair, double dayCount, double delDayCount, double spot, double factor, int basisB, int basisT)
        {

            //returns depos,fwdpts and outright fwd

            double[] retVal = null;

            DataTable dtRateTile = marketData.Tables[ccyPair];
            int jD = dtRateTile.Columns["DayCount"].Ordinal;
            int jB = dtRateTile.Columns["Base"].Ordinal;
            int jP = dtRateTile.Columns["MidPts"].Ordinal;

            List<double> jDays = dtRateTile.AsEnumerable().Select(x => Convert.ToDouble(x[jD])).ToList();
            List<double> jBase = dtRateTile.AsEnumerable().Select(x => Convert.ToDouble(x[jB])).ToList();
            List<double> jPts = dtRateTile.AsEnumerable().Select(x => Convert.ToDouble(x[jP])).ToList();

            double fwdPts = FwdPtsPerDay(jPts, jDays, delDayCount);
            double outRight = spot + (fwdPts / factor);
            double forDepo;
            double domDepo;

            if (fwdPts == 0)
            {
                forDepo = 0;
                domDepo = 0;
            }
            else
            {
                forDepo = LinearInterp(jDays, jBase, delDayCount);
                domDepo = SolveTermsDepo(outRight, spot, delDayCount, forDepo, basisB, basisT);
            }

            retVal = new double[] { fwdPts, outRight, forDepo, domDepo };
            return retVal;

        }

        private  bool sameSign(double num1,double num2)
        {return num1>=0&&num2>=0||num1<0&&num2<0;}

        private double reCalibrateWingControl(double atmVol, double rr, double fly, double spot, DateTime tod, DateTime expiryDate, double dfDom, double dfFor, int premoInc, double smileFlyMult, double rrMult, int upDn)
        {

            double wingControl = 1;
            double convergCrit = 0.0003;
            double err = 100;
            double maxRuns = 5000;
            int q = 1;
            double err1 = 0;
            int wingDirection = 1;
            double rnd = 0.0002;

            double smileFly25 = equivalentfly(spot, tod, expiryDate, wingControl * atmVol, atmVol, rr, fly, dfDom, dfFor, premoInc);

            double pVol = atmVol - 0.5 * rr + smileFly25;
            double cVol = atmVol + 0.5 * rr + smileFly25;
            double strike25P = FXStrikeVol(spot, tod, expiryDate, 0.25, pVol, dfDom, dfFor, "p", premoInc);
            double strike25C = FXStrikeVol(spot, tod, expiryDate, 0.25, cVol, dfDom, dfFor, "c", premoInc);
            double strikeAtm = FXATMStrike(spot, tod, expiryDate, atmVol, dfDom, dfFor, premoInc);
            double vol10 = 0;
            double strike10 = 0;
            double vol10Solve = 0;

            if (upDn == 1)
            {
                 vol10 = atmVol + 0.5 * (rr * rrMult) + (smileFly25 * smileFlyMult);
                 strike10 = FXStrikeVol(spot, tod, expiryDate, 0.1, vol10, dfDom, dfFor, "c", premoInc);
            }
            else
            {
                vol10 = atmVol - 0.5 * (rr * rrMult) + (smileFly25 * smileFlyMult);
                strike10 = FXStrikeVol(spot, tod, expiryDate, 0.1, vol10, dfDom, dfFor, "p", premoInc);
            }

            vol10 = Math.Round((vol10) / rnd, 0) * rnd;

            //interates through wingfactors to match broker fly with target in setup. 
            while (Math.Abs(err) >= convergCrit)
            {

                if ((q > maxRuns)) { break; }

                 vol10Solve = smileInterp(spot, tod, expiryDate, wingControl * atmVol, strike10, strike25P, pVol, strikeAtm, atmVol, strike25C, cVol, dfDom, dfFor, dfDom, dfFor);
                // vol10Solve = Math.Round((vol10Solve) / rnd, 0) * rnd;

                err1 = err;
                err = vol10Solve - vol10;

                if (q == 1)
                {
                    if (err < 0)
                    {
                        wingControl = wingControl + 0.001;
                        wingDirection = 1;
                    }
                    else
                    {
                        wingControl = wingControl - 0.001;
                        wingDirection = -1;
                    }

                }


                if (q > 1)
                {
                    if (Math.Abs(err) > Math.Abs(err1))
                    {
                        wingDirection = wingDirection * -1;
                    }

                    wingControl = wingControl + 0.001 * wingDirection;
                }


                q = q + 1;

            }

            string callPut = "put";

            if (upDn == 1){
                callPut = "call";
            }

            if (Math.Abs(vol10Solve - vol10) > convergCrit) { MessageBox.Show("Algo can't solve wing factor - please check inputs for: " + " " + expiryDate.ToString("ddd-dd-MMM-yy") + " " + callPut + " " + strike10.ToString("0.0000") + " modelVol:" + vol10Solve.ToString("0.0000%") + " vs ratioVol:" + vol10.ToString("0.0000%")); }

            return wingControl;
        }

        private void setSurfaceNew(string curr)
        {

            DataTable surface = surfaceDt();

            string ccy1 = curr.Substring(0, 3);
            string ccy2 = curr.Substring(3, 3);
            string cross = ccy1 + ccy2;

            double atmVol = 0;
            double rr = 0;
            double fly = 0;
            double smileFly25 = 0;
            double brokerFly25 = 0;
            double pVol = 0;
            double cVol = 0;
            double strike25P = 0;
            double strike25C = 0;
            double strikeAtm = 0;
            double strike10P = 0;
            double strike10C = 0;
            double vol10P = 0;
            double vol10C = 0;
  

            double rr10 = 0;
            double smileFly10 = 0;
            double brokerFly10 = 0;
            double rrMult = 0;
            double brokerFlyMult = 0;
            double smileFlyMult = 0;

            DateTime expiryDate = new DateTime();
            DateTime deliveryDate = new DateTime();
            double dayCount = 0;
            double delDayCount = 0;

            double pts = 0;
            double fwd = 0;
            double forDepo = 0;
            double domDepo = 0;
            double dfFor = 0;
            double dfDom = 0;


           

    
           

            List<string> crossName = crosses.AsEnumerable().Select(x => x[0].ToString()).ToList();
            List<string> ccyForDom = ccyDets.AsEnumerable().Select(x => x[0].ToString()).ToList();


            //dataset smileMult contains fly and rr multipliers
            DataTable sMult = smileMult.Tables[cross];
            if (sMult == null)
            {
                sMult = smileMult.Tables[0];
                MessageBox.Show("Please Update Smile Multiplers");
            }


            int i = crossName.IndexOf(cross);
            int cF = ccyForDom.IndexOf(ccy1);
            int cD = ccyForDom.IndexOf(ccy2);

            double factor = Convert.ToDouble(crosses.Rows[i].ItemArray[4]);
            Int16 premoInc = Convert.ToInt16(crosses.Rows[i].ItemArray[5]);
            if (premoInc == 2) { premoInc = 0; } //uses delta type from old version if detlatype =1 then that is premo included else 2 = no premo change to 0 

            double smileFactor = factor;
            if (cross == "USDRUB" || cross == "EURRUB" || cross == "USDTRY") { smileFactor = 100; }


            int dayCountBasis1 = Convert.ToInt16(ccyDets.Rows[cF].ItemArray[2]);
            int dayCountBasis2 = Convert.ToInt16(ccyDets.Rows[cD].ItemArray[2]);


            string depo = crosses.Rows[i].ItemArray[1].ToString();
            string solveBase = crosses.Rows[i].ItemArray[2].ToString();

            List<string> mat = new List<string>();
            mat.Add("1d");
            mat.Add("1w");
            mat.Add("2w");
            mat.Add("1m");
            mat.Add("2m");
            mat.Add("3m");
            mat.Add("6m");
            mat.Add("9m");
            mat.Add("1y");
            mat.Add("2y");
            mat.Add("3y");
            mat.Add("4y");
            mat.Add("5y");

            DateTime tod = DateTime.Today;
            DateTime spotDateCross = SpotDate(tod, ccy1, ccy2);

            DataTable dt = marketData.Tables[cross];
            List<string> term = dt.AsEnumerable().Select(x => Convert.ToString(x[0])).ToList();
            List<double> swapPts = dt.AsEnumerable().Select(x => Convert.ToDouble(x[3])).ToList();
            List<double> swapDayCount = dt.AsEnumerable().Select(x => Convert.ToDouble(x[2])).ToList();
            List<double> baseDepo = dt.AsEnumerable().Select(x => Convert.ToDouble(x[5])).ToList();

            DataTable st = fwds.d_data;
            List<string> ccyPair = fwds.d_data.AsEnumerable().Select(x => x[0].ToString().Substring(0, 6)).ToList();
            int ii = ccyPair.IndexOf(cross);
            double spot = Convert.ToDouble(st.Rows[ii]["PX_MID"]);

            int flyInt = 0;

            foreach (string s in mat)
            {
                expiryDate = AutoExpiryDate(s, tod, "USD", ccy1, ccy2);

                if (expiryDate.DayOfWeek == DayOfWeek.Saturday || expiryDate.DayOfWeek == DayOfWeek.Sunday)
                {
                    expiryDate = AddWorkdays(expiryDate, 1);
                }

                deliveryDate = SpotDate(expiryDate, ccy1, ccy2);
                dayCount = (expiryDate - tod).TotalDays;
                delDayCount = (deliveryDate - spotDateCross).TotalDays;

                pts = FwdPtsPerDay(swapPts, swapDayCount, delDayCount);
                fwd = spot + pts / factor;

                if (pts == 0)
                {
                    forDepo = 0;
                    domDepo = 0;

                }
                else
                {
                    forDepo = LinearInterp(swapDayCount, baseDepo, delDayCount);
                    domDepo = SolveTermsDepo(fwd, spot, delDayCount, forDepo, dayCountBasis1, dayCountBasis2);
                }


                dfFor = DiscountFactor(forDepo, Convert.ToInt16((deliveryDate - spotDateCross).TotalDays), dayCountBasis1);
                dfDom = DiscountFactor(domDepo, Convert.ToInt16((deliveryDate - spotDateCross).TotalDays), dayCountBasis2);

                if (s == "2y" || s == "3y" || s == "4y" || s == "5y")
                {
                    double twoYearSpread = smile(cross, dayCount, s) / 100;
                    expiryDate = AutoExpiryDate(s, tod, "USD", ccy1, ccy2);
                    DateTime expiryDate1y = AutoExpiryDate("1y", tod, "USD", ccy1, ccy2);
                    double dCountOne = (expiryDate1y - tod).TotalDays;
                    double oneY = Vol(cross, dCountOne);


                    atmVol = oneY + twoYearSpread;
                  //  targetFlyMult = Convert.ToDouble(sMult.Rows[flyInt]["flyMultiplier"]);

                    //targetFlyMult = 3.6;
                }

                else
                {
                    atmVol = Vol(cross, dayCount);
                //    targetFlyMult = Convert.ToDouble(sMult.Rows[flyInt]["flyMultiplier"]);

                }

                rr = smile(cross, dayCount, "rr") / 100;
                fly = smile(cross, dayCount, "fly") / 100; //broker fly from bbg  
                smileFlyMult = Convert.ToDouble(sMult.Rows[flyInt]["flyMultiplier"]);
                rrMult = Convert.ToDouble(sMult.Rows[flyInt]["rrMultiplier"]);


                brokerFly25 = fly;

              //  double wingPut = reCalibrateWingControl(atmVol, rr, fly, spot, tod, expiryDate, dfDom, dfFor, premoInc, smileFlyMult, rrMult,0);
               // double wingCall = reCalibrateWingControl(atmVol, rr, fly, spot, tod, expiryDate, dfDom, dfFor, premoInc, smileFlyMult, rrMult,1);

                smileFly25 = equivalentfly(spot, tod, expiryDate, 1 * atmVol, atmVol, rr, fly, dfDom, dfFor, premoInc);

                pVol = atmVol - 0.5 * rr + smileFly25;
                cVol = atmVol + 0.5 * rr + smileFly25;
                strike25P = FXStrikeVol(spot, tod, expiryDate, 0.25, pVol, dfDom, dfFor, "p", premoInc);
                strike25C = FXStrikeVol(spot, tod, expiryDate, 0.25, cVol, dfDom, dfFor, "c", premoInc);
                strikeAtm = FXATMStrike(spot, tod, expiryDate, atmVol, dfDom, dfFor, premoInc);
                vol10P = atmVol + (smileFly25 * smileFlyMult) - (rr / 2 * rrMult);
                strike10P = FXStrikeVol(spot, tod, expiryDate, 0.1, vol10P, dfDom, dfFor, "p", premoInc);
               // vol10PSolve = smileInterp(spot, tod, expiryDate, wingPut * atmVol, strike10P, strike25P, pVol, strikeAtm, atmVol, strike25C, cVol, dfDom, dfFor, dfDom, dfFor);
                vol10C = atmVol + 0.5 * (rr * rrMult) + (smileFly25 * smileFlyMult);
                strike10C = FXStrikeVol(spot, tod, expiryDate, 0.1, vol10C, dfDom, dfFor, "c", premoInc);
               // vol10CSolve = smileInterp(spot, tod, expiryDate, wingCall * atmVol, strike10C, strike25P, pVol, strikeAtm, atmVol, strike25C, cVol, dfDom, dfFor, dfDom, dfFor);

                rr10 = vol10C - vol10P;
                smileFly10 = (vol10C + vol10P) / 2 - atmVol;
                brokerFly10 = marketfly(spot, tod, expiryDate, .1, atmVol, atmVol, rr, fly, dfDom, dfFor, premoInc, smileFlyMult, rrMult, smileFactor);
                rrMult = rr10 / rr;
                brokerFlyMult = brokerFly10 / fly;
                smileFlyMult = smileFly10 / smileFly25;



                surface.Rows.Add(new Object[] { dayCount, s, expiryDate.ToString("ddd-dd-MMM-yy"), deliveryDate.ToString("ddd-dd-MMM-yy"), atmVol.ToString("0.00%"), rr10.ToString("0.00%"), rr.ToString("0.00%"), brokerFly25.ToString("0.00%"), brokerFly10.ToString("0.00%"), rrMult.ToString("0.00"), smileFlyMult.ToString("0.00"), brokerFlyMult.ToString("0.00"), fwd.ToString("0.0000"), pts.ToString("0.00"), domDepo.ToString("0.00%"), forDepo.ToString("0.00%"), dfDom.ToString("0.0000"), dfFor.ToString("0.0000"), smileFly25.ToString("0.00%"), smileFly10.ToString("0.00%"), vol10P.ToString("0.00%"),pVol.ToString("0.00%"), atmVol.ToString("0.00%"), cVol.ToString("0.00%"),vol10C.ToString("0.00%"), strike10P.ToString("0.0000"), strike25P.ToString("0.0000"), strikeAtm.ToString("0.0000"), strike25C.ToString("0.0000"), strike10C.ToString("0.0000") });

             
                flyInt = flyInt + 1;
            }



            surface.TableName = cross;

            //check to see if there is a table for current ccy - if so delete old table and add new data else just at new table
            if (pricingData.Tables.Contains(cross))
            {
                pricingData.Tables.Remove(cross);

                pricingData.Tables.Add(surface);
            }
            else
            {
                pricingData.Tables.Add(surface);
            }


            dataGridView8.DataSource = pricingData.Tables[cross];
            formatSurfaceView();
        }

        private void refreshSurface(string curr)
        {

            //this void will refresh current surface with fresh marketData as well as user inputs for rr, fly and fly multiplier
            // string curr = ((DataTable)dataGridView8.DataSource).TableName;
            DataSet ds = pricingData;
            DataTable dt_surf = ds.Tables[curr];
            DataTable surface = new DataTable();
            surface = dt_surf.Clone();

            string ccy1 = curr.Substring(0, 3);
            string ccy2 = curr.Substring(3, 3);
            string cross = ccy1 + ccy2;

            double atmVol = 0;
            double rr = 0;
            double fly = 0;
            double smileFly25 = 0;
            double pVol = 0;
            double cVol = 0;
            double strike25P = 0;
            double strike25C = 0;
            double strikeAtm = 0;
            double strike10P = 0;
            double strike10C = 0;
            double vol10P = 0;
            double vol10C = 0;
      
            double rr10 = 0;
            double smileFly10 = 0;
            double brokerFly10 = 0;
            double rrMult = 0;
            double brokerFlyMult = 0;
            double smileFlyMult = 0;

            DateTime expiryDate = new DateTime();
            DateTime deliveryDate = new DateTime();
            double dayCount = 0;
            double delDayCount = 0;
            double pts = 0;
            double fwd = 0;
            double forDepo = 0;
            double domDepo = 0;
            double dfFor = 0;
            double dfDom = 0;
        





            List<string> crossName = crosses.AsEnumerable().Select(x => x[0].ToString()).ToList();
            List<string> ccyForDom = ccyDets.AsEnumerable().Select(x => x[0].ToString()).ToList();

            int i = crossName.IndexOf(cross);
            int cF = ccyForDom.IndexOf(ccy1);
            int cD = ccyForDom.IndexOf(ccy2);

            double factor = Convert.ToDouble(crosses.Rows[i].ItemArray[4]);
            Int16 premoInc = Convert.ToInt16(crosses.Rows[i].ItemArray[5]);
            if (premoInc == 2) { premoInc = 0; } //uses delta type from old version if detlatype =1 then that is premo included else 2 = no premo change to 0 
            double smileFactor = factor;

            if (cross == "USDRUB" || cross == "EURRUB" || cross == "USDTRY") { smileFactor = 100; }

            int dayCountBasis1 = Convert.ToInt16(ccyDets.Rows[cF].ItemArray[2]);
            int dayCountBasis2 = Convert.ToInt16(ccyDets.Rows[cD].ItemArray[2]);


            string depo = crosses.Rows[i].ItemArray[1].ToString();
            string solveBase = crosses.Rows[i].ItemArray[2].ToString();

            DateTime tod = DateTime.Today;
            DateTime spotDateCross = SpotDate(tod, ccy1, ccy2);

            DataTable dt = marketData.Tables[cross];
            List<string> term = dt.AsEnumerable().Select(x => Convert.ToString(x[0])).ToList();
            List<double> swapPts = dt.AsEnumerable().Select(x => Convert.ToDouble(x[3])).ToList();
            List<double> swapDayCount = dt.AsEnumerable().Select(x => Convert.ToDouble(x[2])).ToList();
            List<double> baseDepo = dt.AsEnumerable().Select(x => Convert.ToDouble(x[5])).ToList();

            DataTable st = fwds.d_data;
            List<string> ccyPair = fwds.d_data.AsEnumerable().Select(x => x[0].ToString().Substring(0, 6)).ToList();
            int ii = ccyPair.IndexOf(cross);
            double spot = Convert.ToDouble(st.Rows[ii]["PX_MID"]);

            int flyInt = 0;

            foreach (DataRow row in dt_surf.Rows)
            {


                string s = row["Maturity"].ToString();

                expiryDate = AutoExpiryDate(s, tod, "USD", ccy1, ccy2);

                if (expiryDate.DayOfWeek == DayOfWeek.Saturday || expiryDate.DayOfWeek == DayOfWeek.Sunday)
                {
                    expiryDate = AddWorkdays(expiryDate, 1);
                }

                deliveryDate = SpotDate(expiryDate, ccy1, ccy2);

                dayCount = (expiryDate - tod).TotalDays;
                delDayCount = (deliveryDate - spotDateCross).TotalDays;

                pts = FwdPtsPerDay(swapPts, swapDayCount, delDayCount);
                fwd = spot + pts / factor;

                if (pts == 0)
                {
                    forDepo = 0;
                    domDepo = 0;

                }
                else
                {
                    forDepo = LinearInterp(swapDayCount, baseDepo, delDayCount);
                    domDepo = SolveTermsDepo(fwd, spot, delDayCount, forDepo, dayCountBasis1, dayCountBasis2);
                }


                dfFor = DiscountFactor(forDepo, Convert.ToInt16((deliveryDate - spotDateCross).TotalDays), dayCountBasis1);
                dfDom = DiscountFactor(domDepo, Convert.ToInt16((deliveryDate - spotDateCross).TotalDays), dayCountBasis2);

                if (s == "2y" || s == "3y" || s == "4y" || s == "5y")
                {
                    double twoYearSpread = smile(cross, dayCount, s) / 100;
                    expiryDate = AutoExpiryDate(s, tod, "USD", ccy1, ccy2);
                    DateTime expiryDate1y = AutoExpiryDate("1y", tod, "USD", ccy1, ccy2);
                    double dCountOne = (expiryDate1y - tod).TotalDays;
                    double oneY = Vol(cross, dCountOne);


                    atmVol = oneY + twoYearSpread;

                }
                else
                {
                    atmVol = Vol(cross, dayCount);
                }

                rr = convertPercent(row["25DR"].ToString());

                fly = convertPercent(row["25D_BrokerFly"].ToString());

                smileFlyMult = Convert.ToDouble(row["SmileFly_Multiplier"]);
                rrMult = Convert.ToDouble(row["RR_Multiplier"]);

                //  double wingPut = reCalibrateWingControl(atmVol, rr, fly, spot, tod, expiryDate, dfDom, dfFor, premoInc, smileFlyMult, rrMult, 0);
                // double wingCall = reCalibrateWingControl(atmVol, rr, fly, spot, tod, expiryDate, dfDom, dfFor, premoInc, smileFlyMult, rrMult, 1);

                smileFly25 = equivalentfly(spot, tod, expiryDate, 1 * atmVol, atmVol, rr, fly, dfDom, dfFor, premoInc);

                pVol = atmVol - 0.5 * rr + smileFly25;
                cVol = atmVol + 0.5 * rr + smileFly25;
                strike25P = FXStrikeVol(spot, tod, expiryDate, 0.25, pVol, dfDom, dfFor, "p", premoInc);
                strike25C = FXStrikeVol(spot, tod, expiryDate, 0.25, cVol, dfDom, dfFor, "c", premoInc);
                strikeAtm = FXATMStrike(spot, tod, expiryDate, atmVol, dfDom, dfFor, premoInc);
                vol10P = atmVol + (smileFly25 * smileFlyMult) - (rr / 2 * rrMult);
                strike10P = FXStrikeVol(spot, tod, expiryDate, 0.1, vol10P, dfDom, dfFor, "p", premoInc);
                //vol10PSolve = smileInterp(spot, tod, expiryDate, wingPut * atmVol, strike10P, strike25P, pVol, strikeAtm, atmVol, strike25C, cVol, dfDom, dfFor, dfDom, dfFor);
                vol10C = atmVol + 0.5 * (rr * rrMult) + (smileFly25 * smileFlyMult);
                strike10C = FXStrikeVol(spot, tod, expiryDate, 0.1, vol10C, dfDom, dfFor, "c", premoInc);
                //vol10CSolve = smileInterp(spot, tod, expiryDate, wingCall * atmVol, strike10C, strike25P, pVol, strikeAtm, atmVol, strike25C, cVol, dfDom, dfFor, dfDom, dfFor);

                rr10 = vol10C - vol10P;
                smileFly10 = (vol10C + vol10P) / 2 - atmVol;
                brokerFly10 = marketfly(spot, tod, expiryDate, .1, atmVol, atmVol, rr, fly, dfDom, dfFor, premoInc, smileFlyMult, rrMult, smileFactor);
                rrMult = rr10 / rr;
                brokerFlyMult = brokerFly10 / fly;
                smileFlyMult = smileFly10 / smileFly25;


                row["DayCount"] = dayCount;
                row["Maturity"] = s;
                row["ExpiryDate"] = expiryDate.ToString("ddd-dd-MMM-yy");
                row["DeliveryDate"] = deliveryDate.ToString("ddd-dd-MMM-yy");
                row["ATM"] = atmVol.ToString("0.00%");
                row["10DR"] = rr10.ToString("0.00%");
                row["25DR"] = rr.ToString("0.00%");
                row["25D_BrokerFly"] = fly.ToString("0.00%");
                row["10D_BrokerFly"] = brokerFly10.ToString("0.00%");
                row["RR_Multiplier"] = rrMult.ToString("0.00");
                row["BrokerFly_Multiplier"] = brokerFlyMult.ToString("0.00"); // brokerFlyMult.ToString("0.00");
                //row["WingCtrlDn"] = wingPut.ToString("0.0000");
                // row["WingCtrlUp"] = wingCall.ToString("0.0000");
                row["Forward"] = fwd.ToString("0.0000");
                row["Points"] = pts.ToString("0.00");
                row["DepoDom"] = domDepo.ToString("0.00%");
                row["DepoFor"] = forDepo.ToString("0.00%");
                row["DFDom"] = dfDom.ToString("0.0000");
                row["DFFor"] = dfFor.ToString("0.0000");
                row["25D_SmileFly"] = smileFly25.ToString("0.00%");
                row["10D_SmileFly"] = smileFly10.ToString("0.00%");
                row["SmileFly_Multiplier"] = smileFlyMult.ToString("0.00");
                row["25dPutVol"] = pVol.ToString("0.00%");
                row["ATMVol"] = atmVol.ToString("0.00%");
                row["25dCallVol"] = cVol.ToString("0.00%");
                row["10dPutStrike"] = strike10P.ToString("0.0000");
                row["25dPutStrike"] = strike25P.ToString("0.0000");
                row["ATMStrike"] = strikeAtm.ToString("0.0000");
                row["25dCallStrike"] = strike25C.ToString("0.0000");
                row["10dCallStrike"] = strike10C.ToString("0.0000");
                row["10dCallVol"] = vol10C.ToString("0.00%");
                row["10dPutVol"] = vol10P.ToString("0.00%");



             
                flyInt = flyInt + 1;
            }


            dataGridView8.DataSource = pricingData.Tables[cross];
            formatSurfaceView();
        }

        private void displaySurface(string ccyPair)
        {

            if (pricingData.Tables.Contains(ccyPair))
            {
                refreshSurface(ccyPair);
            }
            else
            {
                setSurfaceNew(ccyPair);
            }



        }

        private void marketMakeRun(string ccyPair)
        {
            //set spreads
            ccyPair = ((DataTable)dataGridView8.DataSource).TableName;
            DataSet ds = pricingData;
            DataTable dt = ds.Tables[ccyPair];
            List<string> marketRun = new List<string>();

            DataSet brs = brokerRun;
            DataTable br = new DataTable();
            br.Columns.Add("Maturity");
            br.Columns.Add("BID");
            br.Columns.Add("OFFER");

            string cPaste = "";



            foreach (DataRow row in dt.Rows)
            {
                double vol = Convert.ToDouble(convertPercent(row["ATM"].ToString()));
                double dayCount = Convert.ToDouble(row["DayCount"]);
                string mat = row["Maturity"].ToString();

                List<string> crossList = spreadSet.Tables[0].AsEnumerable().Select(x => x[0].ToString()).ToList();
                int rowInt = crossList.IndexOf(ccyPair);

                //load pillar daycounts from datatable columns then lookup daycount of current option
                var dayArr = spreadSet.Tables[0].Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray();
                int dayCol = 0;
                double volSpread = 0;

                for (int q = 2; q < dayArr.Count(); q++)
                {
                    if (dayCount >= Convert.ToDouble(dayArr[q]))
                    {
                        dayCol = q;
                    }
                }


                volSpread = Convert.ToDouble(spreadSet.Tables[0].Rows[rowInt][dayCol]);
                volSpread = volSpread / 2;

                double bidVol = Math.Round((vol * 100 - volSpread) / .05, 0) * .05;
                double askVol = Math.Round((vol * 100 + volSpread) / .05, 0) * .05;
                string bidOffer = mat + "  " + bidVol.ToString("0.00") + " / " + askVol.ToString("0.00");
                br.Rows.Add(mat, bidVol.ToString("0.00"), askVol.ToString("0.00"));


                if (mat != "1d" && mat != "2w" && mat != "9m" && mat != "2y")
                {
                    marketRun.Add(bidOffer);
                }


            }



            //now copy to clipboard
            foreach (string s in marketRun)
            {
                cPaste = cPaste + s + Environment.NewLine;
            }

            br.TableName = ccyPair;

            if (brs.Tables.Contains(ccyPair))
            {
                brs.Tables.Remove(ccyPair);
            }
            brs.Tables.Add(br);

            Clipboard.SetText(cPaste);
            MessageBox.Show(cPaste);
            comboBox1.Text = ccyPair;
            comboBox1_SelectedIndexChanged(comboBox1, new EventArgs());
        }

        private void murexSmile(string ccyPair)
        {
            //set spreads
            ccyPair = ((DataTable)dataGridView8.DataSource).TableName;
            DataSet ds = pricingData;
            DataTable dt = ds.Tables[ccyPair];
            DataTable temp = new DataTable();
            temp.Columns.Add();
            temp.Columns.Add();
            temp.Columns.Add();
            temp.Columns.Add();

            foreach (DataRow row in dt.Rows)
            {

                double rr10 = Convert.ToDouble(convertPercent(row["10DR"].ToString())) * 100;
                double rr25 = Convert.ToDouble(convertPercent(row["25DR"].ToString())) * 100;
                double sf25 = Convert.ToDouble(convertPercent(row["25D_SmileFly"].ToString())) * 100;
                double sf10 = Convert.ToDouble(convertPercent(row["10D_SmileFly"].ToString())) * 100;

                temp.Rows.Add(rr10.ToString("0.00"), rr25.ToString("0.00"), sf25.ToString("0.00"), sf10.ToString("0.00"));

            }

            StringBuilder Output = new StringBuilder();
            //Generate Cell Value Data
            foreach (DataRow Row in temp.Rows)
            {
                for (int i = 0; i < Row.ItemArray.Length; i++)
                {
                    //Handling the last cell of the line.
                    if (i == (Row.ItemArray.Length - 1))
                    {

                        Output.Append(Row.ItemArray[i].ToString() + "\n");
                    }
                    else
                    {

                        Output.Append(Row.ItemArray[i].ToString() + "\t");
                    }
                }
            }

            Clipboard.SetText(Output.ToString());
            MessageBox.Show(Output.ToString());
        }

        private void setSmileMultFirstTime(string ccyPair)
        {

            //this was only needed to create xml file..
            if (smileMult == null)
                smileMult = new DataSet();

            if (smileMult != null)
            {
                smileMult.Reset();
            }

            DataSet ds = smileMult;

            foreach (DataRow r in crosses.Rows)
            {
                ccyPair = r["Cross"].ToString();


                DataTable sMult = new DataTable();

                sMult.Columns.Add("Maturity");
                sMult.Columns.Add("rrMultiplier");
                sMult.Columns.Add("flyMultiplier");

                sMult.Clear();

                List<string> mat = new List<string>();
                mat.Add("1d");
                mat.Add("1w");
                mat.Add("2w");
                mat.Add("1m");
                mat.Add("2m");
                mat.Add("3m");
                mat.Add("6m");
                mat.Add("9m");
                mat.Add("1y");
                mat.Add("2y");

                foreach (string s in mat)
                {

                    sMult.Rows.Add(s, "2.0", "3.6");

                }

                sMult.TableName = ccyPair;
                smileMult.Tables.Add(sMult);

            }


            string xml = "smileMultipliers";
            saveDatasetXml(xml, ds);
            MessageBox.Show("Smiles Updated");

        }

        private void saveSmileMult()
        {
            string ccyPair = ((DataTable)dataGridView8.DataSource).TableName;
            DataSet ds = smileMult;
            DataSet dp = pricingData;

            DataTable dsT = ds.Tables[ccyPair];
            DataTable dpT = dp.Tables[ccyPair];

            if (dsT == null)
            {
                dsT = ds.Tables[0].Clone();
                dsT.TableName = ccyPair;
            }

         //   dsT.Clear();

            int rowNum = 0;
            foreach (DataRow r in dpT.Rows)
            {

                //dsT.Rows[rowNum]["flyMultiplier"] = r["BrokerFly_Multiplier"].ToString();
                //dsT.Rows[rowNum]["rrMultiplier"] = r["RR_Multiplier"].ToString();

                dsT.Rows.Add(new Object[] { r["Maturity"].ToString(), r["RR_Multiplier"].ToString(), r["SmileFly_Multiplier"].ToString() });
               

                rowNum++;
            }

            saveXmlFile(dsT, xmlFilePath, "smileMultipliers");
                 

            //string xml = "smileMultipliers";
            //saveDatasetXml(xml, ds);
            MessageBox.Show("Smiles Updated");

        }

        private void loadSmileMult()
        {
            if (smileMult == null)
                smileMult = new DataSet();

            smileMult.Reset();

            xmlFileName = "smileMultipliers";

            string myXMLfile = xmlFilePath + xmlFileName + ".xml";

            //check if current file exist and if so load tables into dataset
            if (File.Exists(myXMLfile))
            {
                // Create new FileStream with which to read the schema.
                System.IO.FileStream fsReadXml = new System.IO.FileStream
                    (myXMLfile, System.IO.FileMode.Open);
                try
                {
                    smileMult.ReadXml(fsReadXml);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {
                    fsReadXml.Close();
                }
            }


        }

        public void saveXmlFile(DataTable dt, string xmlFp, string xmlFn)
        {


            string myXMLfile = xmlFp + xmlFn + ".xml";

            DataTable dtCopy = new DataTable();
            dtCopy = dt.Copy();


            DataSet ds = new DataSet();


            //check if current file exist and if so load tables into dataset
            if (File.Exists(myXMLfile))
            {
                // Create new FileStream with which to read the schema.
                System.IO.FileStream fsReadXml = new System.IO.FileStream
                    (myXMLfile, System.IO.FileMode.Open);
                try
                {
                    ds.ReadXml(fsReadXml);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {
                    fsReadXml.Close();
                }

                //   MessageBox.Show(xmlFn + " saved to: " + xmlFp);
            }

            string dtName = dtCopy.TableName;
            //  d_data.TableName = dtName;

            //check to see if there is a table for current ccy - if so delete old table and add new data else just at new table
            if (ds.Tables.Contains(dtName))
            {
                ds.Tables.Remove(dtName);



                ds.Tables.Add(dtCopy);
            }
            else
            {
                ds.Tables.Add(dtCopy);
            }

            //write xml file.

            ds.WriteXml(myXMLfile);

        }

        private void setBbgInterfacefwds()
        {
            addBbgInterface(fwds, tabPage4);

            foreach (DataRow row in crosses.Rows)
            {
                fwds.d_data.Rows.Add(row["Cross"].ToString() + " " + row["Source"].ToString() + " CURNCY");
            }

            fwds.d_data.Columns.Add("PX_MID");
            fwds.d_data.Columns.Add("FWD_CURVE");
            fwds.d_data.Columns.Add("TIME");

            fwds.buttonSendRequest.Enabled = true;


        }

        private void setBbgInterfaceDepos()
        {
            addBbgInterface(depos, tabPage5);

            foreach (DataRow row in ccyDets.Rows)
            {
                string curveCode = row["YieldCurveCode"].ToString();

                if (row["YieldCurveCode"] != DBNull.Value)
                {
                    depos.d_data.Rows.Add(curveCode + " INDEX");
                }
            }

            depos.d_data.Columns.Add("CURRENCY_RT");
            depos.d_data.Columns.Add("PAR_CURVE");


            depos.buttonSendRequest.Enabled = true;


        }

        private void checkBbgSource()
        {
            //this void will check current used swap source against the default source which is saved in xml file and loaded into defaultbbgSource datatable at initialization. If used source is not the default than bbgsource cell will be hilighted in greeen. This void will get called on recalc. 

            int curCol = dataGridView1.CurrentCell.ColumnIndex;
            string ccyPair = pricer.Rows[CcyPairR][curCol].ToString().ToUpper();

            List<string> defSource = defaultBbgSource.AsEnumerable().Select(x => x[0].ToString()).ToList();
            List<string> curSource = crosses.AsEnumerable().Select(x => x[0].ToString()).ToList();


            int i = defSource.IndexOf(ccyPair);
            int ii = curSource.IndexOf(ccyPair);

            string bbgSourceDefault = defaultBbgSource.Rows[i]["bbgSource"].ToString();
            string bbgSourceCurr = crosses.Rows[ii]["Source"].ToString();


            if (bbgSourceCurr != bbgSourceDefault)
                dataGridView1.Rows[BbgSourceR].Cells[curCol].Style.BackColor = Color.LightGreen;
            else
                dataGridView1.Rows[BbgSourceR].Cells[curCol].Style.BackColor = Color.White;

        }

        private void dataSetFwds()
        {
            if (fwdsSet == null)
                fwdsSet = new DataSet();

            DataTable dt = new DataTable();
            dt = fwds.d_data;

            try
            {
                foreach (DataRow cross in dt.Rows)
                {
                    string security = cross["security"].ToString();
                    string field = "FWD_CURVE";
                    string cellData = cross[field].ToString();

                    if (cellData != "Bulk Data...")
                    {
                        return;
                    }
                    // create bulk data table for display
                    DataTable bulkTable = fwds.d_bulkData.Tables[field].Clone();

                    string tName = security.Substring(0, 6);

                    bulkTable.TableName = tName;
                    // Get bulk data
                    DataRow[] rows = fwds.d_bulkData.Tables[field].Select("security = '" + security + "'");
                    foreach (DataRow row in rows)
                        bulkTable.ImportRow(row);


                    if (fwdsSet.Tables.Contains(tName))
                    {
                        fwdsSet.Tables.Remove(tName);
                    }

                    fwdsSet.Tables.Add(bulkTable);

                }
            }
            catch (Exception ex)
            {
                fwds.toolStripStatusLabel1.Text = ex.Message.ToString();
            }

        }

        private void dataSetDepos()
        {
            if (deposSet == null)
                deposSet = new DataSet();

            DataTable dt = new DataTable();
            dt = depos.d_data;

            try
            {
                foreach (DataRow cross in dt.Rows)
                {
                    string security = cross["security"].ToString();
                    string field = "PAR_CURVE";
                    string cellData = cross[field].ToString();

                    if (cellData != "Bulk Data...")
                    {
                        return;
                    }
                    // create bulk data table for display
                    DataTable bulkTable = depos.d_bulkData.Tables[field].Clone();

                    string tName = cross["CURRENCY_RT"].ToString();

                    bulkTable.TableName = tName;
                    // Get bulk data
                    DataRow[] rows = depos.d_bulkData.Tables[field].Select("security = '" + security + "'");
                    foreach (DataRow row in rows)
                        bulkTable.ImportRow(row);


                    if (deposSet.Tables.Contains(tName))
                    {
                        deposSet.Tables.Remove(tName);
                    }


                    deposSet.Tables.Add(bulkTable);

                }
            }
            catch (Exception ex)
            {
                fwds.toolStripStatusLabel1.Text = ex.Message.ToString();
            }

        }

        private void setRateTile(string cross)
        {
            //creates the rate tile that is used in pricing and stores tile in Marketdata dataset
            if (marketData == null)
                marketData = new DataSet();


            DataTable rateTile = new DataTable();

            rateTile.Columns.Add("Term");
            rateTile.Columns.Add("Maturity");
            rateTile.Columns.Add("DayCount");
            rateTile.Columns.Add("MidPts");
            rateTile.Columns.Add("Outright");
            rateTile.Columns.Add("Base");

            // crosses datatable itemArray 0) cross 1)depo 2)sovlbase 3)source 4)factor 5)deltatype
            rateTile.Clear();

            string ccy1 = cross.Substring(0, 3);
            string ccy2 = cross.Substring(3, 3);

            List<string> crossName = crosses.AsEnumerable().Select(x => x[0].ToString()).ToList();
            List<string> ccyPair = fwds.d_data.AsEnumerable().Select(x => x[0].ToString().Substring(0, 6)).ToList();
            List<string> ccyName = ccyDets.AsEnumerable().Select(x => x[0].ToString()).ToList();

            int i = crossName.IndexOf(cross);
            int ii = ccyPair.IndexOf(cross);
            int z = ccyName.IndexOf(ccy1);
            int j = ccyName.IndexOf(ccy2);

            DataTable dt = fwdsSet.Tables[cross];
            DataTable st = fwds.d_data;

            double factor = Convert.ToDouble(crosses.Rows[i].ItemArray[4]);
            string depo = crosses.Rows[i].ItemArray[1].ToString();
            string solveBase = crosses.Rows[i].ItemArray[2].ToString();

            DataTable ot = deposSet.Tables[depo];

            int baseBasis = Convert.ToInt32(ccyDets.Rows[z]["DayCountBasis"]);
            int termsBasis = Convert.ToInt32(ccyDets.Rows[j]["DayCountBasis"]);

            double spot = Convert.ToDouble(st.Rows[ii]["PX_MID"]);

            int r = 0;

            //looksup bbg code from rawrate datatable - if cross includes USD then int =3 because bbg code for usd pairs is ie RUB1M CURNCY. When its a cross then bbg code is EURRUB1M CURNCY. 

            //default to cross. 

            int isUSD = 6;

            if (ccy1 == "USD" || ccy2 == "USD")
            {
                isUSD = 3;
            }

            DateTime tod = DateTime.Today;
            DateTime spotDateDepo = SpotDate(tod, depo, depo);
            DateTime spotDateCross = SpotDate(tod, ccy1, ccy2);

            int rateCol = ot.Columns["Rate"].Ordinal;
            List<double> depoRates = ot.AsEnumerable().Select(x => Convert.ToDouble(x[rateCol])).ToList();
            List<double> depoDayCount = new List<double>();

            //goes through each row of deposit table and calculates the day count from maturity date to spot date and adds result to list that will be used to interplote the depo for the rate tile 

            foreach (DataRow rowDepo in ot.Rows)
            {
                DateTime mat = Convert.ToDateTime(rowDepo.ItemArray[4]);
                double dcount = (mat - spotDateDepo).TotalDays;
                depoDayCount.Add(dcount);

            }


            //goes through each row from fwd_curve and will calculate outright fwd and base rate

            foreach (DataRow row in dt.Rows)
            {
                string tenor = "";

                if (row.ItemArray[2] != DBNull.Value)
                {

                    tenor = row.ItemArray[2].ToString();
                    int len = tenor.IndexOf(' '); //splits the string at space to get bbgcode.
                    tenor = tenor.Substring(isUSD, len - isUSD); //splits string to isolate term. ie 1m, 2m 3m etc. 


                    DateTime matFwd = Convert.ToDateTime(row.ItemArray[3]);

                    double pts = 0;
                    if (row.ItemArray[5] != DBNull.Value)
                    {
                        pts = Convert.ToDouble(row.ItemArray[5]);
                    }

                    double fwd = spot + (pts / factor);

                    double deldays = (matFwd - spotDateCross).TotalDays;


                    if (deldays <= 0)
                    {
                        deldays = 1;
                    }


                    double baseRate = LinearInterp(depoDayCount, depoRates, deldays) / 100;


                    if (solveBase == "Solve")
                    {
                        baseRate = SolveBaseDepo(fwd, spot, deldays, baseRate, baseBasis, termsBasis);

                    }

                    if (tenor != "ON" && tenor != "SN")
                    {


                        rateTile.Rows.Add();
                        rateTile.Rows[r]["Term"] = tenor;
                        rateTile.Rows[r]["Maturity"] = matFwd;
                        rateTile.Rows[r]["MidPts"] = pts;
                        rateTile.Rows[r]["DayCount"] = deldays;
                        rateTile.Rows[r]["Outright"] = fwd;
                        rateTile.Rows[r]["Base"] = baseRate;
                        r++;
                    }


                }

            }

            rateTile.TableName = cross;

            //check to see if there is a table for current ccy - if so delete old table and add new data else just at new table
            if (marketData.Tables.Contains(cross))
            {
                marketData.Tables.Remove(cross);

                marketData.Tables.Add(rateTile);
            }
            else
            {
                marketData.Tables.Add(rateTile);
            }

        }

        private void spreadsFromExcel()
        {
            if (spreadSet == null)
                spreadSet = new DataSet();

            spreadSet.Reset();


            String filePath = systemFiles + @"spreads.xlsx";
            string strExcelConn;
            bool hasHeaders = false;
            string HDR = hasHeaders ? "Yes" : "No";

            if (filePath.Substring(filePath.LastIndexOf('.')).ToLower() == ".xlsx")
                strExcelConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";IMEX=0\"";
            else
                strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties=\"Excel 8.0;HDR=" + HDR + ";IMEX=0\"";


            string[] arr = new string[] { "ATM", "twenty", "ten" };

            for (int i = 0; i < arr.Count(); i++)
            {
                using (OleDbConnection connExcel = new OleDbConnection(strExcelConn))
                {
                    string table = arr[i];

                    //string selectString = "SELECT * FROM [table]";
                    string selectString = "SELECT * FROM [" + table + "]";
                    //string selectString = "SELECT * FROM [SPREADS$J36:Y81]";
                    using (OleDbCommand cmdExcel = new OleDbCommand(selectString, connExcel))
                    {
                        cmdExcel.Connection = connExcel;
                        connExcel.Open();
                        DataTable dt = new DataTable();
                        OleDbDataAdapter adp = new OleDbDataAdapter();
                        adp.SelectCommand = cmdExcel;
                        adp.FillSchema(dt, SchemaType.Source);
                        adp.Fill(dt);
                        int range = dt.Columns.Count;
                        int row = dt.Rows.Count;
                        dt.TableName = table;
                        spreadSet.Tables.Add(dt);
                    }
                }
            }
        }

        private void setRows_code()
        {

            List<string> temp = new List<string>();

            foreach (string s in rowNames)
            {
                //string t = s + "R";
                //string form = "int" + " " + t + "=0;";//+" = rowNames.IndexOf(" + "'" + s + "'" + ");";
                //temp.Add(form);


                string t = "pricer.Rows[" + s + "R" + "][curCol];";
                temp.Add(t);

            }



            string path = @"\\msk.trd.ru\shares\FX_desk\Options\Pricer\temp.txt";
            File.WriteAllLines(path, temp);

        }

        private void setMarketData()
        {
            //loops through each currency and creates rate tile that get stored in dataset marketdata
            foreach (DataRow row in crosses.Rows)
            {
                string cross = row["Cross"].ToString();

                setRateTile(cross);
            }
        }

        private void loadSpot()
        {
            //refrehese spot from bbg. Use bbg class instance called single spot
            int curCol = dataGridView1.CurrentCell.ColumnIndex;
            string cross;

            if (pricer.Rows[CcyPairR][curCol] != DBNull.Value)
            {
                cross = pricer.Rows[CcyPairR][curCol].ToString();
            }
            else
            {
                cross = pricer.Rows[CcyPairR][curCol - 1].ToString();
            }

            //
            if (singleSpot.d_data.Rows.Count < 1)
            {
                singleSpot.d_data.Columns.Add("PX_MID");
            }

            singleSpot.d_data.Rows.Clear();
            singleSpot.d_data.Rows.Add(cross + " CURNCY");
            singleSpot.sendRequest();

            DataRow lastRow = singleSpot.d_data.Rows[singleSpot.d_data.Rows.Count - 1];
            bool endload = false;

            do
            {
                if (lastRow["PX_MID"] != DBNull.Value)
                    endload = true;

            } while (endload == false);

            string spot = singleSpot.d_data.Rows[0]["PX_MID"].ToString();

            List<string> sec = fwds.d_data.AsEnumerable().Select(x => x[0].ToString().Substring(0, 6)).ToList();
            int i = sec.IndexOf(cross);

            fwds.d_data.Rows[i]["PX_MID"] = spot;
            pricer.Rows[SpotR][curCol] = spot;

        }

        private void riskSim()
        {
            DataTable dt = new DataTable();
            dt = skewSheet.Tables["optList"];
             int i = 0;
            foreach (DataRow row in dt.Rows)
            {
                string mat = row["Maturity"].ToString();
                string strike = row["strike"].ToString();
                string pC = row["C/P"].ToString().Substring(0, 1);
              

                optPricer(mat, strike, pC);
               // tradeSim.Tables[0].TableName = "opt." +  i.ToString();
                i++;


            }

        }

        private static void CombineRows(DataTable combinedTable, DataTable table4, int rowNumber, int columnNumber)
        {

            if (table4.Rows[rowNumber][columnNumber] != System.DBNull.Value)
            {
                combinedTable.Rows[rowNumber][columnNumber] = ((Convert.ToDouble(combinedTable.Rows[rowNumber][columnNumber]) + Convert.ToDouble(table4.Rows[rowNumber][columnNumber]))).ToString("#,##0");
            }
            
        }

        private int getRowInt(DataTable dt, string rowHeader)
        {
            List<string> rName = dt.AsEnumerable().Select(x => x[0].ToString()).ToList();
            int r = rName.IndexOf(rowHeader);
            return r;
        }

        private void combineSpreadSim()
        {
            //combines any amonunt of options into a simulation

            Int32 selectedCellCount = dataGridView1.GetCellCount(DataGridViewElementStates.Selected);
            int curCol = dataGridView1.SelectedCells[selectedCellCount - 1].ColumnIndex;
            int colCount = dataGridView1.CurrentCell.ColumnIndex - curCol + 1;
            int colSum = dataGridView1.CurrentCell.ColumnIndex + 1;
            DataSet ds = new DataSet();

            //add datatable into new dataset which will then be combined into new datatable
            for (int i = 0; i < colCount; i++)
            {
                DataTable dt = new DataTable();
                string tName = (curCol + i).ToString();
                dt = tradeSim.Tables[tName].Copy();
                ds.Tables.Add(dt);
            }

            //will add each tabel to the next starting with table 0 in dataset
            DataTable sumDt = new DataTable();
            sumDt = ds.Tables[0].Copy();
            


            for (int i = 1; i < ds.Tables.Count; i++)
            {
                DataTable dt = ds.Tables[i];
                for (var rowNumber = 2; rowNumber < sumDt.Rows.Count; rowNumber++)
                {
                    for (var columnNumber = 1; columnNumber < sumDt.Columns.Count; columnNumber++)
                    {
                        CombineRows(sumDt, dt, rowNumber, columnNumber);
                    }
                }
                
            }

            //add blank row
            sumDt.Rows.Add("VegaSummary");

           //add row with each options vega 
            int j = getRowInt(sumDt, "sVega");
            int v = 0;
            foreach (DataTable dt in ds.Tables)
            {
                sumDt.Rows.Add(dt.Rows[j].ItemArray);
                int x = sumDt.Rows.Count-1;
                sumDt.Rows[x][0] =  pricer.Rows[ExpiryR][curCol + v].ToString();
                v++;
            }

            //add back total vega
            sumDt.Rows.Add(sumDt.Rows[j].ItemArray);

           
            //delete vol and duplicate vega rows
            int r = getRowInt(sumDt, "Vol");
            sumDt.Rows[r].Delete();

            r = getRowInt(sumDt, "sVega");
            sumDt.Rows[r].Delete();
            
            //display
            showTradeSim(sumDt);

        }

        private void combineSpreadSimOLD()
        {
            Int32 selectedCellCount = dataGridView1.GetCellCount(DataGridViewElementStates.Selected);
            int curCol = dataGridView1.SelectedCells[selectedCellCount - 1].ColumnIndex;
            int colCount = dataGridView1.CurrentCell.ColumnIndex - curCol + 1;
            int colSum = dataGridView1.CurrentCell.ColumnIndex + 1;
            DataSet ds = new DataSet();

            //add datatable into new dataset which will then be combined into new datatable
            for (int i = 0; i < colCount; i++)
            {
                DataTable dt = new DataTable();
                string tName = (curCol + i).ToString();    
                dt = tradeSim.Tables[tName].Copy() ;
                ds.Tables.Add(dt);
            }


            DataTable sumDt = new DataTable();
            sumDt = ds.Tables[0].Clone();
            sumDt.Rows.Add(tradeSim.Tables[0].Rows[0].ItemArray);
            sumDt.Rows.Add("sDelta");
            sumDt.Rows.Add("sDeltaCCent");
            sumDt.Rows.Add("SDeltaCCol");
            sumDt.Rows.Add("SVegaTotal");


            for (int i = 1; i < sumDt.Rows.Count; i++)
            
                foreach (DataColumn dc in sumDt.Columns)
                {
                    if (dc.Ordinal != 0)
                    {
            
                        sumDt.Rows[i][dc] = 0;
                    }
                   
                }

            int j = 5;
            int q = 0;
            foreach (DataTable dt in ds.Tables)
            {

                for (int i = 1; i <sumDt.Columns.Count ; i++)
               
                {
                        double d = Convert.ToDouble(dt.Rows[2][i]) + Convert.ToDouble(sumDt.Rows[1][i]);
                        double d1 = Convert.ToDouble(dt.Rows[3][i]) + Convert.ToDouble(sumDt.Rows[2][i]);
                        double d2 = Convert.ToDouble(dt.Rows[4][i]) + Convert.ToDouble(sumDt.Rows[3][i]);
                        double d3 = Convert.ToDouble(dt.Rows[5][i]) + Convert.ToDouble(sumDt.Rows[4][i]);

                        sumDt.Rows[1][i] = d.ToString("#,##0"); 
                        sumDt.Rows[2][i] = d1.ToString("#,##0"); 
                        sumDt.Rows[3][i] = d2.ToString("#,##0");
                        sumDt.Rows[4][i] = d3.ToString("#,##0"); 
                }

               
                sumDt.Rows.Add(dt.Rows[5].ItemArray);
                sumDt.Rows[j][0] = pricer.Rows[ExpiryR][curCol +q].ToString();
                j++;
                q++;
               
            }


            showTradeSim(sumDt);

        }

        private void broadCast()
        {
            //will calculated option spread values of highlighted columns
            Int32 selectedCellCount = dataGridView1.GetCellCount(DataGridViewElementStates.Selected);
            int curCol = dataGridView1.SelectedCells[selectedCellCount - 1].ColumnIndex;
            int colCount = dataGridView1.CurrentCell.ColumnIndex - curCol + 1;
            int colSum = dataGridView1.CurrentCell.ColumnIndex + 1;

            double firstNotional = Convert.ToDouble(pricer.Rows[NotionalR][curCol]) * 1000000;
            double premo = 0;
            double delta = 0;
            double sdelta = 0;
            double vega = 0;
            double svega = 0;
            double rega25 = 0;
            double sega25 = 0;
            double rega10 = 0;
            double sega10 = 0;
            double vanna = 0;
            double volga = 0;
            double gamma = 0;
            double sgamma = 0;
            double dvBase = 0;
            double dvTerms = 0;
      
            double premiumFromMid = 0;
            double theta = 0;
            double vol = 0;

            this.dataGridView1.CellValueChanged -= this.dataGridView1_CellValueChanged;

            for (int i = 0; i < colCount; i++)
            {
                premo = premo + Convert.ToDouble(pricer.Rows[PremiumAR][i + curCol]);
                delta = delta + Convert.ToDouble(pricer.Rows[DeltaAR][i + curCol]);
                sdelta = sdelta + Convert.ToDouble(pricer.Rows[sDeltaR][i + curCol]);
                vega = vega + Convert.ToDouble(pricer.Rows[VegaAR][i + curCol]);
                svega = svega + Convert.ToDouble(pricer.Rows[sVegaR][i + curCol]);
                rega25 = rega25 + Convert.ToDouble(pricer.Rows[Rega25R][i + curCol]);
                sega25 = sega25 + Convert.ToDouble(pricer.Rows[Sega25R][i + curCol]);
                rega10 = rega10 + Convert.ToDouble(pricer.Rows[Rega10R][i + curCol]);
                sega10 = sega25 + Convert.ToDouble(pricer.Rows[Sega10R][i + curCol]);
                vanna = vanna + Convert.ToDouble(pricer.Rows[VannaAR][i + curCol]);
                volga = volga + Convert.ToDouble(pricer.Rows[VolgaAR][i + curCol]);
                gamma = gamma + Convert.ToDouble(pricer.Rows[GammaAR][i + curCol]);
                sgamma = sgamma + Convert.ToDouble(pricer.Rows[sGammaR][i + curCol]);
                dvBase = dvBase + Convert.ToDouble(pricer.Rows[DV01_BaseAR][i + curCol]);
                dvTerms = dvTerms + Convert.ToDouble(pricer.Rows[DV01_TermsAR][i + curCol]);
                premiumFromMid = premiumFromMid + Convert.ToDouble(pricer.Rows[PremiumFromMidR][i + curCol]);
                theta = theta + Convert.ToDouble(pricer.Rows[ThetaAR][i + curCol]);

            }

            if (colCount == 2)
            {

                vol = convertPercent(pricer.Rows[VolR][curCol].ToString()) - convertPercent(pricer.Rows[VolR][curCol + 1].ToString());
            }
            // pricePct = premo / firstNotional;

            pricer.Rows[VolR][colSum] = vol.ToString("0.00%");
            pricer.Rows[PremiumFromMidR][colSum] = premiumFromMid.ToString("#,##0");
            pricer.Rows[PremiumAR][colSum] = premo.ToString("#,##0");
            pricer.Rows[DeltaAR][colSum] = delta.ToString("#,##0");
            pricer.Rows[VegaAR][colSum] = vega.ToString("#,##0");

            pricer.Rows[sVegaR][colSum] = svega.ToString("#,##0");
            pricer.Rows[sDeltaR][colSum] = sdelta.ToString("#,##0");
            pricer.Rows[Rega25R][colSum] = rega25.ToString("#,##0");
            pricer.Rows[Sega25R][colSum] = sega25.ToString("#,##0");
            pricer.Rows[Rega10R][colSum] = rega10.ToString("#,##0");
            pricer.Rows[Sega10R][colSum] = sega10.ToString("#,##0");
            pricer.Rows[VannaAR][colSum] = vanna.ToString("#,##0");
            pricer.Rows[VolgaAR][colSum] = volga.ToString("#,##0");
            pricer.Rows[ThetaAR][colSum] = theta.ToString("#,##0");
            pricer.Rows[GammaAR][colSum] = gamma.ToString("#,##0");
            pricer.Rows[sGammaR][colSum] = sgamma.ToString("#,##0");
            pricer.Rows[DV01_BaseAR][colSum] = dvBase.ToString("#,##0");
            pricer.Rows[DV01_TermsAR][colSum] = dvTerms.ToString("#,##0");


            dataGridView1.Refresh();




            this.dataGridView1.CellValueChanged += this.dataGridView1_CellValueChanged;
        }

        private void refreshData(string instrument)
        {
            //refreshes fwdpts/base depo/vols/smile when expiry date is changed.

            int curCol = dataGridView1.CurrentCell.ColumnIndex;

            //from datatable
            string ccyPair = pricer.Rows[CcyPairR][curCol].ToString();
            double dayCount = Convert.ToDouble(pricer.Rows[Expiry_DaysR][curCol]);
            double delDayCount = Convert.ToDouble(pricer.Rows[Deliver_DaysR][curCol]);
            double spot = Convert.ToDouble(pricer.Rows[SpotR][curCol]);
            string exp = pricer.Rows[ExpiryR][curCol].ToString();
            string strike = pricer.Rows[StrikeR][curCol].ToString();
            string pC = pricer.Rows[Put_CallR][curCol].ToString();


            int[] arr = dayCountBasis(ccyPair);
            int basisB = arr[0];
            int basisT = arr[1];

            // calls method to get cross info 
            object[] crossInfo = crossDtData(ccyPair);
            int volType = Convert.ToInt16(crossInfo[0]);
            double factor = Convert.ToDouble(crossInfo[1]);
            string bbgSource = crossInfo[2].ToString();

            double[] rateComponents = rateBuilder(ccyPair, dayCount, delDayCount, spot, factor, basisB, basisT);
            //fwdPts, outRight, forDepo, domDepo

            double fwdPts = rateComponents[0];
            double outRight = rateComponents[1]; ;
            double forDepo = rateComponents[2];
            double domDepo = rateComponents[3];

            //get fly, rr from pricerData displayed on main pricer screen 
            double[] volComponents = null;

            volComponents = volBuilder(ccyPair, dayCount);
            double atmVol = volComponents[0];
            double rr = volComponents[1];
            double fly = volComponents[2];
            double wingControl = volComponents[3];
            double targetFlyMult = volComponents[4];

            //fill grid
            if (instrument == "all")
            {
                pricer.Rows[Swap_PtsR][curCol] = fwdPts;
                pricer.Rows[FwdR][curCol] = outRight;
                pricer.Rows[Depo_BaseR][curCol] = forDepo;
                pricer.Rows[Depo_TermsR][curCol] = domDepo;
                pricer.Rows[ATM_VOLR][curCol] = atmVol;
                pricer.Rows[RRR][curCol] = rr;
                pricer.Rows[FLYR][curCol] = fly;
            }

            if (instrument == "atm")
            { pricer.Rows[ATM_VOLR][curCol] = atmVol; }

            if (instrument == "fly")
            { pricer.Rows[FLYR][curCol] = fly; }


            if (instrument == "rr")
            { pricer.Rows[RRR][curCol] = rr; }

            if (instrument == "fwdPts")
            { pricer.Rows[Swap_PtsR][curCol] = fwdPts; }

            if (instrument == "outRight")
            { pricer.Rows[FwdR][curCol] = outRight; }

            if (instrument == "depoB")
            { pricer.Rows[Depo_BaseR][curCol] = forDepo; }

            if (instrument == "depoT")
            { pricer.Rows[Depo_TermsR][curCol] = domDepo; }

            optPricer(exp, strike, pC);

        }

        private double convertPercent(string p)
        {
            //Functions converts strings with % into doubles

            double functionReturnValue = 0;

            if (p.Contains("%"))
            {
                functionReturnValue = double.Parse(p.Split(new char[] { '%' })[0]) / 100;
            }
            else
            { functionReturnValue = Convert.ToDouble(p); }


            return functionReturnValue;
        }

        private string cellValidationPct(string p)
        {
            //Functions converts strings with % into doubles
            string functionReturnValue = p;

            if (p.Contains("%"))
            {
                return functionReturnValue;
            }
            else
            { functionReturnValue = functionReturnValue + "%"; }


            return functionReturnValue;
        }

        private void clearAllColumns()
        {
            int curRow = dataGridView1.CurrentCell.RowIndex;
            int curCol = dataGridView1.CurrentCell.ColumnIndex;



            this.dataGridView1.CellValueChanged -= this.dataGridView1_CellValueChanged;

            for (int i = 1; i < dataGridView1.Columns.Count; i++)
            {


                foreach (DataGridViewRow myRow in dataGridView1.Rows)
                {
                    myRow.Cells[i].Value = DBNull.Value; // assuming you want to clear the first column
                    myRow.Cells[i].Style.BackColor = Color.White;
                }
            }

            dataGridView1.Refresh();

            dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[1];

            this.dataGridView1.CellValueChanged += this.dataGridView1_CellValueChanged;

        }

        private void addCombobox()
        {
            List<string> crossName = crosses.AsEnumerable().Select(x => x[0].ToString()).ToList();

            for (int i = 1; i < dataGridView1.Columns.Count; i++)
            {
                DataGridViewComboBoxCell comboBoxCell = new DataGridViewComboBoxCell();
                comboBoxCell.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
                comboBoxCell.DataSource = crossName;
                dataGridView1.Rows[0].Cells[i] = comboBoxCell;
            }


            List<String> premoType = new List<string>();
            premoType.Add("Base %");
            premoType.Add("Terms %");
            premoType.Add("Base Pips");
            premoType.Add("Terms Pips");

            int rowNum = Premium_TypeR;

            for (int i = 1; i < dataGridView1.Columns.Count; i++)
            {
                DataGridViewComboBoxCell comboBoxCell = new DataGridViewComboBoxCell();
                comboBoxCell.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
                comboBoxCell.DataSource = premoType;
                dataGridView1.Rows[rowNum].Cells[i] = comboBoxCell;
            }


            List<String> put_call = new List<string>();
            put_call.Add("p");
            put_call.Add("c");


            int pRow = Put_CallR;

            for (int i = 1; i < dataGridView1.Columns.Count; i++)
            {
                DataGridViewComboBoxCell comboBoxCell = new DataGridViewComboBoxCell();
                comboBoxCell.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
                comboBoxCell.DataSource = put_call;
                dataGridView1.Rows[pRow].Cells[i] = comboBoxCell;
            }
        }

        public double Vol(string curr, double dayCount)
        {
            double functionValue = 0;

            string fN = volPath + curr + ".vols.txt";
            DataTable dt = new DataTable();

            int z = 10;

            for (int i = 0; i <= z; i++)
            {
                dt.Columns.Add(i.ToString());
            }

            loadDataFullFile(dt, fN, z);

            List<double> days = dt.AsEnumerable().Select(x => Convert.ToDouble(x[0])).ToList();
            int y = days.IndexOf(dayCount);

            functionValue = Convert.ToDouble(dt.Rows[y][7]);
            return functionValue;

        }

        public double smile(string curr, double dayCount, string fly_R)
        {
            double functionValue = 0;

            string fN = volPath + curr + ".smile.txt";
            DataTable dt = new DataTable();

            int z = 6;

            for (int i = 0; i <= z; i++)
            {
                dt.Columns.Add(i.ToString());
            }

            loadDataFullFile(dt, fN, z);

            List<double> days = dt.AsEnumerable().Select(x => Convert.ToDouble(x[0])).ToList();
            List<double> rr = dt.AsEnumerable().Select(x => Convert.ToDouble(x[6])).ToList();
            List<double> fly = dt.AsEnumerable().Select(x => Convert.ToDouble(x[5])).ToList();
            List<double> tY = dt.AsEnumerable().Select(x => Convert.ToDouble(x[3])).ToList();

            double maxVal = days.Max();

            if (dayCount > maxVal)
            {
                MessageBox.Show("DayCount is outside ranged passed from dvi: Daycount = " + dayCount + " MaxVal of range is = " + maxVal);
                dayCount = maxVal;

            }
            //int y = days.IndexOf(dayCount);


            if (fly_R == "fly")
            {
                functionValue = LinearInterp(days, rr, dayCount);
                return functionValue;
            }

            if (fly_R == "rr")
            {
                functionValue = LinearInterp(days, fly, dayCount);
                return functionValue;
            }

            if (fly_R == "2y" || fly_R == "3y" || fly_R == "4y" || fly_R == "5y")
            {
                functionValue = LinearInterp(days, tY, dayCount);
                return functionValue;
            }
            functionValue = 0;
            return functionValue;
        }

        private void fill_Cross_Box()
        {

            //Build a list
            List<string> colA = new List<string>();

            foreach (DataRow row in crosses.Rows)
            {
                colA.Add(row["Cross"].ToString().ToUpper());

            }

            //Setup data binding
            this.comboBox1.DataSource = colA;

            // toolStripComboBox1.ComboBox.BindingContext = this.BindingContext;
            //toolStripComboBox1.ComboBox.DataSource = colA;

            // make it readonly
            this.comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            //this.toolStripComboBox1.DropDownStyle = ComboBoxStyle.DropDownList;


        }

        private void formatGrid()
        {

            int curCol = dataGridView1.CurrentCell.ColumnIndex;

            dataGridView1.Rows[Spot_DeltaR].DefaultCellStyle.Font = new Font(this.Font, FontStyle.Bold);
            dataGridView1.Rows[ATM_VOLR].DefaultCellStyle.Font = new Font(this.Font, FontStyle.Bold);
            dataGridView1.Rows[RRR].DefaultCellStyle.Font = new Font(this.Font, FontStyle.Bold);
            dataGridView1.Rows[FLYR].DefaultCellStyle.Font = new Font(this.Font, FontStyle.Bold);
            // dataGridView1.Rows[Price_Pct_BaseR].DefaultCellStyle.Font = new Font(this.Font, FontStyle.Bold);
            dataGridView1.Rows[VegaR].DefaultCellStyle.Font = new Font(this.Font, FontStyle.Bold);
            dataGridView1.Rows[VolR].DefaultCellStyle.Font = new Font(this.Font, FontStyle.Bold);
            dataGridView1.Rows[SpotR].DefaultCellStyle.Font = new Font(this.Font, FontStyle.Bold);
            dataGridView1.Rows[AutoStrikeR].DefaultCellStyle.Font = new Font(this.Font, FontStyle.Bold);
            dataGridView1.Rows[Swap_PtsR].DefaultCellStyle.Font = new Font(this.Font, FontStyle.Bold);
            dataGridView1.Rows[FwdR].DefaultCellStyle.Font = new Font(this.Font, FontStyle.Bold);
            dataGridView1.Rows[BloombergVolR].DefaultCellStyle.Font = new Font(this.Font, FontStyle.Bold);
            dataGridView1.Rows[Bid_OfferR].DefaultCellStyle.Font = new Font(this.Font, FontStyle.Bold);


            dataGridView1.Rows[Spot_DeltaR].DefaultCellStyle.ForeColor = Color.Red;
            dataGridView1.Rows[AutoStrikeR].DefaultCellStyle.ForeColor = Color.Red;
            dataGridView1.Rows[ExpiryDateR].DefaultCellStyle.ForeColor = Color.Red;
            dataGridView1.Rows[DeliveryDateR].DefaultCellStyle.ForeColor = Color.Red;

            dataGridView1.Rows[ATM_VOLR].DefaultCellStyle.ForeColor = Color.Navy;
            dataGridView1.Rows[RRR].DefaultCellStyle.ForeColor = Color.Navy;
            dataGridView1.Rows[FLYR].DefaultCellStyle.ForeColor = Color.Navy;
            dataGridView1.Rows[Bid_OfferR].DefaultCellStyle.ForeColor = Color.Navy;
            dataGridView1.Rows[BloombergVolR].DefaultCellStyle.ForeColor = Color.Navy;


            dataGridView1.Rows[Vol_Spread_to_AtmR].DefaultCellStyle.ForeColor = Color.Navy;
            dataGridView1.Rows[Bps_to_AtmR].DefaultCellStyle.ForeColor = Color.Navy;
            dataGridView1.Rows[PremiumFromMidR].DefaultCellStyle.ForeColor = Color.Navy;

            dataGridView1.Rows[SystemVolR].DefaultCellStyle.ForeColor = Color.Navy;
            dataGridView1.Rows[BpsFromMidR].DefaultCellStyle.ForeColor = Color.Navy;


            //dataGridView1.Rows[ATM_VOLR].DefaultCellStyle.BackColor = Color.PaleTurquoise;
            //dataGridView1.Rows[RRR].DefaultCellStyle.BackColor = Color.PaleTurquoise;
            //dataGridView1.Rows[FLYR].DefaultCellStyle.BackColor = Color.PaleTurquoise;

            dataGridView1.Rows[BloombergVolR].Cells[curCol].Style.BackColor = Color.LightGray;
            dataGridView1.Rows[ATM_VOLR].Cells[curCol].Style.BackColor = Color.PaleTurquoise;
            dataGridView1.Rows[RRR].Cells[curCol].Style.BackColor = Color.PaleTurquoise;
            dataGridView1.Rows[FLYR].Cells[curCol].Style.BackColor = Color.PaleTurquoise;


            //dataGridView1.Rows[ExpiryR].Cells[curCol].Style.BackColor = Color.LightGray;
            //dataGridView1.Rows[Put_CallR].Cells[curCol].Style.BackColor = Color.LightGray;
            //dataGridView1.Rows[StrikeR].Cells[curCol].Style.BackColor = Color.LightGray;

            for (int i = 1; i <= dataGridView1.Columns.Count - 1; i++)
            {
                dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            }
        }

        private void formatPricerDt()
        {
            dataGridView1.Columns["OptionNumber"].Frozen = true;
            dataGridView1.Columns["OptionNumber"].ReadOnly = true;

            foreach (DataGridViewColumn var in this.dataGridView1.Columns) { var.SortMode = DataGridViewColumnSortMode.NotSortable; }

            for (int i = 1; i <= dataGridView1.Columns.Count - 1; i++)
            {
                dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Rows[CcyPairR].Cells[i].Style.BackColor = Color.Aqua;
                dataGridView1.Rows[SpotR].Cells[i].Style.BackColor = Color.Cyan;
                dataGridView1.Rows[ExpiryR].Cells[i].Style.BackColor = Color.Cyan;
                dataGridView1.Rows[StrikeR].Cells[i].Style.BackColor = Color.Cyan;
                dataGridView1.Rows[Put_CallR].Cells[i].Style.BackColor = Color.Cyan;
                dataGridView1.Rows[NotionalR].Cells[i].Style.BackColor = Color.Cyan;


                dataGridView1.Rows[VolR].Cells[i].Style.BackColor = Color.White;
                dataGridView1.Rows[Bid_OfferR].Cells[i].Style.BackColor = Color.White;


                dataGridView1.Rows[SystemVolR].Cells[i].Style.BackColor = Color.White;
                dataGridView1.Rows[BpsFromMidR].Cells[i].Style.BackColor = Color.White;
                dataGridView1.Rows[PremiumFromMidR].Cells[i].Style.BackColor = Color.White;
                dataGridView1.Rows[Vol_Spread_to_AtmR].Cells[i].Style.BackColor = Color.White;
                dataGridView1.Rows[BloombergVolR].Cells[i].Style.BackColor = Color.LightGreen;



                dataGridView1.Rows[Spot_DeltaR].Cells[i].Style.BackColor = Color.White;
                dataGridView1.Rows[AutoStrikeR].Cells[i].Style.BackColor = Color.White;
                dataGridView1.Rows[ExpiryDateR].Cells[i].Style.BackColor = Color.Lavender;
                dataGridView1.Rows[DeliveryDateR].Cells[i].Style.BackColor = Color.Lavender;
                dataGridView1.Rows[Expiry_DaysR].Cells[i].Style.BackColor = Color.Lavender;
                dataGridView1.Rows[Deliver_DaysR].Cells[i].Style.BackColor = Color.Lavender;


                dataGridView1.Rows[ATM_VOLR].Cells[i].Style.BackColor = Color.LavenderBlush;
                dataGridView1.Rows[RRR].Cells[i].Style.BackColor = Color.LavenderBlush;
                dataGridView1.Rows[FLYR].Cells[i].Style.BackColor = Color.LavenderBlush;
                dataGridView1.Rows[BreakEvenR].Cells[i].Style.BackColor = Color.LavenderBlush;

                dataGridView1.Rows[Swap_PtsR].Cells[i].Style.BackColor = Color.LavenderBlush;
                dataGridView1.Rows[FwdR].Cells[i].Style.BackColor = Color.LavenderBlush;
                dataGridView1.Rows[Depo_BaseR].Cells[i].Style.BackColor = Color.LavenderBlush;
                dataGridView1.Rows[Depo_TermsR].Cells[i].Style.BackColor = Color.LavenderBlush;
                dataGridView1.Rows[BbgSourceR].Cells[i].Style.BackColor = Color.LavenderBlush;

                dataGridView1.Rows[Premium_TypeR].Cells[i].Style.BackColor = Color.White;
                dataGridView1.Rows[PriceR].Cells[i].Style.BackColor = Color.White;
               // dataGridView1.Rows[GammaR].Cells[i].Style.BackColor = Color.White;
                dataGridView1.Rows[VegaR].Cells[i].Style.BackColor = Color.White;
               // dataGridView1.Rows[RegaR].Cells[i].Style.BackColor = Color.White;
               // dataGridView1.Rows[SegaR].Cells[i].Style.BackColor = Color.White;
               // dataGridView1.Rows[VannaR].Cells[i].Style.BackColor = Color.White;
                //dataGridView1.Rows[VolgaR].Cells[i].Style.BackColor = Color.White;
                //dataGridView1.Rows[ThetaR].Cells[i].Style.BackColor = Color.White;
                //dataGridView1.Rows[DV01_BaseR].Cells[i].Style.BackColor = Color.White;
                //dataGridView1.Rows[DV01_TermsR].Cells[i].Style.BackColor = Color.White;
                

                dataGridView1.Rows[PremiumAR].Cells[i].Style.BackColor = Color.Azure;
                dataGridView1.Rows[sVegaR].Cells[i].Style.BackColor = Color.Azure;
                dataGridView1.Rows[sDeltaR].Cells[i].Style.BackColor = Color.Azure;
                dataGridView1.Rows[Rega25R].Cells[i].Style.BackColor = Color.Azure;
                dataGridView1.Rows[Rega10R].Cells[i].Style.BackColor = Color.Azure;
                dataGridView1.Rows[Sega25R].Cells[i].Style.BackColor = Color.Azure;
                dataGridView1.Rows[Sega10R].Cells[i].Style.BackColor = Color.Azure;
                


                dataGridView1.Rows[DeltaAR].Cells[i].Style.BackColor = Color.Azure;
                dataGridView1.Rows[GammaAR].Cells[i].Style.BackColor = Color.Azure;
                dataGridView1.Rows[VegaAR].Cells[i].Style.BackColor = Color.Azure;
               // dataGridView1.Rows[RegaAR].Cells[i].Style.BackColor = Color.Azure;
                //dataGridView1.Rows[SegaAR].Cells[i].Style.BackColor = Color.Azure;
                dataGridView1.Rows[VannaAR].Cells[i].Style.BackColor = Color.Azure;
                dataGridView1.Rows[VolgaAR].Cells[i].Style.BackColor = Color.Azure;
                dataGridView1.Rows[ThetaAR].Cells[i].Style.BackColor = Color.Azure;
                dataGridView1.Rows[DV01_BaseAR].Cells[i].Style.BackColor = Color.Azure;
                dataGridView1.Rows[DV01_TermsAR].Cells[i].Style.BackColor = Color.Azure;

            }


            dataGridView1.Columns[0].DefaultCellStyle.BackColor = Color.DarkGray;
            dataGridView1.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
            {
                dataGridView1.Rows[i].Height = 17;
            }

            for (int i = 0; i <= dataGridView1.Columns.Count - 1; i++)
            {
                dataGridView1.Rows[NotionalR + 1].Cells[i].Style.BackColor = Color.LightGray;
                dataGridView1.Rows[NotionalR + 1].Height = 11;
                dataGridView1.Rows[Bid_OfferR + 1].Cells[i].Style.BackColor = Color.LightGray;
                dataGridView1.Rows[Bid_OfferR + 1].Height = 11;
                dataGridView1.Rows[BloombergVolR + 1].Cells[i].Style.BackColor = Color.LightGray;
                dataGridView1.Rows[BloombergVolR + 1].Height = 11;
                dataGridView1.Rows[Deliver_DaysR + 1].Cells[i].Style.BackColor = Color.LightGray;
                dataGridView1.Rows[Deliver_DaysR + 1].Height = 11;
                dataGridView1.Rows[BreakEvenR + 1].Cells[i].Style.BackColor = Color.LightGray;
                dataGridView1.Rows[BreakEvenR + 1].Height = 11;
                dataGridView1.Rows[BbgSourceR + 1].Cells[i].Style.BackColor = Color.LightGray;
                dataGridView1.Rows[BbgSourceR + 1].Height = 11;
            
                dataGridView1.Rows[DV01_TermsAR + 1].Cells[i].Style.BackColor = Color.LightGray;
                dataGridView1.Rows[DV01_TermsAR + 1].Height = 11;


                dataGridView1.Rows[Bps_to_AtmR + 1].Cells[i].Style.BackColor = Color.LightGray;
                dataGridView1.Rows[Bps_to_AtmR + 1].Height = 11;


                dataGridView1.Rows[AutoStrikeR + 1].Cells[i].Style.BackColor = Color.LightGray;
                dataGridView1.Rows[AutoStrikeR + 1].Height = 11;
            }


        }

        private void formatColumn()
        {

            int i = dataGridView1.CurrentCell.ColumnIndex;


            dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Rows[CcyPairR].Cells[i].Style.BackColor = Color.Aqua;
            dataGridView1.Rows[SpotR].Cells[i].Style.BackColor = Color.Cyan;
            dataGridView1.Rows[ExpiryR].Cells[i].Style.BackColor = Color.Cyan;
            dataGridView1.Rows[StrikeR].Cells[i].Style.BackColor = Color.Cyan;
            dataGridView1.Rows[Put_CallR].Cells[i].Style.BackColor = Color.Cyan;
            dataGridView1.Rows[NotionalR].Cells[i].Style.BackColor = Color.Cyan;


            dataGridView1.Rows[VolR].Cells[i].Style.BackColor = Color.White;
            dataGridView1.Rows[Bid_OfferR].Cells[i].Style.BackColor = Color.White;


            dataGridView1.Rows[SystemVolR].Cells[i].Style.BackColor = Color.White;
            dataGridView1.Rows[BpsFromMidR].Cells[i].Style.BackColor = Color.White;
            dataGridView1.Rows[PremiumFromMidR].Cells[i].Style.BackColor = Color.White;
            dataGridView1.Rows[Vol_Spread_to_AtmR].Cells[i].Style.BackColor = Color.White;
            dataGridView1.Rows[BloombergVolR].Cells[i].Style.BackColor = Color.LightGreen;



            dataGridView1.Rows[Spot_DeltaR].Cells[i].Style.BackColor = Color.White;
            dataGridView1.Rows[AutoStrikeR].Cells[i].Style.BackColor = Color.White;
            dataGridView1.Rows[ExpiryDateR].Cells[i].Style.BackColor = Color.Lavender;
            dataGridView1.Rows[DeliveryDateR].Cells[i].Style.BackColor = Color.Lavender;
            dataGridView1.Rows[Expiry_DaysR].Cells[i].Style.BackColor = Color.Lavender;
            dataGridView1.Rows[Deliver_DaysR].Cells[i].Style.BackColor = Color.Lavender;


            dataGridView1.Rows[ATM_VOLR].Cells[i].Style.BackColor = Color.LavenderBlush;
            dataGridView1.Rows[RRR].Cells[i].Style.BackColor = Color.LavenderBlush;
            dataGridView1.Rows[FLYR].Cells[i].Style.BackColor = Color.LavenderBlush;

            dataGridView1.Rows[Swap_PtsR].Cells[i].Style.BackColor = Color.LavenderBlush;
            dataGridView1.Rows[FwdR].Cells[i].Style.BackColor = Color.LavenderBlush;
            dataGridView1.Rows[Depo_BaseR].Cells[i].Style.BackColor = Color.LavenderBlush;
            dataGridView1.Rows[Depo_TermsR].Cells[i].Style.BackColor = Color.LavenderBlush;
            dataGridView1.Rows[BbgSourceR].Cells[i].Style.BackColor = Color.LavenderBlush;

            dataGridView1.Rows[Premium_TypeR].Cells[i].Style.BackColor = Color.White;
            dataGridView1.Rows[PriceR].Cells[i].Style.BackColor = Color.White;
            dataGridView1.Rows[GammaR].Cells[i].Style.BackColor = Color.White;
            dataGridView1.Rows[VegaR].Cells[i].Style.BackColor = Color.White;
           // dataGridView1.Rows[RegaR].Cells[i].Style.BackColor = Color.White;
            //dataGridView1.Rows[SegaR].Cells[i].Style.BackColor = Color.White;
            dataGridView1.Rows[VannaR].Cells[i].Style.BackColor = Color.White;
            dataGridView1.Rows[VolgaR].Cells[i].Style.BackColor = Color.White;
            dataGridView1.Rows[ThetaR].Cells[i].Style.BackColor = Color.White;
            dataGridView1.Rows[DV01_BaseR].Cells[i].Style.BackColor = Color.White;
            dataGridView1.Rows[DV01_TermsR].Cells[i].Style.BackColor = Color.White;
            dataGridView1.Rows[BreakEvenR].Cells[i].Style.BackColor = Color.White;

            dataGridView1.Rows[PremiumAR].Cells[i].Style.BackColor = Color.Azure;
            dataGridView1.Rows[DeltaAR].Cells[i].Style.BackColor = Color.Azure;
            dataGridView1.Rows[GammaAR].Cells[i].Style.BackColor = Color.Azure;
            dataGridView1.Rows[VegaAR].Cells[i].Style.BackColor = Color.Azure;
           // dataGridView1.Rows[RegaAR].Cells[i].Style.BackColor = Color.Azure;
           // dataGridView1.Rows[SegaAR].Cells[i].Style.BackColor = Color.Azure;
            dataGridView1.Rows[VannaAR].Cells[i].Style.BackColor = Color.Azure;
            dataGridView1.Rows[VolgaAR].Cells[i].Style.BackColor = Color.Azure;
            dataGridView1.Rows[ThetaAR].Cells[i].Style.BackColor = Color.Azure;
            dataGridView1.Rows[DV01_BaseAR].Cells[i].Style.BackColor = Color.Azure;
            dataGridView1.Rows[DV01_TermsAR].Cells[i].Style.BackColor = Color.Azure;



            dataGridView1.Rows[NotionalR + 1].Cells[i].Style.BackColor = Color.LightGray;
            dataGridView1.Rows[NotionalR + 1].Height = 11;
            dataGridView1.Rows[Bid_OfferR + 1].Cells[i].Style.BackColor = Color.LightGray;
            dataGridView1.Rows[Bid_OfferR + 1].Height = 11;
            dataGridView1.Rows[BloombergVolR + 1].Cells[i].Style.BackColor = Color.LightGray;
            dataGridView1.Rows[BloombergVolR + 1].Height = 11;
            dataGridView1.Rows[Deliver_DaysR + 1].Cells[i].Style.BackColor = Color.LightGray;
            dataGridView1.Rows[Deliver_DaysR + 1].Height = 11;
            dataGridView1.Rows[FLYR + 1].Cells[i].Style.BackColor = Color.LightGray;
            dataGridView1.Rows[FLYR + 1].Height = 11;
            dataGridView1.Rows[BbgSourceR + 1].Cells[i].Style.BackColor = Color.LightGray;
            dataGridView1.Rows[BbgSourceR + 1].Height = 11;
            dataGridView1.Rows[BreakEvenR + 1].Cells[i].Style.BackColor = Color.LightGray;
            dataGridView1.Rows[BreakEvenR + 1].Height = 11;
            dataGridView1.Rows[DV01_TermsAR + 1].Cells[i].Style.BackColor = Color.LightGray;
            dataGridView1.Rows[DV01_TermsAR + 1].Height = 11;



        }

        private void formatSurfaceView()
        {

            for (int i = 0; i <= dataGridView8.Columns.Count - 1; i++)
            {
                dataGridView8.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView8.Columns[i].DefaultCellStyle.BackColor = Color.LightGray;

            }


            dataGridView8.Columns[0].DefaultCellStyle.BackColor = Color.DarkGray;
            dataGridView8.Columns[4].DefaultCellStyle.BackColor = Color.Cyan;
            dataGridView8.Columns[6].DefaultCellStyle.BackColor = Color.Cyan;
            dataGridView8.Columns[7].DefaultCellStyle.BackColor = Color.Cyan;
            dataGridView8.Columns[9].DefaultCellStyle.BackColor = Color.Cyan;
            dataGridView8.Columns[10].DefaultCellStyle.BackColor = Color.Cyan;


            for (int i = 0; i <= dataGridView8.Rows.Count - 1; i++)
            {
                dataGridView8.Rows[i].Height = 18;
            }







        }

        int control = 0;

        private void toggleSurfData()
        {
            if (control == 0)
            {
                control = 1;

                // splitContainer1.Panel2Collapsed = false;

                splitContainer1.Panel1Collapsed = true;

            }
            else if (control == 1)
            {
                control = 0;

                //splitContainer1.Panel2Collapsed = true;

                splitContainer1.Panel1Collapsed = false;
            }
        }

        private string cleanStrike(string strikeText)
        {


            string functionReturnValue = strikeText.ToLower();
            int TextLength = strikeText.Length;
            string dateEnd = strikeText.Substring(TextLength - 1);
            int textlen = strikeText.Length;
            int[] arr = Enumerable.Range(1, 100).ToArray();
            List<string> deltaRange = new List<string>();

            foreach (int i in arr)
                deltaRange.Add(i + "d");

            deltaRange.Add("a");


            if (IsNumeric(strikeText) == true)
            {
                return functionReturnValue;
            }
            else if (deltaRange.Contains(functionReturnValue))
            {
                return functionReturnValue;
            }
            else
            {
                //err
                functionReturnValue = "CHECK STRIKE";
            }

            return functionReturnValue;
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            int curRow = dataGridView1.CurrentCell.RowIndex;
            int curCol = dataGridView1.CurrentCell.ColumnIndex;
            string exp = pricer.Rows[ExpiryR][curCol].ToString();
            string strike = pricer.Rows[StrikeR][curCol].ToString();
            string pC = pricer.Rows[Put_CallR][curCol].ToString();


            //user changes ccyPair will load spot and clear current column
            if (curRow == CcyPairR)
            {
                string p = pricer.Rows[curRow][curCol].ToString().ToUpper();

                comboBox1.Text = p;
                comboBox1_SelectedIndexChanged(comboBox1, new EventArgs());

                List<string> crossName = crosses.AsEnumerable().Select(x => x[0].ToString()).ToList();
                int j = crossName.IndexOf(p);

                List<string> cross = fwds.d_data.AsEnumerable().Select(x => x[0].ToString().Substring(0, 6)).ToList();
                int i = cross.IndexOf(p);

                if (j < 0)
                {
                    MessageBox.Show("Check Cross");
                }
                else
                {
                    pricer.Rows[curRow][curCol] = p; //convert ccypair to caps
                    pricer.Rows[SpotR][curCol] = fwds.d_data.Rows[i]["PX_MID"].ToString();//get spot from d_data table
                    displaySurface(p);
                }

                //skip spot row and move to expiry date
                BeginInvoke((Action)delegate
                {
                    DataGridViewCell cell = dataGridView1.Rows[2].Cells[curCol];
                    dataGridView1.CurrentCell = cell;
                });

                //clear bbg vol 
                pricer.Rows[BloombergVolR][curCol] = "";
            }

            if (curRow == Put_CallR)
            {

                optPricer(exp, strike, pC);
            }


            if (curRow == StrikeR)
            {
                string p = pricer.Rows[curRow][curCol].ToString();
                p = cleanStrike(p);
                pricer.Rows[curRow][curCol] = p;
                if (p == "CHECK STRIKE")
                    MessageBox.Show("CHECK STRIKE");

                else
                {

                    optPricer(exp, strike, pC);
                    pricer.Rows[BloombergVolR][curCol] = "";
                }
            }

            if (curRow == Premium_TypeR)
            {

                optPricer(exp, strike, pC);
            }
            if (curRow == NotionalR)
            {
                //check to make sure number is being input.
                string p = pricer.Rows[curRow][curCol].ToString();

                if (IsNumeric(p) == false)
                {
                    MessageBox.Show("Please Check Input. Must be Numeric");

                }
                else
                {
                    dataGridView1.SelectedCells[0].Style.BackColor = Color.Red;
                    pricer.Rows[curRow][curCol] = p;
                    optPricer(exp, strike, pC);

                }

                optPricer(exp, strike, pC);
            }
            if (curRow == ExpiryR)
            {
                string p = pricer.Rows[curRow][curCol].ToString();
                p = AutoExpiryString(p);

                if (p == "CHECK DATE")
                {
                    MessageBox.Show("CHECK DATE");

                }

                else
                {
                    optPricer(exp, strike, pC);
                    pricer.Rows[curRow][curCol] = p;
                    pricer.Rows[BloombergVolR][curCol] = "";
                }


            }

            if (curRow == SpotR || curRow == Swap_PtsR || curRow == FwdR)
            {
                dataGridView1.SelectedCells[0].Style.BackColor = Color.Red;
                optPricer(exp, strike, pC);

            }


            if (curRow == ATM_VOLR || curRow == RRR || curRow == FLYR || curRow == Depo_BaseR || curRow == Depo_TermsR || curRow == VolR || curRow == SystemVolR)
            {

                //check to make sure number is being input.
                string p = pricer.Rows[curRow][curCol].ToString();

                if (IsNumeric(p) == false)
                {
                    MessageBox.Show("Please Check Input. Must be Numeric");

                }
                else
                {
                    dataGridView1.SelectedCells[0].Style.BackColor = Color.Red;
                    p = cellValidationPct(p);
                    pricer.Rows[curRow][curCol] = p;
                    optPricer(exp, strike, pC);

                }
            }

            //if (curRow == BbgSourceR)
            //{
            //    //if user changes bbgSource than source will be copied to crosses datatable and new bbgdata will be called from bbg api
            //    List<string> crossName = crosses.AsEnumerable().Select(x => x[0].ToString()).ToList();
            //    string ccyPair = pricer.Rows[CcyPairR][curCol].ToString().ToUpper();
            //    int i = crossName.IndexOf(ccyPair);
            //    string bbgSource = pricer.Rows[BbgSourceR][curCol].ToString().ToUpper();
            //    pricer.Rows[BbgSourceR][curCol] = bbgSource;

            //    crosses.Rows[i]["Source"] = bbgSource;
            //    dataGridView3.Refresh();

            //    fwds.d_data.Rows.Clear();

            //    foreach (DataRow row in crosses.Rows)
            //    {
            //        if (row["Cross"].ToString() == ccyPair)
            //        {
            //            fwds.d_data.Rows.Add(row["Cross"].ToString() + " " + bbgSource + " CURNCY");
            //        }
            //        else
            //        {
            //            fwds.d_data.Rows.Add(row["Cross"].ToString() + " " + row["Source"].ToString() + " CURNCY");
            //        }

            //    }

            //    fwds.sendRequest();

            //    DataRow lastFwds = fwds.d_data.Rows[fwds.d_data.Rows.Count - 1];     
            //    bool endload = false;

            //    do
            //    {
            //        if (lastFwds["FWD_CURVE"] != DBNull.Value)
            //        {
            //            endload = true;
            //        }
            //    } while (endload == false);

            //    dataSetFwds();

            //    ////sets rateTiles
            //    setMarketData();
            //    refreshData("fwdPts");                
            //    reCalc();

            //}



        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            int curRow = dataGridView1.CurrentCell.RowIndex;
            int curCol = dataGridView1.CurrentCell.ColumnIndex;


            if (e.KeyCode == Keys.Delete)
            {
                if (curRow == SpotR)
                {
                    //user can press delete key on empty column to quickly intiliaze new column of same currency 
                    if (pricer.Rows[CcyPairR][curCol] == DBNull.Value)
                    {
                        pricer.Rows[CcyPairR][curCol] = dataGridView1.Rows[0].Cells[curCol - 1].Value;
                        loadSpot();

                    }

                    else
                    {
                        loadSpot();
                    }
                }


                if (curRow == NotionalR)
                {
                    double flipAmt = Convert.ToDouble(dataGridView1.CurrentCell.Value) * -1;
                    dataGridView1.CurrentCell.Value = flipAmt.ToString("0.00");
                }

                if (curRow == ATM_VOLR)
                { refreshData("atm"); dataGridView1.SelectedCells[0].Style.BackColor = Color.LavenderBlush; }

                if (curRow == RRR)
                { refreshData("rr"); dataGridView1.SelectedCells[0].Style.BackColor = Color.LavenderBlush; }

                if (curRow == FLYR)
                { refreshData("fly"); dataGridView1.SelectedCells[0].Style.BackColor = Color.LavenderBlush; }

                if (curRow == Swap_PtsR)
                { refreshData("fwdPts"); dataGridView1.SelectedCells[0].Style.BackColor = Color.LavenderBlush; }

                if (curRow == FwdR)
                { refreshData("outRight"); dataGridView1.SelectedCells[0].Style.BackColor = Color.LavenderBlush; }

                if (curRow == Depo_BaseR)
                { refreshData("depoB"); dataGridView1.SelectedCells[0].Style.BackColor = Color.LavenderBlush; }

                if (curRow == Depo_TermsR)
                { refreshData("depoT"); dataGridView1.SelectedCells[0].Style.BackColor = Color.LavenderBlush; }

                if (curRow == SystemVolR) { dataGridView1.SelectedCells[0].Style.BackColor = Color.White; refreshData(""); }


                if (curRow == VolR) { dataGridView1.SelectedCells[0].Style.BackColor = Color.White; refreshData(""); }


                if (curRow == CcyPairR)
                {
                    //if delete is press on ccyPair entire coloumn will reset. 
                    dataGridView1.Rows[VolR].Cells[curCol].Style.BackColor = Color.White;
                    dataGridView1.Rows[SpotR].Cells[curCol].Style.BackColor = Color.Cyan;
                    dataGridView1.Rows[NotionalR].Cells[curCol].Style.BackColor = Color.Cyan;
                    pricer.Rows[NotionalR][curCol] = 10;
                    dataGridView1.Rows[Swap_PtsR].Cells[curCol].Style.BackColor = Color.LavenderBlush;
                    dataGridView1.Rows[FwdR].Cells[curCol].Style.BackColor = Color.LavenderBlush;
                    dataGridView1.Rows[ATM_VOLR].Cells[curCol].Style.BackColor = Color.LavenderBlush;
                    dataGridView1.Rows[FLYR].Cells[curCol].Style.BackColor = Color.LavenderBlush;
                    dataGridView1.Rows[RRR].Cells[curCol].Style.BackColor = Color.LavenderBlush;
                    dataGridView1.Rows[Depo_BaseR].Cells[curCol].Style.BackColor = Color.LavenderBlush;
                    dataGridView1.Rows[Depo_TermsR].Cells[curCol].Style.BackColor = Color.LavenderBlush;

                    loadSpot();
                    refreshData("all");
                }

            }


            //this will delete current column
            if (e.Control && e.KeyCode == Keys.P)
            {
                //dataGridView1.Columns.Remove(dataGridView1.Columns[curCol]);

                if (curCol > 0)
                {

                    this.dataGridView1.CellValueChanged -= this.dataGridView1_CellValueChanged;


                    foreach (DataGridViewRow myRow in dataGridView1.Rows)
                    {
                        myRow.Cells[curCol].Value = DBNull.Value; // assuming you want to clear the first column
                        formatColumn();
                    }

                    dataGridView1.Refresh();

                    dataGridView1.CurrentCell = dataGridView1.Rows[curRow].Cells[curCol - 1];

                    this.dataGridView1.CellValueChanged += this.dataGridView1_CellValueChanged;
                }

            }


            //clones current column 
            if (e.Control && e.KeyCode == Keys.Q)
            {

                if (curCol > 0)
                {

                    this.dataGridView1.CellValueChanged -= this.dataGridView1_CellValueChanged;


                    foreach (DataGridViewRow myRow in dataGridView1.Rows)
                    {
                        myRow.Cells[curCol + 1].Value = myRow.Cells[curCol].Value; // assuming you want to clear the first column

                    }

                    dataGridView1.Refresh();

                    dataGridView1.CurrentCell = dataGridView1.Rows[curRow].Cells[curCol + 1];

                    this.dataGridView1.CellValueChanged += this.dataGridView1_CellValueChanged;
                }



            }

            if (e.Control && e.KeyCode == Keys.W)
            {
                toggleSurfData();
            }

            if (e.Control && e.KeyCode == Keys.S)
            {
                broadCast();
                
            }


            if (e.Control && e.KeyCode == Keys.B)
            {
                setStrikeVol();
            }

            if ((e.Control && e.KeyCode == Keys.Space))
            {
                if (dataGridView1.SelectedCells.Count > 1)
                {
                    combineSpreadSim();
                }

                else
                {
                
                string tName = curCol.ToString();
                DataTable dt = new DataTable();
                dt = tradeSim.Tables[tName];
                showTradeSim(dt);
                }
            }

            if ((e.Control && e.KeyCode == Keys.Enter))
            {
                hide_rows();
            }
              

            if ((e.Control && e.KeyCode == Keys.F))
            {

            this.dataGridView1.CellValueChanged -= this.dataGridView1_CellValueChanged;




             string ccyPair = pricer.Rows[CcyPairR][curCol].ToString();
             string pC = pricer.Rows[Put_CallR][curCol].ToString();
             string exp = pricer.Rows[ExpiryR][curCol].ToString();
             string strike = pricer.Rows[StrikeR][curCol].ToString();
             double delta = Convert.ToDouble(strike.Substring(0, strike.Length - 1))/100;
            
             double dayCount = Convert.ToDouble(pricer.Rows[Expiry_DaysR][curCol]);
             double delDayCount = Convert.ToDouble(pricer.Rows[Deliver_DaysR][curCol]);
             DateTime expiry = today.AddDays(dayCount);
             double s = Convert.ToDouble(pricer.Rows[SpotR][curCol]);
             
            

            double[] volComponents = null;

            volComponents = volBuilder(ccyPair, dayCount);
            double atmVol = convertPercent(pricer.Rows[ATM_VOLR][curCol].ToString());
            double rr = convertPercent(pricer.Rows[RRR][curCol].ToString());
            double fly = convertPercent(pricer.Rows[FLYR][curCol].ToString());
    
            double smileFlyMult = volComponents[5];
            double rrMult = volComponents[6];

                 //gets daycount basis from ccydets dataTable

            int[] arr = dayCountBasis(ccyPair);
            int basisB = arr[0];
            int basisT = arr[1];

            // calls method to get cross info 
            object[] crossInfo = crossDtData(ccyPair);
            int volType = Convert.ToInt16(crossInfo[0]);
            double factor = Convert.ToDouble(crossInfo[1]);
            string bbgSource = crossInfo[2].ToString();

            //need to change smile factor for usdrub to control number of strikes are calcuted with cubic spline function
            double smileFactor = factor;
            if (ccyPair == "USDRUB" || ccyPair == "EURRUB" || ccyPair == "USDTRY") { smileFactor = 100; }


            //is needed to convert old delatype to new premo included. old was 1 = premo 2 = no, now 1 = premo 0 = no
            int premoInc = 0;


            if (volType == 1)
            {
                premoInc = 1;
            }
            else
            {
                premoInc = 0;
            }

            double[] rateComponents = rateBuilder(ccyPair, dayCount, delDayCount, s, factor, basisB, basisT);
            //fwdPts, outRight, forDepo, domDepo

            double fwdPts = rateComponents[0];
            double outRight = rateComponents[1];
            double forDepo = rateComponents[2];
            double domDepo = rateComponents[3];

            double dfForExp = DiscountFactor(forDepo, Convert.ToInt16(dayCount), Convert.ToInt16(basisB));
            double dfForDel = DiscountFactor(forDepo, Convert.ToInt16(delDayCount), Convert.ToInt16(basisB));

            double dfDomExp = DiscountFactor(domDepo, Convert.ToInt16(dayCount), Convert.ToInt16(basisT));
            double dfDomDel = DiscountFactor(domDepo, Convert.ToInt16(delDayCount), Convert.ToInt16(basisT));

            double flyVol = marketfly(s, today, expiry, delta, atmVol, atmVol, rr, fly, dfDomExp, dfForExp,  premoInc, smileFlyMult, rrMult, smileFactor);
            flyVol = atmVol + flyVol;
            dataGridView1.Rows[VolR].Cells[curCol].Value = flyVol.ToString("0.00%");
            dataGridView1.Rows[VolR].Cells[curCol].Style.BackColor = Color.Red;
            optPricer(exp, strike, pC);

            this.dataGridView1.CellValueChanged += this.dataGridView1_CellValueChanged;


                
            }
        }

        private void dataGridView8_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            int curCol = dataGridView8.CurrentCell.ColumnIndex;
            int curRow = dataGridView8.CurrentCell.RowIndex;

            string ccyPair = ((DataTable)dataGridView8.DataSource).TableName;
            DataTable dt = pricingData.Tables[ccyPair];
            int rrCol = dt.Columns["25DR"].Ordinal;
            int flyCol = dt.Columns["25D_BrokerFly"].Ordinal;
            int fMultCol = dt.Columns["SmileFly_Multiplier"].Ordinal;
            int rMultCol = dt.Columns["RR_Multiplier"].Ordinal;

            if (curCol == fMultCol || curCol == rMultCol)
            {
                string p = dt.Rows[curRow][curCol].ToString();

                if (IsNumeric(p) == false)
                {
                    MessageBox.Show("Please Check Input. Must be Numeric");
                    return;

                }
                else
                {
                    dataGridView8.SelectedCells[0].Style.BackColor = Color.Red;
                    dt.Rows[curRow][curCol] = p;
                    displaySurface(ccyPair);

                }
            }

            if (curCol == rrCol || curCol == flyCol)
            {


                string p = dt.Rows[curRow][curCol].ToString();


                if (IsNumeric(p) == false)
                {
                    MessageBox.Show("Please Check Input. Must be Numeric");
                    return;

                }
                else
                {
                    dataGridView8.SelectedCells[0].Style.BackColor = Color.Red;
                    p = cellValidationPct(p);
                    dt.Rows[curRow][curCol] = p;
                    displaySurface(ccyPair);

                }

            }




        }

        private void dataGridView1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                //broadCast();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (endIni == 1)
            {
                // setMarketData();
                string cross = comboBox1.Text;


                List<string> crossName = crosses.AsEnumerable().Select(x => x[0].ToString()).ToList();
                int j = crossName.IndexOf(cross);

                string depo = crosses.Rows[j].ItemArray[1].ToString();

                DataTable dt = marketData.Tables[cross];
                DataTable dt1 = deposSet.Tables[depo];

                string tableBase = cross.Substring(0, 3);
                string tableTerms = cross.Substring(3, 3);

                DataTable dtBase = holidaySet.Tables[tableBase];
                DataTable dtTerms = holidaySet.Tables[tableTerms];

                DataTable holsBase = new DataTable();
                holsBase.Columns.Add(tableBase);

                DataTable holsTerms = new DataTable();
                holsTerms.Columns.Add(tableTerms);


                foreach (DataRow row in dtBase.Rows)
                    holsBase.Rows.Add(row["Holiday Date"]);

                foreach (DataRow row in dtTerms.Rows)
                    holsTerms.Rows.Add(row["Holiday Date"]);

                dataGridView5.DataSource = holsBase;
                dataGridView7.DataSource = holsTerms;

                dataGridView4.DataSource = dt;
                dataGridView6.DataSource = dt1;

                dataGridView10.DataSource = smileMult.Tables[cross];

                dataGridView9.DataSource = brokerRun.Tables[cross];



            }

        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

            int curCol = dataGridView1.CurrentCell.ColumnIndex;
            string tName = curCol.ToString();
            DataTable dt = new DataTable();
            dt = tradeSim.Tables[tName];
            showTradeSim(dt);
      
        }

        private void setStrikeVol()
        {
            //refrehese spot from bbg. Use bbg class instance called single spot
            int curCol = dataGridView1.CurrentCell.ColumnIndex;

            string cross = pricer.Rows[CcyPairR][curCol].ToString() + " CURNCY";

            DateTime exp = Convert.ToDateTime(pricer.Rows[ExpiryDateR][curCol].ToString().Substring(4));
            double autostrike = Convert.ToDouble(pricer.Rows[AutoStrikeR][curCol]);

            if (strikeVol.d_data.Rows.Count < 1)
            {
                strikeVol.d_data.Columns.Add("sp vol surf mid");
            }

            strikeVol.d_data.Rows.Clear();
            strikeVol.d_data.Rows.Add(cross);

            foreach (ListViewItem ovr in strikeVol.listViewOverrides.Items)
            {
                ovr.Remove();
            }


            ListViewItem item = strikeVol.listViewOverrides.Items.Add("vol surf delta ovr");
            ListViewItem item1 = strikeVol.listViewOverrides.Items.Add("vol surf strike ovr");
            ListViewItem item2 = strikeVol.listViewOverrides.Items.Add("vol_surf_expiry_ovr");

            string strike = autostrike.ToString("0.0000");
            string expiry = exp.ToString("yyyMMdd");

            item.SubItems.Add("0");
            item1.SubItems.Add(strike);
            item2.SubItems.Add(expiry);

            strikeVol.sendRequest();

            DataRow lastRow = strikeVol.d_data.Rows[strikeVol.d_data.Rows.Count - 1];
            bool endload = false;

            do
            {
                if (lastRow["sp vol surf mid"] != DBNull.Value)
                    endload = true;

            } while (endload == false);

            string vol = strikeVol.d_data.Rows[0]["sp vol surf mid"].ToString();


            pricer.Rows[BloombergVolR][curCol] = vol;

        }

        private void skewAnalysis(string curr)
        {
            //create dataSet 
            if (skewBps == null)
                skewBps = new DataSet();

            DataSet ds = pricingData;
            DataTable dt_surf = ds.Tables[curr];

            string ccy1 = curr.Substring(0, 3);
            string ccy2 = curr.Substring(3, 3);
            string cross = ccy1 + ccy2;

            DataTable st = fwds.d_data;
            List<string> ccyPair1 = fwds.d_data.AsEnumerable().Select(x => x[0].ToString().Substring(0, 6)).ToList();
            int ii = ccyPair1.IndexOf(cross);
            double spot = Convert.ToDouble(st.Rows[ii]["PX_MID"]);

            //  setMarketData();

            string ccyPair = curr;
            string baseCcy = ccyPair.Substring(0, 3);
            string termsCcy = ccyPair.Substring(3, 3);


            //add tenors and deltas

            List<string> mat = new List<string>();
            mat.Add("1d");
            mat.Add("1w");
            mat.Add("2w");
            mat.Add("1m");
            mat.Add("2m");
            mat.Add("3m");
            mat.Add("6m");
            mat.Add("9m");
            mat.Add("1y");
            mat.Add("2y");
            mat.Add("3y");
            mat.Add("4y");
            mat.Add("5y");


            List<string> deltas = new List<string>();
            deltas.Add("5dp");
            deltas.Add("10dp");
            deltas.Add("15dp");
            deltas.Add("20dp");
            deltas.Add("25dp");
            deltas.Add("30dp");
            deltas.Add("35dp");
            deltas.Add("40dp");
            deltas.Add("45dp");
            deltas.Add("ATM");
            deltas.Add("45dc");
            deltas.Add("40dc");
            deltas.Add("35dc");
            deltas.Add("30dc");
            deltas.Add("25dc");
            deltas.Add("20dc");
            deltas.Add("15dc");
            deltas.Add("10dc");
            deltas.Add("5dc");

            DataTable skewDt = new DataTable();

            skewDt.Columns.Add("Term");

            foreach (string x in deltas)
            {
                skewDt.Columns.Add(x);
            }

            int rAtm = 0;
            foreach (string m in mat)
            {
                skewDt.Rows.Add(m);
                skewDt.Rows[rAtm]["ATM"] = 0;
                rAtm++;
            }

            DateTime dayStart = today;
            DateTime sptDate = SpotDate(dayStart, baseCcy, termsCcy);

            int dtRow = 0;
          
            foreach (string s in mat)
            {
                DateTime autoExp = AutoExpiryDate(s, dayStart, homeCcy, baseCcy, termsCcy);
                DateTime delDate = SpotDate(autoExp, baseCcy, termsCcy);
                double dayCount = (autoExp - dayStart).TotalDays; //expiry to trade date
                double delDayCount = (delDate - sptDate).TotalDays; // delivery to spot date             
                //gets daycount basis from ccydets dataTable

                int[] arr = dayCountBasis(ccyPair);
                int basisB = arr[0];
                int basisT = arr[1];

                // calls method to get cross info 
                object[] crossInfo = crossDtData(ccyPair);
                int volType = Convert.ToInt16(crossInfo[0]);
                double factor = Convert.ToDouble(crossInfo[1]);

             


                //is needed to convert old delatype to new premo included. old was 1 = premo 2 = no, now 1 = premo 0 = no
                int premoInc = 0;


                if (volType == 1)
                {
                    premoInc = 1;
                }
                else
                {
                    premoInc = 0;
                }

                double[] rateComponents = rateBuilder(ccyPair, dayCount, delDayCount, spot, factor, basisB, basisT);
                //fwdPts, outRight, forDepo, domDepo

                double fwdPts = rateComponents[0];
                double outRight = rateComponents[1];
                double forDepo = rateComponents[2];
                double domDepo = rateComponents[3];

                double dfForExp = DiscountFactor(forDepo, Convert.ToInt16(dayCount), Convert.ToInt16(basisB));
                double dfForDel = DiscountFactor(forDepo, Convert.ToInt16(delDayCount), Convert.ToInt16(basisB));

                double dfDomExp = DiscountFactor(domDepo, Convert.ToInt16(dayCount), Convert.ToInt16(basisT));
                double dfDomDel = DiscountFactor(domDepo, Convert.ToInt16(delDayCount), Convert.ToInt16(basisT));

                //get fly, rr from pricerData displayed on main pricer screen 
                double[] volComponents = null;

                volComponents = volBuilder(ccyPair, dayCount);
                double atmVol = volComponents[0];
                double rr = volComponents[1];
                double fly = volComponents[2];
                double wingControl = volComponents[3];
                double targetFlyMult = volComponents[4];

                double smileFlyMult = Convert.ToDouble(dt_surf.Rows[dtRow]["SmileFly_Multiplier"]);
                double rrMult = Convert.ToDouble(dt_surf.Rows[dtRow]["RR_Multiplier"]);

                double wingPut = reCalibrateWingControl(atmVol, rr, fly, spot, dayStart, autoExp, dfDomDel, dfForDel, premoInc, smileFlyMult, rrMult, 0);
                double wingCall = reCalibrateWingControl(atmVol, rr, fly, spot, dayStart, autoExp, dfDomDel, dfForDel, premoInc, smileFlyMult, rrMult, 1); 

               

                //need to calc smilefly then get 25d and atm strikes and vols
                double smileFly = equivalentfly(spot, dayStart, autoExp, atmVol * wingControl, atmVol, rr, fly, dfDomDel, dfForDel, premoInc);
                double putVol = atmVol + smileFly - rr / 2;
                double callVol = atmVol + smileFly + rr / 2;

                double putStrike = FXStrikeVol(spot, dayStart, autoExp, 0.25, putVol, dfDomDel, dfForDel, "p", premoInc);
                double callStrike = FXStrikeVol(spot, dayStart, autoExp, 0.25, callVol, dfDomDel, dfForDel, "c", premoInc);
                double atmStrike = FXATMStrike(spot, dayStart, autoExp, atmVol, dfDomDel, dfForDel, premoInc);
                string premoString = "Base %";


                //get premoInc for  the delta solve. Keep in mind that the surface is built with premo included or not from the setup menu. This will ensure that strikes will have the correct vols despite what premo convention is used for an individual option. 

                double[] premoTypeInfo = premoConventions(premoString, spot, 1);
                int premoIncDeltSolve = Convert.ToInt16(premoTypeInfo[3]);

                double autoSt = 0;

                int dtCol = 1;

                foreach (string d in deltas)
                {

                    int len = d.Length;
                    string delString = d.Substring(0, len - 2);
                    string pC = d.Substring(d.Length - 1);

                    if (pC == "p")
                    {
                        wingControl = wingPut;
                    }
                    else
                    {
                        wingControl = wingCall;
                    }

                    if (d == "ATM")
                    { autoSt = atmStrike; }
                    else
                    {
                        double delt = Convert.ToDouble(delString) / 100;
                       // autoSt = FXStrikeDelta(spot, outRight, dayStart, autoExp, delt, putStrike10, putVol10, putStrike, putVol, atmStrike, atmVol, callStrike, callVol, callStrike10, callVol10, dfDomDel, dfForDel, pC, premoIncDeltSolve); ;
                    }

                    double vol = smileInterp(spot, dayStart, autoExp, wingControl * atmVol, autoSt, putStrike, putVol, atmStrike, atmVol, callStrike, callVol, dfDomExp, dfForExp, dfDomDel, dfForDel);

                    premoTypeInfo = premoConventions(premoString, spot, autoSt);
                    double premoConversion = premoTypeInfo[0];//applies this factor greeks to convert in proper units
                    double premoFactor = premoTypeInfo[1];// will convert greeks to correct units in nominal amounts
                    double notionalFactor = premoTypeInfo[2]; //notional factor is equal to spot or 1 - will convert notional to terms currency if % terms or base pips is selected

                    double[] greeks = FXOpts(spot, dayStart, autoExp, autoSt, vol, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);
                    double premium = greeks[0];
                    double delta = greeks[1];
                    delta = delta - premium / spot * premoIncDeltSolve;
                    premium = premium * premoConversion;
                    double gamma = greeks[2];
                    gamma = gamma * spot;
                    double vega = greeks[3];
                    vega = vega / 100 * premoConversion;

                    //    'atmvol greeks'
                    double[] greeks1 = FXOpts(spot, dayStart, autoExp, autoSt, atmVol, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);
                    double premium1 = greeks1[0];
                    double delta1 = greeks1[1];
                    delta1 = delta1 - premium1 / spot * premoIncDeltSolve;
                    premium1 = premium1 * premoConversion;
                    double gamma1 = greeks1[2];
                    gamma1 = gamma1 * spot;
                    double vega1 = greeks1[3];
                    vega1 = vega1 / 100 * premoConversion;

                    double result = (premium - premium1) * 100; //vol * 100;

                    string format = "+#.##;-#.##;0.00";
                    string resForm = result.ToString(format);

                    skewDt.Rows[dtRow][dtCol] = resForm;

                    dtCol++;

                }

                dtRow++;

            }


            skewDt.TableName = cross;

            //check to see if there is a table for current ccy - if so delete old table and add new data else just at new table
            if (skewBps.Tables.Contains(cross))
            {
                skewBps.Tables.Remove(cross);

                skewBps.Tables.Add(skewDt);
            }
            else
            {
                skewBps.Tables.Add(skewDt);
            }

        }

        private void setBbgInterfaceSurface(string ccy)
        {

            if (mxSurface.d_data.Rows.Count < 1)
            {

                addBbgInterface(mxSurface, tabPage6);
                mxSurface.d_data.Columns.Add("PX_LAST");
            }

            mxSurface.d_data.Rows.Clear();

            List<string> mat = new List<string>();
            mat.Add("ON");
            mat.Add("1W");
            mat.Add("2W");
            mat.Add("1M");
            mat.Add("2M");
            mat.Add("3M");
            mat.Add("6M");
            mat.Add("9M");
            mat.Add("1Y");
            mat.Add("2Y");
            mat.Add("3Y");
            mat.Add("4Y");
            mat.Add("5Y");

            foreach (string c in mat)
            {
                string atm = ccy + "V" + c + " SBER CURNCY";
                string FLY10 = ccy + "10B" + c + " SBER CURNCY";
                string FLY25 = ccy + "25B" + c + " SBER CURNCY";
                string RR25 = ccy + "25R" + c + " SBER CURNCY";
                string RR10 = ccy + "10R" + c + " SBER CURNCY";



                mxSurface.d_data.Rows.Add(FLY10);
                mxSurface.d_data.Rows.Add(FLY25);
                mxSurface.d_data.Rows.Add(atm);
                mxSurface.d_data.Rows.Add(RR25);
                mxSurface.d_data.Rows.Add(RR10);
            }

            mxSurface.buttonSendRequest.Enabled = true;
            mxSurface.sendRequest();

        }

        private void showMxVols(DataTable dt)
        {
            WindowsFormsApplication1.mxVols bulkData = new WindowsFormsApplication1.mxVols(dt);
            bulkData.ShowDialog(this);
        }

        private void showRisk(DataTable dt, DataTable dt1)
        {

            //pricer.risk bulkData = new pricer.risk(dt, dt1);
            //bulkData.ShowDialog(this);
        }

        private void showRiskNew(DataSet ds)
        {

            pricer.risk bulkData = new pricer.risk(ds);
            bulkData.ShowDialog(this);
        }
        private void mxVolsDisplay(string ccy)
        {
            List<string> mat = new List<string>();
            mat.Add("ON");
            mat.Add("1W");
            mat.Add("2W");
            mat.Add("1M");
            mat.Add("2M");
            mat.Add("3M");
            mat.Add("6M");
            mat.Add("9M");
            mat.Add("1Y");
            mat.Add("2Y");
            mat.Add("3Y");
            mat.Add("4Y");
            mat.Add("5Y");

            DataTable output = mxVolsDt();
            DataTable input = mxSurface.d_data;
            DataTable engineVol = pricingData.Tables[ccy];


            int j = 0;
            int i = 0;
            foreach (string c in mat)
            {

                double mx10F = Convert.ToDouble(input.Rows[i][1]);
                double mx25F = Convert.ToDouble(input.Rows[i + 1][1]);
                double mxAtm = Convert.ToDouble(input.Rows[i + 2][1]);
                double mx25R = Convert.ToDouble(input.Rows[i + 3][1]);
                double mx10R = Convert.ToDouble(input.Rows[i + 4][1]);

                double p10F = convertPercent(engineVol.Rows[j]["10D_SmileFly"].ToString()) * 100;
                double p25F = convertPercent(engineVol.Rows[j]["25D_SmileFly"].ToString()) * 100;
                double pAtm = convertPercent(engineVol.Rows[j]["ATM"].ToString()) * 100;
                double p25R = convertPercent(engineVol.Rows[j]["25DR"].ToString()) * 100;
                double p10R = convertPercent(engineVol.Rows[j]["10DR"].ToString()) * 100;

                double chg10F = mx10F - p10F;
                double chg25F = mx25F - p25F;
                double chgAtm = mxAtm - pAtm;
                double chg25R = mx25R - p25R;
                double chg10R = mx10R - p10R;

                output.Rows.Add(new Object[] { c, mx10F.ToString("0.00"), mx25F.ToString("0.00"), mxAtm.ToString("0.00"), mx25R.ToString("0.00"), mx10R.ToString("0.00"), c, p10F.ToString("0.00"), p25F.ToString("0.00"), pAtm.ToString("0.00"), p25R.ToString("0.00"), p10R.ToString("0.00"), c, chg10F.ToString("0.00"), chg25F.ToString("0.00"), chgAtm.ToString("0.00"), chg25R.ToString("0.00"), chg10R.ToString("0.00") });

                j++;
                i += 5;
            }


            showMxVols(output);


        }

        #endregion

        #region NewSmileFunctions

        private static double erf(double x)
        {
            //A&S formula 7.1.26
            double a1 = 0.254829592;
            double a2 = -0.284496736;
            double a3 = 1.421413741;
            double a4 = -1.453152027;
            double a5 = 1.061405429;
            double p = 0.3275911;
            x = Math.Abs(x);
            double t = 1 / (1 + p * x);
            //Direct calculation using formula 7.1.26 is absolutely correct
            //But calculation of nth order polynomial takes O(n^2) operations
            //return 1 - (a1 * t + a2 * t * t + a3 * t * t * t + a4 * t * t * t * t + a5 * t * t * t * t * t) * Math.Exp(-1 * x * x);

            //Horner's method, takes O(n) operations for nth order polynomial
            return 1 - ((((((a5 * t + a4) * t) + a3) * t + a2) * t) + a1) * t * Math.Exp(-1 * x * x);
        }

        public static double NORMSDIST(double z)
        {
            double sign = 1;
            if (z < 0) sign = -1;
            return 0.5 * (1.0 + sign * erf(Math.Abs(z) / Math.Sqrt(2)));
        }

        public double NormsDens(double x)
        {
            const double PI = 3.14159265358979;
            return 1 / Math.Sqrt(2 * PI) * Math.Exp(-0.5 * Math.Pow(x, 2));
        }

        public double DiscountFactor(double SimpleYield, int DelDays, int Basis)
        {
            return 1 / (1 + SimpleYield * DelDays / Basis);
        }

        public double[] FXOpts(double s, DateTime today, DateTime Expiry, double strike, double Vol, double Dfe, double Dfe2, double Dfd, double Dfd2, string typeo)
        {
            double[] functionReturnValue = null;

            long verso = 0;
            double TimeExp = 0;
            double d1 = 0;
            double d2 = 0;
            double Nd1 = 0;
            double Nd2 = 0;
            double k = 0;
            double Fw = 0;
            double Fwe = 0;
            double r = 0;

            double Q = 0;
            double sig = 0;

          

            try
            {
                TimeExp = (Expiry - today).TotalDays / 365;

               // if (TimeExp > 1.05) { Dfd2 = 1; }

                //forward price at delivery
                Fw = s * Dfd2 / Dfd;
                //forward price at expiry
                Fwe = s * Dfe2 / Dfe;

               

                k = strike;

                if (typeo == "c")
                {
                    verso = 1;
                }
                else
                {
                    verso = -1;
                }

               

                if (TimeExp == 0)
                {
                    if (typeo == "c" & (s - k) > 0)
                    {
                        functionReturnValue = new double[] { s - k, 1, 0, 0, 0 };
                        return functionReturnValue;
                    }
                    else if (typeo != "c" & (k - s) > 0)
                    {
                        functionReturnValue = new double[] { k - s, -1, 0, 0, 0 };
                        return functionReturnValue;
                    }
                    else
                    {
                        functionReturnValue = new double[] { 0, 0, 0, 0, 0 };
                        return functionReturnValue;
                    }
                }

                r = -Math.Log(Dfe) / TimeExp;
                Q = -Math.Log(Dfe2) / TimeExp;

                sig = Vol;
                d1 = (Math.Log(Fw / k) + 0.5 * (Math.Pow(sig, 2.0)) * TimeExp) / (sig * (Math.Pow(TimeExp, (0.5))));
                d2 = d1 - sig * (Math.Pow(TimeExp, (0.5)));
                Nd1 = NORMSDIST(verso * d1);
                Nd2 = NORMSDIST(verso * d2);


                double price = Dfd * (verso * Fw * Nd1 - verso * k * Nd2);
                double fwd_price = (verso * Fw * Nd1 - verso * k * Nd2);

                double Delta = verso * Nd1 *Dfd2;
                double Fwd_Delta = verso * Nd1;
                double gamma = Dfd2 * NormsDens(d1) / (s * sig * (Math.Pow(TimeExp, (0.5))));
                double Vega = Dfd2 * s * (Math.Pow(TimeExp, (0.5))) * NormsDens(d1);
                double theta = (-0.5 * Math.Pow(sig, 2) * Math.Pow(s, 2) * gamma + r * price - (r - Q) * s * Delta);

                //if (Vol <= 0)
                //{
                //    if (typeo == "c")
                //    {
                //        price = Math.Max((Fw - k) * Dfd,0);
                //        fwd_price = Math.Max((Fw - k),0);
                //    }
                //    else
                //    {
                //        price = Math.Max((k- Fw) * Dfd,0);
                //        fwd_price = Math.Max((k- Fw),0);
                //    }
                    
                   

                //    functionReturnValue = new double[] { price, 0, 0, 0, 0, 0,fwd_price };
                //    return functionReturnValue;
                //}

                //else
                {
                    functionReturnValue = new double[] { price, Delta, gamma, Vega, theta, Fwd_Delta, fwd_price };
                    return functionReturnValue;
                }
              

            }
            catch
            {
                functionReturnValue = new double[] { 0, 0, 0, 0, 0,0,0 };
                return functionReturnValue;
            }



        }

        public double FXStrikeVol(double s, DateTime today, DateTime Expiry, double Delta, double Vol, double dF, double df2, string typeo, int princ)
        {
            double functionReturnValue = 0;

            double K1 = 0;
            double TimeExp = 0;
            double k = 0;
            double sig = 0;

            try
            {
                double Fw = s * df2 / dF;
                k = 0;
                TimeExp = (Expiry - today).TotalDays / 365;
                if (TimeExp > 370.00 / 365.00) { Delta = Delta * df2; } //converts to fwd delta for maturities longer than 1y;

               
                sig = Vol;
                if (typeo == "c")
                {

                  
                    double Kc = Fw * Math.Exp(-sig * Math.Sqrt(TimeExp) * (QNorm(Delta / df2, 0, 1, true, false) - 0.5 * sig * Math.Sqrt(Math.Sqrt(TimeExp))));

                    for (int j = 1; j <= 5 * princ; j++)
                    {
                        K1 = Kc;
                        double[] temp1 = FXOpts(s, today, Expiry, K1, sig, dF, df2, dF, df2, "c");
                        double c = temp1[0];
                        double deltac = temp1[1];
                        double pr = c;
                        temp1 = FXOpts(s, today, Expiry, K1 * 1.000001, sig, dF, df2, dF, df2, "c");
                        double c1 = temp1[0];
                        double deltac1 = temp1[1];
                        double de = (deltac1 - c1 / s - deltac + pr / s) / (K1 * 1E-06);
                        Kc = K1 - ((deltac1 - Delta) - c1 / s) / de;
                    }
                    k = Kc;

                }
                else
                {

                    double KP = Fw * Math.Exp(sig * Math.Sqrt(TimeExp) * (QNorm((Delta) / df2, 0, 1, true, false) + 0.5 * sig * Math.Sqrt(TimeExp)));

                  

                    for (int j = 1; j <= 5 * princ; j++)
                    {
                        K1 = KP;
                        double[] temp1 = FXOpts(s, today, Expiry, K1, sig, dF, df2, dF, df2, "p");
                        double p = temp1[0];
                        double deltap = temp1[1];
                        double pr = p;
                        temp1 = FXOpts(s, today, Expiry, K1 * 1.000001, sig, dF, df2, dF, df2, "p");
                        double p1 = temp1[0];
                        double deltap1 = temp1[1];
                        double de = (deltap1 - p1 / s - deltap + pr / s) / (K1 * 1E-06);
                        KP = K1 - ((deltap1 + Delta) - p1 / s) / de;
                    }
                    k = KP;
                }

                functionReturnValue = k;
                return functionReturnValue;

            }

            catch
            {
                functionReturnValue = 0;
                return functionReturnValue;
            }

        }

        public double FXStrikeDeltaOLD(double s, DateTime today, DateTime Expiry, double Delta, double sigWing, double K25p, double sig25p, double KA, double sigatm, double K25c, double sig25c, double Dfe, double Dfe2, string typeo, int princ)
        {
            double functionReturnValue = 0;
            double TimeExp = 0;
            double Dfd = 0;
            double Dfd2 = 0;
            double sigdc = 0;
            double sigdp = 0;
            double k = 0;

            try
            {

                Dfd = Dfe;
                Dfd2 = Dfe2;

                TimeExp = (Expiry - today).TotalDays / 365;


                double KU = KA;
                if (typeo == "c")
                {
                    double K0 = KA * (1 + 4 * sigatm * Math.Sqrt(TimeExp));
                    for (int i = 1; i <= 70; i++)
                    {
                        k = KU + (K0 - KU) / 2;
                        sigdc = smileInterp(s, today, Expiry, sigWing, k, K25p, sig25p, KA, sigatm, K25c, sig25c, Dfe, Dfe2, Dfd, Dfd2);
                        double[] price = FXOpts(s, today, Expiry, k, sigdc, Dfe, Dfe2, Dfd, Dfd2, typeo);
                        double deltac1 = price[1] - price[0] / s * princ;
                        double esc = deltac1 - Delta;
                        if (esc > 0)
                        {
                            KU = k;
                        }
                        else
                        {
                            K0 = k;
                        }
                        if (Math.Abs(esc) < 1E-05)
                            goto exitloop;

                    }

                }
                else
                {

                    double K0 = KA * (1 - 4 * sigatm * Math.Sqrt(TimeExp));
                    for (int i = 1; i <= 70; i++)
                    {
                        k = KU + (K0 - KU) / 2;
                        sigdp = smileInterp(s, today, Expiry, sigWing, k, K25p, sig25p, KA, sigatm, K25c, sig25c, Dfe, Dfe2, Dfd, Dfd2);
                        double[] price = FXOpts(s, today, Expiry, k, sigdp, Dfe, Dfe2, Dfd, Dfd2, typeo);
                        double deltap1 = price[1] - price[0] / s * princ;
                        double esc = -deltap1 - Delta;
                        if (esc > 0)
                        {
                            KU = k;
                        }
                        else
                        {
                            K0 = k;
                        }
                        if (Math.Abs(esc) < 1E-05)
                            goto exitloop;

                    }
                }



            exitloop:
                functionReturnValue = k;
                return functionReturnValue;
            }

            catch
            {
                functionReturnValue = 0;
                return functionReturnValue;
            }





        }

        public List<double> vKs(double atmVol, double rr, double fly, double rrMult, double smileFlyMult, double spot, DateTime dayStart, DateTime autoExp, double dfDomExp, double dfForExp, double dfDomDel, double dfForDel, int premoInc)
        {


            double fwd = spot * dfForDel / dfDomDel;
            //need to calc smilefly then get 25d and atm strikes and vols
            double smileFly = equivalentfly(spot, dayStart, autoExp, atmVol, atmVol, rr, fly, dfDomDel, dfForDel, premoInc);


            double v10p = atmVol - 0.5 * (rr * rrMult) + (smileFly * smileFlyMult);
            double v25p = atmVol + smileFly - rr / 2;
            double vatm = atmVol;
            double v25c = atmVol + smileFly + rr / 2;
            double v10c = atmVol + 0.5 * (rr * rrMult) + (smileFly * smileFlyMult);
            double k10p = FXStrikeVol(spot, dayStart, autoExp, 0.1, v10p, dfDomDel, dfForDel, "p", premoInc);
            double k25p = FXStrikeVol(spot, dayStart, autoExp, 0.25, v25p, dfDomDel, dfForDel, "p", premoInc);
            double katm = FXATMStrike(spot, dayStart, autoExp, atmVol, dfDomDel, dfForDel, premoInc);
            double k25c = FXStrikeVol(spot, dayStart, autoExp, 0.25, v25c, dfDomDel, dfForDel, "c", premoInc);
            double k10c = FXStrikeVol(spot, dayStart, autoExp, 0.1, v10c, dfDomDel, dfForDel, "c", premoInc);


            List<double> retList = new List<double> { v10p, v25p, vatm, v25c, v10c, k10p, k25p, katm, k25c, k10c };

            return retList;



        }

        public double FXStrikeDelta(double Delta, string typeo, double fwd, double atmVol, double rr, double fly, double rrMult, double smileFlyMult, double spot, DateTime dayStart, DateTime autoExp, double Dfe, double Dfe2, double Dfd, double Dfd2, int premoInc, double splineFactor)
        {
            List<double> vK = vKs(atmVol, rr, fly, rrMult, smileFlyMult, spot, dayStart, autoExp, Dfe, Dfe2, Dfd, Dfd2, premoInc);
            double K10P, K25P, KATM, K25C, K10C, V10P, V25P, VATM, V25C, V10C;
            V10P = vK[0]; V25P = vK[1]; VATM = vK[2]; V25C = vK[3]; V10C = vK[4]; K10P = vK[5]; K25P = vK[6]; KATM = vK[7]; K25C = vK[8]; K10C = vK[9];


            double functionReturnValue = 0;
            double TimeExp = 0;

            double sigdc = 0;
            double sigdp = 0;
            double k = 0;

            try
            {

                Dfd = Dfe;
                Dfd2 = Dfe2;

                double fwdCut = 370.00 / 365.00;

                TimeExp = (autoExp - dayStart).TotalDays / 365;
                if (TimeExp > fwdCut) { Delta = Delta * Dfe2; } //converts to fwd delta for maturities longer than 1y;

                double KU = KATM;
                if (typeo == "c")
                {
                    double K0 = KATM * (1 + 10 * VATM * Math.Sqrt(TimeExp));
                    for (int i = 1; i <= 70; i++)
                    {
                        k = KU + (K0 - KU) / 2;

                        sigdc = fxVol(k, atmVol, rr, fly, rrMult, smileFlyMult, spot, dayStart, autoExp, Dfe, Dfe2, Dfd, Dfd2, premoInc, splineFactor);

                        double[] price = FXOpts(spot, dayStart, autoExp, k, sigdc, Dfe, Dfe2, Dfd, Dfd2, typeo);
                        double deltac1 = price[1] - price[0] / spot * premoInc;
                        double esc = deltac1 - Delta;
                        if (esc > 0)
                        {
                            KU = k;
                        }
                        else
                        {
                            K0 = k;
                        }
                        if (Math.Abs(esc) < 1E-05)
                            goto exitloop;

                    }

                }
                else
                {

                    double K0 = KATM * (1 - 10 * VATM * Math.Sqrt(TimeExp));
                    for (int i = 1; i <= 70; i++)
                    {
                        k = KU + (K0 - KU) / 2;

                        sigdp = fxVol(k, atmVol, rr, fly, rrMult, smileFlyMult, spot, dayStart, autoExp, Dfe, Dfe2, Dfd, Dfd2, premoInc, splineFactor);

                        double[] price = FXOpts(spot, dayStart, autoExp, k, sigdp, Dfe, Dfe2, Dfd, Dfd2, typeo);
                        double deltap1 = price[1] - price[0] / spot * premoInc;
                        double esc = -deltap1 - Delta;
                        if (esc > 0)
                        {
                            KU = k;
                        }
                        else
                        {
                            K0 = k;
                        }
                        if (Math.Abs(esc) < 1E-05)
                            goto exitloop;

                    }
                }



            exitloop:
                functionReturnValue = k;
                return functionReturnValue;
            }

            catch
            {
                functionReturnValue = 0;
                return functionReturnValue;
            }



        }


        public double FXStrikeDeltaOLDD(double s, double fwd, DateTime today, DateTime Expiry, double Delta, double K10p, double sig10p, double K25p, double sig25p, double KA, double sigatm, double K25c, double sig25c, double K10c, double sig10c, double Dfe, double Dfe2, string typeo, int princ, double factor)
        {
            double functionReturnValue = 0;
            double TimeExp = 0;
            double Dfd = 0;
            double Dfd2 = 0;
            double sigdc = 0;
            double sigdp = 0;
            double k = 0;


            try
            {

                Dfd = Dfe;
                Dfd2 = Dfe2;

                double fwdCut = 370.00 / 365.00;

                TimeExp = (Expiry - today).TotalDays / 365;
                if (TimeExp > fwdCut) { Delta = Delta * Dfe2; } //converts to fwd delta for maturities longer than 1y;

                double KU = KA;
                if (typeo == "c")
                {
                    double K0 = KA * (1 + 10 * sigatm * Math.Sqrt(TimeExp));
                    for (int i = 1; i <= 70; i++)
                    {
                        k = KU + (K0 - KU) / 2;
      
                        sigdc = smileSpline(K10p, K25p, KA, K25c, K10c, sig10p, sig25p, sigatm, sig25c, sig10c, fwd, k,factor);

                        double[] price = FXOpts(s, today, Expiry, k, sigdc, Dfe, Dfe2, Dfd, Dfd2, typeo);
                        double deltac1 = price[1] - price[0] / s * princ;
                        double esc = deltac1 - Delta;
                        if (esc > 0)
                        {
                            KU = k;
                        }
                        else
                        {
                            K0 = k;
                        }
                        if (Math.Abs(esc) < 1E-05)
                            goto exitloop;

                    }

                }
                else
                {

                    double K0 = KA * (1 - 10 * sigatm * Math.Sqrt(TimeExp));
                    for (int i = 1; i <= 70; i++)
                    {
                        k = KU + (K0 - KU) / 2;
                        sigdp = smileSpline(K10p, K25p, KA, K25c, K10c, sig10p, sig25p, sigatm, sig25c, sig10c, fwd, k, factor); 
                        double[] price = FXOpts(s, today, Expiry, k, sigdp, Dfe, Dfe2, Dfd, Dfd2, typeo);
                        double deltap1 = price[1] - price[0] / s * princ;
                        double esc = -deltap1 - Delta;
                        if (esc > 0)
                        {
                            KU = k;
                        }
                        else
                        {
                            K0 = k;
                        }
                        if (Math.Abs(esc) < 1E-05)
                            goto exitloop;

                    }
                }



            exitloop:
                functionReturnValue = k;
                return functionReturnValue;
            }

            catch
            {
                functionReturnValue = 0;
                return functionReturnValue;
            }





        }

        public double FXImpvol(double s, DateTime today, DateTime Expiry, double strike, double price, double dF, double df2, string typeo)
        {

            double functionReturnValue = 0;
            double sig = 0;
            double premio = 0;
            int MAX = 0;
            int i = 0;
            double k = 0;
            double ConvergeCrit = 0;
            double esc = 0;

            try
            {

                if (price <= 0)
                {
                    functionReturnValue = 0;
                    return functionReturnValue;
                }
                else
                {
                    MAX = 35;
                    ConvergeCrit = 0.000001;
                    double Fw = s * df2 / dF;
                    k = strike;

                    //starting point set a 15%, for smarter choices see chapter 2

                    sig = 0.15;
                    if ((sig == 0))
                        sig = 0.15;
                    esc = 1;
                    i = 1;

                    while (Math.Abs(esc) > ConvergeCrit)
                    {
                        if ((i > MAX)) { functionReturnValue = 0; return functionReturnValue; }

                        double[] opt = FXOpts(s, today, Expiry, k, sig, dF, df2, dF, df2, typeo);
                        premio = opt[0];
                        double Vega = opt[3];
                        if ((Vega <= 0.00000001))
                        {
                            functionReturnValue = 0;
                            return functionReturnValue;
                        }

                        esc = (premio - price);
                        sig = sig - esc / Vega;
                        i = i + 1;

                    }

                    functionReturnValue = sig;
                    return functionReturnValue;

                }
            }
            catch
            {
                functionReturnValue = 0;
                return functionReturnValue;
            }
        }

        public double smileInterp(double s, DateTime today, DateTime Expiry, double vWing, double k, double K1, double v1, double K2, double v2, double K3, double v3, double Dfe, double Dfe2, double Dfd, double Dfd2)
        {
            double functionReturnValue = 0;
            double sig = 0;

            try
            {

                double TimeExp = (Expiry - today).TotalDays / 365;
                double Fw = s * Dfe2 / Dfe;

                double x1 = (Math.Log(K2 / k) * Math.Log(K3 / k)) / (Math.Log(K2 / K1) * Math.Log(K3 / K1));
                double x2 = (Math.Log(k / K1) * Math.Log(K3 / k)) / (Math.Log(K2 / K1) * Math.Log(K3 / K2));
                double x3 = (Math.Log(k / K1) * Math.Log(k / K2)) / (Math.Log(K3 / K1) * Math.Log(K3 / K2));
                double d1 = (Math.Log(Fw / k) + 0.5 * Math.Pow(vWing, 2) * TimeExp) / (vWing * Math.Sqrt(TimeExp));
                double d2 = d1 - vWing * Math.Sqrt(TimeExp);


                double d11 = (Math.Log(Fw / K1) + 0.5 * Math.Pow(vWing, 2) * TimeExp) / (vWing * Math.Sqrt(TimeExp));
                double d21 = d11 - vWing * Math.Sqrt(TimeExp);
                double d12 = (Math.Log(Fw / K2) + 0.5 * Math.Pow(vWing, 2) * TimeExp) / (vWing * Math.Sqrt(TimeExp));
                double d22 = d12 - vWing * Math.Sqrt(TimeExp);
                double d13 = (Math.Log(Fw / K3) + 0.5 * Math.Pow(vWing, 2) * TimeExp) / (vWing * Math.Sqrt(TimeExp));
                double d23 = d13 - vWing * Math.Sqrt(TimeExp);

                double dk1 = x1 * v1 + x2 * v2 + x3 * v3 - vWing;

                double dk2 = x1 * d11 * d21 * Math.Pow((v1 - vWing), 2) + x3 * d13 * d23 * Math.Pow((v3 - vWing), 2) + x2 * d12 * d22 * Math.Pow((v2 - vWing), 2);


                if ((Math.Pow(vWing, 2) + d1 * d2 * (2 * vWing * dk1 + dk2)) < 0)
                {
                    if (k > K2)
                    {
                        sig = FXImpvol(s, today, Expiry, k, 1E-05 * s, Dfe, Dfe2, "c");
                    }
                    else
                    {
                        sig = FXImpvol(s, today, Expiry, k, 1E-05 * s, Dfe, Dfe2, "p");
                    }

                    functionReturnValue = sig;
                    return functionReturnValue;
                }

                sig = vWing + (-vWing + Math.Sqrt(Math.Pow(vWing, 2) + d1 * d2 * (2 * vWing * dk1 + dk2))) / (d1 * d2);

                functionReturnValue = sig;
                return functionReturnValue;

            }
            catch
            {
                functionReturnValue = 0;
                return functionReturnValue;
            }

        }

        private double smileSpline(double p10, double p25, double atm, double c25, double c10, double pv10, double pv25, double vatm, double cv25, double cv10, double F, double target, double factor)
        {
            double retVal = 0;
            int n = 5;
            float[] logSF = new float[n];
            float[] strikeVol = new float[n];


            logSF[0] = Convert.ToSingle(Math.Log(p10 / F));
            strikeVol[0] = Convert.ToSingle(pv10);

            logSF[1] = Convert.ToSingle(Math.Log(p25 / F));
            strikeVol[1] = Convert.ToSingle(pv25);

            logSF[2] = Convert.ToSingle(Math.Log(atm / F));
            strikeVol[2] = Convert.ToSingle(vatm);

            logSF[3] = Convert.ToSingle(Math.Log(c25 / F));
            strikeVol[3] = Convert.ToSingle(cv25);

            logSF[4] = Convert.ToSingle(Math.Log(c10 / F));
            strikeVol[4] = Convert.ToSingle(cv10);



            double lbound = p10;
            double ubound = c10;
            double tVal = Math.Log(target / F);

            double step = 1 / factor;

            int nStrikes = Convert.ToInt16((ubound - lbound) / step) + 50;

            double strike = lbound;

            float[] logFF = new float[nStrikes];

            for (int q = 0; q < nStrikes; q++)
            {

                logFF[q] = Convert.ToSingle(Math.Log(strike / F));

                strike = strike + step;
            }


            TestMySpline.CubicSpline spline = new TestMySpline.CubicSpline();

            float[] ys = spline.FitAndEval(logSF, strikeVol, logFF);


            if (target >= lbound && target <= ubound)
            {
                int arrayNum = 0;

                do
                {
                    arrayNum++;
                } while (tVal >= logFF[arrayNum]);

                retVal = ys[Convert.ToInt16(arrayNum)];
            }

            else
            {
                List<double> logStrike = new List<double>();
                foreach (double d in logFF)
                {
                    logStrike.Add(d);
                }

                List<double> strikeVolList = new List<double>();

                foreach (double d in ys)
                {
                    strikeVolList.Add(d);
                }

                retVal = LinearInterp(logStrike, strikeVolList, tVal);
            }

            return retVal;


        }

        private double combinedInterp(double s, DateTime today, DateTime Expiry, double vWing, double k,  double K10P, double V10P, double K25P, double V25P, double KATM, double VATM, double K25C, double V25C, double K10C, double V10C, double Dfe, double Dfe2, double Dfd, double Dfd2, double F, double factor)
        {
            double vol = 0;

            if (k >= K25P && k <= K25C)
            {
                vol = smileInterp(s, today, Expiry, vWing, k, K25P, V25P, KATM, VATM, K25C, V25C, Dfe, Dfe2, Dfd, Dfd2);
            }

            else
            {
                vol = smileSpline(K10P, K25P, KATM, K25C, K10C, V10P, V25P, VATM, V25C, V10C, F, k, factor);
            }

            return vol;
        }

        public double FXATMStrike(double s, DateTime today, DateTime Expiry, double Vol, double dF, double df2, int deltaprinc)
        {
            double functionReturnValue = 0;
            double TimeExp = 0;
            double k = 0;

            // ERROR: Not supported in C#: OnErrorStatement

            try
            {
                double Fw = s * df2 / dF;
                double pif = deltaprinc;

                TimeExp = (Expiry - today).TotalDays / 365;

                k = Fw * Math.Exp((0.5 - pif) * Math.Pow(Vol, 2) * TimeExp);

                functionReturnValue = k;
                return functionReturnValue;
            }

            catch
            {
                functionReturnValue = 0;
                return functionReturnValue;

            }


        }

        public double equivalentfly(double s, DateTime today, DateTime Expiry, double sigWing, double sigatm, double rr, double Bfly, double Dfe, double Dfe2, int premimuincluded)
        {
            double functionReturnValue = 0;

            double TimeExp = 0;
            double KA = 0;
            double K25c = 0;
            double K25p = 0;
            double K25bc = 0;
            double K25bp = 0;
            double Dfd = 0;
            double Dfd2 = 0;
            double sig25c = 0;
            double sig25p = 0;
            double sig25b = 0;
            double sig25bc = 0;
            double sig25bp = 0;

            try
            {


                Dfd = Dfe;
                Dfd2 = Dfe2;
                TimeExp = (Expiry - today).TotalDays / 365;
                double Fw = s * Dfe2 / Dfe;

                if (TimeExp < 2)
                {
                    KA = Fw * Math.Exp((0.5 - premimuincluded) * Math.Pow(sigatm, 2) * TimeExp);
                }
                else
                {
                    KA = Fw;
                }

                sig25p = sigatm + Bfly - rr / 2;
                sig25c = sigatm + Bfly + rr / 2;


                //market covention for the fly volatility
                sig25b = sigatm + Bfly;

                //Formula 2.46 chapter 2, for 25D call
                K25c = FXStrikeVol(s, today, Expiry, 0.25, sig25c, Dfe, Dfe2, "c", premimuincluded);

                //same formula and procedure as above for 25D put
                K25p = FXStrikeVol(s, today, Expiry, 0.25, sig25p, Dfe, Dfe2, "p", premimuincluded);

                K25bc = FXStrikeVol(s, today, Expiry, 0.25, sig25b, Dfe, Dfe2, "c", premimuincluded);

                K25bp = FXStrikeVol(s, today, Expiry, 0.25, sig25b, Dfe, Dfe2, "p", premimuincluded);

                sig25bc = smileInterp(s, today, Expiry, sigWing, K25bc, K25p, sig25p, KA, sigatm, K25c,
                sig25c, Dfe, Dfe2, Dfd, Dfd2);
                sig25bp = smileInterp(s, today, Expiry, sigWing, K25bp, K25p, sig25p, KA, sigatm, K25c,
                sig25c, Dfe, Dfe2, Dfd, Dfd2);

                double[] call25c = FXOpts(s, today, Expiry, K25bc, sig25bc, Dfe, Dfe2, Dfd, Dfd2, "c");
                double[] call25b = FXOpts(s, today, Expiry, K25bc, sig25b, Dfe, Dfe2, Dfd, Dfd2, "c");
                double[] put25p = FXOpts(s, today, Expiry, K25bp, sig25bp, Dfe, Dfe2, Dfd, Dfd2, "p");
                double[] put25b = FXOpts(s, today, Expiry, K25bp, sig25b, Dfe, Dfe2, Dfd, Dfd2, "p");

                double f0 = (call25c[0] + put25p[0]) - (call25b[0] + put25b[0]);

                if (Math.Abs(f0) < 0.0000001 * s)
                {
                    functionReturnValue = Bfly;
                    return functionReturnValue;
                }

                Bfly = Bfly + 0.0001;
                double dfly = 0.0001;

                //while loop starts here
                while (Math.Abs(f0) > 0.0000001 * s)
                {
                    sig25p = sigatm + Bfly - rr / 2;
                    sig25c = sigatm + Bfly + rr / 2;

                    K25c = FXStrikeVol(s, today, Expiry, 0.25, sig25c, Dfe, Dfe2, "c", premimuincluded);

                    K25p = FXStrikeVol(s, today, Expiry, 0.25, sig25p, Dfe, Dfe2, "p", premimuincluded);

                    K25bc = FXStrikeVol(s, today, Expiry, 0.25, sig25b, Dfe, Dfe2, "c", premimuincluded);

                    K25bp = FXStrikeVol(s, today, Expiry, 0.25, sig25b, Dfe, Dfe2, "p", premimuincluded);

                    sig25bc = smileInterp(s, today, Expiry, sigWing, K25bc, K25p, sig25p, KA, sigatm, K25c,
                    sig25c, Dfe, Dfe2, Dfd, Dfd2);
                    sig25bp = smileInterp(s, today, Expiry, sigWing, K25bp, K25p, sig25p, KA, sigatm, K25c,
                    sig25c, Dfe, Dfe2, Dfd, Dfd2);

                    call25c = FXOpts(s, today, Expiry, K25bc, sig25bc, Dfe, Dfe2, Dfd, Dfd2, "c");
                    call25b = FXOpts(s, today, Expiry, K25bc, sig25b, Dfe, Dfe2, Dfd, Dfd2, "c");
                    put25p = FXOpts(s, today, Expiry, K25bp, sig25bp, Dfe, Dfe2, Dfd, Dfd2, "p");
                    put25b = FXOpts(s, today, Expiry, K25bp, sig25b, Dfe, Dfe2, Dfd, Dfd2, "p");

                    double F = (call25c[0] + put25p[0]) - (call25b[0] + put25b[0]);
                    double dF = (F - f0) / dfly;

                    Bfly = Bfly - F / dF;
                    dfly = -F / dF;
                    f0 = F;
                }
                //while loop ends here


                functionReturnValue = Bfly;
                return functionReturnValue;
            }
            catch
            {
                functionReturnValue = Convert.ToDouble("-");
                return functionReturnValue;

            }
        }

        public double equivalentflyNew(double s, DateTime today, DateTime Expiry, double sigWing, double sigatm, double rr, double Bfly, double Dfe, double Dfe2, int premimuincluded)
        {
            double functionReturnValue = 0;

            double TimeExp = 0;
            double KA = 0;
            double K25c = 0;
            double K25p = 0;
            double K25bc = 0;
            double K25bp = 0;
            double Dfd = 0;
            double Dfd2 = 0;
            double sig25c = 0;
            double sig25p = 0;
            double sig25b = 0;
            double sig25bc = 0;
            double sig25bp = 0;

            try
            {


                Dfd = Dfe;
                Dfd2 = Dfe2;
                TimeExp = (Expiry - today).TotalDays / 365;
                double Fw = s * Dfe2 / Dfe;

                if (TimeExp < 2)
                {
                    KA = Fw * Math.Exp((0.5 - premimuincluded) * Math.Pow(sigatm, 2) * TimeExp);
                }
                else
                {
                    KA = Fw;
                }

                sig25p = sigatm + Bfly - rr / 2;
                sig25c = sigatm + Bfly + rr / 2;


                //market covention for the fly volatility
                sig25b = sigatm + Bfly;

                //Formula 2.46 chapter 2, for 25D call
                K25c = FXStrikeVol(s, today, Expiry, 0.25, sig25c, Dfe, Dfe2, "c", premimuincluded);

                //same formula and procedure as above for 25D put
                K25p = FXStrikeVol(s, today, Expiry, 0.25, sig25p, Dfe, Dfe2, "p", premimuincluded);

                K25bc = FXStrikeVol(s, today, Expiry, 0.25, sig25b, Dfe, Dfe2, "c", premimuincluded);

                K25bp = FXStrikeVol(s, today, Expiry, 0.25, sig25b, Dfe, Dfe2, "p", premimuincluded);

                sig25bc = smileInterp(s, today, Expiry, sigWing, K25bc, K25p, sig25p, KA, sigatm, K25c,
                sig25c, Dfe, Dfe2, Dfd, Dfd2);
                sig25bp = smileInterp(s, today, Expiry, sigWing, K25bp, K25p, sig25p, KA, sigatm, K25c,
                sig25c, Dfe, Dfe2, Dfd, Dfd2);

                double[] call25c = FXOpts(s, today, Expiry, K25bc, sig25bc, Dfe, Dfe2, Dfd, Dfd2, "c");
                double[] call25b = FXOpts(s, today, Expiry, K25bc, sig25b, Dfe, Dfe2, Dfd, Dfd2, "c");
                double[] put25p = FXOpts(s, today, Expiry, K25bp, sig25bp, Dfe, Dfe2, Dfd, Dfd2, "p");
                double[] put25b = FXOpts(s, today, Expiry, K25bp, sig25b, Dfe, Dfe2, Dfd, Dfd2, "p");

                double f0 = (call25c[0] + put25p[0]) - (call25b[0] + put25b[0]);

                if (Math.Abs(f0) < 0.0000001 * s)
                {
                    functionReturnValue = Bfly;
                    return functionReturnValue;
                }

                Bfly = Bfly + 0.0001;
                double dfly = 0.0001;

                //while loop starts here
                while (Math.Abs(f0) > 0.0000001 * s)
                {
                    sig25p = sigatm + Bfly - rr / 2;
                    sig25c = sigatm + Bfly + rr / 2;

                    K25c = FXStrikeVol(s, today, Expiry, 0.25, sig25c, Dfe, Dfe2, "c", premimuincluded);

                    K25p = FXStrikeVol(s, today, Expiry, 0.25, sig25p, Dfe, Dfe2, "p", premimuincluded);

                    K25bc = FXStrikeVol(s, today, Expiry, 0.25, sig25b, Dfe, Dfe2, "c", premimuincluded);

                    K25bp = FXStrikeVol(s, today, Expiry, 0.25, sig25b, Dfe, Dfe2, "p", premimuincluded);

                    sig25bc = smileInterp(s, today, Expiry, sigWing, K25bc, K25p, sig25p, KA, sigatm, K25c,
                    sig25c, Dfe, Dfe2, Dfd, Dfd2);
                    sig25bp = smileInterp(s, today, Expiry, sigWing, K25bp, K25p, sig25p, KA, sigatm, K25c,
                    sig25c, Dfe, Dfe2, Dfd, Dfd2);

                    call25c = FXOpts(s, today, Expiry, K25bc, sig25bc, Dfe, Dfe2, Dfd, Dfd2, "c");
                    call25b = FXOpts(s, today, Expiry, K25bc, sig25b, Dfe, Dfe2, Dfd, Dfd2, "c");
                    put25p = FXOpts(s, today, Expiry, K25bp, sig25bp, Dfe, Dfe2, Dfd, Dfd2, "p");
                    put25b = FXOpts(s, today, Expiry, K25bp, sig25b, Dfe, Dfe2, Dfd, Dfd2, "p");

                    double F = (call25c[0] + put25p[0]) - (call25b[0] + put25b[0]);
                    double dF = (F - f0) / dfly;

                    Bfly = Bfly - F / dF;
                    dfly = -F / dF;
                    f0 = F;
                }
                //while loop ends here


                functionReturnValue = Bfly;
                return functionReturnValue;
            }
            catch
            {
                functionReturnValue = Convert.ToDouble("-");
                return functionReturnValue;

            }
        }

        public double marketflyOLD(double s, DateTime today, DateTime Expiry, double Delta, double sigWing, double sigatm, double rr, double fly, double Dfe, double Dfe2,
        int premimuincluded, double smileFlyMult, double rrMult)
        {
            double functionReturnValue = 0;

            double TimeExp = 0;
            double KA = 0;
            double K25c = 0;
            double K25p = 0;
            double Kdc = 0;
            double Kdp = 0;
            double sigdp = 0;
            double Kdbc = 0;
            double Kdbp = 0;
            double Dfd = 0;
            double Dfd2 = 0;
            double sig25c = 0;
            double sig25p = 0;
            double sig25b = 0;
            double sigdc = 0;
            double sigdb = 0;
            double sigdbc = 0;
            double sigdbp = 0;
            double Bfly = equivalentfly(s, today, Expiry, sigatm, sigatm, rr, fly, Dfe, Dfe2, premimuincluded);

            try
            {
                Dfd = Dfe;
                Dfd2 = Dfe2;
                TimeExp = (Expiry - today).TotalDays / 365;
                double Fw = s * Dfe2 / Dfe;


                KA = FXATMStrike(s, today, Expiry, sigatm, Dfe, Dfe2, premimuincluded);

                sig25p = sigatm + Bfly - rr / 2;
                sig25c = sigatm + Bfly + rr / 2;

                
                sig25b = sigatm + fly;

                //Formula 2.46 chapter 2, for 25D call
                K25c = FXStrikeVol(s, today, Expiry, 0.25, sig25c, Dfe, Dfe2, "c", premimuincluded);

                //same formula and procedure as above for 25D put
                K25p = FXStrikeVol(s, today, Expiry, 0.25, sig25p, Dfe, Dfe2, "p", premimuincluded);


                sigdc = sig25c;
                sigdp = sig25p;
 

                double wingPut = reCalibrateWingControl(sigatm, rr, fly, s, today, Expiry, Dfe, Dfe2, premimuincluded, smileFlyMult, rrMult,0) * sigatm;

                double wingCall = reCalibrateWingControl(sigatm, rr, fly, s, today, Expiry, Dfe, Dfe2, premimuincluded, smileFlyMult, rrMult,1) * sigatm;
                



                for (int i = 1; i <= 5; i++)
                {

                    Kdc = FXStrikeVol(s, today, Expiry, Delta, sigdc, Dfe, Dfe2, "c", premimuincluded);

                    Kdp = FXStrikeVol(s, today, Expiry, Delta, sigdp, Dfe, Dfe2, "p", premimuincluded);

                    sigdc = smileInterp(s, today, Expiry, wingCall, Kdc, K25p, sig25p, KA, sigatm, K25c,
                    sig25c, Dfe, Dfe2, Dfd, Dfd2);

                    sigdp = smileInterp(s, today, Expiry, wingPut, Kdp, K25p, sig25p, KA, sigatm, K25c,
                    sig25c, Dfe, Dfe2, Dfd, Dfd2);
                }

                sigdb = (sigdc + sigdp) / 2;


                Kdbc = FXStrikeVol(s, today, Expiry, Delta, sigdb, Dfe, Dfe2, "c", premimuincluded);

                Kdbp = FXStrikeVol(s, today, Expiry, Delta, sigdb, Dfe, Dfe2, "p", premimuincluded);

                sigdbc = smileInterp(s, today, Expiry, wingCall, Kdbc, K25p, sig25p, KA, sigatm, K25c,
                sig25c, Dfe, Dfe2, Dfd, Dfd2);
                sigdbp = smileInterp(s, today, Expiry, wingPut, Kdbp, K25p, sig25p, KA, sigatm, K25c,
                sig25c, Dfe, Dfe2, Dfd, Dfd2);

                double[] calldc = FXOpts(s, today, Expiry, Kdbc, sigdc, Dfe, Dfe2, Dfd, Dfd2, "c");
                double[] calldb = FXOpts(s, today, Expiry, Kdbc, sigdb, Dfe, Dfe2, Dfd, Dfd2, "c");
                double[] putdp = FXOpts(s, today, Expiry, Kdbp, sigdp, Dfe, Dfe2, Dfd, Dfd2, "p");
                double[] putdb = FXOpts(s, today, Expiry, Kdbp, sigdb, Dfe, Dfe2, Dfd, Dfd2, "p");

                double f0 = (calldc[0] + putdp[0]) - (calldb[0] + putdb[0]);

                double dfly = sigdb - sigatm;

                if (Math.Abs(f0) < 0.0000001 * s)
                {
                    functionReturnValue = dfly;
                    return functionReturnValue;
                }

                sigdb = sigdb + 0.00005;

                int j = 0;

                //loop
                double dsig = 0.00005;

                while (Math.Abs(f0) > 0.00000001 * s)
                {
                    j = j + 1;

                    Kdbc = FXStrikeVol(s, today, Expiry, Delta, sigdb, Dfe, Dfe2, "c", premimuincluded);

                    Kdbp = FXStrikeVol(s, today, Expiry, Delta, sigdb, Dfe, Dfe2, "p", premimuincluded);

                    sigdbc = smileInterp(s, today, Expiry, wingCall, Kdbc, K25p, sig25p, KA, sigatm, K25c, sig25c, Dfe, Dfe2, Dfd, Dfd2);
                    sigdbp = smileInterp(s, today, Expiry, wingPut, Kdbp, K25p, sig25p, KA, sigatm, K25c, sig25c, Dfe, Dfe2, Dfd, Dfd2);

                    calldc = FXOpts(s, today, Expiry, Kdbc, sigdbc, Dfe, Dfe2, Dfd, Dfd2, "c");
                    calldb = FXOpts(s, today, Expiry, Kdbc, sigdb, Dfe, Dfe2, Dfd, Dfd2, "c");
                    putdp = FXOpts(s, today, Expiry, Kdbp, sigdbp, Dfe, Dfe2, Dfd, Dfd2, "p");
                    putdb = FXOpts(s, today, Expiry, Kdbp, sigdb, Dfe, Dfe2, Dfd, Dfd2, "p");

                    double F = (calldc[0] + putdp[0]) - (calldb[0] + putdb[0]);
                    double dF = (F - f0) / dsig;

                    sigdb = sigdb - F / dF;
                    dsig = -F / dF;
                    f0 = F;

                    if (j * 0.00001 >= dfly)
                    {
                        functionReturnValue = sigdb - sigatm;
                        return functionReturnValue;
                    }


                }

                functionReturnValue = sigdb - sigatm;
                return functionReturnValue;


            }
            catch
            {

                functionReturnValue = sigdb - sigatm;
                return functionReturnValue;
            }



        }

        public double marketfly(double s, DateTime today, DateTime Expiry, double Delta, double sigWing, double sigatm, double rr, double fly, double Dfe, double Dfe2,
       int premimuincluded, double smileFlyMult, double rrMult, double factor)
        {
            double functionReturnValue = 0;

            double TimeExp = 0;
            double KA = 0;
            double K25c = 0;
            double K25p = 0;
         
            double sigdp = 0;
            double Kdbc = 0;
            double Kdbp = 0;
            double Dfd = 0;
            double Dfd2 = 0;
            double sig25c = 0;
            double sig25p = 0;
            double sig25b = 0;
            double sigdc = 0;
            double sigdb = 0;
            double sigdbc = 0;
            double sigdbp = 0;
            double Bfly = equivalentfly(s, today, Expiry, sigatm, sigatm, rr, fly, Dfe, Dfe2, premimuincluded);

          

            try
            {
                Dfd = Dfe;
                Dfd2 = Dfe2;
                TimeExp = (Expiry - today).TotalDays / 365;
                double Fw = s * Dfe2 / Dfe;

                KA = FXATMStrike(s, today, Expiry, sigatm, Dfe, Dfe2, premimuincluded);

                sig25p = sigatm + Bfly - rr / 2;
                sig25c = sigatm + Bfly + rr / 2;
                sig25b = sigatm + fly;      
                K25c = FXStrikeVol(s, today, Expiry, 0.25, sig25c, Dfe, Dfe2, "c", premimuincluded);          
                K25p = FXStrikeVol(s, today, Expiry, 0.25, sig25p, Dfe, Dfe2, "p", premimuincluded);

                double sig10C = sigatm + 0.5 * (rr * rrMult) + (Bfly * smileFlyMult);
                double K10c = FXStrikeVol(s, today, Expiry, 0.1, sig10C, Dfe, Dfe2, "c", premimuincluded);

                double sig10P = sigatm - 0.5 * (rr * rrMult) + (Bfly * smileFlyMult);
                double K10p = FXStrikeVol(s, today, Expiry, 0.1, sig10P, Dfe, Dfe2, "p", premimuincluded);

                sigdc = sig25c;
                sigdp = sig25p;

                //for (int i = 1; i <= 1; i++)
                //{

                //    Kdc = FXStrikeVol(s, today, Expiry, Delta, sigdc, Dfe, Dfe2, "c", premimuincluded);

                //    Kdp = FXStrikeVol(s, today, Expiry, Delta, sigdp, Dfe, Dfe2, "p", premimuincluded);

                //    sigdc = combinedInterp(s, today, Expiry, sigWing, Kdc, K10p, sig10P, K25p, sig25p, KA, sigatm, K25c, sig25c, K10c, sig10C, Dfe, Dfe2, Dfd, Dfd2, Fw, factor);

                //    sigdp = combinedInterp(s, today, Expiry, sigWing, Kdp, K10p, sig10P, K25p, sig25p, KA, sigatm, K25c, sig25c, K10c, sig10C, Dfe, Dfe2, Dfd, Dfd2, Fw, factor);
                //   // sigdc = smileSpline(K10p, K25p, KA, K25c, K10c, sig10P, sig25p, sigatm, sig25c, sig10C, Fw, Kdc, factor);
                //    //sigdp = smileSpline(K10p, K25p, KA, K25c, K10c, sig10P, sig25p, sigatm, sig25c, sig10C, Fw, Kdp, factor);

                //}


                sigdb = (sigdc + sigdp) / 2;


                Kdbc = FXStrikeVol(s, today, Expiry, Delta, sigdb, Dfe, Dfe2, "c", premimuincluded);

                Kdbp = FXStrikeVol(s, today, Expiry, Delta, sigdb, Dfe, Dfe2, "p", premimuincluded);

                //sigdbc = smileSpline(K10p, K25p, KA, K25c, K10c, sig10P, sig25p, sigatm, sig25c, sig10C, Fw, Kdbc, factor);

                sigdbc = combinedInterp(s, today, Expiry, sigWing, Kdbc, K10p, sig10P, K25p, sig25p, KA, sigatm, K25c, sig25c, K10c, sig10C, Dfe, Dfe2, Dfd, Dfd2, Fw, factor);


                //sigdbp = smileSpline(K10p, K25p, KA, K25c, K10c, sig10P, sig25p, sigatm, sig25c, sig10C, Fw, Kdbp, factor);
                sigdbp = combinedInterp(s, today, Expiry, sigWing, Kdbp, K10p, sig10P, K25p, sig25p, KA, sigatm, K25c, sig25c, K10c, sig10C, Dfe, Dfe2, Dfd, Dfd2, Fw, factor);

                double[] calldc = FXOpts(s, today, Expiry, Kdbc, sigdc, Dfe, Dfe2, Dfd, Dfd2, "c");
                double[] calldb = FXOpts(s, today, Expiry, Kdbc, sigdb, Dfe, Dfe2, Dfd, Dfd2, "c");
                double[] putdp = FXOpts(s, today, Expiry, Kdbp, sigdp, Dfe, Dfe2, Dfd, Dfd2, "p");
                double[] putdb = FXOpts(s, today, Expiry, Kdbp, sigdb, Dfe, Dfe2, Dfd, Dfd2, "p");

                double f0 = (calldc[0] + putdp[0]) - (calldb[0] + putdb[0]);

                double dfly = sigdb - sigatm;

                if (Math.Abs(f0) < 0.0000001 * s)
                {
                    functionReturnValue = dfly;
                    return functionReturnValue;
                }

                sigdb = sigdb + 0.00005;

                int j = 0;

                //loop
                double dsig = 0.00005;

                while (Math.Abs(f0) > 0.00000001 * s)
                {
                    j = j + 1;

                    Kdbc = FXStrikeVol(s, today, Expiry, Delta, sigdb, Dfe, Dfe2, "c", premimuincluded);

                    Kdbp = FXStrikeVol(s, today, Expiry, Delta, sigdb, Dfe, Dfe2, "p", premimuincluded);

                   // sigdbc = smileSpline(K10p, K25p, KA, K25c, K10c, sig10P, sig25p, sigatm, sig25c, sig10C, Fw, Kdbc, factor);
                    //sigdbp = smileSpline(K10p, K25p, KA, K25c, K10c, sig10P, sig25p, sigatm, sig25c, sig10C, Fw, Kdbp, factor);

                    sigdbc = combinedInterp(s, today, Expiry, sigWing, Kdbc, K10p, sig10P, K25p, sig25p, KA, sigatm, K25c, sig25c, K10c, sig10C, Dfe, Dfe2, Dfd, Dfd2, Fw, factor);
                    sigdbp = combinedInterp(s, today, Expiry, sigWing, Kdbp, K10p, sig10P, K25p, sig25p, KA, sigatm, K25c, sig25c, K10c, sig10C, Dfe, Dfe2, Dfd, Dfd2, Fw, factor);
                    
                   

                     calldc = FXOpts(s, today, Expiry, Kdbc, sigdbc, Dfe, Dfe2, Dfd, Dfd2, "c");
                     calldb = FXOpts(s, today, Expiry, Kdbc, sigdb, Dfe, Dfe2, Dfd, Dfd2, "c");
                     putdp = FXOpts(s, today, Expiry, Kdbp, sigdbp, Dfe, Dfe2, Dfd, Dfd2, "p");
                     putdb = FXOpts(s, today, Expiry, Kdbp, sigdb, Dfe, Dfe2, Dfd, Dfd2, "p");

                    double F = (calldc[0] + putdp[0]) - (calldb[0] + putdb[0]);
                    double dF = (F - f0) / dsig;

                    sigdb = sigdb - F / dF;
                    dsig = -F / dF;
                    f0 = F;

                    if (j * 0.00001 >= dfly)
                    {
                        functionReturnValue = sigdb - sigatm;
                        return functionReturnValue;
                    }


                }

                functionReturnValue = sigdb - sigatm;
                return functionReturnValue;


            }
            catch
            {

                functionReturnValue = sigdb - sigatm;
                return functionReturnValue;
            }



        }

        #endregion

        #region Options Functions
        public double CcyOptionGamma(double ExpDays, double DelDays, double strike, double Spot, double Fwd, double BaseDepo, double Vol, int TermsBasis, int BaseBasis, int terms)
        {
            double functionReturnValue = 0;

            if (Vol == 0)
            {
                functionReturnValue = 0;
                return functionReturnValue;
            }

            double Time = 0;
            double d1 = 0;
   
            double SpotGamma = 0;
            double FwdGamma = 0;
            double TermsDepo = 0;
            Time = ExpDays / 365;
            //*** should amend to account for leap years ***
            BaseDepo = ContinuousRate(BaseDepo, DelDays, BaseBasis);
            TermsDepo = ContinuousRate(SolveTermsDepo(Fwd, Spot, DelDays, BaseDepo, BaseBasis, TermsBasis), DelDays, TermsBasis);
            d1 = (Math.Log(Fwd / strike) + 0.5 * Vol * Vol * Time) / (Vol * Math.Sqrt(Time));
            SpotGamma = 0.01 * Math.Exp(-BaseDepo * DelDays / BaseBasis) * NPrime(d1) / (Vol * Math.Sqrt(Time));
            FwdGamma = 0.01 * NPrime(d1) / (Vol * Math.Sqrt(Time));
            switch (terms)
            {
                case 1:
                    functionReturnValue = SpotGamma;
                    // Spot Gamma in Pct Base Ccy
                    break;
                case 2:
                    functionReturnValue = SpotGamma * Spot / strike;
                    // Spot Gamma in Pct Terms Ccy
                    break;
                case 3:
                    functionReturnValue = FwdGamma;
                    // Fwd Gamma in Pct Base Ccy
                    break;
                case 4:
                    functionReturnValue = FwdGamma * Spot / strike;
                    // Fwd Gamma in Pct Terms Ccy
                    break;
                default:
                    functionReturnValue = 0;
                    // Error
                    break;
            }
            return functionReturnValue;
        }

        public double CumNorm(double x)
        {
            double functionReturnValue = 0;
            double k = 0;
            double nx = 0;
            k = 1 / (1 + GAMMA * x);
            nx = Math.Exp(-x * x / 2) / Math.Sqrt(2 * PI);
            if ((x >= 0))
            {
                functionReturnValue = 1 - nx * (k * (A1 + k * (A2 + k * (A3 + k * (A4 + k * A5)))));
            }
            else
            {
                functionReturnValue = 1 - CumNorm(-x);
            }
            return functionReturnValue;
        }

        public double CcyOptionPriceFromFwd(string PutCall, double ExpDays, double DelDays, double strike, double Spot, double Fwd, double BaseDepo, double Vol, int BaseBasis, int TermsBasis,
int terms)
        {
            double functionReturnValue = 0;

            if (Vol == 0)
            {
                functionReturnValue = 0;
                return functionReturnValue;
            }

            double d1 = 0;
            double d2 = 0;
            double Time = 0;
            double CallPrice = 0;
            double PutPrice = 0;
            double TermsDepo = 0;
            Time = ExpDays / 365;
            //*** should amend to account for leap years ***
            TermsDepo = ContinuousRate(SolveTermsDepo(Fwd, Spot, DelDays, BaseDepo, BaseBasis, TermsBasis), DelDays, TermsBasis);
            d1 = (Math.Log(Fwd / strike) + 0.5 * Vol * Vol * Time) / (Vol * Math.Sqrt(Time));
            d2 = d1 - Vol * Math.Sqrt(Time);
            CallPrice = Math.Exp(-TermsDepo * DelDays / TermsBasis) * (Fwd * CumNorm(d1) - strike * CumNorm(d2));
            PutPrice = Math.Exp(-TermsDepo * DelDays / TermsBasis) * (strike * CumNorm(-d2) - Fwd * CumNorm(-d1));
            if (PutCall == "Put" | PutCall == "put" | PutCall == "P" | PutCall == "p")
            {
                switch (terms)
                {
                    case 1:
                        functionReturnValue = PutPrice / Spot;
                        // Pct Base Ccy
                        break;
                    case 2:
                        functionReturnValue = PutPrice / strike;
                        // Pct Terms Ccy
                        break;
                    case 3:
                        functionReturnValue = PutPrice;
                        // Terms Pts
                        break;
                    case 4:
                        functionReturnValue = PutPrice / (strike * Spot);
                        // Base Pts
                        break;
                    default:
                        functionReturnValue = 9999;
                        // Error
                        break;
                }
            }
            else if (PutCall == "Call" | PutCall == "call" | PutCall == "C" | PutCall == "c")
            {
                switch (terms)
                {
                    case 1:
                        functionReturnValue = CallPrice / Spot;
                        // Pct Base Ccy
                        break;
                    case 2:
                        functionReturnValue = CallPrice / strike;
                        // Pct Terms Ccy
                        break;
                    case 3:
                        functionReturnValue = CallPrice;
                        // Terms Points
                        break;
                    case 4:
                        functionReturnValue = CallPrice / (strike * Spot);
                        // Base Points
                        break;
                    default:
                        functionReturnValue = 9999;
                        // Error
                        break;
                }
            }
            else if (PutCall == "S" | PutCall == "s")
            {
                switch (terms)
                {
                    case 1:
                        functionReturnValue = (PutPrice + CallPrice) / Spot;
                        // Pct Base Ccy
                        break;
                    case 2:
                        functionReturnValue = (PutPrice + CallPrice) / strike;
                        // Pct Terms Ccy
                        break;
                    case 3:
                        functionReturnValue = (PutPrice + CallPrice);
                        // Terms Points
                        break;
                    case 4:
                        functionReturnValue = (PutPrice + CallPrice) / (strike * Spot);
                        // Base Points
                        break;
                    default:
                        functionReturnValue = 9999;
                        // Error
                        break;
                }
            }
            else
            {
                functionReturnValue = 9999;
                // Error
            }
            return functionReturnValue;
        }


        //-----------------------------------------------------------------------------------------------------------------------------------'
        //   Returns the European-style vanilla currency option delta using the fwd, where put/call is on the base ccy.                      '
        //   Black-Scholes gives the delta in percent of base ccy exclusive of premium paid or received.                                     '
        //   The BS delta is calculated that way because an equity option premium is paid in dollars (terms ccy).                            '
        //   CcyOptionDeltaFromFwd calculates delta eeeeeeeeeeeeeee including premium (as traded in ccy option markets), since the premium is paid in the    '
        //   same units as the base ccy (equivalent to paying an equity option premium in shares instead of dollars).                        '
        //   Terms:                                                                                                                          '
        //   1 = spot delta % base ccy                                                                                                       '
        //   2 = spot delta % terms ccy                                                                                                      '
        //   3 = fwd delta % base ccy                                                                                                        '
        //   4 = fwd delta % terms ccy                                                                                                       '
        //-----------------------------------------------------------------------------------------------------------------------------------'
        public double CcyOptionDeltaFromFwd(string PutCall, double ExpDays, double DelDays, double strike, double Spot, double Fwd, double BaseDepo, double Vol, int BaseBasis, int TermsBasis,
        int terms)
        {
            double functionReturnValue = 0;

            if (Vol == 0)
            {
                functionReturnValue = 0;
                return functionReturnValue;
            }

            double Time = 0;
            double d1 = 0;
            double TermsDepo = 0;
            double BasePrem = 0;
            double TermsPrem = 0;
            BasePrem = CcyOptionPriceFromFwd(PutCall, ExpDays, DelDays, strike, Spot, Fwd, BaseDepo, Vol, BaseBasis, TermsBasis,
            1);
            TermsPrem = CcyOptionPriceFromFwd(PutCall, ExpDays, DelDays, strike, Spot, Fwd, BaseDepo, Vol, BaseBasis, TermsBasis,
            2);
            Time = ExpDays / 365;
            //*** should amend to account for leap years ***
            TermsDepo = ContinuousRate(SolveTermsDepo(Fwd, Spot, DelDays, BaseDepo, BaseBasis, TermsBasis), DelDays, TermsBasis);
            BaseDepo = ContinuousRate(BaseDepo, DelDays, BaseBasis);
            d1 = (Math.Log(Fwd / strike) + 0.5 * Vol * Vol * Time) / (Vol * Math.Sqrt(Time));

            if (PutCall == "Call" | PutCall == "call" | PutCall == "C" | PutCall == "c")
            {
                switch (terms)
                {
                    case 1:
                        functionReturnValue = Math.Exp(-BaseDepo * DelDays / BaseBasis) * CumNorm(d1) - BasePrem;
                        // Spot Delta Pct Base Ccy
                        break;
                    case 2:
                        functionReturnValue = Math.Exp(-BaseDepo * DelDays / BaseBasis) * CumNorm(d1);
                        //CcyOptionDeltaFromFwd = Math.Exp(-BaseDepo * DelDays / BaseBasis) * CumNorm(d1) * (Spot / strike) - TermsPrem        ' Spot Delta Pct Terms Ccy
                        break;
                    case 3:
                        functionReturnValue = CumNorm(d1) - BasePrem * Math.Exp(BaseDepo * DelDays / BaseBasis);
                        // Fwd Delta Pct Base Ccy
                        break;
                    case 4:
                        functionReturnValue = CumNorm(d1) * Math.Exp(TermsDepo * DelDays / TermsBasis);
                        // Fwd Delta Pct Terms Ccy
                        break;
                    default:
                        functionReturnValue = 0;
                        break;
                }
            }
            else if (PutCall == "Put" | PutCall == "put" | PutCall == "P" | PutCall == "p")
            {
                switch (terms)
                {
                    case 1:
                        functionReturnValue = Math.Exp(-BaseDepo * DelDays / BaseBasis) * (CumNorm(d1) - 1) - BasePrem;
                        // Spot Delta Pct Base Ccy
                        break;
                    case 2:
                        functionReturnValue = Math.Exp(-BaseDepo * DelDays / BaseBasis) * (CumNorm(d1) - 1);
                        //Math.Exp(-BaseDepo * DelDays / BaseBasis) * (CumNorm(d1) - 1) * (Spot / strike) - TermsPrem  ' Spot Delta Pct Terms Ccy
                        break;
                    case 3:
                        functionReturnValue = CumNorm(d1) - 1 - BasePrem * Math.Exp(BaseDepo * DelDays / BaseBasis);
                        // Fwd Delta Pct Base Ccy
                        break;
                    case 4:
                        functionReturnValue = CumNorm(d1) - 1 * Math.Exp(BaseDepo * DelDays / BaseBasis);
                        //CcyOptionDeltaFromFwd = (CumNorm(d1) - 1) * (Fwd / strike) - TermsPrem * Math.Exp(TermsDepo * DelDays / TermsBasis)  ' Fwd Delta Pct Terms Ccy
                        break;
                    default:
                        functionReturnValue = 0;
                        break;
                }
            }
            else if (PutCall == "S" | PutCall == "s")
            {
                double BasePremPut = 0;
                double BasePremCall = 0;
                double TermsPremPut = 0;
                double TermsPremCall = 0;
                BasePremPut = CcyOptionPriceFromFwd("p", ExpDays, DelDays, strike, Spot, Fwd, BaseDepo, Vol, BaseBasis, TermsBasis,
                1);
                BasePremCall = CcyOptionPriceFromFwd("c", ExpDays, DelDays, strike, Spot, Fwd, BaseDepo, Vol, BaseBasis, TermsBasis,
                1);
                TermsPremPut = CcyOptionPriceFromFwd("p", ExpDays, DelDays, strike, Spot, Fwd, BaseDepo, Vol, BaseBasis, TermsBasis,
                2);
                TermsPremCall = CcyOptionPriceFromFwd("p", ExpDays, DelDays, strike, Spot, Fwd, BaseDepo, Vol, BaseBasis, TermsBasis,
                2);
                switch (terms)
                {
                    case 1:
                        functionReturnValue = Math.Exp(-BaseDepo * DelDays / BaseBasis) * (2 * CumNorm(d1) - 1) - BasePremCall - BasePremPut;
                        // Spot Delta Pct Base Ccy
                        break;
                    case 2:
                        functionReturnValue = Math.Exp(-BaseDepo * DelDays / BaseBasis) * (2 * CumNorm(d1) - 1);
                        // Spot Delta Pct Terms Ccy
                        break;
                    case 3:
                        functionReturnValue = 2 * CumNorm(d1) - 1 - BasePremPut * Math.Exp(BaseDepo * DelDays / BaseBasis) - BasePremCall * Math.Exp(BaseDepo * DelDays / BaseBasis);
                        // Fwd Delta Pct Base Ccy
                        break;
                    case 4:
                        functionReturnValue = (2 * CumNorm(d1) - 1) * Math.Exp(TermsDepo * DelDays / TermsBasis);
                        // Fwd Delta Pct Terms Ccy
                        break;
                    default:
                        functionReturnValue = 0;
                        break;
                }
            }
            else
            {
                functionReturnValue = 0;
            }
            return functionReturnValue;
        }


        public double CcyOptionTheta(string PutCall, double ExpDays, double DelDays, double strike, double Spot, double Fwd, double BaseDepo, double Vol, int BaseBasis, int TermsBasis,
        int terms)
        {
            double functionReturnValue = 0;
            double d1 = 0;
            double d2 = 0;
            double Time = 0;
            double CallTheta = 0;
            double PutTheta = 0;
            double TermsDepo = 0;
            Time = ExpDays / 365;
            //*** should amend to account for leap years ***
            BaseDepo = ContinuousRate(BaseDepo, DelDays, BaseBasis);
            TermsDepo = ContinuousRate(SolveTermsDepo(Fwd, Spot, DelDays, BaseDepo, BaseBasis, TermsBasis), DelDays, TermsBasis);
            d1 = (Math.Log(Fwd / strike) + 0.5 * Vol * Vol * Time) / (Vol * Math.Sqrt(Time));
            d2 = d1 - Vol * Math.Sqrt(Time);
            CallTheta = -(Spot * NPrime(d1) * Vol * Math.Exp(-BaseDepo * DelDays / BaseBasis)) / (2 * Math.Sqrt(Time)) + BaseDepo * Spot * CumNorm(d1) * Math.Exp(-BaseDepo * DelDays / BaseBasis) - TermsDepo * strike * Math.Exp(-TermsDepo * DelDays / TermsBasis) * CumNorm(d2);
            PutTheta = -(Spot * NPrime(d1) * Vol * Math.Exp(-BaseDepo * DelDays / BaseBasis)) / (2 * Math.Sqrt(Time)) - BaseDepo * Spot * CumNorm(-d1) * Math.Exp(-BaseDepo * DelDays / BaseBasis) + TermsDepo * strike * Math.Exp(-TermsDepo * DelDays / TermsBasis) * CumNorm(-d2);
            if (PutCall == "Call" | PutCall == "call" | PutCall == "C" | PutCall == "c")
            {
                switch (terms)
                {
                    case 1:
                        functionReturnValue = CallTheta / Spot / 365;
                        // Theta in Pct Base Ccy
                        break;
                    case 2:
                        functionReturnValue = CallTheta / 365;
                        // Theta in Terms Pts per Base Ccy
                        break;
                }
            }
            else if (PutCall == "Put" | PutCall == "put" | PutCall == "P" | PutCall == "p")
            {
                switch (terms)
                {
                    case 1:
                        functionReturnValue = PutTheta / Spot / 365;
                        // Theta in Pct Base Ccy
                        break;
                    case 2:
                        functionReturnValue = PutTheta / 365;
                        // Theta in Terms Pts per Base Ccy
                        break;
                }
            }
            else if (PutCall == "s")
            {
                switch (terms)
                {
                    case 1:
                        functionReturnValue = (0.5 * (CallTheta + PutTheta)) / Spot / 365;
                        break;
                    case 2:
                        functionReturnValue = (0.5 * (CallTheta + PutTheta)) / 365;
                        break;
                }
            }
            else
            {
                functionReturnValue = 0;
            }
            return functionReturnValue;
        }

        public double CcyOptionVega(double ExpDays, double DelDays, double strike, double Spot, double Fwd, double BaseDepo, double Vol, int BaseBasis, int TermsBasis, int terms)
        {
            double functionReturnValue = 0;
            if (Vol == 0)
            {
                functionReturnValue = 0;
                return functionReturnValue;
            }
            double Time = 0;
            double d1 = 0;
            
            Time = ExpDays / 365;
            //*** should amend to account for leap years ***
            BaseDepo = ContinuousRate(BaseDepo, DelDays, BaseBasis);
            d1 = (Math.Log(Fwd / strike) + 0.5 * Vol * Vol * Time) / (Vol * Math.Sqrt(Time));
            functionReturnValue = 0.01 * Math.Sqrt(Time) * NPrime(d1) * Math.Exp(-BaseDepo * DelDays / BaseBasis);
            switch (terms)
            {
                //Need to multiply by 0.01 below to convert result for a 1% move in vol instead of 100%
                case 1:
                    functionReturnValue = 0.01 * Math.Sqrt(Time) * NPrime(d1) * Math.Exp(-BaseDepo * DelDays / BaseBasis);
                    // Vega in Pct Base Ccy
                    break;
                //Need to multiply by spot below to convert the pct of base ccy into terms pts
                case 2:
                    functionReturnValue = 0.01 * Spot * Math.Sqrt(Time) * NPrime(d1) * Math.Exp(-BaseDepo * DelDays / BaseBasis);
                    // Vega in Terms Pts per Base Ccy
                    break;
            }
            return functionReturnValue;
        }


        //-----------------------------------------------------------------------------------------------------------------------------------'
        //   Black-Scholes expresses rho as the instantaneous price change for an interest rate change scaled to 100%.                       '
        //   CcyOptionRho expresses rho for either interest rate scaled to an interest rate change of 1%.                                    '
        //   **** N.B. This function's output does not exactly match that of FENICS -- formulae should be reviewed. ****                     '
        //-----------------------------------------------------------------------------------------------------------------------------------'
        public double CcyOptionRho(string PutCall, double ExpDays, double DelDays, double strike, double Spot, double Fwd, double BaseDepo, double Vol, int BaseBasis, int TermsBasis,
        int terms)
        {
            double functionReturnValue = 0;

            if (Vol == 0)
            {
                functionReturnValue = 0;
                return functionReturnValue;
            }

            double Time = 0;
            double d1 = 0;
            double d2 = 0;
            double TermsDepo = 0;
            Time = ExpDays / 365;
            //*** should amend to account for leap years ***
            TermsDepo = ContinuousRate(SolveTermsDepo(Fwd, Spot, DelDays, BaseDepo, BaseBasis, TermsBasis), DelDays, TermsBasis);
            BaseDepo = ContinuousRate(BaseDepo, DelDays, BaseBasis);
            d1 = (Math.Log(Fwd / strike) + 0.5 * Vol * Vol * Time) / (Vol * Math.Sqrt(Time));
            d2 = d1 - Vol * Math.Sqrt(Time);
            if (PutCall == "c" | PutCall == "C" | PutCall == "call" | PutCall == "Call")
            {
                switch (terms)
                {
                    //(1) base ccy rho in pct base ccy, and (2) terms ccy rho in pct base ccy
                    case 1:
                        functionReturnValue = 0.01 * -Time * Math.Exp(-BaseDepo * DelDays / BaseBasis) * CumNorm(d1);
                        break;
                    case 2:
                        functionReturnValue = 0.01 * Time * Math.Exp(-TermsDepo * DelDays / TermsBasis) * CumNorm(d2) * strike / Spot;
                        break;
                }
            }
            else if (PutCall == "p" | PutCall == "P" | PutCall == "put" | PutCall == "Put")
            {
                switch (terms)
                {
                    //(1) base ccy rho in pct base ccy, and (2) terms ccy rho in pct base ccy
                    case 1:
                        functionReturnValue = 0.01 * Time * Math.Exp(-BaseDepo * DelDays / BaseBasis) * CumNorm(-d1);
                        break;
                    case 2:
                        functionReturnValue = 0.01 * -Time * Math.Exp(-TermsDepo * DelDays / TermsBasis) * CumNorm(-d2) * strike / Spot;
                        break;
                }
            }
            return functionReturnValue;
        }

        public double CcyOptionVanna(double ExpDays, double DelDays, double strike, double Spot, double Fwd, double BaseDepo, double Vol, int BaseBasis, int TermsBasis)
        {
            double functionReturnValue = 0;

            if (Vol == 0)
            {
                functionReturnValue = 0;
                return functionReturnValue;
            }

            double Time = 0;
            double d1 = 0;
            double d2 = 0;
            Time = ExpDays / 365;
            //*** should amend to account for leap years ***
            BaseDepo = ContinuousRate(BaseDepo, DelDays, BaseBasis);
            d1 = (Math.Log(Fwd / strike) + (Vol * Vol * 0.5) * Time) / (Vol * Math.Sqrt(Time));
            d2 = d1 - Vol * Math.Sqrt(Time);
            //the following was "-1 / spot ..." before editing
            functionReturnValue = -0.1 / Spot * Math.Exp(-BaseDepo * DelDays / BaseBasis) * d2 * NPrime(d1);
            return functionReturnValue;
        }

        //-----------------------------------------------------------------------------------------------------------------------------------'
        //   Black-Scholes does not calculate volga.                                                                                         '
        //   CcyOptionVolga calculates the change in vega for a change of 1 % vol in pct base.                                               '
        //-----------------------------------------------------------------------------------------------------------------------------------'
        public double CcyOptionVolga(double ExpDays, double DelDays, double strike, double Spot, double Fwd, double BaseDepo, double Vol, int BaseBasis, int TermsBasis)
        {
            double functionReturnValue = 0;

            if (Vol == 0)
            {
                functionReturnValue = 0;
                return functionReturnValue;
            }

            double Time = 0;
            double d1 = 0;
            double d2 = 0;
            Time = ExpDays / 365;
            //*** should amend to account for leap years ***
            BaseDepo = ContinuousRate(BaseDepo, DelDays, BaseBasis);
            d1 = (Math.Log(Fwd / strike) + (Vol * Vol * 0.5) * Time) / (Vol * Math.Sqrt(Time));
            d2 = d1 - Vol * Math.Sqrt(Time);
            functionReturnValue = 0.0001 * Math.Exp(-BaseDepo * DelDays / BaseBasis) * d2 * d1 * NPrime(d1) / Vol;
            return functionReturnValue;
        }

        public double VolSmileFromDelta(double DeltaNeutralVol, double RiskRev, double Bfly, double ExpDays, double Delta, double Accuracy = 1E-06)
        {
            double functionReturnValue = 0;
            double Time = 0;
            double Vega = 0;
            double Vanna = 0;
            double Volga = 0;
            double VannaFactor = 0;
            double VolgaFactor = 0;
            double dOption = 0;
            double Vol = 0;


            if (RiskRev > 0.25 | Bfly > 0.1)
            {
                functionReturnValue = 0;
                return functionReturnValue;
            }
            else if (Delta <= 0)
            {
                functionReturnValue = 0;
                return functionReturnValue;
            }

            const double d = -0.674489525667987;
            //number of std devs below mean for 25% of standard normal distribution

            Time = ExpDays / 365;
            //*** Should amend to account for leap years ***

            Vega = Math.Sqrt(Time) * NPrime(d);
            //calculate VannaFactor & VolgaFactor based on 25-delta
            Vanna = -d * NPrime(d);
            Volga = -d * -d * NPrime(d);
            VannaFactor = RiskRev * Vega / (2 * Vanna);
            VolgaFactor = Bfly * Vega / Volga;



            dOption = QNorm(Delta, 0, 1, true, false);
            //now re-calculate vega, vanna, volga based on actual delta
            Vega = Math.Sqrt(Time) * NPrime(dOption);
            Vanna = -dOption * NPrime(dOption);
            Volga = -dOption * -dOption * NPrime(dOption);

            Vol = 1 / Vega * (Volga * VolgaFactor + Vanna * VannaFactor);
            //now calculate vol to add/subtract from 50 delta vol
            functionReturnValue = Vol + DeltaNeutralVol;
            return functionReturnValue;
        }


        //-----------------------------------------------------------------------------------------------------------------------------------'
        //   Returns strike given the delta (and implicitly given the implied vol) using Newton's method.                                    '
        //   Based on "Numerical Recipes in C," 2nd ed., pp. 365-366.                                                                        '
        //   Our calculation of [dx] is basically equivalent to:                                                                             '
        //       (Delta-TargetDelta)/Gamma ==> dDelta/(dDelta/dSpot) ==> dSpot                                                               '
        //       This means [dx] is in terms of SPOT, rather than in terms of STRIKE                                                         '
        //       This is slightly mis-specified but makes almost no difference for convergence purposes                                      '
        //-----------------------------------------------------------------------------------------------------------------------------------'
        public double SolveCcyStrike(double TargetDelta, double Vol, double RiskRev, double Bfly, string PutCall, double ExpDays, double DelDays, double Spot, double Fwd, double BaseDepo,
        int BaseBasis, int TermsBasis, int DeltaType, int VolType, double Accuracy = 0.000001)
        {
            double functionReturnValue = 0;



            double TargetVol = 0;
            int j = 0;
            double Delta = 0;
            double dDelta = 0;
            double dx = 0;
            double AdjVol = 0;
            double DeltaNeutralVol = 0;
            functionReturnValue = Fwd;

            //if vol type is delta neutral (approximated for this function by delta = 0.5)
            if (VolType == 1)
            {
                DeltaNeutralVol = Vol;

                //******************************************
                //if vol type is atmf-THIS IS BEING ADUSTED BUT FOR NOW ONLY WORKS WITH DNS CONVENTION
            }
            else if (VolType == 20)
            {
                Delta = CcyOptionDeltaFromFwd("c", ExpDays, DelDays, Fwd, Spot, Fwd, BaseDepo, Vol, BaseBasis, TermsBasis,
                3);
                //get fwd delta for atmf strike ...
                AdjVol = VolSmileFromDelta(Vol, RiskRev, Bfly, ExpDays, Delta);
                //... get vol for atmf strike ...
                DeltaNeutralVol = 2 * Vol - AdjVol;
                //... and set delta neutral vol (for delta = 0.5)
            }
            else
            {
                functionReturnValue = 0;
                return functionReturnValue;
            }

            //Always convert to delta neutral vol (instead of using atmf vol) to ensure consistency between SolveCcyStrike and VolSmile,
            //  i.e. so that VolSmile returns the same vol for the strike that SolveCcyStrike finds using that vol
            //Multiplying SolveCcyStrike by (1 + dx) increases/decreases strike by dx percentage of strike

            if (PutCall == "c" | PutCall == "call" | PutCall == "C" | PutCall == "Call")
            {
                TargetVol = VolSmileFromDelta(DeltaNeutralVol, RiskRev, Bfly, ExpDays, TargetDelta);
                for (j = 1; j <= 50; j++)
                {
                    Delta = CcyOptionDeltaFromFwd(PutCall, ExpDays, DelDays, functionReturnValue, Spot, Fwd, BaseDepo, TargetVol, BaseBasis, TermsBasis,
                    DeltaType);
                    dDelta = CcyOptionGamma(ExpDays, DelDays, functionReturnValue, Spot, Fwd, BaseDepo, TargetVol, TermsBasis, BaseBasis, 1);
                    dx = 0.01 * (Delta - TargetDelta) / dDelta;
                    functionReturnValue = functionReturnValue * (1 + dx);
                    if (Math.Abs(dx) < Accuracy)

                        return functionReturnValue;
                }
            }
            else if (PutCall == "p" | PutCall == "put" | PutCall == "P" | PutCall == "Put")
            {
                TargetVol = VolSmileFromDelta(DeltaNeutralVol, RiskRev, Bfly, ExpDays, 1 - TargetDelta);
                for (j = 1; j <= 50; j++)
                {
                    Delta = CcyOptionDeltaFromFwd(PutCall, ExpDays, DelDays, functionReturnValue, Spot, Fwd, BaseDepo, TargetVol, BaseBasis, TermsBasis,
                    DeltaType);
                    dDelta = CcyOptionGamma(ExpDays, DelDays, functionReturnValue, Spot, Fwd, BaseDepo, TargetVol, TermsBasis, BaseBasis, 1);
                    dx = 0.01 * (Delta + TargetDelta) / dDelta;
                    functionReturnValue = functionReturnValue * (1 + dx);
                    if (Math.Abs(dx) < Accuracy)
                        return functionReturnValue;
                }
            }
            functionReturnValue = 0;
            return functionReturnValue;
        }






        public double VolSmile(double Vol, double RiskRev, double Bfly, double ExpDays, double DelDays, double Strike, double Spot, double Fwd, double BaseDepo, int BaseBasis,
        int TermsBasis, int VolType, double Accuracy = 1E-07)
        {
            int curCol = dataGridView1.CurrentCell.ColumnIndex;
            string strikeText = pricer.Rows[StrikeR][curCol].ToString();

            double functionReturnValue = 0;
            double Time = 0;
            double PreviousVol = 0;
            double Delta = 0;
            double DeltaNeutralVol = 0;
            int j = 0;


            int dtype = 0;
            if (strikeText.Substring(strikeText.Length - 1) == "f")
            {
                dtype = 2;
            }

            if (IsNumeric(Vol) == false)
            {
                functionReturnValue = 0;
                return functionReturnValue;
            }
            else if (Vol == 0)
            {
                functionReturnValue = 0;
                return functionReturnValue;
            }
            else if (RiskRev > 0.25 | Bfly > 0.1)
            {
                functionReturnValue = 0;
                return functionReturnValue;
            }

            Time = ExpDays / 365;
            //*** should amend to account for leap years ***

            //if vol type is delta neutral (approximated for this function by 50 delta)
            if (VolType == 1 | VolType == 2)
            {
                DeltaNeutralVol = Vol;
                functionReturnValue = DeltaNeutralVol;
                //if vol type is atmf ...
            }
            else if (VolType == 20)
            {
                //... get fwd delta for atmf strike ...
                Delta = CcyOptionDeltaFromFwd("c", ExpDays, DelDays, Fwd, Spot, Fwd, BaseDepo, Vol, BaseBasis, TermsBasis,
                3);
                //... get vol for atmf strike ...
                functionReturnValue = VolSmileFromDelta(Vol, RiskRev, Bfly, ExpDays, Delta);
                //... and set approximated delta neutral vol (for delta = 0.5)
                DeltaNeutralVol = 2 * Vol - functionReturnValue;

                //VolSmile = VolSmileFromDelta(DeltaNeutralVol, RiskRev, Bfly, ExpDays, Delta) 'now set vol to atmf strike ...
                //Delta = CcyOptionDeltaFromFwd("c", ExpDays, DelDays, Strike, Spot, Fwd, BaseDepo, Vol, BaseBasis, TermsBasis, 3) 'set initial delta using atmf vol
                //VolSmile = VolSmileFromDelta(DeltaNeutralVol, RiskRev, Bfly, ExpDays, Delta) '... and set initial vol using this delta (of atmf vol)
            }
            else
            {
                functionReturnValue = 0;
                return functionReturnValue;
            }

            //Ensures delta passed to VolSmileFromDelta correctly by using "1 + Delta(put)", i.e. 25 delta downside = 0.75
            if (Strike < Fwd)
            {
                for (j = 1; j <= 500; j++)
                {
                    PreviousVol = functionReturnValue;

                    if (strikeText == "dns" || strikeText == "dnf" || strikeText == "atm")
                    {
                        Delta = 0.5;

                    }
                    else
                    {
                        Delta = 1 + CcyOptionDeltaFromFwd("p", ExpDays, DelDays, Strike, Spot, Fwd, BaseDepo, PreviousVol, BaseBasis, TermsBasis,
                        VolType + dtype);
                    }

                    functionReturnValue = VolSmileFromDelta(DeltaNeutralVol, RiskRev, Bfly, ExpDays, Delta);
                    if (Math.Abs(functionReturnValue - PreviousVol) < Accuracy)
                        return functionReturnValue;
                }
            }
            else
            {
                for (j = 1; j <= 500; j++)
                {
                    PreviousVol = functionReturnValue;

                    if (strikeText == "dns" || strikeText == "dnf" || strikeText == "atm")
                    {
                        Delta = 0.5;

                    }
                    else
                    {
                        Delta = CcyOptionDeltaFromFwd("c", ExpDays, DelDays, Strike, Spot, Fwd, BaseDepo, PreviousVol, BaseBasis, TermsBasis,
                        VolType + dtype);
                    }

                    functionReturnValue = VolSmileFromDelta(DeltaNeutralVol, RiskRev, Bfly, ExpDays, Delta);
                    if (Math.Abs(functionReturnValue - PreviousVol) < Accuracy)
                        return functionReturnValue;
                }
            }
            functionReturnValue = Vol;
            return functionReturnValue;
        }

        //-----------------------------------------------------------------------------------------------------------------------------------'
        //   Solves for delta neutral strike, where DeltaType is for spot (1) or fwd (2) delta.                                              '
        //-----------------------------------------------------------------------------------------------------------------------------------'
        public double SolveDnStrike(double Vol, double ExpDays, double DelDays, double Spot, double Fwd, double BaseDepo, int BaseBasis, int TermsBasis, int DeltaType, int VolType,
        double Accuracy = 1E-06)
        {
            double functionReturnValue = 0;

            int j = 0;
            double DeltaCall = 0;
            double DeltaPut = 0;
            double dDelta = 0;
            double dx = 0;

            if (DeltaType == 1 || DeltaType == 2)
            {
                functionReturnValue = 0.5 * (SolveCcyStrike(0.5, Vol, 0, 0, "c", ExpDays, DelDays, Spot, Fwd, BaseDepo,
                BaseBasis, TermsBasis, DeltaType, 1) + SolveCcyStrike(0.5, Vol, 0, 0, "p", ExpDays, DelDays, Spot, Fwd, BaseDepo,
                BaseBasis, TermsBasis, DeltaType, 1));

            }
            else if (DeltaType == 3 || DeltaType == 4)
            {
                //fwd delta

                functionReturnValue = 0.5 * (SolveCcyStrike(0.5, Vol, 0, 0, "c", ExpDays, DelDays, Spot, Fwd, BaseDepo,
                BaseBasis, TermsBasis, DeltaType, 1) + SolveCcyStrike(0.5, Vol, 0, 0, "p", ExpDays, DelDays, Spot, Fwd, BaseDepo,
                BaseBasis, TermsBasis, DeltaType, 1));
            }
            else
            {
                functionReturnValue = 0;
                return functionReturnValue;
            }

            for (j = 1; j <= 50; j++)
            {
                DeltaCall = CcyOptionDeltaFromFwd("c", ExpDays, DelDays, functionReturnValue, Spot, Fwd, BaseDepo, Vol, BaseBasis, TermsBasis,
                DeltaType);
                DeltaPut = CcyOptionDeltaFromFwd("p", ExpDays, DelDays, functionReturnValue, Spot, Fwd, BaseDepo, Vol, BaseBasis, TermsBasis,
                DeltaType);
                dDelta = CcyOptionGamma(ExpDays, DelDays, functionReturnValue, Spot, Fwd, BaseDepo, Vol, TermsBasis, BaseBasis, DeltaType);
                dx = 0.01 * (DeltaCall + DeltaPut) / (2 * dDelta);
                functionReturnValue = functionReturnValue * (1 + dx);
                if (Math.Abs(dx) < Accuracy)
                    return functionReturnValue;
            }

            functionReturnValue = 0;
            return functionReturnValue;
        }


        //-----------------------------------------------------------------------------------------------------------------------------------'
        //   Converts entered text to an option strike.                                                                                      '
        //-----------------------------------------------------------------------------------------------------------------------------------'
        public double autostrike(string StrikeText, double Vol, double RiskRev, double Bfly, string PutCall, double ExpDays, double DelDays, double Spot, double Fwd, double BaseDepo,
        int BaseBasis, int TermsBasis, int VolType, double Accuracy = 1E-05, double Guess = 0.1)
        {
            double functionReturnValue = 0;

            int DeltaType = 0;

            if (IsNumeric(StrikeText) == true)
            {
                functionReturnValue = Convert.ToDouble(StrikeText);
                return functionReturnValue;
            }
            else
            {
                int TextLength = 0;
                TextLength = StrikeText.Length;
                if (StrikeText == "atmf")
                {
                    functionReturnValue = Fwd;
                }
                else if (StrikeText == "atms")
                {
                    functionReturnValue = Spot;
                    //out of the money fwd
                }

                else if (StrikeText.Substring(TextLength - 1) == "d")
                {

                    if (VolType == 1)
                    {
                        DeltaType = 1;
                    }
                    else
                    {
                        DeltaType = 2;
                    }

                    functionReturnValue = SolveCcyStrike(Convert.ToDouble(StrikeText.Substring(0, TextLength - 1)) / 100, Vol, RiskRev, Bfly, PutCall, ExpDays, DelDays, Spot, Fwd, BaseDepo,
                    BaseBasis, TermsBasis, DeltaType, 1);

                    //fwd delta
                }
                else if (StrikeText.Substring(TextLength - 2) == "df")
                {

                    if (VolType == 1)
                    {
                        DeltaType = 3;
                    }
                    else
                    {
                        DeltaType = 4;
                    }

                    functionReturnValue = SolveCcyStrike(Convert.ToDouble(StrikeText.Substring(0, TextLength - 2)) / 100, Vol, RiskRev, Bfly, PutCall, ExpDays, DelDays, Spot, Fwd, BaseDepo,
                    BaseBasis, TermsBasis, DeltaType, 1);

                    //need to fix dns strike for proper premo currency

                    //delta neutral fwd strike
                }
                else if (StrikeText.Substring(TextLength - 3) == "dnf")
                {

                    if (VolType == 1)
                    {
                        DeltaType = 3;
                    }
                    else
                    {
                        DeltaType = 4;
                    }

                    functionReturnValue = SolveDnStrike(Vol, ExpDays, DelDays, Spot, Fwd, BaseDepo, BaseBasis, TermsBasis, DeltaType, 1);


                    //delta neutral spot strike
                }
                else if (StrikeText.Substring(TextLength - 3) == "dns" | StrikeText.Substring(TextLength - 3) == "atm")
                {

                    if (VolType == 1)
                    {
                        DeltaType = 1;
                    }
                    else
                    {
                        DeltaType = 2;
                    }


                    functionReturnValue = SolveDnStrike(Vol, ExpDays, DelDays, Spot, Fwd, BaseDepo, BaseBasis, TermsBasis, DeltaType, 1);
                }




                else if (StrikeText.Substring(TextLength - 4) == "otmf")
                {
                    if (PutCall == "c" | PutCall == "C" | PutCall == "call" | PutCall == "Call")
                    {
                        functionReturnValue = (1 + Convert.ToDouble(StrikeText.Substring(0, TextLength - 4)) / 100) * Fwd;
                    }
                    else
                    {
                        functionReturnValue = Fwd / (1 + (Convert.ToDouble(StrikeText.Substring(0, TextLength - 4)) / 100));
                    }
                    //in the money fwd
                }
                else if (StrikeText.Substring(TextLength - 4) == "itmf")
                {
                    if (PutCall == "p" | PutCall == "P" | PutCall == "put" | PutCall == "Put")
                    {
                        functionReturnValue = (1 + Convert.ToDouble(StrikeText.Substring(0, TextLength - 4)) / 100) * Fwd;
                    }
                    else
                    {
                        functionReturnValue = Fwd / (1 + Convert.ToDouble(StrikeText.Substring(0, TextLength - 4)) / 100);
                    }
                    //out of the money spot
                }
                else if (StrikeText.Substring(TextLength - 4) == "otms")
                {
                    if (PutCall == "c" | PutCall == "C" | PutCall == "call" | PutCall == "Call")
                    {
                        functionReturnValue = (1 + Convert.ToDouble(StrikeText.Substring(0, TextLength - 4)) / 100) * Spot;
                    }
                    else
                    {
                        functionReturnValue = Spot / (1 + Convert.ToDouble(StrikeText.Substring(0, TextLength - 4)) / 100);
                    }
                    //in the money spot
                }
                else if (StrikeText.Substring(TextLength - 4) == "itms")
                {
                    if (PutCall == "p" | PutCall == "P" | PutCall == "put" | PutCall == "Put")
                    {
                        functionReturnValue = (1 + Convert.ToDouble(StrikeText.Substring(0, TextLength - 4)) / 100) * Spot;
                    }
                    else
                    {
                        functionReturnValue = Spot / (1 + Convert.ToDouble(StrikeText.Substring(0, TextLength - 4)) / 100);
                    }

                    //spot delta
                }
                else
                {
                    functionReturnValue = 0;
                }

            }
            return functionReturnValue;
        }


        public double SolveCcyVol(double TargetPrice, string PutCall, double ExpDays, double DelDays, double strike, double Spot, double Fwd, double BaseDepo, int BaseBasis, int TermsBasis,
      int PriceTerms, int VegaTerms, double Accuracy = 1E-05, double Guess = 0.1)
        {
            double functionReturnValue = 0;
            //Here we solve for vol given the option price
            int j = 0;
            double p = 0;
            double dP = 0;
            double dx = 0;
            functionReturnValue = Guess;
            //Multiply by 0.01 in loop to scale dx to vol
            //Note that our calculation of [dx] is basically equivalent to:
            //(Price-TargetPrice)/Vega ==> dPrice/(dPrice/dVol) ==> dVol
            for (j = 1; j <= 50; j++)
            {
                p = CcyOptionPriceFromFwd(PutCall, ExpDays, DelDays, strike, Spot, Fwd, BaseDepo, functionReturnValue, BaseBasis, TermsBasis,
                PriceTerms);
                dP = CcyOptionVega(ExpDays, DelDays, strike, Spot, Fwd, BaseDepo, functionReturnValue, TermsBasis, BaseBasis, VegaTerms);
                dx = 0.01 * (p - TargetPrice) / dP;
                functionReturnValue = functionReturnValue - dx;
                if (Math.Abs(dx) < Accuracy)
                    return functionReturnValue;
            }
            functionReturnValue = 9999;
            return functionReturnValue;
        }





        #endregion

        #region Rates Functions

        public double ContinuousRate(double Depo, double DelDays, int Basis)
        {
            int Years = 0;
            Years = Convert.ToInt32(RoundDown(DelDays / Basis, 0));
            return Basis / DelDays * Math.Log((1 + Depo * (DelDays / Basis - Years)) * Math.Pow((1 + Depo), Years));
        }

        public double NPrime(double x)
        {
            const double PI = 3.14159265358979;
            return Math.Exp(0.5 * -x * x) / (Math.Sqrt(2 * PI));
        }



        public double Forward(double Spot, double DelDays, double BaseDepo, int BaseBasis, double TermsDepo, int TermsBasis)
        {
            //int TermsYears = 0;
            //int BaseYears = 0;
            //BaseYears = Convert.ToInt32(RoundDown(DelDays / BaseBasis,0));
            //TermsYears = Convert.ToInt32(RoundDown(DelDays / TermsBasis,0));
            //return Spot * (1 + TermsDepo * (DelDays / TermsBasis - TermsYears)) * Math.Pow((1 + TermsDepo), TermsYears) / ((1 + BaseDepo * (DelDays / BaseBasis - BaseYears)) * Math.Pow((1 + BaseDepo), BaseYears));

            int TermsYears = 0;
            int BaseYears = 0;
            BaseYears = Convert.ToInt32(RoundDown(DelDays / BaseBasis, 0));
            TermsYears = Convert.ToInt32(RoundDown(DelDays / TermsBasis, 0));
            return Spot * (1 + TermsDepo * (DelDays / TermsBasis)) / (1 + BaseDepo * (DelDays / BaseBasis));


            //double dfTerms = DiscountFactor(TermsDepo, Convert.ToInt16(DelDays), TermsBasis);
            //double dfBase = DiscountFactor(BaseDepo, Convert.ToInt16(DelDays), BaseBasis);

            //double f =Spot * dfTerms / dfBase;
            //return f;

        }

        public double SolveTermsDepoOLD(double TargetForward, double Spot, double DelDays, double BaseDepo, int BaseBasis, int TermsBasis, double Accuracy = 1E-06, double Guess = 0.1)
        {
            double functionReturnValue = 0;
            double dx = 0;
            int j = 0;
            double F = 0;
            double dF = 0;
            int BaseYears = 0;
            int TermsYears = 0;
            functionReturnValue = Guess;
            BaseYears = Convert.ToInt32(DelDays / BaseBasis);
            TermsYears = Convert.ToInt32(DelDays / TermsBasis);
            for (j = 1; j <= 50; j++)
            {
                F = Forward(Spot, DelDays, BaseDepo, BaseBasis, functionReturnValue, TermsBasis);
                dF = Spot / ((1 + BaseDepo * (DelDays / TermsBasis - BaseYears)) * Math.Pow((1 + BaseDepo), BaseYears)) * ((DelDays / TermsBasis - TermsYears) * Math.Pow((1 + functionReturnValue), TermsYears) + TermsYears * Math.Pow((1 + functionReturnValue), (TermsYears - 1)) * (1 + functionReturnValue * (DelDays / TermsBasis - TermsYears)));
                dx = (F - TargetForward) / dF;
                functionReturnValue = functionReturnValue - dx;
                if (Math.Abs(dx) < Accuracy)
                    return functionReturnValue;
            }
            functionReturnValue = 9999;
            return functionReturnValue;
        }

        public double RoundDown(double value, int digits)
        {
            if (value >= 0)
                return Math.Floor(value * Math.Pow(10, digits)) / Math.Pow(10, digits);

            return Math.Ceiling(value * Math.Pow(10, digits)) / Math.Pow(10, digits);
        }

        public double SolveTermsDepo(double TargetForward, double Spot, double DelDays, double BaseDepo, int BaseBasis, int TermsBasis, double Accuracy = 1E-06, double Guess = 0.1)
        {
            double functionReturnValue = 0.0;
            double dx = 0;
            int j = 0;
            double F = 0;
            double dF = 0;
            int BaseYears = 0;
            int TermsYears = 0;
            functionReturnValue = Guess;
            BaseYears = Convert.ToInt32(RoundDown(DelDays / BaseBasis, 0));
            TermsYears = Convert.ToInt32(RoundDown(DelDays / TermsBasis, 0));
            for (j = 1; j <= 50; j++)
            {
                F = Forward(Spot, DelDays, BaseDepo, BaseBasis, functionReturnValue, TermsBasis);
                dF = Spot / ((1 + BaseDepo * (DelDays / TermsBasis - BaseYears)) * Math.Pow((1 + BaseDepo), BaseYears)) * ((DelDays / TermsBasis - TermsYears) * Math.Pow((1 + functionReturnValue), TermsYears) + TermsYears * Math.Pow((1 + functionReturnValue), (TermsYears - 1)) * (1 + functionReturnValue * (DelDays / TermsBasis - TermsYears)));
                dx = (F - TargetForward) / dF;
                functionReturnValue = functionReturnValue - dx;
                if (Math.Abs(dx) < Accuracy)
                    return functionReturnValue;
            }
            functionReturnValue = 9999;
            return functionReturnValue;
        }
        public double SolveBaseDepo(double TargetForward, double Spot, double DelDays, double TermsDepo, int BaseBasis, int TermsBasis, double Accuracy = 1E-06, double Guess = 0.1)
        {
            double functionReturnValue = 0;
            double dx = 0;
            int j = 0;
            double F = 0;
            double dF = 0;
            int BaseYears = 0;
            int TermsYears = 0;
            functionReturnValue = Guess;
            BaseYears = Convert.ToInt32(RoundDown(DelDays / BaseBasis, 0));
            TermsYears = Convert.ToInt32(RoundDown(DelDays / TermsBasis, 0));
            for (j = 1; j <= 50; j++)
            {
                F = Forward(Spot, DelDays, functionReturnValue, BaseBasis, TermsDepo, TermsBasis);

                dF = Spot / ((1 + TermsDepo * (DelDays / TermsBasis - TermsYears)) * Math.Pow((1 + TermsDepo), TermsYears)) * ((BaseYears - DelDays / BaseBasis) * Math.Pow((1 + functionReturnValue * (DelDays / BaseBasis - BaseYears)), -2) * Math.Pow((1 + functionReturnValue), -BaseYears) - BaseYears * Math.Pow(1 + functionReturnValue, (-BaseYears - 1)) * Math.Pow((1 + functionReturnValue * (DelDays / BaseBasis - BaseYears)), -1));
                dx = (F - TargetForward) / dF;
                functionReturnValue = functionReturnValue - dx;
                if (Math.Abs(dx) < Accuracy)
                    return functionReturnValue;
            }
            functionReturnValue = 9999;
            return functionReturnValue;
        }

        //-----------------------------------------------------------------------------------------------------------------------------------'
        //   Interpolates forward points, where p = points, d = daycounts, t = target daycount.                                              '
        //   pd = target pts/day, pdLo = lower bracket pts/day, pdHi = upper bracket pts/day                                                 '                              '
        //-----------------------------------------------------------------------------------------------------------------------------------'
        public double FwdPtsPerDay(List<double> fp, List<double> fd, double t)
        {
            double functionReturnValue = 0;

            //exit function if data is bad
            if (fd.Count != fp.Count)
            {
                functionReturnValue = 9999;
                return functionReturnValue;
            }



            //bracket period required
            int k = 0;
            int kLo = 0;
            int kHi = 0;


            kLo = 0;
            kHi = fp.Count() - 1;
            while ((kHi - kLo) > 1)
            {
                k = (kHi + kLo) / 2;
                if (fd[k] > t)
                {
                    kHi = k;
                }
                else
                {
                    kLo = k;
                }
            }

            if (fd[kLo] == 0)
            {
                functionReturnValue = 0;
            }
            else
            {
                //calculate [points per day] for kLo and kHi and interpolate
                double pd = 0;
                double pdLo = 0;
                double pdHi = 0;
                pdLo = fp[kLo] / fd[kLo];
                pdHi = fp[kHi] / fd[kHi];
                pd = ((t - fd[kLo]) / (fd[kHi] - fd[kLo])) * (pdHi - pdLo) + pdLo;
                functionReturnValue = t * pd;
            }
            return functionReturnValue;

        }

        //-----------------------------------------------------------------------------------------------------------------------------------'
        //   Returns linear interpolation, where d = day counts, r = rates, t = target daycount.                                             '
        //-----------------------------------------------------------------------------------------------------------------------------------'
        public double LinearInterp(List<double> fd, List<double> fr, double t)
        {
            double functionReturnValue = 0;

            //exit function if data is bad
            if (fd.Count != fr.Count)
            {
                functionReturnValue = 9999;
                return functionReturnValue;
            }
            //bracket period required
            int k = 0;
            int kLo = 0;
            int kHi = 0;
            kLo = 0;
            kHi = fd.Count - 1;
            while ((kHi - kLo) > 1)
            {
                k = (kHi + kLo) / 2;
                if (fd[k] > t)
                {
                    kHi = k;
                }
                else
                {
                    kLo = k;
                }
            }

            //now interpolate
            functionReturnValue = (fr[kHi] - fr[kLo]) / (fd[kHi] - fd[kLo]) * (t - fd[kLo]) + fr[kLo];
            return functionReturnValue;
        }


        #endregion

        #region Date Functions




        //-----------------------------------------------------------------------------------------------------------------------------------'
        //   Converts entered text to an option expiry date.                                                                                 '
        //-----------------------------------------------------------------------------------------------------------------------------------'
        public System.DateTime AutoExpiryDate(string ExpiryText, System.DateTime StartDate, string HomeCcy, string BaseCcy, string TermsCcy)
        {

            DateTime err = new DateTime(1900, 1, 1);
            System.DateTime functionReturnValue = default(System.DateTime);
            int TextLength = 0;
            TextLength = ExpiryText.Length;
            if (ExpiryText.Substring(TextLength - 1) == "d" || ExpiryText.Substring(TextLength - 1) == "D")
            {
                functionReturnValue = StartDate.AddDays(Convert.ToInt32(ExpiryText.Substring(0, TextLength - 1)));
            }
            else if (ExpiryText.Substring(TextLength - 1) == "w" || ExpiryText.Substring(TextLength - 1) == "W")
            {
                functionReturnValue = ExpiryWeekDate(StartDate, Convert.ToInt32(ExpiryText.Substring(0, TextLength - 1)), HomeCcy);
            }
            else if (ExpiryText.Substring(TextLength - 1) == "m" || ExpiryText.Substring(TextLength - 1) == "M")
            {
                functionReturnValue = ExpiryMonthDate(StartDate, Convert.ToInt32(ExpiryText.Substring(0, TextLength - 1)), HomeCcy, BaseCcy, TermsCcy);
            }
            else if (ExpiryText.Substring(TextLength - 1) == "y" || ExpiryText.Substring(TextLength - 1) == "Y")
            {
                functionReturnValue = ExpiryMonthDate(StartDate, 12 * Convert.ToInt32(ExpiryText.Substring(0, TextLength - 1)), HomeCcy, BaseCcy, TermsCcy);

            }
            else if (IsDate(ExpiryText) == true)
            {
                functionReturnValue = Convert.ToDateTime(ExpiryText);
                double past = (functionReturnValue - DateTime.Now).TotalDays;
                if (past < 0)
                {
                    functionReturnValue = functionReturnValue.AddYears(1);
                }
            }
            else
            {
                functionReturnValue = err;
            }
            return functionReturnValue;
        }

        private string AutoExpiryString(string ExpiryText)
        {


            string functionReturnValue = ExpiryText;
            int TextLength = ExpiryText.Length;
            string dateEnd = ExpiryText.Substring(TextLength - 1);
            int textlen = ExpiryText.Length;

            if (IsDate(ExpiryText) == true)
            {
                DateTime testDate = Convert.ToDateTime(ExpiryText);
                double past = (testDate - DateTime.Now).TotalDays;

                if (past < 0)
                {
                    testDate = testDate.AddYears(1);
                }

                functionReturnValue = testDate.ToString("ddMMMyy");
            }


            else if (dateEnd == "d" && textlen <= 5 || dateEnd == "w" && textlen <= 5 || dateEnd == "m" && textlen <= 5 || dateEnd == "y" && textlen <= 5)
            {
                return functionReturnValue;
            }

            else if (dateEnd == "D" && textlen <= 5 || dateEnd == "W" && textlen <= 5 || dateEnd == "M" && textlen <= 5 || dateEnd == "Y" && textlen <= 5)
            {
                return functionReturnValue;
            }
            else
            {
                //err
                functionReturnValue = "CHECK DATE";
            }
            return functionReturnValue;
        }


        private bool IsDate(string inputDate)
        {
            bool isDate = true;
            try
            {
                DateTime dt = DateTime.Parse(inputDate);

            }
            catch
            {
                isDate = false;
            }

            return isDate;
        }




        //-----------------------------------------------------------------------------------------------------------------------------------'
        //   Returns tomorrow's date (T+1), one business day from the start date.                                                            '
        //-----------------------------------------------------------------------------------------------------------------------------------'
        public System.DateTime TomDate(System.DateTime StartDate, string BaseCcy, string TermsCcy)
        {

            System.DateTime functionReturnValue = StartDate.AddDays(1);

            while (TestHoliday(functionReturnValue, BaseCcy) == false || TestHoliday(functionReturnValue, TermsCcy) == false)
            {
                functionReturnValue = functionReturnValue.AddDays(1);
            }

            return functionReturnValue;
        }

        //-----------------------------------------------------------------------------------------------------------------------------------'
        //   Returns monthly delivery dates for a specified currency pair.                                                                   '
        //   StartDate is the date from which to calculate the monthly date, e.g. 20 Jul 1998.                                               '
        //   StartDay, StartMonth, StartYear are the day, month and year of the StartDate, e.g. 20, 7, 1998.                                 '
        //-----------------------------------------------------------------------------------------------------------------------------------'

        public System.DateTime MonthDate(System.DateTime StartDate, long Months, string BaseCcy, string TermsCcy)
        {
            System.DateTime functionReturnValue = default(System.DateTime);
            long StartDay = 0;
            long StartMonth = 0;
            long StartYear = 0;
            long EndDay = 0;
            long EndMonth = 0;
            long EndYear = 0;
            long MonthPlusOne = 0;
            long YearCount = 0;
            DateTime err = new DateTime(1900, 1, 1);

            StartDay = StartDate.Day;
            StartMonth = StartDate.Month;
            StartYear = StartDate.Year;
            MonthPlusOne = TomDate(StartDate, BaseCcy, TermsCcy).Month;

            //Accounts for possible month-end effects
            YearCount = Convert.ToInt32((StartMonth + Months - 1) / 12);
            EndDay = StartDay;
            EndMonth = StartMonth + Months - YearCount * 12;
            EndYear = StartYear + YearCount;

            //Tests for an invalid input
            if ((Convert.ToInt32(RoundDown(Months, 0)) != Months | Months < 0))
            {
                functionReturnValue = err;

            }
            else if (StartMonth != MonthPlusOne)
            {
                functionReturnValue = StartDate.AddMonths(Convert.ToInt32(Months + 1));
                functionReturnValue = new DateTime(functionReturnValue.Year, functionReturnValue.Month, 1);
                functionReturnValue = functionReturnValue.AddDays(-1);
                //Goes to the first day of the next month, then subtracts one day
                while (TestTwoHolidays(functionReturnValue, BaseCcy, TermsCcy) == false)
                {
                    functionReturnValue = functionReturnValue.AddDays(-1);
                }

            }
            else
            {
                try
                {

                    functionReturnValue = new DateTime(Convert.ToInt32(EndYear), Convert.ToInt32(EndMonth), Convert.ToInt32(EndDay));



                }
                catch
                {


                    EndDay = EndDay - 1;

                    //fed 29
                    if (EndMonth == 2 && EndDay == 29)
                    {
                        EndDay = 28;
                    }

                    functionReturnValue = new DateTime(Convert.ToInt32(EndYear), Convert.ToInt32(EndMonth), Convert.ToInt32(EndDay));

                }
                while (TestTwoHolidays(functionReturnValue, BaseCcy, TermsCcy) == false)
                {
                    functionReturnValue = functionReturnValue.AddDays(1);
                }
            }

            return functionReturnValue;


        }

        //-----------------------------------------------------------------------------------------------------------------------------------'
        //   Returns a specified currency pair's date as True or False ("good" or "bad")                                                     '
        //-----------------------------------------------------------------------------------------------------------------------------------'
        public bool TestTwoHolidays(System.DateTime DateToCheck, string BaseCcy, string TermsCcy)
        {
            bool functionReturnValue = false;
            if ((TestHoliday(DateToCheck, BaseCcy) == true & TestHoliday(DateToCheck, TermsCcy) == true))
            {
                functionReturnValue = true;
            }
            else
            {
                functionReturnValue = false;
            }
            return functionReturnValue;
        }



        //  Returns the expiry date a specified number of months from the start date. 
        public System.DateTime ExpiryMonthDate(System.DateTime StartDate, long Months, string HomeCcy, string BaseCcy, string TermsCcy)
        {
            System.DateTime functionReturnValue = default(System.DateTime);
            functionReturnValue = MonthDate(SpotDate(StartDate, BaseCcy, TermsCcy), Months, BaseCcy, TermsCcy);
            functionReturnValue = expirydate(functionReturnValue, HomeCcy, BaseCcy, TermsCcy);
            return functionReturnValue;
        }





        ////-----------------------------------------------------------------------------------------------------------------------------------'
        ////   Returns the expiry date, a number of specified business days prior to a given value date.                                       '                                                                             '
        ////-----------------------------------------------------------------------------------------------------------------------------------'
        public System.DateTime expirydate(System.DateTime ValueDate, string HomeCcy, string BaseCcy, string TermsCcy)
        {

            DateTime err = new DateTime(1900, 1, 1);

            System.DateTime functionReturnValue = default(System.DateTime);
            functionReturnValue = ValueDate;
            if (TestHoliday(functionReturnValue, BaseCcy) == false | TestHoliday(functionReturnValue, TermsCcy) == false)
            {
                functionReturnValue = err;
            }
            else
            {
                functionReturnValue = ValueDate;
                int i = 0;
                for (i = 1; i <= 20; i++)
                {
                    if (TestHoliday(functionReturnValue, HomeCcy) == true)
                    {
                        if (SpotDate(functionReturnValue, BaseCcy, TermsCcy) == ValueDate)
                        {
                            return functionReturnValue;
                        }
                    }
                    functionReturnValue = functionReturnValue.AddDays(-1);
                }
            }

            int DayCount = 0;
            if (SpotDateFactor(TermsCcy) < SpotDateFactor(BaseCcy))
            {
                DayCount = SpotDateFactor(TermsCcy);
            }
            else
            {
                DayCount = SpotDateFactor(BaseCcy);
            }

            functionReturnValue = ValueDate;
            int Counter = 0;
            Counter = 0;
            while (Counter < DayCount)
            {
                if (TestHoliday(functionReturnValue, BaseCcy) == true & TestHoliday(functionReturnValue, TermsCcy) == true)
                {
                    Counter = Counter + 1;
                }
                functionReturnValue = functionReturnValue.AddDays(-1);
            }
            return functionReturnValue;
        }

        //-----------------------------------------------------------------------------------------------------------------------------------'
        //   Returns the number of business days for spot trading in the currency pair.                                                      '
        //-----------------------------------------------------------------------------------------------------------------------------------'
        public int SpotDateFactor(string Ccy)
        {

            int functionReturnValue = 0;

            List<string> ccyName = new List<string>();
            ccyName = ccyDets.AsEnumerable().Select(x => x[0].ToString().ToUpper().Trim()).ToList();

            Ccy = Ccy.ToUpper().Trim();

            if (ccyName.Contains(Ccy))
            {
                int index = ccyName.IndexOf(Ccy);

                string ccyDt = ccyDets.Rows[index]["CcySpotDays"].ToString();


                functionReturnValue = Convert.ToInt32(ccyDt);
            }

            else
                functionReturnValue = 0;

            return functionReturnValue;

        }


        public System.DateTime SpotDate(System.DateTime StartDate, string BaseCcy, string TermsCcy)
        {
            System.DateTime functionReturnValue = default(System.DateTime);
            functionReturnValue = StartDate;

            int Counter = 0;
            Counter = 0;
            int DayCount = 0;
            DayCount = SpotDateFactor(BaseCcy);
            if (SpotDateFactor(TermsCcy) < DayCount)
            {
                DayCount = SpotDateFactor(TermsCcy);
            }

            if ((TestHoliday(AddWorkdays(functionReturnValue, DayCount), BaseCcy) == true & TestHoliday(AddWorkdays(functionReturnValue, DayCount), TermsCcy) == true))
            {
                functionReturnValue = AddWorkdays(functionReturnValue, DayCount);
                return functionReturnValue;
            }



            while (Counter < DayCount)
            {
                functionReturnValue = functionReturnValue.AddDays(1);
                if ((TestHoliday(functionReturnValue, BaseCcy) == true & TestHoliday(functionReturnValue, TermsCcy) == true))
                {
                    Counter = Counter + 1;
                }
            }



            if ((BaseCcy != "USD" | TermsCcy != "USD"))
            {
                while (TestHoliday(functionReturnValue, "USD") == false)
                {
                    functionReturnValue = functionReturnValue.AddDays(1);

                }
            }




            return functionReturnValue;

        }


        public static DateTime AddWorkdays(DateTime originalDate, int workDays)
        {
            DateTime tmpDate = originalDate;
            while (workDays > 0)
            {
                tmpDate = tmpDate.AddDays(1);
                if (tmpDate.DayOfWeek < DayOfWeek.Saturday &&
                    tmpDate.DayOfWeek > DayOfWeek.Sunday)

                    workDays--;
            }
            return tmpDate;
        }


        //-----------------------------------------------------------------------------------------------------------------------------------'
        //   Returns the correct business day for the specified number of weeks from the start date.
        //-----------------------------------------------------------------------------------------------------------------------------------'
        public System.DateTime ExpiryWeekDate(System.DateTime StartDate, long Weeks, string HomeCcy)
        {

            DateTime err = new DateTime(1900, 1, 1);

            System.DateTime functionReturnValue = default(System.DateTime);
            if ((Convert.ToInt32(RoundDown(Weeks, 0)) != Weeks | Weeks < 0))
            {

                functionReturnValue = err;
            }
            else
            {
                functionReturnValue = StartDate.AddDays(Weeks * 7);
                while (TestHoliday(functionReturnValue, HomeCcy) == false)
                {
                    functionReturnValue = functionReturnValue.AddDays(1);
                }
            }
            return functionReturnValue;
        }


        //        '-----------------------------------------------------------------------------------------------------------------------------------'
        //'   Returns the date for a specified currency as either True or False (a "good" or "bad" day)                                       '
        //'   ByVal -- prevents procedure from changing the value of the variable it receives (passes by value instead of by reference)       '
        //'   CLng -- converts the date to a serial number (long integer), which can then be compared to date numbers on a spreadsheet        '
        //'-----------------------------------------------------------------------------------------------------------------------------------'

        public bool TestHoliday(System.DateTime DateToCheck, string Ccy)
        {
            bool functionReturnValue = false;

            List<string> hols = new List<string>();
            List<DateTime> holsDate = new List<DateTime>();

            hols = holidaySet.Tables[Ccy].AsEnumerable().Select(x => x[2].ToString()).ToList();



            foreach (string s in hols)
            {
                if (s != "")
                {

                    holsDate.Add(Convert.ToDateTime(s));
                }


            }



            if (DateToCheck.DayOfWeek == DayOfWeek.Sunday || DateToCheck.DayOfWeek == DayOfWeek.Saturday)
            {
                functionReturnValue = false;
            }
            else if (!holsDate.Contains(DateToCheck))
            {
                functionReturnValue = true;
            }
            else
            {
                functionReturnValue = false;
            }
            return functionReturnValue;

        }


        # endregion

        #region Misc Functions

        private void LoadXmldt(DataTable dt, string xmlFp, string xmlFn)
        {

            string myXMLfile = xmlFp + xmlFn + ".xml";

            DataSet ds = new DataSet();
            string dtName = dt.TableName;



            //check if current file exist and if so load tables into dataset
            if (File.Exists(myXMLfile))
            {
                // Create new FileStream with which to read the schema.
                System.IO.FileStream fsReadXml = new System.IO.FileStream
                    (myXMLfile, System.IO.FileMode.Open);
                try
                {
                    ds.ReadXml(fsReadXml);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {
                    fsReadXml.Close();
                }
            }



            for (int j = 0; j <= ds.Tables[dtName].Rows.Count - 1; j++)
            {
                //Adds a new row to the DataGridView for each line of text.
                dt.Rows.Add();

                //This for loop loops through the array in order to retrieve each
                //line of text.
                for (int i = 0; i <= dt.Columns.Count - 1; i++)
                {

                    //Sets the value of the cell to the value of the text retreived from the text file.
                    dt.Rows[dt.Rows.Count - 1][i] = ds.Tables[dtName].Rows[j].ItemArray[i];
                }

            }
        }

        private void saveXmlFile(DataTable dt, string xmlFp, string xmlFn, bool show)
        {
            string myXMLfile = xmlFp + xmlFn + ".xml";

            DataSet ds = new DataSet();


            //check if current file exist and if so load tables into dataset
            if (File.Exists(myXMLfile))
            {
                // Create new FileStream with which to read the schema.
                System.IO.FileStream fsReadXml = new System.IO.FileStream
                    (myXMLfile, System.IO.FileMode.Open);
                try
                {
                    ds.ReadXml(fsReadXml);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {
                    fsReadXml.Close();
                }

                // 
            }

            string dtName = dt.TableName;
            //  d_data.TableName = dtName;

            //check to see if there is a table for current ccy - if so delete old table and add new data else just at new table
            if (ds.Tables.Contains(dtName))
            {
                ds.Tables.Remove(dtName);

                DataTable dtCopy = new DataTable();
                dtCopy = dt.Copy();

                ds.Tables.Add(dtCopy);
            }
            else
            {
                ds.Tables.Add(dt);
            }

            //write xml file.

            ds.WriteXml(myXMLfile);

            if (show == true)
            {
                MessageBox.Show(xmlFn + " saved to: " + xmlFp);
            }


        }

        public double CrossVol(double Vol1, double Vol2, double Correl)
        {
            return Math.Pow((Vol1 * Vol1 + Vol2 * Vol2 - 2 * Correl * Vol1 * Vol2), 0.5);
        }

        //-----------------------------------------------------------------------------------------------------------------------------------'
        //   Returns the implied correlation given two vols and the cross volatility between them.                                           '
        //   Make sure to use the correct side of bid/ask.                                                                                   '
        //-----------------------------------------------------------------------------------------------------------------------------------'
        public double ImpliedCorrel(double Vol1, double Vol2, double CrossVol)
        {
            return (Vol1 * Vol1 + Vol2 * Vol2 - CrossVol * CrossVol) * (1 / (2 * Vol1 * Vol2));
        }


        //-----------------------------------------------------------------------------------------------------------------------------------'
        //   Returns the vega hedge ratios given two vols and the correlation between them.                                                  '
        //   Make sure to use the correct side of bid/ask:                                                                                   '
        //       Vol1Bid, Vol2Bid, CorrelAsk                                                                                                 '
        //       Vol1Ask, Vol2Ask, CorrelBid                                                                                                 '
        //   Terms - 1 for Vol1 hedge ratio, 2 for Vol2 hedge ratio                                                                          '
        //-----------------------------------------------------------------------------------------------------------------------------------'
        public double CrossVolHedgeRatio(double Vol1, double Vol2, double Correl, int terms)
        {
            double functionReturnValue = 0;
            if (terms == 1)
            {
                functionReturnValue = (Vol1 - Correl * Vol2) / (Math.Pow((Vol1 * Vol1 + Vol2 * Vol2 - 2 * Correl * Vol1 * Vol2), 0.5));
            }
            else if (terms == 2)
            {
                functionReturnValue = (Vol2 - Correl * Vol1) / (Math.Pow((Vol1 * Vol1 + Vol2 * Vol2 - 2 * Correl * Vol1 * Vol2), 0.5));
            }
            else
            {
                functionReturnValue = 9999;
            }
            return functionReturnValue;
        }

        public static System.Boolean IsNumeric(System.Object Expression)
        {
            if (Expression == null || Expression is DateTime)
                return false;

            if (Expression is Int16 || Expression is Int32 || Expression is Int64 || Expression is Decimal || Expression is Single || Expression is Double || Expression is Boolean)
                return true;

            try
            {
                if (Expression is string)
                    Double.Parse(Expression as string);
                else
                    Double.Parse(Expression.ToString());
                return true;
            }
            catch { } // just dismiss errors but return false
            return false;
        }

        private void loadData(DataTable dt, string fileName, int col)
        {
            String sLine = "";
            try
            {
                System.IO.StreamReader FileStream = new System.IO.StreamReader(fileName);
                sLine = FileStream.ReadLine();

                while (sLine != null)
                {
                    dt.Rows.Add();
                    for (int i = 0; i <= col; i++)
                    {
                        string[] s = sLine.Split(';');
                        dt.Rows[dt.Rows.Count - 1][i] = s[i].ToString();
                    }
                    sLine = FileStream.ReadLine();
                }

                FileStream.Close();
            }
            catch (Exception err)
            {
                System.Windows.Forms.MessageBox.Show("Error:  " + err.Message, "Program Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void loadDataFullFile(DataTable dv, string fileName, int col)
        {


            String sLine = "";

            try
            {

                System.IO.StreamReader FileStream = new System.IO.StreamReader(fileName);

                //You must set the value to false when you are programatically adding rows to
                //a DataGridView.  If you need to allow the user to add rows, you
                //can set the value back to true after you have populated the DataGridView
                // dv.AllowUserToAddRows = false;

                sLine = FileStream.ReadLine();
                //The while loop reads each line of text.
                while (sLine != null)
                {
                    //Adds a new row to the DataGridView for each line of text.
                    dv.Rows.Add();

                    //This for loop loops through the array in order to retrieve each
                    //line of text.
                    for (int i = 0; i <= col; i++)
                    {
                        //Splits each line in the text file into a string array
                        string[] s = sLine.Split(';');
                        //Sets the value of the cell to the value of the text retreived from the text file.
                        dv.Rows[dv.Rows.Count - 1][i] = s[i].ToString();
                    }
                    sLine = FileStream.ReadLine();
                }
                //Close the selected text file.
                FileStream.Close();
            }
            catch (Exception err)
            {
                //Display any errors in a Message Box.
                System.Windows.Forms.MessageBox.Show("Error:  " + err.Message, "Program Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


            //  dv.AllowUserToAddRows = true;
            // end of getsearchterms

        }
        #endregion

        #region statistic functions

        public double QNorm(double p, double mu, double sigma, bool lower_tail, bool log_p)
        {
            if (double.IsNaN(p) || double.IsNaN(mu) || double.IsNaN(sigma)) return (p + mu + sigma);
            double ans;
            bool isBoundaryCase = R_Q_P01_boundaries(p, double.NegativeInfinity, double.PositiveInfinity, lower_tail, log_p, out ans);
            if (isBoundaryCase) return (ans);
            if (sigma < 0) return (double.NaN);
            if (sigma == 0) return (mu);

            double p_ = R_DT_qIv(p, lower_tail, log_p);
            double q = p_ - 0.5;
            double r, val;

            if (Math.Abs(q) <= 0.425)  // 0.075 <= p <= 0.925
            {
                r = .180625 - q * q;
                val = q * (((((((r * 2509.0809287301226727 +
                           33430.575583588128105) * r + 67265.770927008700853) * r +
                         45921.953931549871457) * r + 13731.693765509461125) * r +
                       1971.5909503065514427) * r + 133.14166789178437745) * r +
                     3.387132872796366608)
                / (((((((r * 5226.495278852854561 +
                         28729.085735721942674) * r + 39307.89580009271061) * r +
                       21213.794301586595867) * r + 5394.1960214247511077) * r +
                     687.1870074920579083) * r + 42.313330701600911252) * r + 1.0);
            }
            else
            {
                r = q > 0 ? R_DT_CIv(p, lower_tail, log_p) : p_;
                r = Math.Sqrt(-((log_p && ((lower_tail && q <= 0) || (!lower_tail && q > 0))) ? p : Math.Log(r)));

                if (r <= 5)              // <==> min(p,1-p) >= exp(-25) ~= 1.3888e-11
                {
                    r -= 1.6;
                    val = (((((((r * 7.7454501427834140764e-4 +
                            .0227238449892691845833) * r + .24178072517745061177) *
                          r + 1.27045825245236838258) * r +
                         3.64784832476320460504) * r + 5.7694972214606914055) *
                       r + 4.6303378461565452959) * r +
                      1.42343711074968357734)
                     / (((((((r *
                              1.05075007164441684324e-9 + 5.475938084995344946e-4) *
                             r + .0151986665636164571966) * r +
                            .14810397642748007459) * r + .68976733498510000455) *
                          r + 1.6763848301838038494) * r +
                         2.05319162663775882187) * r + 1.0);
                }
                else                     // very close to  0 or 1 
                {
                    r -= 5.0;
                    val = (((((((r * 2.01033439929228813265e-7 +
                            2.71155556874348757815e-5) * r +
                           .0012426609473880784386) * r + .026532189526576123093) *
                         r + .29656057182850489123) * r +
                        1.7848265399172913358) * r + 5.4637849111641143699) *
                      r + 6.6579046435011037772)
                     / (((((((r *
                              2.04426310338993978564e-15 + 1.4215117583164458887e-7) *
                             r + 1.8463183175100546818e-5) * r +
                            7.868691311456132591e-4) * r + .0148753612908506148525)
                          * r + .13692988092273580531) * r +
                         .59983220655588793769) * r + 1.0);
                }
                if (q < 0.0) val = -val;
            }

            return (mu + sigma * val);
        }
        private static bool R_Q_P01_boundaries(double p, double _LEFT_, double _RIGHT_, bool lower_tail, bool log_p, out double ans)
        {
            if (log_p)
            {
                if (p > 0.0)
                {
                    ans = double.NaN;
                    return (true);
                }
                if (p == 0.0)
                {
                    ans = lower_tail ? _RIGHT_ : _LEFT_;
                    return (true);
                }
                if (p == double.NegativeInfinity)
                {
                    ans = lower_tail ? _LEFT_ : _RIGHT_;
                    return (true);
                }
            }
            else
            {
                if (p < 0.0 || p > 1.0)
                {
                    ans = double.NaN;
                    return (true);
                }
                if (p == 0.0)
                {
                    ans = lower_tail ? _LEFT_ : _RIGHT_;
                    return (true);
                }
                if (p == 1.0)
                {
                    ans = lower_tail ? _RIGHT_ : _LEFT_;
                    return (true);
                }
            }
            ans = double.NaN;
            return (false);
        }

        private static double R_DT_qIv(double p, bool lower_tail, bool log_p)
        {
            return (log_p ? (lower_tail ? Math.Exp(p) : -ExpM1(p)) : R_D_Lval(p, lower_tail));
        }

        private static double R_DT_CIv(double p, bool lower_tail, bool log_p)
        {
            return (log_p ? (lower_tail ? -ExpM1(p) : Math.Exp(p)) : R_D_Cval(p, lower_tail));
        }

        private static double R_D_Lval(double p, bool lower_tail)
        {
            return lower_tail ? p : 0.5 - p + 0.5;
        }

        private static double R_D_Cval(double p, bool lower_tail)
        {
            return lower_tail ? 0.5 - p + 0.5 : p;
        }
        private static double ExpM1(double x)
        {
            if (Math.Abs(x) < 1e-5)
                return x + 0.5 * x * x;
            else
                return Math.Exp(x) - 1.0;
        }


        public double InvesreDistrib(double input)
        {
            int i = 1;
            double t = new double();
            double y = new double();
            double a1 = 0.254829592;
            double a2 = -0.284496736;
            double a3 = 1.421413741;
            double a4 = -1.453152027;
            double a5 = 1.061405429;
            double p = 0.3275911;
            if (input < 0)
            {
                i = -1;
            }
            input = Math.Abs(input) / Math.Pow(input, 0.5);
            t = 1 / (1 + p * input);
            y = 1 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Math.Exp(-input * input);
            return 0.5 * (1 + i * y);

        }
        private double StanNormCumDistr(double x)
        {
            double RT2PI = Math.Pow((2 * 3.14159265359), 0.5);
            double SPLIT = 7.07106781186547;
            double N0 = 220.206867912376;
            double N1 = 221.213596169931;
            double N2 = 112.079291497871;
            double N3 = 33.912866078383;
            double N4 = 6.37396220353165;
            double N5 = 0.700383064443688;
            double N6 = 3.52624965998911e-02;
            double M0 = 440.413735824752;
            double M1 = 793.826512519948;
            double M2 = 637.333633378831;
            double M3 = 296.564248779674;
            double M4 = 86.7807322029461;
            double M5 = 16.064177579207;
            double M6 = 1.75566716318264;
            double M7 = 8.83883476483184e-02;
            double z = Math.Abs(x);
            double c = 0;
            if (z <= 37.0)
            {
                double e = Math.Exp(-z * z / 2.0);
                if (z < SPLIT)
                {
                    double n = (((((N6 * z + N5) * z + N4) * z + N3) * z + N2) * z + N1) * z + N0;
                    double d = ((((((M7 * z + M6) * z + M5) * z + M4) * z + M3) * z + M2) * z + M1) * z + M0;
                    c = e * n / d;
                }
                else
                {
                    double f = z + 1.0 / (z + 2.0 / (z + 3.0 / (z + 4.0 / (z + 13.0 / 20.0))));
                    c = e / (RT2PI * f);
                }
            }
            if (x <= 0)
                return c;
            else
                return 1 - c;

        }

        #endregion

        #region bbgCallFunctions

        private void getholsNew()
        {

            if (holidaySet == null)
                holidaySet = new DataSet();

            if (holidaySet != null)
            {
                holidaySet.Reset();
            }


            if (holidays.d_data.Columns.Count > 1)
                holidays.d_data.Columns.Remove("CALENDAR_NON_SETTLEMENT_DATES");

            List<string> ccy = ccyDets.AsEnumerable().Select(x => x[0].ToString()).ToList();


            string fromDate = DateTime.Now.ToString("yyyyMMdd");
            string toDate = DateTime.Now.AddYears(2).ToString("yyyyMMdd");

            holidays.d_data.Columns.Add("CALENDAR_NON_SETTLEMENT_DATES");

            ListViewItem item1 = holidays.listViewOverrides.Items.Add("CALENDAR_START_DATE");
            ListViewItem item2 = holidays.listViewOverrides.Items.Add("CALENDAR_END_DATE");

            item1.SubItems.Add(fromDate);
            item2.SubItems.Add(toDate);


            foreach (DataRow row in ccyDets.Rows)
            {
                int i = ccy.IndexOf(row["Ccy"].ToString());
                string cdrCode = ccyDets.Rows[i]["CdrCode"].ToString();

                ListViewItem item = holidays.listViewOverrides.Items.Add("SETTLEMENT_CALENDAR_CODE");
                item.SubItems.Add(cdrCode);

                string ccyCal = row["Ccy"].ToString() + " CURNCY";

                holidays.d_data.Rows.Add(ccyCal);

                holidays.sendRequest();



                DataRow lastRow = holidays.d_data.Rows[0];
                bool endload = false;

                do
                {
                    if (lastRow["CALENDAR_NON_SETTLEMENT_DATES"] != DBNull.Value)
                        endload = true;

                } while (endload == false);

                DataTable hols = holidays.d_bulkData.Tables["CALENDAR_NON_SETTLEMENT_DATES"].Copy();
                hols.TableName = ccyCal.Substring(0, 3);

                holidaySet.Tables.Add(hols);

                holidays.d_data.Rows.Clear();
                item.Remove();


            }


            string xml = "holidayDataNew";
            DataSet ds = holidaySet;
            saveDatasetXml(xml, ds);


            MessageBox.Show("Holidays Updated");


        }

        private void LoadHols()
        {
            if (holidaySet == null)
                holidaySet = new DataSet();

            holidaySet.Reset();

            xmlFileName = "holidayDataNew";

            string myXMLfile = xmlFilePath + xmlFileName + ".xml";

            //check if current file exist and if so load tables into dataset
            if (File.Exists(myXMLfile))
            {
                // Create new FileStream with which to read the schema.
                System.IO.FileStream fsReadXml = new System.IO.FileStream
                    (myXMLfile, System.IO.FileMode.Open);
                try
                {
                    holidaySet.ReadXml(fsReadXml);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {
                    fsReadXml.Close();
                }
            }


        }

        private void saveDatasetXml(string xmlFile, DataSet ds)
        {

            string myXMLfile = xmlFilePath + xmlFile + ".xml";

            try
            {
                ds.WriteXml(myXMLfile);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        //private void saveAllHolidays()
        //{

        //    foreach (DataRow row in ccyDets.Rows)
        //    {
        //        string curr = row["Ccy"].ToString();
        //        string cdrCode = row["CdrCode"].ToString();

        //        getHolidays(curr, cdrCode);

        //    }

        //    MessageBox.Show(xmlFileName + ".xml saved to: " + xmlFilePath);


        //}

        //private void getHolidays(string curr, string cdrCode)
        //{
        //    Bloomberglp.Blpapi.Examples.getBbgData holidays = new Bloomberglp.Blpapi.Examples.getBbgData();

        //    string dtName = curr;


        //    System.Text.StringBuilder sec = new System.Text.StringBuilder();
        //    System.Text.StringBuilder field = new System.Text.StringBuilder();
        //    System.Text.StringBuilder over = new System.Text.StringBuilder();
        //    System.Text.StringBuilder overVal = new System.Text.StringBuilder();

        //    string fromDate = DateTime.Now.ToString("yyyyMMdd");
        //    string toDate = DateTime.Now.AddYears(2).ToString("yyyyMMdd");
        //    //string cCode = "US";


        //    string d_sec = curr + " CURNCY"; //textBox1.Text;
        //    string d_field = "CALENDAR_NON_SETTLEMENT_DATES"; //textBox2.Text;
        //    string d_over = "SETTLEMENT_CALENDAR_CODE," + "CALENDAR_START_DATE," + "CALENDAR_END_DATE";
        //    string d_overVal = cdrCode + "," + fromDate + "," + toDate;

        //    sec.AppendLine(d_sec);
        //    field.AppendLine(d_field);
        //    over.AppendLine(d_over);
        //    overVal.AppendLine(d_overVal);

        //    holidays.get_BBG(d_sec, d_field, d_over, d_overVal, xmlFilePath, xmlFileName, dtName);
        //}

        //private void bbgStrikevol(string strike, string expiryDate, string ccyPair)
        //{
        //    string myXMLfile = xmlFilePath + "StrikeVol.xml";

        //    FileInfo info = new FileInfo(myXMLfile);
        //    lastUpdate = info.LastWriteTime;


        //    Bloomberglp.Blpapi.Examples.getBbgData strikeVol = new Bloomberglp.Blpapi.Examples.getBbgData();

        //    string xmlFn = "strikeVol";
        //    string dtName = "dtStrikeVol";

        //    System.Text.StringBuilder sec = new System.Text.StringBuilder();
        //    System.Text.StringBuilder field = new System.Text.StringBuilder();
        //    System.Text.StringBuilder over = new System.Text.StringBuilder();
        //    System.Text.StringBuilder overVal = new System.Text.StringBuilder();

        //    string d_sec = ccyPair + " CURNCY";
        //    string d_field = "sp vol surf mid";
        //    string d_over = "vol surf delta ovr," + "vol surf strike ovr," + "vol_surf_expiry_ovr";
        //    string d_overVal = "0," + strike + "," + expiryDate;

        //    sec.AppendLine(d_sec);
        //    field.AppendLine(d_field);
        //    over.AppendLine(d_over);
        //    overVal.AppendLine(d_overVal);

        //    strikeVol.get_BBG(d_sec, d_field, d_over, d_overVal, xmlFilePath, xmlFn, dtName);





        //}

        //private void setStrikeVol()
        //{

        //    int curCol = dataGridView1.CurrentCell.ColumnIndex;


        //    dtStrikeVol.Rows.Clear();
        //    string ccyPair = pricer.Rows[CcyPairR][curCol].ToString();
        //    DateTime exp = Convert.ToDateTime(pricer.Rows[ExpiryDateR][curCol].ToString().Substring(4));
        //    double autostrike = Convert.ToDouble(pricer.Rows[AutoStrikeR][curCol]);

        //    bbgStrikevol(autostrike.ToString("0.0000"), exp.ToString("yyyMMdd"), ccyPair);



        //    pricer.Rows[BloombergVolR][curCol] = "";


        //    string myXMLfile = xmlFilePath + "StrikeVol.xml";

        //    DateTime dtUpdated;
        //    string vol;

        //    do
        //    {
        //        dtUpdated = File.GetLastWriteTime(myXMLfile);

        //    } while (dtUpdated == lastUpdate);


        //    DataSet ds = new DataSet();
        //    string dtName = "dtStrikeVol";

        //    //check if current file exist and if so load tables into dataset
        //    if (File.Exists(myXMLfile))
        //    {
        //        // Create new FileStream with which to read the schema.
        //        System.IO.FileStream fsReadXml = new System.IO.FileStream
        //            (myXMLfile, System.IO.FileMode.Open);
        //        try
        //        {
        //            ds.ReadXml(fsReadXml);

        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show(ex.ToString());
        //        }
        //        finally
        //        {
        //            fsReadXml.Close();
        //        }
        //    }

        //    vol = ds.Tables[dtName].Rows[0].ItemArray[1].ToString();

        //    pricer.Rows[BloombergVolR][curCol] = vol;

        //}

        private void addBbgInterface(Bloomberglp.Blpapi.Examples.Form1 f, TabPage tp)
        {
            f.TopLevel = false;
            f.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;

            f.Visible = true;

            tp.Controls.Add(f);

        }


        #endregion

        #region buttons

        private void button4_Click(object sender, EventArgs e)
        {
            //save ccy setup
            saveXmlFile(ccyDets, xmlFilePath, "ccySetupDets", true);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            //save cross DataTable
            saveXmlFile(crosses, xmlFilePath, "currSetup", true);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void refreshAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            refreshDataButton();

        }

        private void clearAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            clearAllColumns();
            formatPricerDt();
        }

        private void refreshHolidayDataToolStripMenuItem_Click(object sender, EventArgs e)
        {

            getholsNew();
        }

        private void testToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            marketMakeRun("");
        }

        private void testToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            string ccyPair = ((DataTable)dataGridView8.DataSource).TableName;
            setSurfaceNew(ccyPair);
        }

        private void testToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            saveSmileMult();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

            DataSet ds = smileMult;
            string xml = "smileMultipliers";
            saveDatasetXml(xml, ds);
            MessageBox.Show("Smiles Updated");
        }

        private void testToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toggleSurfData();

        }

        private void copyMxSmileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            murexSmile("");
        }

        private void copyMxSmileToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            murexSmile("");
        }

        #endregion

        #region skewSheet


        private void uploadMxData()
        {
            if (skewSheet == null)
                skewSheet = new DataSet();

            skewSheet.Reset();

            DataTable dt = new DataTable();
            string[] col = new string[] { "ccyPair", "Maturity", "C/P", "strike", "Nom1", "Nom2", "VOL" };
            foreach (string s in col)
            {
                dt.Columns.Add(s);
            }


            dt.TableName = "OptList";
            skewSheet.Tables.Add(dt);


            DataTable dts = new DataTable();
            string[] cols = new string[] { "Change", "Spot", "Atm", "RR", "Fly" };

            foreach (string s in cols)
            {
                dts.Columns.Add(s);
            }

            dts.TableName = "scenarios";
            skewSheet.Tables.Add(dts);


            //load data from text files 
            string Xpath = systemFiles + user + @"\skewSheet\";
            string fn = Xpath + "mxData.txt";
            string fns = Xpath + "scenarios.txt";


            loadData(dt, fn, dt.Columns.Count - 1);
            loadData(dts, fns, dts.Columns.Count - 1);

            //clear emptyrows
            dt = dt.Rows.Cast<DataRow>().Where(row => !row.ItemArray.All(field => field is DBNull || string.IsNullOrWhiteSpace(field as string))).CopyToDataTable();




        }

        private DataTable nSurfDt(String s)
        {

            DataTable nSurf = new DataTable();
            nSurf.Columns.Add("DayCount",typeof(double));
            nSurf.Columns.Add("ATM");
            nSurf.Columns.Add("RR");
            nSurf.Columns.Add("FLY");
            nSurf.AcceptChanges();
            nSurf.TableName = s;

            return nSurf;

        }

        private void skewSheetVols(string ccyPair, bool wgtShift)
        {
            uploadMxData(); //uploads list of options and scenarios from text files. 
            //option data load from excel and saved into skewsheet dataset
            DataTable dt = skewSheet.Tables["OptList"];
            DataTable dtS = skewSheet.Tables["Scenarios"];

            string filter = "[ccyPair] = '" + ccyPair + "'";
            DataRow[] resultFiler = dt.Select(filter);

            if (allSurfs == null)
                allSurfs = new DataSet();

            allSurfs.Tables.Clear();

            scenRun = 0; //need to reset test to add scenarios to dropdown list


            //name each datatable for screnrio pct shift
            foreach (DataRow r in dtS.Rows)
            {
                string tableName = r.ItemArray[0].ToString();

                allSurfs.Tables.Add(nSurfDt(tableName));
            }


            //dt = dt.Rows.Cast<DataRow>().Where(row => !row.ItemArray.All(field => field is DBNull || string.IsNullOrWhiteSpace(field as string))).CopyToDataTable();

            List<string> sec = fwds.d_data.AsEnumerable().Select(x => x[0].ToString().Substring(0, 6)).ToList();
            int s = sec.IndexOf(ccyPair);

            double rSpot = Convert.ToDouble(fwds.d_data.Rows[s]["PX_MID"]);

            //clear emptyrows

            int rowLen = dt.Rows.Count;
            int colLen = 9;

            DataTable volsUnshocked = new DataTable();
            DataTable volsShocked = new DataTable();
            DataTable atmUnshocked = new DataTable();
            DataTable atmShocked = new DataTable();
            DataTable depos = new DataTable();

            volsUnshocked.TableName = "volsUnshocked";
            volsShocked.TableName = "volsShocked";
            atmUnshocked.TableName = "atmUnshocked";
            atmShocked.TableName = "atmShocked";
            depos.TableName = "depos";

            //setupdatatables
            for (int ii = 0; ii < rowLen; ii++)
            {
                volsUnshocked.Rows.Add();
                volsShocked.Rows.Add();
                atmUnshocked.Rows.Add();
                atmShocked.Rows.Add();
                depos.Rows.Add();
            }


            for (int iii = 0; iii < colLen; iii++)
            {
                volsUnshocked.Columns.Add();
                volsShocked.Columns.Add();
                atmShocked.Columns.Add();
            }

            atmUnshocked.Columns.Add();
            volsUnshocked.Columns.Add("Mat");
            volsUnshocked.Columns.Add("Strike");
            depos.Columns.Add("depoFor");
            depos.Columns.Add("depoDom");


            //gets daycount basis from ccydets dataTable

            int[] arr = dayCountBasis(ccyPair);
            int basisB = arr[0];
            int basisT = arr[1];

            // calls method to get cross info 
            object[] crossInfo = crossDtData(ccyPair);
            int volType = Convert.ToInt16(crossInfo[0]);
            double factor = Convert.ToDouble(crossInfo[1]);
            string bbgSource = crossInfo[2].ToString();

            //is needed to convert old delatype to new premo included. old was 1 = premo 2 = no, now 1 = premo 0 = no
            int premoInc = 0;


            if (volType == 1)
            {
                premoInc = 1;
            }
            else
            {
                premoInc = 0;
            }

            string baseCcy = ccyPair.Substring(0, 3);
            string termsCcy = ccyPair.Substring(3, 3);

            List<string> scenList = new List<string>(); // list to populate dropdown menu for weigth vol scenarios

            //cycle through each option to get vols for each sceneario
            int i = 0;
            // foreach (DataRow opt in dt.Rows)
            foreach (DataRow opt in resultFiler)
            {

                string expiryText = opt.ItemArray[1].ToString();
                string pC = opt.ItemArray[2].ToString();
                string strText = opt.ItemArray[3].ToString();

                DateTime dayStart = today;
                DateTime sptDate = SpotDate(dayStart, baseCcy, termsCcy);
                DateTime autoExp = Convert.ToDateTime(expiryText);
                DateTime delDate = SpotDate(autoExp, baseCcy, termsCcy);
                double dayCount = (autoExp - dayStart).TotalDays; //expiry to trade date
                double delDayCount = (delDate - sptDate).TotalDays; // delivery to spot date
                double autoSt = Convert.ToDouble(strText);

                //get fly, rr from pricerData displayed on main pricer screen 
                double[] volComponents = null;

                volComponents = volBuilder(ccyPair, dayCount);
                double atmVol = volComponents[0];
                double rr = volComponents[1];
                double fly = volComponents[2];
                double wingControl = volComponents[3];
                double targetFlyMult = volComponents[4];

                // double Startspot = Convert.ToDouble(dtS.Rows[4][1]);

                double[] rateComponents = rateBuilder(ccyPair, dayCount, delDayCount, rSpot, factor, basisB, basisT);
                //fwdPts, outRight, forDepo, domDepo

                double forDepo = rateComponents[2];
                double domDepo = rateComponents[3];

                depos.Rows[i][0] = domDepo * 100;
                depos.Rows[i][1] = forDepo * 100;

                double dfForExp = DiscountFactor(forDepo, Convert.ToInt16(dayCount), Convert.ToInt16(basisB));
                double dfForDel = DiscountFactor(forDepo, Convert.ToInt16(delDayCount), Convert.ToInt16(basisB));
                double dfDomExp = DiscountFactor(domDepo, Convert.ToInt16(dayCount), Convert.ToInt16(basisT));
                double dfDomDel = DiscountFactor(domDepo, Convert.ToInt16(delDayCount), Convert.ToInt16(basisT));

                int j = 0;
                foreach (DataRow r in dtS.Rows)
                {

                    //double spot = Convert.ToDouble(r.ItemArray[1]);
                    double spot = rSpot * (1 + Convert.ToDouble(r.ItemArray[0]));

                    //need to calc smilefly then get 25d and atm strikes and vols
                    double smileFly = equivalentfly(spot, dayStart, autoExp, atmVol * wingControl, atmVol, rr, fly, dfDomDel, dfForDel, premoInc);
                    double putVol = atmVol + smileFly - rr / 2;
                    double callVol = atmVol + smileFly + rr / 2;

                    double putStrike = FXStrikeVol(spot, dayStart, autoExp, 0.25, putVol, dfDomDel, dfForDel, "p", premoInc);
                    double callStrike = FXStrikeVol(spot, dayStart, autoExp, 0.25, callVol, dfDomDel, dfForDel, "c", premoInc);
                    double atmStrike = FXATMStrike(spot, dayStart, autoExp, atmVol, dfDomDel, dfForDel, premoInc);

                    double vol = smileInterp(spot, dayStart, autoExp, wingControl * atmVol, autoSt, putStrike, putVol, atmStrike, atmVol, callStrike, callVol, dfDomExp, dfForExp, dfDomDel, dfForDel);

                    if (vol == 0) { vol = 0.0001; } //added this as if vol fails at least option can calculate intrisic value if in the money 

                    volsUnshocked.Rows[i][j] = vol * 100;

                    if (j == 0) { atmUnshocked.Rows[i][0] = atmVol * 100; }

                    //parrarel shift in vols 
                    double atmShift = Convert.ToDouble(r.ItemArray[2]) / 100;
                    double rrShift = Convert.ToDouble(r.ItemArray[3]) / 100;
                    double flyShift = Convert.ToDouble(r.ItemArray[4]) / 100;

                    double wgt = 0;

                    if (wgtShift == true)
                    {
                        //set 2m as fixed weight. sub 1m options are weighted flat to 1m and not vega weigthed. 
                        int fixedDate = 60;

                        if (dayCount < 30)
                        {
                            wgt = Math.Sqrt(fixedDate / 30);
                        }
                        else
                        {
                            wgt = Math.Sqrt(fixedDate / dayCount);
                        }

                        atmShift = atmShift * wgt;
                        rrShift = rrShift * wgt;
                        flyShift = flyShift * wgt;
                    }


                    DataTable dtx = allSurfs.Tables[r.ItemArray[0].ToString()];
                    dtx.Rows.Add(new Object[] { dayCount, atmShift.ToString("0.0000%"), rrShift.ToString("0.0000%"), flyShift.ToString("0.0000%") });
 
                    //getShocked Vol
                    double smileFlyS = equivalentfly(spot, dayStart, autoExp, (atmVol + atmShift) * wingControl, (atmVol + atmShift), (rr + rrShift), (fly + flyShift), dfDomDel, dfForDel, premoInc);
                    double putVolS = (atmVol + atmShift) + smileFly - (rr + rrShift) / 2;
                    double callVolS = (atmVol + atmShift) + smileFly + (rr + rrShift) / 2;

                    double putStrikeS = FXStrikeVol(spot, dayStart, autoExp, 0.25, putVolS, dfDomDel, dfForDel, "p", premoInc);
                    double callStrikeS = FXStrikeVol(spot, dayStart, autoExp, 0.25, callVolS, dfDomDel, dfForDel, "c", premoInc);
                    double atmStrikeS = FXATMStrike(spot, dayStart, autoExp, (atmVol + atmShift), dfDomDel, dfForDel, premoInc);

                    double volS = smileInterp(spot, dayStart, autoExp, wingControl * (atmVol + atmShift), autoSt, putStrikeS, putVolS, atmStrikeS, (atmVol + atmShift), callStrikeS, callVolS, dfDomExp, dfForExp, dfDomDel, dfForDel);

                    if (volS == 0) { volS = 0.0001; }
                    volsShocked.Rows[i][j] = volS * 100;
                    atmShocked.Rows[i][j] = (atmVol + atmShift) * 100;

                    j++;
                }

                volsUnshocked.Rows[i][9] = expiryText;
                volsUnshocked.Rows[i][10] = strText;

                i++;

            }

            //add starting spot ref to last cell of atmUnshockedDatatable 
            atmUnshocked.Rows.Add(rSpot);


            //add shift tupe to last cell of volsunshocked
            string shiftType;
            shiftType = "Parallel Vol Shift";
            if (wgtShift == true) {shiftType = "Weighted Vol Shift";}
         
            volsUnshocked.Rows.Add(shiftType);

            //remove duplicate daycounts from vol shift tables
            foreach (DataTable dd in allSurfs.Tables)
            {
                RemoveDuplicateRows(dd, "DayCount");

            }

            string Xpath = systemFiles + user + @"\skewSheet\";

            saveXmlFile(volsUnshocked, Xpath, volsUnshocked.TableName, false);
            saveXmlFile(volsShocked, Xpath, volsShocked.TableName, false);
            saveXmlFile(atmUnshocked, Xpath, atmUnshocked.TableName, false);
            saveXmlFile(atmShocked, Xpath, atmShocked.TableName, false);
            saveXmlFile(depos, Xpath, depos.TableName, false);


            //adds dropdown list to show scenrios of weigth vols shift
            foreach(DataRow rr in dtS.Rows){
                scenList.Add(rr.ItemArray[0].ToString());
            
            }
            //Setup data binding   
            toolStripComboBox3.ComboBox.BindingContext = this.BindingContext;
            toolStripComboBox3.ComboBox.DataSource = scenList;
            // make it readonly
           this.toolStripComboBox3.DropDownStyle = ComboBoxStyle.DropDownList;
           scenRun = 1;

            MessageBox.Show("Vol Calc is done...");

        }

        public DataTable RemoveDuplicateRows(DataTable table, string DistinctColumn)
        {
            try
            {
                ArrayList UniqueRecords = new ArrayList();
                ArrayList DuplicateRecords = new ArrayList();

                // Check if records is already added to UniqueRecords otherwise,
                // Add the records to DuplicateRecords
                foreach (DataRow dRow in table.Rows)
                {
                    if (UniqueRecords.Contains(dRow[DistinctColumn]))
                        DuplicateRecords.Add(dRow);
                    else
                        UniqueRecords.Add(dRow[DistinctColumn]);
                }

                // Remove duplicate rows from DataTable added to DuplicateRecords
                foreach (DataRow dRow in DuplicateRecords)
                {
                    table.Rows.Remove(dRow);
                }

                // Return the clean DataTable which contains unique records.
                return table;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private void showLoadedData(DataTable dt)
        {


            WindowsFormsApplication1.RiskyPopUP bulkData = new WindowsFormsApplication1.RiskyPopUP(dt);
            bulkData.ShowDialog(this);
        }

        private void fill_tool_bar()
        {

            List<string> colA = new List<string>();

            foreach (DataRow row in crosses.Rows)
            {
                colA.Add(row["Cross"].ToString().ToUpper());

            }

            //Setup data binding   
            toolStripComboBox1.ComboBox.BindingContext = this.BindingContext;
            toolStripComboBox1.ComboBox.DataSource = colA;
            // make it readonly

            this.toolStripComboBox1.DropDownStyle = ComboBoxStyle.DropDownList;

            toolStripComboBox2.ComboBox.BindingContext = this.BindingContext;
            toolStripComboBox2.ComboBox.DataSource = colA;
            // make it readonly

            this.toolStripComboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
            toolStripComboBox4.ComboBox.BindingContext = this.BindingContext;
            toolStripComboBox4.ComboBox.DataSource = colA;
            // make it readonly

            this.toolStripComboBox4.DropDownStyle = ComboBoxStyle.DropDownList;

            toolStripComboBox5.ComboBox.BindingContext = this.BindingContext;
            toolStripComboBox5.ComboBox.DataSource = colA;
            // make it readonly

            this.toolStripComboBox5.DropDownStyle = ComboBoxStyle.DropDownList;

        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (endIni == 1)
            {


                string ccy = toolStripComboBox1.Text;
                displaySurface(ccy);

                DialogResult dialogResult = MessageBox.Show("Calculate Vols for " + ccy + "?", "", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    skewSheetVols(ccy, false);
                }
                else if (dialogResult == DialogResult.No)
                {
                    //do something else
                }


            }

        }

        private void toolStripComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (endIni == 1)
            {


                string ccy = toolStripComboBox2.Text;
                displaySurface(ccy);

                DialogResult dialogResult = MessageBox.Show("Calculate Vols for " + ccy + "?", "", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    skewSheetVols(ccy, true);
                }
                else if (dialogResult == DialogResult.No)
                {
                    //do something else
                }


            }
        }

        private void toolStripComboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {


            if (scenRun == 1)
            {


                string tName = toolStripComboBox3.Text;


                DataTable dt = new DataTable();
                dt = allSurfs.Tables[tName];
                showLoadedData(dt);

            }

        }

        private void showOptionListToolStripMenuItem_Click(object sender, EventArgs e)
        {
            uploadMxData();

            DataTable dt = new DataTable();
            dt = skewSheet.Tables["optList"];

            string filter = "[ccyPair] = '" + toolStripComboBox1.Text + "'";
            DataRow[] resultFiler = dt.Select(filter);

            DataTable dtClone = dt.Clone();
            foreach (DataRow r in resultFiler)
            {
                dtClone.ImportRow(r);

            }

            showLoadedData(dtClone);


        }

        private void replicaiton()
        {
            uploadMxData();
            DataTable dt = new DataTable();
            dt = skewSheet.Tables["optList"];

            DataSet ds = new DataSet();

            string filter = "[ccyPair] = '" + toolStripComboBox5.Text + "'";
            DataRow[] resultFiler = dt.Select(filter);

            DataTable dtClone = dt.Clone();
            foreach (DataRow r in resultFiler)
            {
                dtClone.ImportRow(r);

            }
            
            DataTable dtr = new DataTable();
            string[] col = new string[] { "Expiry", "Strike", "Notional", "Delta", "sVega", "sGamma", "sega10","sega25", "rega10", "rega25" };
            foreach (string s in col)
            {
                dtr.Columns.Add(s);
            }

            string ccyPair = dtClone.Rows[0]["ccyPair"].ToString();

            List<string> cross = fwds.d_data.AsEnumerable().Select(x => x[0].ToString().Substring(0, 6)).ToList();
            int i = cross.IndexOf(ccyPair);

            double spot  = Convert.ToDouble(fwds.d_data.Rows[i]["PX_MID"]);//get spot from d_data table

            foreach (DataRow r in dtClone.Rows)
            {
                
                string mat = r["Maturity"].ToString();
                string strike = r["strike"].ToString();
                string pC = r["C/P"].ToString().Substring(0, 1);
                double notional = Convert.ToDouble(r["Nom1"])/1000000;

                object[] retVal =  optPricerRep(ccyPair, mat, strike, pC, spot, notional);
                dtr.Rows.Add(retVal);

            }

            


            DataTable bucket = new DataTable();
            string[] col1 = new string[] { "Term", "sVega", "sGamma", "sega10","sega25","rega10", "rega25" };
           
            foreach (string s in col1)
            {
                bucket.Columns.Add(s);
            }

            DataTable pd = pricingData.Tables[ccyPair];

            for (int j = 0; j < pd.Rows.Count; j++)
            {
                double sVega = 0;
                double sGamma = 0;
                double sega25 = 0;
                double sega10 = 0;
                double rega25 = 0;
                double rega10 = 0;

                DataTable breakDown = new DataTable();
                breakDown = dtr.Clone();

                string term = pd.Rows[j].ItemArray[1].ToString();

                if (j == 0)
                {

                }

                else
                {

                    int dayCount = Convert.ToInt16(pd.Rows[j-1]["DayCount"]);

                    DateTime cDay = DateTime.Now;
                    DateTime lb = cDay.AddDays(dayCount);
                    DateTime ub = cDay;


                    if (j == pd.Rows.Count - 1)
                    {
                        ub = cDay.AddYears(100);
                    }

                    else
                    {
                        int dayCount2 = Convert.ToInt16(pd.Rows[j]["DayCount"]);
                        ub = cDay.AddDays(dayCount2);
                    }

                    foreach (DataRow row in dtr.Rows)
                    {
                        DateTime exp = Convert.ToDateTime(row["Expiry"]);

                        if (exp >= lb && exp < ub)
                        {
                            sVega += Convert.ToDouble(row["sVega"]);
                            sGamma += Convert.ToDouble(row["sGamma"]);
                            sega25 += Convert.ToDouble(row["sega25"]);
                            sega10 += Convert.ToDouble(row["sega10"]);
                            rega25 += Convert.ToDouble(row["rega25"]);
                            rega10 += Convert.ToDouble(row["rega10"]);

                             DataRow newRow = breakDown.NewRow();
                             newRow.ItemArray = row.ItemArray;
                             breakDown.Rows.Add(newRow);                           

                        }

                    }
                }

                string tName = ccyPair + term.ToUpper();
                breakDown.TableName = tName;

                if (regaSega.Tables.Contains(tName))
                {
                    regaSega.Tables.Remove(tName);
                }

                regaSega.Tables.Add(breakDown);

                bucket.Rows.Add(tName, sVega.ToString("#,##0"), sGamma.ToString("#,##0"), sega10.ToString("#,##0"), sega25.ToString("#,##0"), rega10.ToString("#,##0"), rega25.ToString("#,##0"));
            }


            double sVegat = 0;
            double sGammat = 0;
            double sega25t = 0;
            double sega10t = 0;
            double rega25t = 0;
            double rega10t = 0;

            foreach (DataRow row in bucket.Rows)
            {
                sVegat += Convert.ToDouble(row["sVega"]);
                sGammat += Convert.ToDouble(row["sGamma"]);
                sega25t += Convert.ToDouble(row["sega25"]);
                sega10t += Convert.ToDouble(row["sega10"]);
                rega25t += Convert.ToDouble(row["rega25"]);
                rega10t += Convert.ToDouble(row["rega10"]);

            }

             bucket.Rows.Add("TOTAL:", sVegat.ToString("#,##0"), sGammat.ToString("#,##0"), sega10t.ToString("#,##0"), sega25t.ToString("#,##0"), rega10t.ToString("#,##0"), rega25t.ToString("#,##0"));
             bucket.TableName = "Bucketed";
             dtr.TableName = "All";


             ds.Tables.Add(bucket);
             ds.Tables.Add(dtr);

             foreach (DataTable dt1 in regaSega.Tables)
             {
                 DataTable dc = dt1.Copy();
                 ds.Tables.Add(dc);
             }


            //showRisk(bucket,dtr);
             showRiskNew(ds);
           
        }

        private object[] optPricerRep(string ccyPair, string expiryText, string strText, string pC, double spot, double notional)
        {

            //  setMarketData();
           // double spot = Convert.ToDouble(pricer.Rows[SpotR][curCol]);
           // string ccyPair = pricer.Rows[CcyPairR][curCol].ToString();

            object[] retVal = null;
            string baseCcy = ccyPair.Substring(0, 3);
            string termsCcy = ccyPair.Substring(3, 3);


            DateTime dayStart = today;
            DateTime sptDate = SpotDate(dayStart, baseCcy, termsCcy);
            DateTime autoExp = AutoExpiryDate(expiryText, dayStart, homeCcy, baseCcy, termsCcy);
            DateTime delDate = SpotDate(autoExp, baseCcy, termsCcy);
            double dayCount = (autoExp - dayStart).TotalDays; //expiry to trade date
            double delDayCount = (delDate - sptDate).TotalDays; // delivery to spot date


            //gets daycount basis from ccydets dataTable

            int[] arr = dayCountBasis(ccyPair);
            int basisB = arr[0];
            int basisT = arr[1];

            // calls method to get cross info 
            object[] crossInfo = crossDtData(ccyPair);
            int volType = Convert.ToInt16(crossInfo[0]);
            double factor = Convert.ToDouble(crossInfo[1]);
            string bbgSource = crossInfo[2].ToString();

            //need to change smile factor for usdrub to control number of strikes are calcuted with cubic spline function
            double smileFactor = factor;
            if (ccyPair == "USDRUB" || ccyPair == "EURRUB" || ccyPair == "USDTRY") { smileFactor = 100; }


            //is needed to convert old delatype to new premo included. old was 1 = premo 2 = no, now 1 = premo 0 = no
            int premoInc = 0;


            if (volType == 1)
            {
                premoInc = 1;
            }
            else
            {
                premoInc = 0;
            }

            double[] rateComponents = rateBuilder(ccyPair, dayCount, delDayCount, spot, factor, basisB, basisT);
            //fwdPts, outRight, forDepo, domDepo

            double fwdPts = rateComponents[0];
            double outRight = rateComponents[1];
            double forDepo = rateComponents[2];
            double domDepo = rateComponents[3];


            double dfForExp = DiscountFactor(forDepo, Convert.ToInt16(dayCount), Convert.ToInt16(basisB));
            double dfForDel = DiscountFactor(forDepo, Convert.ToInt16(delDayCount), Convert.ToInt16(basisB));

            double dfDomExp = DiscountFactor(domDepo, Convert.ToInt16(dayCount), Convert.ToInt16(basisT));
            double dfDomDel = DiscountFactor(domDepo, Convert.ToInt16(delDayCount), Convert.ToInt16(basisT));


            //get fly, rr from pricerData displayed on main pricer screen 
            double[] volComponents = null;

            volComponents = volBuilder(ccyPair, dayCount);
            double atmVol = volComponents[0];
            double rr = volComponents[1];
            double fly = volComponents[2];
            double wingControl = volComponents[3];
            double targetFlyMult = volComponents[4];
            double smileFlyMult = volComponents[5];
            double rrMult = volComponents[6];

            wingControl = 1;
            //need to calc smilefly then get 25d and atm strikes and vols
            double smileFly = equivalentfly(spot, dayStart, autoExp, atmVol * wingControl, atmVol, rr, fly, dfDomDel, dfForDel, premoInc);
            double putVol = atmVol + smileFly - rr / 2;
            double callVol = atmVol + smileFly + rr / 2;

            double putStrike = FXStrikeVol(spot, dayStart, autoExp, 0.25, putVol, dfDomDel, dfForDel, "p", premoInc);
            double callStrike = FXStrikeVol(spot, dayStart, autoExp, 0.25, callVol, dfDomDel, dfForDel, "c", premoInc);
            double atmStrike = FXATMStrike(spot, dayStart, autoExp, atmVol, dfDomDel, dfForDel, premoInc);
            double callVol10 = atmVol + 0.5 * (rr * rrMult) + (smileFly * smileFlyMult);
            double callStrike10 = FXStrikeVol(spot, dayStart, autoExp, 0.1, callVol10, dfDomDel, dfForDel, "c", premoInc);

            double putVol10 = atmVol - 0.5 * (rr * rrMult) + (smileFly * smileFlyMult);
            double putStrike10 = FXStrikeVol(spot, dayStart, autoExp, 0.1, putVol10, dfDomDel, dfForDel, "p", premoInc);

            string premoString = "";//pricer.Rows[Premium_TypeR][curCol].ToString(); //checks datagridview for value

            //defaults premoInfo to base% for premo included and terms pips if not
            if (premoInc == 1)
            {
                premoString = "Base %";
            }

            else if (premoInc == 0)
            {
                premoString = "Terms Pips";
            }

            //get premoInc for  the delta solve. Keep in mind that the surface is built with premo included or not from the setup menu. This will ensure that strikes will have the correct vols despite what premo convention is used for an individual option. 

            double[] premoTypeInfo = premoConventions(premoString, spot, 1);

            int premoIncDeltSolve = Convert.ToInt16(premoTypeInfo[3]);

            double autoSt = 0;

            if (strText == "a") { autoSt = atmStrike; }

            if (IsNumeric(strText) == true)
            {
                autoSt = Convert.ToDouble(strText);
            }
            else
            {
                int TextLength = 0;
                TextLength = strText.Length;
                if (strText == "atmf")
                {
                    autoSt = outRight;
                }
                else if (strText == "atms")
                {
                    autoSt = spot;
                }
                else if (strText.Substring(TextLength - 1) == "d")
                {
                    double delt = Convert.ToDouble(strText.Substring(0, TextLength - 1)) / 100;

                    autoSt = FXStrikeDelta(delt, pC, outRight, atmVol, rr, fly, rrMult, smileFlyMult, spot, dayStart, autoExp, dfDomExp, dfForExp, dfDomDel, dfForDel, premoInc, smileFactor);

                }
            }



            double vol = combinedInterp(spot, dayStart, autoExp, wingControl * atmVol, autoSt, putStrike10, putVol10, putStrike, putVol, atmStrike, atmVol, callStrike, callVol, callStrike10, callVol10, dfDomExp, dfForExp, dfDomDel, dfForDel, outRight, smileFactor);



            double systemVol = vol;

    

            premoTypeInfo = premoConventions(premoString, spot, autoSt);
            double premoConversion = premoTypeInfo[0];//applies this factor greeks to convert in proper units
            double premoFactor = premoTypeInfo[1];// will convert greeks to correct units in nominal amounts
            double notionalFactor = premoTypeInfo[2]; //notional factor is equal to spot or 1 - will convert notional to terms currency if % terms or base pips is selected


            double[] greeks = FXOpts(spot, dayStart, autoExp, autoSt, vol, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);
            double premium = greeks[0];
            double fpremium = greeks[6];
            double delta = greeks[1];
            double fdelta = greeks[5];
            delta = delta - premium / spot * premoIncDeltSolve;
            fdelta = fdelta - fpremium / outRight * premoIncDeltSolve;
            premium = premium * premoConversion;

            List<string> returnType = new List<string>(new string[] { "Base %", "Terms %", "Base Pips", "Terms Pips" });
            double fwd_premoConversion = 0;

            if (premoString == returnType[0]) { fwd_premoConversion = 1 / outRight * 100; }
            if (premoString == returnType[1]) { fwd_premoConversion = 1 / autoSt * 100; }
            if (premoString == returnType[2]) { fwd_premoConversion = 1 / (outRight * autoSt); }
            if (premoString == returnType[3]) { fwd_premoConversion = 1; }


            fpremium = fpremium * fwd_premoConversion;
            double gamma = greeks[2];
            gamma = gamma * spot;
            double vega = greeks[3];
            vega = vega / 100 * premoConversion;

            //theta - just rolls the day 1 day foward - this wont be real theta  as doesnt roll the vol or depo curves. Can work on this later. 
            DateTime dayTheta = AddWorkdays(dayStart, 1);
            double[] theta = FXOpts(spot, dayTheta, autoExp, autoSt, vol, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);
            double sTheta = theta[0] * premoConversion - premium;

            double[] smileVolga = FXOpts(spot, dayStart, autoExp, autoSt, vol + .01, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);
            double[] smileVanna = FXOpts(spot * 1.01, dayStart, autoExp, autoSt, vol, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);

            double sVolga = 0;
            double sVanna = 0;

            sVolga = (smileVolga[3] / 100 * premoConversion) - vega; //dvega/dvol
            sVanna = (smileVanna[3] / 100 * premoConversion) - vega; //dvega/dspot



            double dvForDepo;
            double dvDomDepo;

            //dv01
            if (fwdPts == 0)
            {
                dvForDepo = 0;
                dvDomDepo = 0;

            }
            else
            {
                dvForDepo = forDepo + .0001;
                dvDomDepo = domDepo + .0001;
            }

            double dvForExp = DiscountFactor(dvForDepo, Convert.ToInt16(dayCount), Convert.ToInt16(basisB));
            double dvForDel = DiscountFactor(dvForDepo, Convert.ToInt16(delDayCount), Convert.ToInt16(basisB));

            double dvDomExp = DiscountFactor(dvDomDepo, Convert.ToInt16(dayCount), Convert.ToInt16(basisT));
            double dvDomDel = DiscountFactor(dvDomDepo, Convert.ToInt16(delDayCount), Convert.ToInt16(basisT));

            double[] dv01For = FXOpts(spot, dayStart, autoExp, autoSt, vol, dfDomExp, dvForExp, dfDomDel, dvForDel, pC);
            double[] dv01Dom = FXOpts(spot, dayStart, autoExp, autoSt, vol, dvDomExp, dfForExp, dvDomDel, dfForDel, pC);

            double sDv01For = dv01For[0] * premoConversion - premium;
            double sDv01Dom = dv01Dom[0] * premoConversion - premium;

            double smileVolSpread = vol - atmVol;
            double[] sVSf = FXOpts(spot, dayStart, autoExp, autoSt, atmVol, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);
            double premoSmileVSflat = premium - sVSf[0] * premoConversion;

            double[] sysVolPremo = FXOpts(spot, dayStart, autoExp, autoSt, systemVol, dfDomExp, dfForExp, dfDomDel, dfForDel, pC);
            double priceFromMid = premium - sysVolPremo[0] * premoConversion;

            double breakEven = atmVol / 24 * Math.Sqrt(dayCount) * spot;

            //greeks in nominal amounts 
          //  double notional = 100.0;

            double premiumA = notional * premium * premoFactor * notionalFactor;
            double DeltaA = notional * delta * 1000000;
            double GammaA = notional * gamma * 10000;
            double VegaA = notional * vega * premoFactor * notionalFactor;
            double VannaA = notional * sVanna * premoFactor * notionalFactor;
            double VolgaA = notional * sVolga * premoFactor * notionalFactor;
            double ThetaA = notional * sTheta * premoFactor * notionalFactor;
            double Dv01_BaseA = notional * sDv01For * premoFactor * notionalFactor;
            double Dv01_TermsA = notional * sDv01Dom * premoFactor * notionalFactor;
            double premoFromMid = notional * priceFromMid * premoFactor * -1 * notionalFactor;

            //call smile greeks {smileVega,smileRega25,smileRega10,smileSega25,smileSega10,smileDelta }; function returns amounts times notional 
            object[] smileGreeksOuput = smileGreeks(autoSt, atmVol, rr, fly, rrMult, smileFlyMult, spot, dayStart, autoExp, dfDomExp, dfForExp, dfDomDel, dfForDel, premoInc, smileFactor, pC, premoString, notional);

            double sVega = Convert.ToDouble(smileGreeksOuput[0]);
            double rega25 = Convert.ToDouble(smileGreeksOuput[1]);
            double rega10 = Convert.ToDouble(smileGreeksOuput[2]);
            double sega25 = Convert.ToDouble(smileGreeksOuput[3]);
            double sega10 = Convert.ToDouble(smileGreeksOuput[4]);
            double sDelta = Convert.ToDouble(smileGreeksOuput[5]);
            double sGamma = Convert.ToDouble(smileGreeksOuput[7]);


            //set spreads
            //spreads are loaded into datatables on market data update . They are saved into dataset spreadset
            //get row of currency in datatable

            List<string> crossList = spreadSet.Tables[0].AsEnumerable().Select(x => x[0].ToString()).ToList();
            int rowInt = crossList.IndexOf(ccyPair);

            //load pillar daycounts from datatable columns then lookup daycount of current option
            var dayArr = spreadSet.Tables[0].Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray();
            int dayCol = 0;
            double volSpread = 0;

            for (int q = 2; q < dayArr.Count(); q++)
            {
                if (dayCount >= Convert.ToDouble(dayArr[q]))
                {
                    dayCol = q;
                }
            }

            //finaly  check delta to get spread from correct datatable. There are 3 (atm, 25delta, 10delta spreads)
            double deltaAdj = Math.Abs(delta);

            if (autoSt >= outRight)
            {
                double[] x = FXOpts(spot, dayStart, autoExp, autoSt, vol, dfDomExp, dfForExp, dfDomDel, dfForDel, "c");
                deltaAdj = x[1];

            }
            else
            {
                double[] x = FXOpts(spot, dayStart, autoExp, autoSt, vol, dfDomExp, dfForExp, dfDomDel, dfForDel, "p");
                deltaAdj = x[1];
            }


            deltaAdj = Math.Abs(deltaAdj);


            if (deltaAdj <= 0.60)
            {
                volSpread = Convert.ToDouble(spreadSet.Tables[0].Rows[rowInt][dayCol]);
            }

            if (deltaAdj <= 0.26)
            {
                volSpread = Convert.ToDouble(spreadSet.Tables[1].Rows[rowInt][dayCol]);
            }

            if (deltaAdj <= 0.11)
            {
                volSpread = Convert.ToDouble(spreadSet.Tables[2].Rows[rowInt][dayCol]);
            }

            volSpread = volSpread / 2;
            double bidVol = Math.Round((vol * 100 - volSpread) / .05, 0) * .05;
            double askVol = Math.Round((vol * 100 + volSpread) / .05, 0) * .05;
            string bidOffer = bidVol.ToString("0.00") + " / " + askVol.ToString("0.00");


            retVal = new object[] { autoExp.ToShortDateString(), autoSt, notional.ToString("#,##0.00"), delta.ToString("#,##0.00"), sVega.ToString("#,##0"), sGamma.ToString("#,##0"), sega10.ToString("#,##0"), sega25.ToString("#,##0"), rega10.ToString("#,##0"), rega25.ToString("#,##0") };
            return retVal;


           
        }

        private void showScenariosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            uploadMxData();
            DataTable dt = new DataTable();
            dt = skewSheet.Tables["scenarios"];
            showLoadedData(dt);
        }

        private void chooseCcyPairToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void runAllSurfacesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Run All Surfaces?", "", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {

                foreach (DataRow row in crosses.Rows)
                {
                    string cross = row["Cross"].ToString();

                    setSurfaceNew(cross);

                }

                setSurfaceNew("USDRUB");

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            MessageBox.Show("All Surfaces Loaded");

        }

        private void showLoadedDataSkew(DataTable dt)
        {
            WindowsFormsApplication1.SkewPopUP bulkData = new WindowsFormsApplication1.SkewPopUP(dt);
            bulkData.ShowDialog(this);
        }

        private void showTradeSim(DataTable dt)
        {
            WindowsFormsApplication1.tradeSim bulkData = new WindowsFormsApplication1.tradeSim(dt);
            bulkData.ShowDialog(this);
        }

        private void toolStripComboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (endIni == 1)
            {

                string tName = toolStripComboBox4.Text;

                skewAnalysis(tName);

                DataTable dt = new DataTable();
                dt = skewBps.Tables[tName];
                showLoadedDataSkew(dt);
            }
        }

        private void simulationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            riskSim();
        }

        private void uSDRUBToolStripMenuItem_Click(object sender, EventArgs e)
        {
            setBbgInterfaceSurface("USDRUB");
            //Task.Delay(3);
            mxVolsDisplay("USDRUB");
        }

        private void replicationPortToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void toolStripComboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (endIni == 1)
            {


                replicaiton();

            }

           
        }

        //private void skewSheetDataFromExcel()
        //{
        //    if (skewSheet == null)
        //        skewSheet = new DataSet();

        //    skewSheet.Reset();


        //    String filePath = systemFiles + @"skewSheet.xls";
        //    string strExcelConn;
        //    bool hasHeaders = false;
        //    string HDR = hasHeaders ? "Yes" : "No";

        //    if (filePath.Substring(filePath.LastIndexOf('.')).ToLower() == ".xls")
        //        strExcelConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";IMEX=0\"";
        //    else
        //        strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties=\"Excel 8.0;HDR=" + HDR + ";IMEX=0\"";


        //    string[] arr = new string[] { "OptList", "Scenarios" };

        //    for (int i = 0; i < arr.Count(); i++)
        //    {
        //        using (OleDbConnection connExcel = new OleDbConnection(strExcelConn))
        //        {
        //            string table = arr[i];


        //            string selectString = "SELECT * FROM [" + table + "]";
        //            using (OleDbCommand cmdExcel = new OleDbCommand(selectString, connExcel))
        //            {
        //                cmdExcel.Connection = connExcel;
        //                connExcel.Open();
        //                DataTable dt = new DataTable();
        //                OleDbDataAdapter adp = new OleDbDataAdapter();
        //                adp.SelectCommand = cmdExcel;
        //                adp.FillSchema(dt, SchemaType.Source);
        //                adp.Fill(dt);
        //                int range = dt.Columns.Count;
        //                int row = dt.Rows.Count;
        //                dt.TableName = table;
        //                skewSheet.Tables.Add(dt);
        //            }
        //        }
        //    }

        //    MessageBox.Show("Data Loaded");
        //}



        #endregion



    }
}

