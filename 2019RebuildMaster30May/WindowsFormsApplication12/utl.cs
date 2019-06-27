using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Threading;
using System.Text.RegularExpressions;
using DateTime = System.DateTime;
using DayOfWeek = System.DayOfWeek;
using System.Globalization;
using System.Data.OleDb;

namespace SharedClasses
{
    class Utility
    {
        public List<string> pathString(string fileName)
        {
            List<string> retVal;

            //string filename = @"C:\DVIUSER.txt";
            string[] ReadFile = File.ReadAllLines(fileName).ToArray();
            string directory = ReadFile[0];
            string user = ReadFile[1];

            string systemFiles = directory + user + @"\systemFiles\";

            string xmlFilePath = directory + user + @"\xmlFiles\";

            retVal = new List<string> { systemFiles, xmlFilePath };
            return retVal;


        }

        public DataSet masterSet(string xmlFp, string xmlFn)
        {

            DataSet ds = new DataSet();

            string myXMLfile = xmlFp + xmlFn + ".xml";
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

            return ds;

        }

        public void LoadXmldt(DataTable dt, string xmlFp, string xmlFn)
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


            if (ds.Tables.Contains(dtName))
            {

                if (dt.Columns.Count == 0)
                {
                    dt = ds.Tables[dtName].Clone();
                }

                for (int j = 0; j <= ds.Tables[dtName].Rows.Count - 1; j++)
                {
                    //Adds a new row to the DataGridView for each line of text.
                    dt.Rows.Add();

                    //This for loop loops through the array in order to retrieve each
                    //line of text.
                    for (int i = 0; i <= dt.Columns.Count - 1; i++)
                    {
                        try
                        {
                            //Sets the value of the cell to the value of the text retreived from the text file.
                            dt.Rows[dt.Rows.Count - 1][i] = ds.Tables[dtName].Rows[j].ItemArray[i];
                        }
                        catch { }

                    }

                }
            }
            else
            {
                //MessageBox.Show("Please Setup Run");
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

        public int getRowInt(DataTable dt, string rowHeader, string colName)
        {
            var c = dt.Columns[colName];
            int col = c.Ordinal;


            List<string> rName = dt.AsEnumerable().Select(x => x[col].ToString()).ToList();
            int r = rName.IndexOf(rowHeader);
            return r;
        }

        public void addTableToDs(DataSet ds, DataTable dt)
        {
            string dtName = dt.TableName;
            //  d_data.TableName = dtName;

            DataTable dtCopy = new DataTable();
            dtCopy = dt.Copy();

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
        }

        public void saveDatasetXml(string xmlFilePath,string xmlFile, DataSet ds)
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

        public List<double> colToList(DataTable dt, string colName)
        {
            List<double> outPut = dt.AsEnumerable().Select(x => Convert.ToDouble(x[dt.Columns[colName].Ordinal])).ToList();

            return outPut;

        }

        public double searchCol(DataTable dt, string colName, string sColName, string search)
        {
            List<string> sList = dt.AsEnumerable().Select(x => Convert.ToString(x[dt.Columns[sColName].Ordinal])).ToList();
            List<double> outPut = dt.AsEnumerable().Select(x => Convert.ToDouble(x[dt.Columns[colName].Ordinal])).ToList();


            int i = sList.IndexOf(search);
            double retVal = outPut[i];

            return retVal;

        }

        public string searchColString(DataTable dt, string colName, string sColName, string search)
        {
            List<string> sList = dt.AsEnumerable().Select(x => Convert.ToString(x[dt.Columns[sColName].Ordinal])).ToList();
            List<string> outPut = dt.AsEnumerable().Select(x => Convert.ToString(x[dt.Columns[colName].Ordinal])).ToList();



            int i = sList.IndexOf(search);
            string retVal = outPut[i];

            return retVal;

        }

        public DataTable LoadXmlFast(string xmlFp, string xmlFn)
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
            }



            return ds.Tables[0];
        }


    }

    class optionFunctions
    {
        rateFunctions rf = new rateFunctions();

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

        public static bool R_Q_P01_boundaries(double p, double _LEFT_, double _RIGHT_, bool lower_tail, bool log_p, out double ans)
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

        public static double R_DT_qIv(double p, bool lower_tail, bool log_p)
        {
            return (log_p ? (lower_tail ? Math.Exp(p) : -ExpM1(p)) : R_D_Lval(p, lower_tail));
        }

        public static double R_DT_CIv(double p, bool lower_tail, bool log_p)
        {
            return (log_p ? (lower_tail ? -ExpM1(p) : Math.Exp(p)) : R_D_Cval(p, lower_tail));
        }

        public static double R_D_Lval(double p, bool lower_tail)
        {
            return lower_tail ? p : 0.5 - p + 0.5;
        }

        public static double R_D_Cval(double p, bool lower_tail)
        {
            return lower_tail ? 0.5 - p + 0.5 : p;
        }

        public static double ExpM1(double x)
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

        public double StanNormCumDistr(double x)
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

        public Boolean IsNumeric(System.Object Expression)
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

        public static double erf(double x)
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

                double Delta = verso * Nd1 * Dfd2;
                double Fwd_Delta = verso * Nd1;
                double gamma = Dfd2 * NormsDens(d1) / (s * sig * (Math.Pow(TimeExp, (0.5))));
                double Vega = Dfd2 * s * (Math.Pow(TimeExp, (0.5))) * NormsDens(d1);
                double theta = (-0.5 * Math.Pow(sig, 2) * Math.Pow(s, 2) * gamma + r * price - (r - Q) * s * Delta);

                if (Vol <= 0)
                {
                    functionReturnValue = new double[] { 0, 0, 0, 0, 0 };
                    return functionReturnValue;
                }

                else
                {
                    functionReturnValue = new double[] { price, Delta, gamma, Vega, theta, Fwd_Delta, fwd_price };
                    return functionReturnValue;
                }


            }
            catch
            {
                functionReturnValue = new double[] { 0, 0, 0, 0, 0 };
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

        public double equivalentfly(double s, DateTime today, DateTime Expiry, double sigatm, double rr, double Bfly, double Dfe, double Dfe2, int premimuincluded)
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

                sig25bc = smileInterp(s, today, Expiry, K25bc, K25p, sig25p, KA, sigatm, K25c,
                sig25c, Dfe, Dfe2, Dfd, Dfd2);
                sig25bp = smileInterp(s, today, Expiry, K25bp, K25p, sig25p, KA, sigatm, K25c,
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

                    sig25bc = smileInterp(s, today, Expiry, K25bc, K25p, sig25p, KA, sigatm, K25c,
                    sig25c, Dfe, Dfe2, Dfd, Dfd2);
                    sig25bp = smileInterp(s, today, Expiry, K25bp, K25p, sig25p, KA, sigatm, K25c,
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

        public double marketfly(double Delta, double atmVol, double rr, double fly, double rrMult, double smileFlyMult, double spot, DateTime dayStart, DateTime autoExp, double Dfe, double Dfe2, double Dfd, double Dfd2, int premoInc, double splineFactor)
        {

            List<double> vK = vKs(atmVol, rr, fly, rrMult, smileFlyMult, spot, dayStart, autoExp, Dfe, Dfe2, Dfd, Dfd2, premoInc);

            double K10P, K25P, KATM, K25C, K10C, V10P, V25P, VATM, V25C, V10C;

            V10P = vK[0]; V25P = vK[1]; VATM = vK[2]; V25C = vK[3]; V10C = vK[4]; K10P = vK[5]; K25P = vK[6]; KATM = vK[7]; K25C = vK[8]; K10C = vK[9];


            double functionReturnValue = 0;

            double TimeExp = 0;
            double KA = 0;
            double K25c = 0;
            double K25p = 0;

            double sigdp = 0;
            double Kdbc = 0;
            double Kdbp = 0;
         
            double sig25c = 0;
            double sig25p = 0;
            double sig25b = 0;
            double sigdc = 0;
            double sigdb = 0;
            double sigdbc = 0;
            double sigdbp = 0;


            double Bfly = equivalentfly(spot, dayStart, autoExp, atmVol,rr, fly, Dfe, Dfe2, premoInc);



            try
            {
                Dfd = Dfe;
                Dfd2 = Dfe2;
                TimeExp = (autoExp - dayStart).TotalDays / 365;
                double Fw = spot * Dfe2 / Dfe;

                KA = FXATMStrike(spot, dayStart, autoExp, atmVol, Dfe, Dfe2, premoInc);

                sig25p = V25P;
                sig25c= V25C;
                sig25b = VATM + fly;
                K25c = K25C;
                K25p = K25P;

                double sig10C = V10C;
                double K10c = K10C;

                double sig10P = V10P; 
                double K10p = K10C; 

                sigdc = sig25c;
                sigdp = sig25p;

                sigdb = (sigdc + sigdp) / 2;


                Kdbc = FXStrikeVol(spot, dayStart, autoExp, Delta, sigdb, Dfe, Dfe2, "c", premoInc);

                Kdbp = FXStrikeVol(spot, dayStart, autoExp, Delta, sigdb, Dfe, Dfe2, "p", premoInc);


                sigdbc = fxVol(Kdbc, atmVol, rr, fly, rrMult, smileFlyMult, spot, dayStart, autoExp, Dfe, Dfe2, Dfd, Dfd2, premoInc, splineFactor);
                sigdbp = fxVol(Kdbp, atmVol, rr, fly, rrMult, smileFlyMult, spot, dayStart, autoExp, Dfe, Dfe2, Dfd, Dfd2, premoInc, splineFactor);

               // sigdbc = combinedInterp(s, today, Expiry, sigWing, Kdbc, K10p, sig10P, K25p, sig25p, KA, sigatm, K25c, sig25c, K10c, sig10C, Dfe, Dfe2, Dfd, Dfd2, Fw, factor);
              //  sigdbp = combinedInterp(s, today, Expiry, sigWing, Kdbp, K10p, sig10P, K25p, sig25p, KA, sigatm, K25c, sig25c, K10c, sig10C, Dfe, Dfe2, Dfd, Dfd2, Fw, factor);

                double[] calldc = FXOpts(spot, dayStart, autoExp, Kdbc, sigdc, Dfe, Dfe2, Dfd, Dfd2, "c");
                double[] calldb = FXOpts(spot, dayStart, autoExp, Kdbc, sigdb, Dfe, Dfe2, Dfd, Dfd2, "c");
                double[] putdp = FXOpts(spot, dayStart, autoExp, Kdbp, sigdp, Dfe, Dfe2, Dfd, Dfd2, "p");
                double[] putdb = FXOpts(spot, dayStart, autoExp, Kdbp, sigdb, Dfe, Dfe2, Dfd, Dfd2, "p");

                double f0 = (calldc[0] + putdp[0]) - (calldb[0] + putdb[0]);

                double dfly = sigdb - VATM;

                if (Math.Abs(f0) < 0.0000001 * spot)
                {
                    functionReturnValue = dfly;
                    return functionReturnValue;
                }

                sigdb = sigdb + 0.00005;

                int j = 0;

                //loop
                double dsig = 0.00005;

                while (Math.Abs(f0) > 0.00000001 * spot)
                {
                    j = j + 1;

                    Kdbc = FXStrikeVol(spot, dayStart, autoExp, Delta, sigdb, Dfe, Dfe2, "c", premoInc);

                    Kdbp = FXStrikeVol(spot, dayStart, autoExp, Delta, sigdb, Dfe, Dfe2, "p", premoInc);

                    sigdbc = fxVol(Kdbc, atmVol, rr, fly, rrMult, smileFlyMult, spot, dayStart, autoExp, Dfe, Dfe2, Dfd, Dfd2, premoInc, splineFactor);
                    sigdbp = fxVol(Kdbp, atmVol, rr, fly, rrMult, smileFlyMult, spot, dayStart, autoExp, Dfe, Dfe2, Dfd, Dfd2, premoInc, splineFactor);

                    calldc = FXOpts(spot, dayStart, autoExp, Kdbc, sigdbc, Dfe, Dfe2, Dfd, Dfd2, "c");
                    calldb = FXOpts(spot, dayStart, autoExp, Kdbc, sigdb, Dfe, Dfe2, Dfd, Dfd2, "c");
                    putdp = FXOpts(spot, dayStart, autoExp, Kdbp, sigdbp, Dfe, Dfe2, Dfd, Dfd2, "p");
                    putdb = FXOpts(spot, dayStart, autoExp, Kdbp, sigdb, Dfe, Dfe2, Dfd, Dfd2, "p");

                    double F = (calldc[0] + putdp[0]) - (calldb[0] + putdb[0]);
                    double dF = (F - f0) / dsig;

                    sigdb = sigdb - F / dF;
                    dsig = -F / dF;
                    f0 = F;

                    if (j * 0.00001 >= dfly)
                    {
                        functionReturnValue = sigdb - VATM;
                        return functionReturnValue;
                    }


                }

                functionReturnValue = sigdb - VATM;
                return functionReturnValue;


            }
            catch
            {

                functionReturnValue = sigdb - VATM;
                return functionReturnValue;
            }



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

     
        public List<double> vKs(double atmVol, double rr, double fly, double rrMult, double smileFlyMult, double spot, DateTime dayStart, DateTime autoExp, double dfDomExp, double dfForExp, double dfDomDel, double dfForDel, int premoInc)
        {


            double fwd = spot * dfForDel / dfDomDel;
            //need to calc smilefly then get 25d and atm strikes and vols
            double smileFly = equivalentfly(spot, dayStart, autoExp, atmVol, rr, fly, dfDomDel, dfForDel, premoInc);


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

        public double smileInterp(double s, DateTime today, DateTime Expiry, double k, double K1, double v1, double K2, double v2, double K3, double v3, double Dfe, double Dfe2, double Dfd, double Dfd2)
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
                double d1 = (Math.Log(Fw / k) + 0.5 * Math.Pow(v2, 2) * TimeExp) / (v2 * Math.Sqrt(TimeExp));
                double d2 = d1 - v2 * Math.Sqrt(TimeExp);


                double d11 = (Math.Log(Fw / K1) + 0.5 * Math.Pow(v2, 2) * TimeExp) / (v2 * Math.Sqrt(TimeExp));
                double d21 = d11 - v2 * Math.Sqrt(TimeExp);
                double d12 = (Math.Log(Fw / K2) + 0.5 * Math.Pow(v2, 2) * TimeExp) / (v2 * Math.Sqrt(TimeExp));
                double d22 = d12 - v2 * Math.Sqrt(TimeExp);
                double d13 = (Math.Log(Fw / K3) + 0.5 * Math.Pow(v2, 2) * TimeExp) / (v2 * Math.Sqrt(TimeExp));
                double d23 = d13 - v2 * Math.Sqrt(TimeExp);

                double dk1 = x1 * v1 + x2 * v2 + x3 * v3 - v2;

                double dk2 = x1 * d11 * d21 * Math.Pow((v1 - v2), 2) + x3 * d13 * d23 * Math.Pow((v3 - v2), 2) + x2 * d12 * d22 * Math.Pow((v2 - v2), 2);


                if ((Math.Pow(v2, 2) + d1 * d2 * (2 * v2 * dk1 + dk2)) < 0)
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

                sig = v2 + (-v2 + Math.Sqrt(Math.Pow(v2, 2) + d1 * d2 * (2 * v2 * dk1 + dk2))) / (d1 * d2);

                functionReturnValue = sig;
                return functionReturnValue;

            }
            catch
            {
                functionReturnValue = 0;
                return functionReturnValue;
            }

        }

        public double smileSpline(double p10, double p25, double atm, double c25, double c10, double pv10, double pv25, double vatm, double cv25, double cv10, double F, double target, double splineFactor)
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

            double step = 1 / splineFactor;

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

                retVal = rf.LinearInterp(logStrike, strikeVolList, tVal);
            }

            return retVal;


        }

        public double fxVol(double k, double atmVol, double rr, double fly, double rrMult, double smileFlyMult, double spot, DateTime dayStart, DateTime autoExp, double Dfe, double Dfe2, double Dfd, double Dfd2, int premoInc, double splineFactor)
        {
            double fwd = spot * Dfd2 / Dfd;

            List<double> vK = vKs(atmVol, rr, fly, rrMult, smileFlyMult, spot, dayStart, autoExp, Dfe, Dfe2, Dfd, Dfd2, premoInc);

            double V10P, V25P, VATM, V25C, V10C,K10P, K25P, KATM, K25C, K10C;

            V10P = vK[0]; V25P = vK[1]; VATM = vK[2]; V25C = vK[3]; V10C = vK[4]; K10P = vK[5]; K25P = vK[6]; KATM = vK[7]; K25C = vK[8]; K10C = vK[9];


            double vol = 0;

            if (k >= K25P && k <= K25C)
            {
                vol = smileInterp(spot, dayStart, autoExp, k, K25P, V25P, KATM, VATM, K25C, V25C, Dfe, Dfe2, Dfd, Dfd2);
            }

            else
            {
                vol = smileSpline(K10P, K25P, KATM, K25C, K10C, V10P, V25P, VATM, V25C, V10C, fwd, k, splineFactor);
            }

            return vol;
        }

        public double CrossVol(double Vol1, double Vol2, double Correl)
        {
            return Math.Pow((Vol1 * Vol1 + Vol2 * Vol2 - 2 * Correl * Vol1 * Vol2), 0.5);
        }

        public double ImpliedCorrel(double Vol1, double Vol2, double CrossVol)
        {
            return (Vol1 * Vol1 + Vol2 * Vol2 - CrossVol * CrossVol) * (1 / (2 * Vol1 * Vol2));
        }

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
    }

    class dateFunctions
    {

       

        public double RoundDown(double value, int digits)
        {
            if (value >= 0)
                return Math.Floor(value * Math.Pow(10, digits)) / Math.Pow(10, digits);

            return Math.Ceiling(value * Math.Pow(10, digits)) / Math.Pow(10, digits);
        }

        public  DateTime AddWorkdays(DateTime originalDate, int workDays)
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

        
        //   Converts entered text to an option expiry date.                                                                                 '
        
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
                if (ExpiryText.Substring(TextLength - 2) == "ay")
                {

                    functionReturnValue = Convert.ToDateTime(ExpiryText);
                    double past = (functionReturnValue - DateTime.Today).TotalDays;
                    if (past < 0)
                    {
                        functionReturnValue = functionReturnValue.AddYears(1);
                    }
                }
                else
                {
                    functionReturnValue = ExpiryMonthDate(StartDate, 12 * Convert.ToInt32(ExpiryText.Substring(0, TextLength - 1)), HomeCcy, BaseCcy, TermsCcy);
                }

            }
            else if (IsDate(ExpiryText) == true)
            {
                functionReturnValue = Convert.ToDateTime(ExpiryText);
                double past = (functionReturnValue - DateTime.Today).TotalDays;
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

        public string AutoExpiryString(string ExpiryText)
        {


            string functionReturnValue = ExpiryText;
            int TextLength = ExpiryText.Length;
            string dateEnd = ExpiryText.Substring(TextLength - 1);
            int textlen = ExpiryText.Length;

            if (IsDate(ExpiryText) == true)
            {
                DateTime testDate = Convert.ToDateTime(ExpiryText);
                double past = (testDate - DateTime.Today).TotalDays;

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

        public bool IsDate(string inputDate)
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


        public System.DateTime TomDate(System.DateTime StartDate, string BaseCcy, string TermsCcy)
        {

            System.DateTime functionReturnValue = StartDate.AddDays(1);

            while (TestHoliday(functionReturnValue, BaseCcy) == false || TestHoliday(functionReturnValue, TermsCcy) == false)
            {
                functionReturnValue = functionReturnValue.AddDays(1);
            }

            return functionReturnValue;
        }

        
      
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

                if (EndMonth == 2 && EndDay > 28)
                {
                    EndDay = 28;
                }

                functionReturnValue = new DateTime(Convert.ToInt32(EndYear), Convert.ToInt32(EndMonth), Convert.ToInt32(EndDay));

                while (TestTwoHolidays(functionReturnValue, BaseCcy, TermsCcy) == false)
                {
                    functionReturnValue = functionReturnValue.AddDays(1);
                }
            }

            return functionReturnValue;


        }

        
        //   Returns a specified currency pair's date as True or False ("good" or "bad")                                                     '
        
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

        


        Utility u = new Utility();
        PriceEngine2019.shared sh = new PriceEngine2019.shared();
       
        private DataTable ccyDetsdt()
        {


            string s = sh.sharedDataSet();     
            DataSet msDs = u.masterSet(s, "masterSetupNew");
            DataTable dt = msDs.Tables["ccyDets"].Copy();
            return dt;

        }

        public DataSet hSet()
        {
             string s = sh.sharedDataSet();   //directory
             DataSet ds = u.masterSet(s, "holidayDataNew");
             return ds;
        }

        public bool TestHoliday(System.DateTime DateToCheck, string Ccy)
        {

            DataSet holidaySet = hSet();
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

        public int SpotDateFactor(string Ccy)
        {


            DataTable ccyDets = ccyDetsdt();

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

        //   Returns the correct business day for the specified number of weeks from the start date.
        
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

  


       

    }

    class rateFunctions
    {
        public double DiscountFactor(double SimpleYield, int DelDays, int Basis)
        {
            return 1 / (1 + SimpleYield * DelDays / Basis);
        }

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
            int TermsYears = 0;
            int BaseYears = 0;
            BaseYears = Convert.ToInt32(RoundDown(DelDays / BaseBasis, 0));
            TermsYears = Convert.ToInt32(RoundDown(DelDays / TermsBasis, 0));
            return Spot * (1 + TermsDepo * (DelDays / TermsBasis - TermsYears)) * Math.Pow((1 + TermsDepo), TermsYears) / ((1 + BaseDepo * (DelDays / BaseBasis - BaseYears)) * Math.Pow((1 + BaseDepo), BaseYears));
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


     

    }

}
