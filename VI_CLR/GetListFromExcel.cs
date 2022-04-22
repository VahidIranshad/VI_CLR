//using ExcelDataReader;
using Microsoft.SqlServer.Server;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace VI_CLR
{
    public partial class GetListFromExcel
    {
        public SqlString A;
        public SqlString B;
        public SqlString C;
        public SqlString D;
        public SqlString E;
        public SqlString F;
        public SqlString G;
        public SqlString H;
        public SqlString I;
        public SqlString J;
        public SqlString K;
        public SqlString L;
        public SqlString M;
        public SqlString N;
        public SqlString O;
        public SqlString P;
        public SqlString Q;
        public SqlString R;
        public SqlString S;
        public SqlString T;
        public SqlString U;
        public SqlString V;
        public SqlString W;
        public SqlString X;
        public SqlString Y;
        public SqlString Z;
        public SqlString AA;
        public SqlString AB;
        public SqlString AC;
        public SqlString AD;
        public SqlString AE;
        public SqlString AF;
        public SqlString AG;
        public SqlString AH;
        public SqlString AI;
        public SqlString AJ;
        public SqlString AK;
        public SqlString AL;
        public SqlString AM;
        public SqlString AN;
        public SqlString AO;
        public SqlString AP;
        public SqlString AQ;
        public SqlString AR;
        public SqlString _AS;
        public SqlString AT;
        public SqlString AU;
        public SqlString AV;
        public SqlString AW;
        public SqlString AX;
        public SqlString AY;
        public SqlString AZ;
        public GetListFromExcel(
            SqlString a,
            SqlString b,
            SqlString c,
            SqlString d,
            SqlString e,
            SqlString f,
            SqlString g,
            SqlString h,
            SqlString i,
            SqlString j,
            SqlString k,
            SqlString l,
            SqlString m,
            SqlString n,
            SqlString o,
            SqlString p,
            SqlString q,
            SqlString r,
            SqlString s,
            SqlString t,
            SqlString u,
            SqlString v,
            SqlString w,
            SqlString x,
            SqlString y,
            SqlString z,
            SqlString aa,
            SqlString ab,
            SqlString ac,
            SqlString ad,
            SqlString ae,
            SqlString af,
            SqlString ag,
            SqlString ah,
            SqlString ai,
            SqlString aj,
            SqlString ak,
            SqlString al,
            SqlString am,
            SqlString an,
            SqlString ao,
            SqlString ap,
            SqlString aq,
            SqlString ar,
            SqlString _as,
            SqlString at,
            SqlString au,
            SqlString av,
            SqlString aw,
            SqlString ax,
            SqlString ay,
            SqlString az
            )
        {
            A = a;
            B = b;
            C = c;
            D = d;
            E = e;
            F = f;
            G = g;
            H = h;
            I = i;
            J = j;
            K = k;
            L = l;
            M = m;
            N = n;
            O = o;
            P = p;
            Q = q;
            R = r;
            S = s;
            T = t;
            U = u;
            V = v;
            W = w;
            X = x;
            Y = y;
            Z = z; ;
            AA = aa;
            AB = ab;
            AC = ac;
            AD = ad;
            AE = ae;
            AF = af;
            AG = ag;
            AH = ah;
            AI = ai;
            AJ = aj;
            AK = ak;
            AL = al;
            AM = am;
            AN = an;
            AO = ao;
            AP = ap;
            AQ = aq;
            AR = ar;
            _AS = _as;
            AT = at;
            AU = au;
            AV = av;
            AW = aw;
            AX = ax;
            AY = ay;
            AZ = az;
        }
        //The SqlFunction attribute tells Visual Studio to register this 
        //code as a user defined function
        [Microsoft.SqlServer.Server.SqlFunction(
            FillRowMethodName = "FindFiles",
            TableDefinition = @"
a nvarchar(max), 
b nvarchar(max), 
c nvarchar(max), 
d nvarchar(max), 
e nvarchar(max), 
f nvarchar(max), 
g nvarchar(max), 
h nvarchar(max), 
i nvarchar(max), 
j nvarchar(max), 
k nvarchar(max), 
l nvarchar(max), 
m nvarchar(max), 
n nvarchar(max), 
o nvarchar(max), 
p nvarchar(max),
q nvarchar(max),
r nvarchar(max),
s nvarchar(max),
t nvarchar(max),
u nvarchar(max),
v nvarchar(max),
w nvarchar(max),
x nvarchar(max),
y nvarchar(max),
z nvarchar(max),
aa nvarchar(max), 
ab nvarchar(max), 
ac nvarchar(max), 
ad nvarchar(max), 
ae nvarchar(max), 
af nvarchar(max), 
ag nvarchar(max), 
ah nvarchar(max), 
ai nvarchar(max), 
aj nvarchar(max), 
ak nvarchar(max), 
al nvarchar(max), 
am nvarchar(max), 
an nvarchar(max), 
ao nvarchar(max), 
ap nvarchar(max),
aq nvarchar(max),
ar nvarchar(max),
_as nvarchar(max),
at nvarchar(max),
au nvarchar(max),
av nvarchar(max),
aw nvarchar(max),
ax nvarchar(max),
ay nvarchar(max),
az nvarchar(max)
",
            DataAccess = DataAccessKind.Read)]
        public static IEnumerable GetExelDataByPath(string targetDirectory, string sheetName)
        {
            try
            {
                ArrayList FilePropertiesCollection = new ArrayList();
                //DirectoryInfo dirInfo = new DirectoryInfo(targetDirectory);
                //FileInfo[] files = dirInfo.GetFiles(searchPattern);
                //foreach (FileInfo fileInfo in files)
                //{
                //    //I'm adding to the colection the properties (FileProperties) 
                //    //of each file I've found  
                //    //FilePropertiesCollection.Add(new GetListFromExcel(fileInfo.Name,
                //    //fileInfo.Length, fileInfo.CreationTime));
                //}


                //var connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", fileName);
                var connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;OLE DB Services=-4;Data Source={0}; Extended Properties=Excel 12.0;", targetDirectory);

                var adapter = new OleDbDataAdapter(string.Format("SELECT * FROM [{0}$]", sheetName), connectionString);
                var ds = new DataSet();

                adapter.Fill(ds, "anyNameHere");
                var isFirstRow = true;

                foreach (DataRow item in ds.Tables[0].Rows)
                {
                    //if (isFirstRow)
                    //{
                    //    isFirstRow = false;
                    //    continue;
                    //}
                    int columnCount = item.Table.Columns.Count;
                    var a = getCellValue(columnCount, 0, item);
                    var b = getCellValue(columnCount, 1, item);
                    var c = getCellValue(columnCount, 2, item);
                    var d = getCellValue(columnCount, 3, item);
                    var e = getCellValue(columnCount, 4, item);
                    var f = getCellValue(columnCount, 5, item);
                    var g = getCellValue(columnCount, 6, item);
                    var h = getCellValue(columnCount, 7, item);
                    var i = getCellValue(columnCount, 8, item);
                    var j = getCellValue(columnCount, 9, item);
                    var k = getCellValue(columnCount, 10, item);
                    var l = getCellValue(columnCount, 11, item);
                    var m = getCellValue(columnCount, 12, item);
                    var n = getCellValue(columnCount, 13, item);
                    var o = getCellValue(columnCount, 14, item);
                    var p = getCellValue(columnCount, 15, item);
                    var q = getCellValue(columnCount, 16, item);
                    var r = getCellValue(columnCount, 17, item);
                    var s = getCellValue(columnCount, 18, item);
                    var t = getCellValue(columnCount, 19, item);
                    var u = getCellValue(columnCount, 20, item);
                    var v = getCellValue(columnCount, 21, item);
                    var w = getCellValue(columnCount, 22, item);
                    var x = getCellValue(columnCount, 23, item);
                    var y = getCellValue(columnCount, 24, item);
                    var z = getCellValue(columnCount, 25, item);
                    var aa = getCellValue(columnCount, 26, item);
                    var ab = getCellValue(columnCount, 27, item);
                    var ac = getCellValue(columnCount, 28, item);
                    var ad = getCellValue(columnCount, 29, item);
                    var ae = getCellValue(columnCount, 30, item);
                    var af = getCellValue(columnCount, 31, item);
                    var ag = getCellValue(columnCount, 32, item);
                    var ah = getCellValue(columnCount, 33, item);
                    var ai = getCellValue(columnCount, 34, item);
                    var aj = getCellValue(columnCount, 35, item);
                    var ak = getCellValue(columnCount, 36, item);
                    var al = getCellValue(columnCount, 37, item);
                    var am = getCellValue(columnCount, 38, item);
                    var an = getCellValue(columnCount, 39, item);
                    var ao = getCellValue(columnCount, 40, item);
                    var ap = getCellValue(columnCount, 41, item);
                    var aq = getCellValue(columnCount, 42, item);
                    var ar = getCellValue(columnCount, 43, item);
                    var _as = getCellValue(columnCount, 44, item);
                    var at = getCellValue(columnCount, 45, item);
                    var au = getCellValue(columnCount, 46, item);
                    var av = getCellValue(columnCount, 47, item);
                    var aw = getCellValue(columnCount, 48, item);
                    var ax = getCellValue(columnCount, 49, item);
                    var ay = getCellValue(columnCount, 50, item);
                    var az = getCellValue(columnCount, 51, item);
                    FilePropertiesCollection.Add(new GetListFromExcel(
                        a,
                        b,
                        c,
                        d,
                        e,
                        f,
                        g,
                        h,
                        i,
                        j,
                        k,
                        l,
                        m,
                        n,
                        o,
                        p,
                        q,
                        r,
                        s,
                        t,
                        u,
                        v,
                        w,
                        x,
                        y,
                        z,
                        aa,
                        ab,
                        ac,
                        ad,
                        ae,
                        af,
                        ag,
                        ah,
                        ai,
                        aj,
                        ak,
                        al,
                        am,
                        an,
                        ao,
                        ap,
                        aq,
                        ar,
                        _as,
                        at,
                        au,
                        av,
                        aw,
                        ax,
                        ay,
                        az
                        ));
                }
                return FilePropertiesCollection;
                //3. DataSet - Create column names from first row
                //excelReader.IsFirstRowAsColumnNames = false;
            }
            catch (Exception ex)
            {
                ArrayList FilePropertiesCollection = new ArrayList();
                FilePropertiesCollection.Add(new GetListFromExcel(ex.Message, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""));
                return FilePropertiesCollection;
            }
        }
        private static SqlString getCellValue(int columnCount, int position, DataRow item)
        {
            SqlString result = null;
            try
            {
                result = columnCount < position + 1 || item[position] == null ? SqlString.Null : item[position].ToString();
            }
            catch (Exception exp)
            {

                throw exp;
            }
            return result;
        }
        //FillRow method. The method name has been specified above as 
        //a SqlFunction attribute property
        public static void FindFiles(object objFileProperties, out SqlString a, out SqlString b, out SqlString c, out SqlString d, out SqlString e,
            out SqlString f, out SqlString g, out SqlString h, out SqlString i,
           out SqlString j, out SqlString k, out SqlString l, out SqlString m,
           out SqlString n, out SqlString o, out SqlString p,
           out SqlString q,
           out SqlString r,
           out SqlString s,
           out SqlString t,
           out SqlString u,
           out SqlString v,
           out SqlString w,
           out SqlString x,
           out SqlString y,
           out SqlString z,
           out SqlString aa,
           out SqlString ab,
           out SqlString ac,
           out SqlString ad,
           out SqlString ae,
           out SqlString af,
           out SqlString ag,
           out SqlString ah,
           out SqlString ai,
           out SqlString aj,
           out SqlString ak,
           out SqlString al,
           out SqlString am,
           out SqlString an,
           out SqlString ao,
           out SqlString ap,
           out SqlString aq,
           out SqlString ar,
           out SqlString _as,
           out SqlString at,
           out SqlString au,
           out SqlString av,
           out SqlString aw,
           out SqlString ax,
           out SqlString ay,
           out SqlString az
           )
        {
            //I'm using here the FileProperties class defined above
            GetListFromExcel fileProperties = (GetListFromExcel)objFileProperties;
            //fileName = fileProperties.FileName;
            a = fileProperties.A;
            b = fileProperties.B;
            c = fileProperties.C;
            d = fileProperties.D;
            e = fileProperties.E;
            f = fileProperties.F;
            g = fileProperties.G;
            h = fileProperties.H;
            i = fileProperties.I;
            j = fileProperties.J;
            k = fileProperties.K;
            l = fileProperties.L;
            m = fileProperties.M;
            n = fileProperties.N;
            o = fileProperties.O;
            p = fileProperties.P;
            q = fileProperties.Q;
            r = fileProperties.R;
            s = fileProperties.S;
            t = fileProperties.T;
            u = fileProperties.U;
            v = fileProperties.V;
            w = fileProperties.W;
            x = fileProperties.X;
            y = fileProperties.Y;
            z = fileProperties.Z;
            aa = fileProperties.AA;
            ab = fileProperties.AB;
            ac = fileProperties.AC;
            ad = fileProperties.AD;
            ae = fileProperties.AE;
            af = fileProperties.AF;
            ag = fileProperties.AG;
            ah = fileProperties.AH;
            ai = fileProperties.AI;
            aj = fileProperties.AJ;
            ak = fileProperties.AK;
            al = fileProperties.AL;
            am = fileProperties.AM;
            an = fileProperties.AN;
            ao = fileProperties.AO;
            ap = fileProperties.AP;
            aq = fileProperties.AQ;
            ar = fileProperties.AR;
            _as = fileProperties._AS;
            at = fileProperties.AT;
            au = fileProperties.AU;
            av = fileProperties.AV;
            aw = fileProperties.AW;
            ax = fileProperties.AX;
            ay = fileProperties.AY;
            az = fileProperties.AZ;
        }

    }

}

/*
 EXEC sp_configure 'clr enabled', 1
RECONFIGURE  
EXEC sp_configure 'show advanced options', 1
RECONFIGURE
EXEC sp_configure 'clr strict security', 0
RECONFIGURE
GO
--use Test;  
go  

ALTER AUTHORIZATION ON DATABASE::Test TO [sa]; 
GO
ALTER DATABASE Test SET TRUSTWORTHY ON; 
GO

IF EXISTS (SELECT name FROM sysobjects WHERE name = 'GetExelDataByPath')  
   DROP FUNCTION GetExelDataByPath;  
go  
IF EXISTS (SELECT name FROM sys.assemblies WHERE name = 'VI_CLR')  
   DROP ASSEMBLY VI_CLR;  
go  

CREATE ASSEMBLY VI_CLR FROM 'C:\D\Project\VI_CLR.dll'
WITH PERMISSION_SET = unsafe --SAFE -- EXTERNAL_ACCESS;  
GO  

CREATE OR ALTER FUNCTION GetExelDataByPath(@targetDirectory nvarchar(4000), @sheetName nvarchar(4000))   
RETURNS TABLE (  a nvarchar(max), b nvarchar(max), c nvarchar(max), d nvarchar(max), e nvarchar(max), 
           f nvarchar(max), g nvarchar(max), h nvarchar(max), i nvarchar(max), 
          j nvarchar(max), k nvarchar(max), l nvarchar(max), m nvarchar(max), n nvarchar(max), o nvarchar(max), p nvarchar(max),
q  nvarchar(max),r  nvarchar(max),s  nvarchar(max),t  nvarchar(max),u  nvarchar(max),v  nvarchar(max),w  nvarchar(max),
x  nvarchar(max),y  nvarchar(max),z  nvarchar(max),aa nvarchar(max),ab nvarchar(max),ac nvarchar(max),ad nvarchar(max),ae nvarchar(max),
af nvarchar(max),ag nvarchar(max),ah nvarchar(max),ai nvarchar(max),aj nvarchar(max),ak nvarchar(max),
al nvarchar(max),am nvarchar(max),an nvarchar(max),ao nvarchar(max),ap nvarchar(max),aq nvarchar(max),ar nvarchar(max),_as nvarchar(max),at nvarchar(max),
au nvarchar(max),av nvarchar(max),aw nvarchar(max),ax nvarchar(max),ay nvarchar(max),az nvarchar(max)
)  
AS EXTERNAL NAME VI_CLR.[VI_CLR.GetListFromExcel].[GetExelDataByPath];  
go  



SELECT *
FROM GetExelDataByPath('C:\D\Daily_CheckList.xlsx','sheet1')
     */
