using Feng.Excel.Interfaces;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Feng.Excel.Functions
{
    public class ScriptSample
    {
        public ScriptSample()
        {

        }

        public string Name { get; set; }
        public string Descript { get; set; }
        public string Script { get; set; }
    }

    public class ScriptSampleCollection
    {
        public ScriptSampleCollection()
        {
            Samples = new List<ScriptSample>();
        }
        public string Name = "程序结构";
        public string Descript { get; set; }
        public void Init()
        {
            ScriptSample model = new ScriptSample();
            model.Descript = "循环100次";
            model.Name = "循环100次";
            model.Script = @"var i=0; 
while i<100 
    i=i+1;
endwhile";
            Samples.Add(model);

            model = new ScriptSample();
            model.Descript = "FOREACH";
            model.Name = "FOREACH";
            model.Script = @"
FOREACH ITEM IN LIST
    
ENDFOREACH 
";
            Samples.Add(model);


            model = new ScriptSample();
            model.Descript = "多字符联接";
            model.Name = "多字符联接";
            model.Script = @"
VAR TEXT=""########""
TEXT=TEXT+""########"" 
TEXT=TEXT+""########"" 
TEXT=TEXT+""########"" 
TEXT=TEXT+""########"" 
TEXT=TEXT+""########"" 
";
            Samples.Add(model);

            model = new ScriptSample();
            model.Descript = "单元格FOREACH";
            model.Name = "单元格FOREACH";
            model.Script = @"
VAR I=0; 
VAR CELLS=CELLREANGE(""A1:B10"")
FOREACH ITEM IN CELLS
    I=CONVERTTOINT(ITEM)+I
ENDFOREACH
CELLVALUE(""CELLSUM"",I)
";

            model = new ScriptSample();
            model.Descript = "For循环";
            model.Name = "For循环";
            model.Script = @"
var i=0;
for i=0;i<10;i=i+1;
    if i%5==1 
        sum=sum+i; 
    endif
endfor
";
            Samples.Add(model);
            model = new ScriptSample();
            model.Descript = "IF判断";
            model.Name = "IF判断";
            model.Script = @"var i=0; 
//判断
IF i<100 
    //
ELSE i<200

ELSE I<300

ENDIF";
            Samples.Add(model);

            model = new ScriptSample();
            model.Descript = "单元格循环附值";
            model.Name = "单元格循环附值";
            model.Script = @"var i=0;
var currentcell=me;
while i<100
    currentcell=celldown(me,currentcell);
    cellvalue(currentcell,i);
    i=i+1;
endwhile";
            Samples.Add(model);



            model = new ScriptSample();
            model.Descript = "数据库查询语句";
            model.Name = "数据库查询语句";
            model.Script = @"SqlExecuteScalar(""server=.\zsql2016;database=TEST;USER=SA;PWD=sql2016"",
""SELECT * FROM DBO.T_TABLE1 WHERE 1=1 --STATE=@STATE"",ScriptArg(""hash""))";
            Samples.Add(model);


            model = new ScriptSample();
            model.Descript = "设置Table";
            model.Name = "设置Table";
            model.Script = @"
VAR CELLRANGE=CELLRANGE(""A1:A11"")
VAR INDEX=1;
FOREACH CELL IN CELLRANGE
    CellTable(CELL,""TableName"",INDEX,""COLUMNAME"")
    INDEX=INDEX+1;
ENDFOREACH
";
            Samples.Add(model);
        }
        public List<ScriptSample> Samples { get; set; }

    }


    public class ScriptSampleGroup
    {
        public ScriptSampleGroup()
        {
            SampleGroupes = new List<ScriptSampleCollection>();
        }
        public List<ScriptSampleCollection> SampleGroupes { get; set; }
        public void Init()
        {

        }
    }
}