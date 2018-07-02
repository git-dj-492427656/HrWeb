using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    public class Hr_Midysrs
    {
        public int id;
        public string easComCode;

        //部门，岗位 保存岗位配备规则编码
        public string ruleDeptCode;  
        public string rulePostCode;
        public int coreQuota;
        public int coreActual;
        public int boneQuota;
        public int boneActual;
        public int floatQuota;
        public int floatActual;
        public int floatFore;
        public int floatghys;
        public int floattzys;

        public int yearly;
        public int monthly;

        public string deptName;
        public string postName;            
        public string postLevel;

    }
}
