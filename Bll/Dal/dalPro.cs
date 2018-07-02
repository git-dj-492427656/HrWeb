using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Model;

namespace Bll.Dal
{
    public class dalPro
    {

        //GIS20170814数据库
        public SqlConnection SqlCon2()
        {
            try
            {
                string constr = "server=192.168.11.58;database=GIS20170814;User ID = pubadmin; Password = 88911522;";
                //string connString = "server=192.168.38.61;Initial Catalog= BaisData;User ID = sa; Password = ***;";

                SqlConnection con = new SqlConnection(constr);

                return con;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        

        #region //PcMaker 数据源
        /*
         * 添加PCMaker 规划预算
         * 判断公司同一年度月度数据是否存在
         * 存在则覆盖，反之新增
         */
        public void AddHr_PcMaker_ghys(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                /*判断是否公司月度数据是否存在：存在则修改，否则新增*/
                string sql = @"if not exists(select id from Hr_PcMaker_ghys where pcComCode=@pcComCode and yearly=@yearly and monthly=@monthly)
                                    begin
                                        insert into Hr_PcMaker_ghys(pcComCode,pcComName,cxNum,yieEffic,gjEffic,workDays,yearly,monthly,mid,addPer,addTime) 
                                        values(@pcComCode,@pcComName,@cxNum,@yieEffic,@gjEffic,@workDays,@yearly,@monthly,@mid,@addPer,@addTime)
                                    end
                                else
                                    begin
                                        update Hr_PcMaker_ghys set pcComCode=@pcComCode,pcComName=@pcComName,cxNum=@cxNum,yieEffic=@yieEffic,gjEffic=@gjEffic,
                                        workDays=@workDays,mid=@mid,addPer=@addPer,addTime=@addTime
                                        where pcComCode=@pcComCode and yearly=@yearly and monthly=@monthly
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@pcComCode", diction["pcComCode"]));
                cmd.Parameters.Add(new SqlParameter("@pcComName", diction["pcComName"]));
                cmd.Parameters.Add(new SqlParameter("@cxNum", double.Parse(diction["cxNum"])));
                cmd.Parameters.Add(new SqlParameter("@yieEffic", double.Parse(diction["yieEffic"])));
                cmd.Parameters.Add(new SqlParameter("@gjEffic", double.Parse(diction["gjEffic"])));
                cmd.Parameters.Add(new SqlParameter("@workDays", double.Parse(diction["workDays"])));
                cmd.Parameters.Add(new SqlParameter("@monthly", int.Parse(diction["monthly"])));
                cmd.Parameters.Add(new SqlParameter("@yearly", int.Parse(diction["yearly"])));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
         * 获取公司下的PCMaker规划预算数据
         * 开始月 到 结束月
         */
        public DataTable GetHr_PcMaker_ghys(DateTime start, DateTime end, int mid)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                int yearly1 = start.Year;
                int monthly1 = start.Month;

                int yearly2 = end.Year;
                int monthly2 = end.Month;

                string sql = @"select id,pcComCode,pcComName,cxNum,yieEffic,gjEffic,workDays,monthly,yearly,addTime from Hr_PcMaker_ghys ";

                if (yearly1 < yearly2)
                {
                    sql += " where ((yearly=" + yearly1 + " and monthly>=" + monthly1 + ") or (yearly=" + yearly2 + " and monthly<=" + monthly2 + "))";
                }
                else
                {
                    sql += " where yearly=" + yearly1 + " and monthly>=" + monthly1 + " and monthly<=" + monthly2;
                }

                sql += " and mid=" + mid;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        //删除PcMaker规划预算数据
        public int DelHr_PcMaker_ghysById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "delete Hr_PcMaker_ghys where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                int result = cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }



        /*
         * 添加PCMaker 项目进展
         * 判断公司同一年度月度数据是否存在
         * 存在则覆盖，反之新增
         */
        public void AddHr_PcMaker_xmjz(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                /*判断是否相同的公司，月度数据是否存在：存在则修改，否则新增*/
                string sql = @"if not exists(select id from Hr_PcMaker_xmjz where pcComCode=@pcComCode and yearly=@yearly and monthly=@monthly)
                                    begin
                                        insert into Hr_PcMaker_xmjz(pcComCode,pcComName,cxNum,proBudget,progjBudget,yieEffic,gjEffic,workDays,yearly,monthly,mid,addPer,addTime) 
                                        values(@pcComCode,@pcComName,@cxNum,@proBudget,@progjBudget,@yieEffic,@gjEffic,@workDays,@yearly,@monthly,@mid,@addPer,@addTime)
                                    end
                                else
                                    begin
                                        update Hr_PcMaker_xmjz set pcComCode=@pcComCode,pcComName=@pcComName,cxNum=@cxNum,proBudget=@proBudget,
                                        progjBudget=@progjBudget,yieEffic=@yieEffic,gjEffic=@gjEffic,workDays=@workDays,mid=@mid,addPer=@addPer,addTime=@addTime 
                                        where pcComCode=@pcComCode and yearly=@yearly and monthly=@monthly
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@pcComCode", diction["pcComCode"]));
                cmd.Parameters.Add(new SqlParameter("@pcComName", diction["pcComName"]));
                cmd.Parameters.Add(new SqlParameter("@cxNum", double.Parse(diction["cxNum"])));
                cmd.Parameters.Add(new SqlParameter("@proBudget", double.Parse(diction["proBudget"])));
                cmd.Parameters.Add(new SqlParameter("@progjBudget", double.Parse(diction["progjBudget"])));
                cmd.Parameters.Add(new SqlParameter("@yieEffic", double.Parse(diction["yieEffic"])));
                cmd.Parameters.Add(new SqlParameter("@gjEffic", double.Parse(diction["gjEffic"])));
                cmd.Parameters.Add(new SqlParameter("@workDays", double.Parse(diction["workDays"])));
                cmd.Parameters.Add(new SqlParameter("@yearly", int.Parse(diction["yearly"])));
                cmd.Parameters.Add(new SqlParameter("@monthly", int.Parse(diction["monthly"])));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
        * 获取公司下的PCMaker 项目进展数据
        * 开始月 到 结束月
        */
        public DataTable GetHr_PcMaker_xmjz(DateTime start, DateTime end, int mid)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                int yearly1 = start.Year;
                int monthly1 = start.Month;

                int yearly2 = end.Year;
                int monthly2 = end.Month;

                string sql = @"select id,pcComCode,pcComName,cxNum,proBudget,progjBudget,yieEffic,gjEffic,workDays,monthly,yearly,addTime from Hr_PcMaker_xmjz ";

                if (yearly1 < yearly2)
                {
                    sql += " where ((yearly=" + yearly1 + " and monthly>=" + monthly1 + ") or (yearly=" + yearly2 + " and monthly<=" + monthly2 + "))";
                }
                else
                {
                    sql += " where yearly=" + yearly1 + " and monthly>=" + monthly1 + " and monthly<=" + monthly2;
                }

                sql += " and mid=" + mid;
                

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        //删除PcMaker 项目进展
        public int DelHr_PcMaker_xmjzById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "delete Hr_PcMaker_xmjz where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                int result = cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /*
        * 添加PCMaker 实际
        * 判断公司同一年度月度数据是否存在
        * 存在则覆盖，反之新增
        */
        public void AddHr_PcMaker_fact(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                /*判断是否公司月度数据是否存在：存在则修改，否则新增*/
                string sql = @"if not exists(select id from Hr_PcMaker_fact where pcComCode=@pcComCode and yearly=@yearly and monthly=@monthly)
                                    begin
                                        insert into Hr_PcMaker_fact(pcComCode,pcComName,cxNum,yieEffic,gjEffic,yearly,monthly,mid,addPer,addTime) 
                                        values(@pcComCode,@pcComName,@cxNum,@yieEffic,@gjEffic,@yearly,@monthly,@mid,@addPer,@addTime)
                                    end
                                else
                                    begin
                                        update Hr_PcMaker_fact set pcComCode=@pcComCode,pcComName=@pcComName,cxNum=@cxNum,yieEffic=@yieEffic,
                                        gjEffic=@gjEffic,mid=@mid,addPer=@addPer,addTime=@addTime
                                        where pcComCode=@pcComCode and yearly=@yearly and monthly=@monthly
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@pcComCode", diction["pcComCode"]));
                cmd.Parameters.Add(new SqlParameter("@pcComName", diction["pcComName"]));
                cmd.Parameters.Add(new SqlParameter("@cxNum", double.Parse(diction["cxNum"])));
                cmd.Parameters.Add(new SqlParameter("@yieEffic", double.Parse(diction["yieEffic"])));
                cmd.Parameters.Add(new SqlParameter("@gjEffic", double.Parse(diction["gjEffic"])));
                cmd.Parameters.Add(new SqlParameter("@monthly", int.Parse(diction["monthly"])));
                cmd.Parameters.Add(new SqlParameter("@yearly", int.Parse(diction["yearly"])));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
         * 获取公司下的PCMaker 实际数据
         * 开始月 到 结束月
         */
        public DataTable GetHr_PcMaker_fact(DateTime start, DateTime end, int mid)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                int yearly1 = start.Year;
                int monthly1 = start.Month;

                int yearly2 = end.Year;
                int monthly2 = end.Month;

                string sql = @"select id,pcComCode,pcComName,cxNum,yieEffic,gjEffic,monthly,yearly,addTime from Hr_PcMaker_fact ";

                if (yearly1 < yearly2)
                {
                    sql += " where ((yearly=" + yearly1 + " and monthly>=" + monthly1 + ") or (yearly=" + yearly2 + " and monthly<=" + monthly2 + "))";
                }
                else
                {
                    sql += " where yearly=" + yearly1 + " and monthly>=" + monthly1 + " and monthly<=" + monthly2;
                }

                sql += " and mid=" + mid;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        //删除PcMaker 实际数据
        public int DelHr_PcMaker_factById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "delete Hr_PcMaker_fact where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                int result = cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        #endregion

        #region //Bais 数据源
        /*
       * 添加Bais 规划预算
       * 判断公司同一年度月度数据是否存在
       * 存在则覆盖，反之新增
       */
        public void AddHr_Bais_ghys(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                /*判断是否相同的公司，月度数据是否存在：存在则修改，否则新增*/
                string sql = @"if not exists(select id from Hr_Bais_ghys where baisComCode=@baisComCode and yearly=@yearly and monthly=@monthly)
                                    begin
                                        insert into Hr_Bais_ghys(baisComCode,baisComName,htAmount,ysAmount,lrAmount,yield,yearly,monthly,mid,addPer,addTime) 
                                        values(@baisComCode,@baisComName,@htAmount,@ysAmount,@lrAmount,@yield,@yearly,@monthly,@mid,@addPer,@addTime)
                                    end
                                else
                                    begin
                                        update Hr_Bais_ghys set baisComCode=@baisComCode,baisComName=@baisComName,htAmount=@htAmount,ysAmount=@ysAmount,
                                        lrAmount=@lrAmount,yield=@yield,mid=@mid,addPer=@addPer,addTime=@addTime 
                                        where baisComCode=@baisComCode and yearly=@yearly and monthly=@monthly
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@baisComCode", diction["baisComCode"]));
                cmd.Parameters.Add(new SqlParameter("@baisComName", diction["baisComName"]));
                cmd.Parameters.Add(new SqlParameter("@htAmount", double.Parse(diction["htAmount"])));
                cmd.Parameters.Add(new SqlParameter("@ysAmount", double.Parse(diction["ysAmount"])));
                cmd.Parameters.Add(new SqlParameter("@lrAmount", double.Parse(diction["lrAmount"])));
                cmd.Parameters.Add(new SqlParameter("@yield", double.Parse(diction["yield"])));
                cmd.Parameters.Add(new SqlParameter("@yearly", int.Parse(diction["yearly"])));
                cmd.Parameters.Add(new SqlParameter("@monthly", int.Parse(diction["monthly"])));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
        * 获取公司下的Bais 规划预算数据
        * 开始月 到 结束月
        */
        public DataTable GetHr_Bais_ghys(DateTime start, DateTime end, int mid)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                int yearly1 = start.Year;
                int monthly1 = start.Month;

                int yearly2 = end.Year;
                int monthly2 = end.Month;

                string sql = @"select id,baisComCode,baisComName,htAmount,ysAmount,lrAmount,yield,monthly,yearly,addTime from Hr_Bais_ghys ";

                if (yearly1 < yearly2)
                {
                    sql += " where ((yearly=" + yearly1 + " and monthly>=" + monthly1 + ") or (yearly=" + yearly2 + " and monthly<=" + monthly2 + "))";
                }
                else
                {
                    sql += " where yearly=" + yearly1 + " and monthly>=" + monthly1 + " and monthly<=" + monthly2;
                }

                sql += " and mid=" + mid;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //删除Bais 规划预算
        public int DelHr_Bais_ghysById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "delete Hr_Bais_ghys where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                int result = cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /*
        * 添加Bais 项目进展
        * 判断公司同一年度月度数据是否存在
        * 存在则覆盖，反之新增
        */
        public void AddHr_Bais_xmjz(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                /*判断是否相同的公司，月度数据是否存在：存在则修改，否则新增*/
                string sql = @"if not exists(select id from Hr_Bais_xmjz where baisComCode=@baisComCode and yearly=@yearly and monthly=@monthly)
                                    begin
                                        insert into Hr_Bais_xmjz(baisComCode,baisComName,htAmount,ysAmount,lrAmount,yield,yearly,monthly,mid,addPer,addTime) 
                                        values(@baisComCode,@baisComName,@htAmount,@ysAmount,@lrAmount,@yield,@yearly,@monthly,@mid,@addPer,@addTime)
                                    end
                                else
                                    begin
                                        update Hr_Bais_xmjz set baisComCode=@baisComCode,baisComName=@baisComName,htAmount=@htAmount,ysAmount=@ysAmount,
                                        lrAmount=@lrAmount,yield=@yield,mid=@mid,addPer=@addPer,addTime=@addTime 
                                        where baisComCode=@baisComCode and yearly=@yearly and monthly=@monthly
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@baisComCode", diction["baisComCode"]));
                cmd.Parameters.Add(new SqlParameter("@baisComName", diction["baisComName"]));
                cmd.Parameters.Add(new SqlParameter("@htAmount", double.Parse(diction["htAmount"])));
                cmd.Parameters.Add(new SqlParameter("@ysAmount", double.Parse(diction["ysAmount"])));
                cmd.Parameters.Add(new SqlParameter("@lrAmount", double.Parse(diction["lrAmount"])));
                cmd.Parameters.Add(new SqlParameter("@yield", double.Parse(diction["yield"])));
                cmd.Parameters.Add(new SqlParameter("@yearly", int.Parse(diction["yearly"])));
                cmd.Parameters.Add(new SqlParameter("@monthly", int.Parse(diction["monthly"])));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
        * 获取公司下的Bais 项目进展数据
        * 开始月 到 结束月
        */
        public DataTable GetHr_Bais_xmjz(DateTime start, DateTime end, int mid)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                int yearly1 = start.Year;
                int monthly1 = start.Month;

                int yearly2 = end.Year;
                int monthly2 = end.Month;

                string sql = @"select id,baisComCode,baisComName,htAmount,ysAmount,lrAmount,yield,monthly,yearly,addTime from Hr_Bais_xmjz ";

                if (yearly1 < yearly2)
                {
                    sql += " where ((yearly=" + yearly1 + " and monthly>=" + monthly1 + ") or (yearly=" + yearly2 + " and monthly<=" + monthly2 + "))";
                }
                else
                {
                    sql += " where yearly=" + yearly1 + " and monthly>=" + monthly1 + " and monthly<=" + monthly2;
                }

                sql += " and mid=" + mid;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //删除Bais 项目进展
        public int DelHr_Bais_xmjzById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "delete Hr_Bais_xmjz where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                int result = cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /*
        * 添加Bais 实际
        * 判断公司同一年度月度数据是否存在
        * 存在则覆盖，反之新增
        */
        public void AddHr_Bais_fact(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                /*判断是否相同的公司，月度数据是否存在：存在则修改，否则新增*/
                string sql = @"if not exists(select id from Hr_Bais_fact where baisComCode=@baisComCode and yearly=@yearly and monthly=@monthly)
                                    begin
                                        insert into Hr_Bais_fact(baisComCode,baisComName,htAmount,ysAmount,lrAmount,yield,yearly,monthly,mid,addPer,addTime) 
                                        values(@baisComCode,@baisComName,@htAmount,@ysAmount,@lrAmount,@yield,@yearly,@monthly,@mid,@addPer,@addTime)
                                    end
                                else
                                    begin
                                        update Hr_Bais_fact set baisComCode=@baisComCode,baisComName=@baisComName,htAmount=@htAmount,ysAmount=@ysAmount,
                                        lrAmount=@lrAmount,yield=@yield,mid=@mid,addPer=@addPer,addTime=@addTime 
                                        where baisComCode=@baisComCode and yearly=@yearly and monthly=@monthly
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@baisComCode", diction["baisComCode"]));
                cmd.Parameters.Add(new SqlParameter("@baisComName", diction["baisComName"]));
                cmd.Parameters.Add(new SqlParameter("@htAmount", double.Parse(diction["htAmount"])));
                cmd.Parameters.Add(new SqlParameter("@ysAmount", double.Parse(diction["ysAmount"])));
                cmd.Parameters.Add(new SqlParameter("@lrAmount", double.Parse(diction["lrAmount"])));
                cmd.Parameters.Add(new SqlParameter("@yield", double.Parse(diction["yield"])));
                cmd.Parameters.Add(new SqlParameter("@yearly", int.Parse(diction["yearly"])));
                cmd.Parameters.Add(new SqlParameter("@monthly", int.Parse(diction["monthly"])));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
        * 获取公司下的Bais 实际数据
        * 开始月 到 结束月
        */
        public DataTable GetHr_Bais_fact(DateTime start, DateTime end, int mid)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                int yearly1 = start.Year;
                int monthly1 = start.Month;

                int yearly2 = end.Year;
                int monthly2 = end.Month;

                string sql = @"select id,baisComCode,baisComName,htAmount,ysAmount,lrAmount,yield,monthly,yearly,addTime from Hr_Bais_fact ";

                if (yearly1 < yearly2)
                {
                    sql += " where ((yearly=" + yearly1 + " and monthly>=" + monthly1 + ") or (yearly=" + yearly2 + " and monthly<=" + monthly2 + "))";
                }
                else
                {
                    sql += " where yearly=" + yearly1 + " and monthly>=" + monthly1 + " and monthly<=" + monthly2;
                }

                sql += " and mid=" + mid;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //删除Bais 实际
        public int DelHr_Bais_factById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "delete Hr_Bais_fact where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                int result = cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /*
        * 添加Bais 人工成本-收入
        * 判断公司同一年度月度数据是否存在
        * 存在则覆盖，反之新增
        */
        public void AddHr_Bais_rgsr(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                /*判断是否相同的公司、部门，月度数据是否存在：存在则修改，否则新增*/
                string sql = @"if not exists(select id from Hr_Bais_rgsr where baisComCode=@baisComCode and yearly=@yearly and monthly=@monthly)
                                    begin
                                        insert into HR_Bais_rgsr(baisComCode,baisComName,costType,planBudget,adjustBudget,proBudget,quotaLabor,proportion,yearly,monthly,mid,addPer,addTime) 
                                        values(@baisComCode,@baisComName,@costType,@planBudget,@adjustBudget,@proBudget,@quotaLabor,@proportion,@yearly,@monthly,@mid,@addPer,@addTime)
                                    end
                                else
                                    begin
                                        update Hr_Bais_rgsr set baisComCode=@baisComCode,baisComName=@baisComName,costType=@costType,planBudget=@planBudget,adjustBudget=@adjustBudget,
                                        proBudget=@proBudget,quotaLabor=@quotaLabor,proportion=@proportion,mid=@mid,addPer=@addPer,addTime=@addTime 
                                        where baisComCode=@baisComCode and yearly=@yearly and monthly=@monthly
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@baisComCode", diction["baisComCode"]));
                cmd.Parameters.Add(new SqlParameter("@baisComName", diction["baisComName"]));
                cmd.Parameters.Add(new SqlParameter("@costType", diction["costType"]));
                cmd.Parameters.Add(new SqlParameter("@planBudget", double.Parse(diction["planBudget"])));
                cmd.Parameters.Add(new SqlParameter("@adjustBudget", double.Parse(diction["adjustBudget"])));
                cmd.Parameters.Add(new SqlParameter("@proBudget", double.Parse(diction["proBudget"])));
                cmd.Parameters.Add(new SqlParameter("@quotaLabor", double.Parse(diction["quotaLabor"])));
                cmd.Parameters.Add(new SqlParameter("@proportion", double.Parse(diction["proportion"])));
                cmd.Parameters.Add(new SqlParameter("@yearly", int.Parse(diction["yearly"])));
                cmd.Parameters.Add(new SqlParameter("@monthly", int.Parse(diction["monthly"])));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
        * 获取公司下的Bais 人工收入数据
        * 开始月 到 结束月
        */
        public DataTable GetHr_Bais_rgsr(DateTime start, DateTime end, int mid)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                int yearly1 = start.Year;
                int monthly1 = start.Month;

                int yearly2 = end.Year;
                int monthly2 = end.Month;

                string sql = @"select id,baisComCode,baisComName,costType,planBudget,adjustBudget,proBudget,quotaLabor,proportion,yearly,monthly,mid,addTime from Hr_Bais_rgsr ";

                if (yearly1 < yearly2)
                {
                    sql += " where ((yearly=" + yearly1 + " and monthly>=" + monthly1 + ") or (yearly=" + yearly2 + " and monthly<=" + monthly2 + "))";
                }
                else
                {
                    sql += " where yearly=" + yearly1 + " and monthly>=" + monthly1 + " and monthly<=" + monthly2;
                }

                sql += " and mid=" + mid;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //删除Bais 人工收入
        public int DelHr_Bais_rgsrById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "delete Hr_Bais_rgsr where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                int result = cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /*
       * 添加Bais 人工成本-支出
       * 判断公司同一年度月度数据是否存在
       * 存在则覆盖，反之新增
       */
        public void AddHr_Bais_rgzc(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                /*判断是否相同的公司、部门，月度数据是否存在：存在则修改，否则新增*/
                string sql = @"if not exists(select id from Hr_Bais_rgzc where baisComCode=@baisComCode and yearly=@yearly and monthly=@monthly)
                                    begin
                                        insert into Hr_Bais_rgzc(baisComCode,baisComName,costType,quotaLabor,yearly,monthly,mid,addPer,addTime) 
                                        values(@baisComCode,@baisComName,@costType,@quotaLabor,@yearly,@monthly,@mid,@addPer,@addTime)
                                    end
                                else
                                    begin
                                        update Hr_Bais_rgzc set baisComCode=@baisComCode,baisComName=@baisComName,costType=@costType,
                                        quotaLabor=@quotaLabor,mid=@mid,addPer=@addPer,addTime=@addTime 
                                        where baisComCode=@baisComCode and yearly=@yearly and monthly=@monthly
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@baisComCode", diction["baisComCode"]));
                cmd.Parameters.Add(new SqlParameter("@baisComName", diction["baisComName"]));
                cmd.Parameters.Add(new SqlParameter("@costType", diction["costType"]));
                cmd.Parameters.Add(new SqlParameter("@quotaLabor", double.Parse(diction["quotaLabor"])));
                cmd.Parameters.Add(new SqlParameter("@yearly", int.Parse(diction["yearly"])));
                cmd.Parameters.Add(new SqlParameter("@monthly", int.Parse(diction["monthly"])));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
        * 获取公司下的Bais 人工支出数据
        * 开始月 到 结束月
        */
        public DataTable GetHr_Bais_rgzc(DateTime start, DateTime end, int mid)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                int yearly1 = start.Year;
                int monthly1 = start.Month;

                int yearly2 = end.Year;
                int monthly2 = end.Month;

                string sql = @"select id,baisComCode,baisComName,costType,quotaLabor,yearly,monthly,mid,addTime from Hr_Bais_rgzc ";

                if (yearly1 < yearly2)
                {
                    sql += " where ((yearly=" + yearly1 + " and monthly>=" + monthly1 + ") or (yearly=" + yearly2 + " and monthly<=" + monthly2 + "))";
                }
                else
                {
                    sql += " where yearly=" + yearly1 + " and monthly>=" + monthly1 + " and monthly<=" + monthly2;
                }

                sql += " and mid=" + mid;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //删除Bais 人工支出
        public int DelHr_Bais_rgzcById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "delete Hr_Bais_rgzc where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                int result = cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        #endregion

        #region //Crm数据源
        /*
         * 添加Crm 规划预算
         * 判断公司同一年度月度数据是否存在
         * 存在则覆盖，反之新增
         */
        public void AddHr_Crm_ghys(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                /*判断是否相同的公司，部门，月度数据是否存在：存在则修改，否则新增*/
                string sql = @"if not exists(select id from Hr_Crm_ghys where crmComCode=@crmComCode and crmDeptCode=@crmDeptCode and yearly=@yearly and monthly=@monthly)
                                    begin
                                        insert into Hr_Crm_ghys(crmComCode,crmComName,crmDeptCode,crmDeptName,htAmount,yearly,monthly,mid,addPer,addTime) 
                                        values(@crmComCode,@crmComName,@crmDeptCode,@crmDeptName,@htAmount,@yearly,@monthly,@mid,@addPer,@addTime)
                                    end
                                else
                                    begin
                                        update Hr_Crm_ghys set crmComCode=@crmComCode,crmComName=@crmComName,crmDeptCode=@crmDeptCode,crmDeptName=@crmDeptName,
                                        htAmount=@htAmount,mid=@mid,addPer=@addPer,addTime=@addTime 
                                        where crmComCode=@crmComCode and crmDeptCode=@crmDeptCode and yearly=@yearly and monthly=@monthly
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@crmComCode", diction["crmComCode"]));
                cmd.Parameters.Add(new SqlParameter("@crmComName", diction["crmComName"]));
                cmd.Parameters.Add(new SqlParameter("@crmDeptCode", diction["crmDeptCode"]));
                cmd.Parameters.Add(new SqlParameter("@crmDeptName", diction["crmDeptName"]));
                cmd.Parameters.Add(new SqlParameter("@htAmount", double.Parse(diction["htAmount"])));
                cmd.Parameters.Add(new SqlParameter("@yearly", int.Parse(diction["yearly"])));
                cmd.Parameters.Add(new SqlParameter("@monthly", int.Parse(diction["monthly"])));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
         * 获取公司下的Crm 规划预算数据
         * 开始月 到 结束月
         */
        public DataTable GetHr_Crm_ghys(DateTime start, DateTime end, int mid)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                int yearly1 = start.Year;
                int monthly1 = start.Month;

                int yearly2 = end.Year;
                int monthly2 = end.Month;

                string sql = @"select id,crmComCode,crmComName,crmDeptCode,crmDeptName,htAmount,yearly,monthly,mid,addTime from Hr_Crm_ghys ";

                if (yearly1 < yearly2)
                {
                    sql += " where ((yearly=" + yearly1 + " and monthly>=" + monthly1 + ") or (yearly=" + yearly2 + " and monthly<=" + monthly2 + "))";
                }
                else
                {
                    sql += " where yearly=" + yearly1 + " and monthly>=" + monthly1 + " and monthly<=" + monthly2;
                }

                sql += " and mid=" + mid;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //删除Crm 规划预算数据
        public int DelHr_Crm_ghysById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "delete Hr_Crm_ghys where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                int result = cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /*
         * 添加Crm 实际合同
         * 判断公司同一年度月度数据是否存在
         * 存在则覆盖，反之新增
         */
        public void AddHr_Crm_fact(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                /*判断是否相同的公司，部门，月度数据是否存在：存在则修改，否则新增*/
                string sql = @"if not exists(select id from Hr_Crm_fact where crmComCode=@crmComCode and crmDeptCode=@crmDeptCode and yearly=@yearly and monthly=@monthly)
                                    begin
                                        insert into Hr_Crm_fact(crmComCode,crmComName,crmDeptCode,crmDeptName,syAmount,yearly,monthly,mid,addPer,addTime) 
                                        values(@crmComCode,@crmComName,@crmDeptCode,@crmDeptName,@syAmount,@yearly,@monthly,@mid,@addPer,@addTime)
                                    end
                                else
                                    begin
                                        update Hr_Crm_fact set crmComCode=@crmComCode,crmComName=@crmComName,crmDeptCode=@crmDeptCode,
                                        crmDeptName=@crmDeptName,syAmount=@syAmount,mid=@mid,addPer=@addPer,addTime=@addTime 
                                        where crmComCode=@crmComCode and crmDeptCode=@crmDeptCode and yearly=@yearly and monthly=@monthly
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@crmComCode", diction["crmComCode"]));
                cmd.Parameters.Add(new SqlParameter("@crmComName", diction["crmComName"]));
                cmd.Parameters.Add(new SqlParameter("@crmDeptCode", diction["crmDeptCode"]));
                cmd.Parameters.Add(new SqlParameter("@crmDeptName", diction["crmDeptName"]));
                cmd.Parameters.Add(new SqlParameter("@syAmount", double.Parse(diction["syAmount"])));
                cmd.Parameters.Add(new SqlParameter("@yearly", int.Parse(diction["yearly"])));
                cmd.Parameters.Add(new SqlParameter("@monthly", int.Parse(diction["monthly"])));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
         * 获取公司下的Crm 实际数据
         * 开始月 到 结束月
         */
        public DataTable GetHr_Crm_fact(DateTime start, DateTime end, int mid)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                int yearly1 = start.Year;
                int monthly1 = start.Month;

                int yearly2 = end.Year;
                int monthly2 = end.Month;

                string sql = @"select id,crmComCode,crmComName,crmDeptCode,crmDeptName,syAmount,yearly,monthly,mid,addTime from Hr_Crm_fact ";

                if (yearly1 < yearly2)
                {
                    sql += " where ((yearly=" + yearly1 + " and monthly>=" + monthly1 + ") or (yearly=" + yearly2 + " and monthly<=" + monthly2 + "))";
                }
                else
                {
                    sql += " where yearly=" + yearly1 + " and monthly>=" + monthly1 + " and monthly<=" + monthly2;
                }

                sql += " and mid=" + mid;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //删除Crm 实际数据
        public int DelHr_Crm_factById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "delete Hr_Crm_fact where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                int result = cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /*
         * 
         *  获取当前公司 月度Crm 规划合同额/实际合同额  数据源
         *  baisComCode 当前公司 bais公司编码
         */
        public DataTable GetHr_Crm_tzysFrom(int yearly, int monthly, string baisComCode)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select t1.crmComCode,t1.crmComName,t1.crmDeptCode,t1.crmDeptName,t1.htAmount,t1.yearly,t1.monthly,ISNULL(t2.syAmount,0)syAmount from Hr_Crm_ghys t1
                            left join Hr_Crm_fact t2 on t2.crmComCode=t1.crmComCode and t2.crmDeptCode=t1.crmDeptCode and t2.yearly=t1.yearly and t2.monthly=t1.monthly 
                            left join Hr_Middle_sys t3 on t3.crmCode=t1.crmComCode and t3.mType='公司'";

                sql += " where t3.baisCode=" + baisComCode + " and t1.yearly=" + yearly + " and t1.monthly=" + monthly;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        
        /*
         * 添加Crm 调整预算
         * 判断公司同一年度月度数据是否存在
         * 存在则覆盖，反之新增
         */
        public int AddHr_Crm_tzys(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                /*判断是否相同的公司，部门，月度数据是否存在：存在则修改，否则新增*/
                string sql = @"if not exists(select id from Hr_Crm_tzys where crmComCode=@crmComCode and crmDeptCode=@crmDeptCode and yearly=@yearly and monthly=@monthly)
                                    begin
                                        insert into Hr_Crm_tzys(crmComCode,crmComName,crmDeptCode,crmDeptName,htAmount,syAmount,ghsyAmount,tzsyAmount,yearly,monthly,mid,addPer,addTime) 
                                        values(@crmComCode,@crmComName,@crmDeptCode,@crmDeptName,@htAmount,@syAmount,@ghsyAmount,@tzsyAmount,@yearly,@monthly,@mid,@addPer,@addTime)
                                    end
                                else
                                    begin
                                        update Hr_Crm_tzys set crmComCode=@crmComCode,crmComName=@crmComName,crmDeptCode=@crmDeptCode,crmDeptName=@crmDeptName,htAmount=@htAmount,
                                        syAmount=@syAmount,ghsyAmount=@ghsyAmount,tzsyAmount=@tzsyAmount,mid=@mid,addPer=@addPer,addTime=@addTime 
                                        where crmComCode=@crmComCode and crmDeptCode=@crmDeptCode and yearly=@yearly and monthly=@monthly
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@crmComCode", diction["crmComCode"]));
                cmd.Parameters.Add(new SqlParameter("@crmComName", diction["crmComName"]));
                cmd.Parameters.Add(new SqlParameter("@crmDeptCode", diction["crmDeptCode"]));
                cmd.Parameters.Add(new SqlParameter("@crmDeptName", diction["crmDeptName"]));
                cmd.Parameters.Add(new SqlParameter("@htAmount", double.Parse(diction["htAmount"])));
                cmd.Parameters.Add(new SqlParameter("@syAmount", double.Parse(diction["syAmount"])));
                cmd.Parameters.Add(new SqlParameter("@ghsyAmount", double.Parse(diction["ghsyAmount"])));
                cmd.Parameters.Add(new SqlParameter("@tzsyAmount", double.Parse(diction["tzsyAmount"])));
                cmd.Parameters.Add(new SqlParameter("@yearly", int.Parse(diction["yearly"])));
                cmd.Parameters.Add(new SqlParameter("@monthly", int.Parse(diction["monthly"])));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                int result = cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
         * 获取公司下的Crm 调整预算数据
         * 开始月 到 结束月
         */
        public DataTable GetHr_Crm_tzys(DateTime start, DateTime end, int mid)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                int yearly1 = start.Year;
                int monthly1 = start.Month;

                int yearly2 = end.Year;
                int monthly2 = end.Month;

                string sql = @"select id,crmComCode,crmComName,crmDeptCode,crmDeptName,htAmount,syAmount,ghsyAmount,tzsyAmount,yearly,monthly,mid,addTime from Hr_Crm_tzys ";

                if (yearly1 < yearly2)
                {
                    sql += " where ((yearly=" + yearly1 + " and monthly>=" + monthly1 + ") or (yearly=" + yearly2 + " and monthly<=" + monthly2 + "))";
                }
                else
                {
                    sql += " where yearly=" + yearly1 + " and monthly>=" + monthly1 + " and monthly<=" + monthly2;
                }

                sql += " and mid=" + mid;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //删除Crm 调整预算数据
        public void DelHr_Crm_tzysById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "delete Hr_Crm_tzys where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();
                
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
         * 修改公司下的Crm 调整预算合同额
         */
        public int EditHr_Crm_tzysById(int id, double htAmount, double tzsyAmount, string addPer)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "update Hr_Crm_tzys set htAmount=@htAmount,tzsyAmount=@tzsyAmount,addPer=@addPer,addTime=@addTime where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                cmd.Parameters.Add(new SqlParameter("@htAmount", htAmount));
                cmd.Parameters.Add(new SqlParameter("@tzsyAmount", tzsyAmount));
                cmd.Parameters.Add(new SqlParameter("@addPer", addPer));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));
                int result = cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        #endregion

        #region //Bhr 数据源
        /*
        * 添加Bhr 实际
        * 判断公司同一年度月度数据是否存在
        * 存在则覆盖，反之新增
        */
        public void AddHr_Bhr_fact(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                /*判断是否相同的公司，部门，岗位，员工，月度数据是否存在：存在则修改，否则新增*/
                string sql = @"if not exists(select id from Hr_Bhr_fact where easComCode=@easComCode and easDeptCode=@easDeptCode and easPostCode=@easPostCode and staffId=@staffId and yearly=@yearly and monthly=@monthly)
                                    begin
                                        insert into Hr_Bhr_fact(easComCode,easComName,easDeptCode,easDeptName,easPostCode,easPostName,staffId,postLevel,postType,wage,workDays,
                                        yearly,monthly,mid,addPer,addTime) values(@easComCode,@easComName,@easDeptCode,@easDeptName,@easPostCode,@easPostName,@staffId,@postLevel,
                                        @postType,@wage,@workDays,@yearly,@monthly,@mid,@addPer,@addTime)
                                    end
                                else
                                    begin
                                        update Hr_Bhr_fact set easComCode=@easComCode,easComName=@easComName,easDeptCode=@easDeptCode,easDeptName=@easDeptName,easPostCode=@easPostCode,
                                        easPostName=@easPostName,staffId=@staffId,postLevel=@postLevel,postType=@postType,wage=@wage,workDays=@workDays,
                                        mid=@mid,addPer=@addPer,addTime=@addTime where easComCode=@easComCode and easDeptCode=@easDeptCode and easPostCode=@easPostCode and staffId=@staffId and yearly=@yearly and monthly=@monthly
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@easComCode", diction["easComCode"]));
                cmd.Parameters.Add(new SqlParameter("@easComName", diction["easComName"]));
                cmd.Parameters.Add(new SqlParameter("@easDeptCode", diction["easDeptCode"]));
                cmd.Parameters.Add(new SqlParameter("@easDeptName", diction["easDeptName"]));
                cmd.Parameters.Add(new SqlParameter("@easPostCode", diction["easPostCode"]));
                cmd.Parameters.Add(new SqlParameter("@easPostName", diction["easPostName"]));
                cmd.Parameters.Add(new SqlParameter("@staffId", diction["staffId"]));
                cmd.Parameters.Add(new SqlParameter("@postLevel", diction["postLevel"]));
                cmd.Parameters.Add(new SqlParameter("@postType", diction["postType"]));
                cmd.Parameters.Add(new SqlParameter("@wage", double.Parse(diction["wage"])));
                cmd.Parameters.Add(new SqlParameter("@workDays", double.Parse(diction["workDays"])));
                cmd.Parameters.Add(new SqlParameter("@yearly", int.Parse(diction["yearly"])));
                cmd.Parameters.Add(new SqlParameter("@monthly", int.Parse(diction["monthly"])));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
       * 获取公司下的Bhr 实际数据
       * 开始月 到 结束月
       */
        public DataTable GetHr_Bhr_fact(DateTime start, DateTime end, int mid)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                int yearly1 = start.Year;
                int monthly1 = start.Month;

                int yearly2 = end.Year;
                int monthly2 = end.Month;

                string sql = @"select id,easComCode,easComName,easDeptCode,easDeptName,easPostCode,easPostName,staffId,postLevel,postType,wage,workDays,
                            yearly,monthly,mid,addTime from Hr_Bhr_fact ";

                if (yearly1 < yearly2)
                {
                    sql += " where ((yearly=" + yearly1 + " and monthly>=" + monthly1 + ") or (yearly=" + yearly2 + " and monthly<=" + monthly2 + "))";
                }
                else
                {
                    sql += " where yearly=" + yearly1 + " and monthly>=" + monthly1 + " and monthly<=" + monthly2;
                }

                sql += " and mid=" + mid;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //删除Bhr 实际数据
        public int DelHr_Bhr_factById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "delete Hr_Bhr_fact where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                int result = cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        #endregion


        #region //岗位配备规则

        
        /*
         * 岗位配备规则
         * 营收规则
         * 营收规则-岗位人数配备
         * 营收规则-岗位浮动人数
         * 同一公司，同一年度规则表唯一
         */
        public void AddHr_Rule_ys(List<Hr_Rule_yshead> list1, List<Hr_Rule_ysgw> list2, List<Hr_Rule_ysfd> list3, string addPer, int mid)
        {
            SqlTransaction trans = null;
            SqlConnection sqlCon = SqlCon2();

            try
            {
                sqlCon.Open();
                trans = sqlCon.BeginTransaction();

                string sql = "delete Hr_Rule_yshead where baisComCode=@baisComCode and yearly=@yearly;";
                foreach (var item1 in list1)
                {
                    sql += "insert into Hr_Rule_yshead(id,baisComCode,ysAmount,yield,effic,dhbEffic,nqEffic,wqEffic,smzEffic,workDays,yearly,mid,addPer,addTime)";
                    sql += "values('" + item1.id + "','" + item1.baisComCode + "'," + item1.ysAmount + "," + item1.yield + ","
                        + item1.effic + "," + item1.dhbEffic + "," + item1.nqEffic + "," + item1.wqEffic + ","
                        + item1.smzEffic + "," + item1.workDays + "," + item1.yearly + ","
                        + mid + ",'" + addPer + "','" + DateTime.Now + "');";
                }

                sql += "delete Hr_Rule_ysgw where baisComCode=@baisComCode and yearly=@yearly;";
                foreach (var item2 in list2)
                {
                    sql += "insert into Hr_Rule_ysgw(id,coreNum,boneNum,baisComCode,deptCode,deptName,postCode,postName,postLevel,costType,quotaWage,yearly,mid,addPer,addTime)";
                    sql += "values('" + item2.id + "'," + item2.coreNum + "," + item2.boneNum + ",'" + item2.baisComCode + "','"
                        + item2.deptCode + "','" + item2.deptName + "','" + item2.postCode + "','" + item2.postName + "','"
                        + item2.postLevel + "','" + item2.costType + "'," + item2.quotaWage + "," + item2.yearly + ","
                        + mid + ",'" + addPer + "','" + DateTime.Now + "');";
                }

                sql += "delete Hr_Rule_ysfd where baisComCode=@baisComCode and yearly=@yearly;";
                foreach (var item3 in list3)
                {
                    sql += "insert into Hr_Rule_ysfd(rid,gwid,floatNum,baisComCode,yearly,mid,addPer,addTime)";
                    sql += "values('" + item3.rid + "','" + item3.gwid + "','" + item3.floatNum + "','" + item3.baisComCode + "','" + item3.yearly + "','"
                        + mid + "','" + addPer + "','" + DateTime.Now + "')";
                }

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@baisComCode", list1[0].baisComCode));
                cmd.Parameters.Add(new SqlParameter("@yearly", list1[0].yearly));

                cmd.Transaction = trans;
                cmd.ExecuteNonQuery();

                trans.Commit();//执行提交事务 

                cmd.Dispose();
            }
            catch (Exception e)
            {
                //如果前面有异常则事务回滚  
                trans.Rollback();
                throw e;
            }
            finally
            {
                trans = null;
                sqlCon.Close();
            }
        }


        //获取 营收规则数据
        public DataTable GetHr_Rule_yshead(int yearly, string baisComCode)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select id,baisComCode,ysAmount,yield,effic,dhbEffic,nqEffic,wqEffic,smzEffic,workDays,yearly from Hr_Rule_yshead ";
                sql += " where baisComCode =" + baisComCode + " and yearly=" + yearly;
                sql += " order by yield ";

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //获取 营收规则 岗位人数数据
        public DataTable GetHr_Rule_ysgw(int yearly, string baisComCode)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select id,coreNum,boneNum,baisComCode,deptCode,deptName,postCode,postName,postLevel,costType,quotaWage,yearly from Hr_Rule_ysgw ";
                sql += " where baisComCode =" + baisComCode + " and yearly=" + yearly;
                sql += " order by deptCode ";

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //获取 营收规则 岗位浮动人数数据
        public DataTable GetHr_Rule_ysfd(int yearly, string baisComCode)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select id,rid,gwid,baisComCode,floatNum,yearly from Hr_Rule_ysfd ";
                sql += " where baisComCode =" + baisComCode + " and yearly=" + yearly;
                sql += " order by gwid ";

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /*
         * 岗位配备规则
         * 合同规则
         * 合同规则-岗位配备
         * 合同规则-岗位人数
         * 同一公司，同一年度规则表唯一
         */
        public void AddHr_Rule_ht(List<Hr_Rule_hthead> list1, List<Hr_Rule_htgw> list2, List<Hr_Rule_htfd> list3, string addPer, int mid)
        {
            SqlTransaction trans = null;
            SqlConnection sqlCon = SqlCon2();

            try
            {
                sqlCon.Open();
                trans = sqlCon.BeginTransaction();

                string sql = "delete Hr_Rule_hthead where baisComCode=@baisComCode and yearly=@yearly and htType=@htType;";
                foreach (var item1 in list1)
                {
                    sql += "insert into Hr_Rule_hthead(id,baisComCode,htType,htTitle,htAmount,yearly,mid,addPer,addTime)";
                    sql += "values('" + item1.id + "','" + item1.baisComCode + "','" + item1.htType + "','" + item1.htTitle + "'," + item1.htAmount + "," + item1.yearly + ","
                        + mid + ",'" + addPer + "','" + DateTime.Now + "');";
                }

                sql += "delete Hr_Rule_htgw where baisComCode=@baisComCode and yearly=@yearly and htType=@htType;";
                foreach (var item2 in list2)
                {
                    sql += "insert into Hr_Rule_htgw(id,htType,coreNum,boneNum,baisComCode,deptCode,deptName,postCode,postName,postLevel,costType,quotaWage,yearly,mid,addPer,addTime)";
                    sql += "values('" + item2.id + "','" + item2.htType + "'," + item2.coreNum + "," + item2.boneNum + ",'" + item2.baisComCode + "','"
                        + item2.deptCode + "','" + item2.deptName + "','" + item2.postCode + "','" + item2.postName + "','"
                        + item2.postLevel + "','" + item2.costType + "'," + item2.quotaWage + "," + item2.yearly + ","
                        + mid + ",'" + addPer + "','" + DateTime.Now + "');";
                }

                sql += "delete Hr_Rule_htfd where baisComCode=@baisComCode and yearly=@yearly and htType=@htType;";
                foreach (var item3 in list3)
                {
                    sql += "insert into Hr_Rule_htfd(rid,gwid,htType,floatNum,baisComCode,yearly,mid,addPer,addTime)";
                    sql += "values('" + item3.rid + "','" + item3.gwid + "','" + item3.htType + "','" + item3.floatNum + "','" + item3.baisComCode + "','" + item3.yearly + "','"
                        + mid + "','" + addPer + "','" + DateTime.Now + "')";
                }

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@baisComCode", list1[0].baisComCode));
                cmd.Parameters.Add(new SqlParameter("@yearly", list1[0].yearly));
                cmd.Parameters.Add(new SqlParameter("@htType", list1[0].htType));

                cmd.Transaction = trans;
                cmd.ExecuteNonQuery();

                trans.Commit();//执行提交事务 

                cmd.Dispose();
            }
            catch (Exception e)
            {
                //如果前面有异常则事务回滚  
                trans.Rollback();
                throw e;
            }
            finally
            {
                trans = null;
                sqlCon.Close();
            }
        }

        //获取 合同规则数据
        public DataTable GetHr_Rule_hthead(int yearly, string htType, string baisComCode)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select id,baisComCode,htType,htTitle,htAmount,yearly from Hr_Rule_hthead ";
                sql += " where baisComCode ='" + baisComCode + "' and yearly=" + yearly + " and htType='" + htType + "'";
                sql += " order by htAmount ";

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //获取 合同规则 岗位数据
        public DataTable GetHr_Rule_htgw(int yearly, string htType, string baisComCode)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select id,htType,coreNum,boneNum,baisComCode,deptCode,deptName,postCode,postName,postLevel,costType,quotaWage,yearly from Hr_Rule_htgw ";
                sql += " where baisComCode =" + baisComCode + " and yearly=" + yearly;

                if ("" != htType)
                {
                    sql += " and htType='" + htType + "' ";
                }

                sql += " order by deptCode ";

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //获取 合同规则 岗位人数数据
        public DataTable GetHr_Rule_htfd(int yearly, string htType, string baisComCode)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select id,rid,gwid,htType,baisComCode,floatNum,yearly from Hr_Rule_htfd ";
                sql += " where baisComCode =" + baisComCode + " and yearly=" + yearly + " and htType='" + htType + "'";
                sql += " order by gwid ";

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }



        #endregion


        #region //编码对照

        //添加 中间表
        public int AddHr_Middle_sys(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                /*判断是否相同的easCode数据是否存在：存在则修改，否则新增*/
                string sql = @"if not exists(select id from Hr_Middle_sys where mType=@mType and easCode=@easCode)
                                    begin
                                        insert into Hr_Middle_sys(mType,easCode,easName,baisCode,baisName,pcMakerCode,pcMakerName,crmCode,crmName,mid,addPer,addTime)  
                                        values(@mType,@easCode,@easName,@baisCode,@baisName,@pcMakerCode,@pcMakerName,@crmCode,@crmName,@mid,@addPer,@addTime)
                                    end
                                else
                                    begin
                                        update Hr_Middle_sys set mType=@mType,easCode=@easCode,easName=@easName,baisCode=@baisCode,baisName=@baisName,
                                        pcMakerCode=@pcMakerCode,pcMakerName=@pcMakerName,crmCode=@crmCode,crmName=@crmName,mid=@mid,addPer=@addPer,addTime=@addTime 
                                        where mType=@mType and easCode=@easCode
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@mType", diction["mType"]));
                cmd.Parameters.Add(new SqlParameter("@easCode", diction["easCode"]));
                cmd.Parameters.Add(new SqlParameter("@easName", diction["easName"]));
                cmd.Parameters.Add(new SqlParameter("@baisCode", diction["baisCode"]));
                cmd.Parameters.Add(new SqlParameter("@baisName", diction["baisName"]));
                cmd.Parameters.Add(new SqlParameter("@pcMakerCode", diction["pcMakerCode"]));
                cmd.Parameters.Add(new SqlParameter("@pcMakerName", diction["pcMakerName"]));
                cmd.Parameters.Add(new SqlParameter("@crmCode", diction["crmCode"]));
                cmd.Parameters.Add(new SqlParameter("@crmName", diction["crmName"]));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                int result = cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

                return result;

            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //获取中间表数据
        public DataTable GetHr_Middle_sys(int mid)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select id,mType,easCode,easName,baisCode,baisName,pcMakerCode,pcMakerName,crmCode,crmName,mid,addTime from Hr_Middle_sys ";
                sql += " where mid =" + mid;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //通过id 获取中间表数据【未启用】
        public DataTable GetHR_Middle_sysById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select id,mType,easCode,easName,baisCode,baisName,pcMakerCode,pcMakerName,crmCode,crmName,mid,addPer,addTime from Hr_Middle_sys where id=" + id;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //通过id 修改中间表数据【未启用】
        public int EditHR_Middle(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                /*判断是否相同的gisCode数据是否存在：存在则修改，否则新增*/
                string sql = @"if exists(select id from HR_Middle where id=@id)
                                    begin
                                        update HR_Middle set mType=@mType,pcMakerCode=@pcMakerCode,baisCode=@baisCode,crmCode=@crmCode,pcMakerName=@pcMakerName,
                                        baisName=@baisName,crmName=@crmName,gisCode=@gisCode,gisName=@gisName,mid=@mid,addPer=@addPer,addTime=@addTime 
                                        where id=@id
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@id", int.Parse(diction["id"])));
                cmd.Parameters.Add(new SqlParameter("@mType", diction["mType"]));
                cmd.Parameters.Add(new SqlParameter("@pcMakerCode", diction["pcMakerCode"]));
                cmd.Parameters.Add(new SqlParameter("@baisCode", diction["baisCode"]));
                cmd.Parameters.Add(new SqlParameter("@crmCode", diction["crmCode"]));
                cmd.Parameters.Add(new SqlParameter("@gisCode", diction["gisCode"]));
                cmd.Parameters.Add(new SqlParameter("@pcMakerName", diction["pcMakerName"]));
                cmd.Parameters.Add(new SqlParameter("@baisName", diction["baisName"]));
                cmd.Parameters.Add(new SqlParameter("@crmName", diction["crmName"]));
                cmd.Parameters.Add(new SqlParameter("@gisName", diction["gisName"]));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                int result = cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

                return result;

            }
            catch (Exception e)
            {
                throw e;
            }
        }
        
        //删除 中间编码对照表
        public int DelHr_Middle_sysById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "delete Hr_Middle_sys where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                int result = cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }



        //添加 岗位配备规则中间编码对照
        public int AddHr_Middle_rule(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                /*判断是否相同的ruleCode/easCode数据是否存在：存在则修改，否则新增*/
                string sql = @"if not exists(select id from Hr_Middle_rule where mType=@mType and ruleCode=@ruleCode and easCode=@easCode)
                                    begin
                                        insert into Hr_Middle_rule(mType,ruleCode,ruleName,easCode,easName,mid,addPer,addTime)  
                                        values(@mType,@ruleCode,@ruleName,@easCode,@easName,@mid,@addPer,@addTime)
                                    end
                                else
                                    begin
                                        update Hr_Middle_rule set mType=@mType,ruleCode=@ruleCode,ruleName=@ruleName,easCode=@easCode,easName=@easName,
                                        mid=@mid,addPer=@addPer,addTime=@addTime 
                                        where mType=@mType and ruleCode=@ruleCode and easCode=@easCode
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@mType", diction["mType"]));
                cmd.Parameters.Add(new SqlParameter("@ruleCode", diction["ruleCode"]));
                cmd.Parameters.Add(new SqlParameter("@ruleName", diction["ruleName"]));
                cmd.Parameters.Add(new SqlParameter("@easCode", diction["easCode"]));
                cmd.Parameters.Add(new SqlParameter("@easName", diction["easName"]));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                int result = cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

                return result;

            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //获取 岗位配备规则中间编码对照数据
        public DataTable GetHr_Middle_rule(int mid)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select id,mType,ruleCode,ruleName,easCode,easName from Hr_Middle_rule ";
                sql += " where mid =" + mid;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //删除 岗位配备规则中间编码对照
        public int DelHr_Middle_ruleById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "delete Hr_Middle_rule where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                int result = cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        #endregion



        /*
         * 
         *  获取当前公司 月度PCMaker规划预算
         *  baisComCode 当前公司 bais公司编码
         */
        public DataTable GetHr_PcMaker_ghysByMonth(int yearly, int monthly, string baisComCode)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select t1.cxNum,t1.yieEffic,t1.gjEffic,t1.workDays,t1.yearly,t1.monthly,t2.easCode from Hr_PcMaker_ghys t1
                            left join Hr_Middle_sys t2 on t2.pcMakerCode=t1.pcComCode and t2.mType='公司' ";
                sql += " where t2.baisCode=" + baisComCode + " and t1.yearly=" + yearly + " and t1.monthly=" + monthly;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
         * 
         *  获取当前公司 月度bais规划预算
         *  baisComCode 当前公司 bais公司编码
         */
        public DataTable GetHr_Bais_ghysByMonth(int yearly, int monthly, string baisComCode)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select t1.htAmount,t1.ysAmount,t1.lrAmount,t1.yield,t1.yearly,t1.monthly,t2.easCode from Hr_Bais_ghys t1 
                            left join Hr_Middle_sys t2 on t2.baisCode=t1.baisComCode and t2.mType='公司' ";
                sql += " where t2.baisCode=" + baisComCode + " and t1.yearly=" + yearly + " and t1.monthly=" + monthly;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
         * 
         *  获取当前公司 月度Crm规划预算
         *  baisComCode 当前公司 bais公司编码
         */
        public DataTable GetHr_Crm_ghysByMonth(int yearly, int monthly, string baisComCode)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select t1.crmComCode,t1.crmDeptCode,t1.htAmount,t1.yearly,t1.monthly,t2.easCode from Hr_Crm_ghys t1 
                            left join Hr_Middle_sys t2 on t2.crmCode=t1.crmComCode and t2.mType='公司' ";
                sql += " where t2.baisCode=" + baisComCode + " and t1.yearly=" + yearly + " and t1.monthly=" + monthly;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /*
         * 
         *  获取当前公司 月度PCMaker项目进展
         *  baisComCode 当前公司 bais公司编码
         */
        public DataTable GetHr_PcMaker_xmjzByMonth(int yearly, int monthly, string baisComCode)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select t1.cxNum,t1.proBudget,t1.progjBudget,t1.yieEffic,t1.gjEffic,t1.workDays,t1.yearly,t1.monthly,t2.easCode from Hr_PcMaker_xmjz t1
                            left join Hr_Middle_sys t2 on t2.pcMakerCode=t1.pcComCode and t2.mType='公司' ";

                sql += " where t2.baisCode=" + baisComCode + " and t1.yearly=" + yearly + " and t1.monthly=" + monthly;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
         * 
         *  获取当前公司 月度bais项目进展
         *  baisComCode 当前公司 bais公司编码
         */
        public DataTable GetHr_Bais_xmjzByMonth(int yearly, int monthly, string baisComCode)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select t1.htAmount,t1.ysAmount,t1.lrAmount,t1.yearly,t1.monthly,t2.easCode from Hr_Bais_xmjz t1
                            left join Hr_Middle_sys t2 on t2.baisCode=t1.baisComCode and t2.mType='公司' ";
                sql += " where t2.baisCode=" + baisComCode + " and t1.yearly=" + yearly + " and t1.monthly=" + monthly;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /*
         * 
         *  获取当前公司 月度PCMaker实际
         *  baisComCode 当前公司 bais公司编码
         */
        public DataTable GetHr_PcMaker_factByMonth(int yearly, int monthly, string baisComCode)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select t1.cxNum,t1.yieEffic,t1.gjEffic,t1.yearly,t1.monthly,t2.easCode from Hr_PcMaker_fact t1
                            left join Hr_Middle_sys t2 on t2.pcMakerCode=t1.pcComCode and t2.mType='公司' ";

                sql += " where t2.baisCode=" + baisComCode + " and t1.yearly=" + yearly + " and t1.monthly=" + monthly;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
         * 
         *  获取当前公司 月度bais实际
         *  baisComCode 当前公司 bais公司编码
         */
        public DataTable GetHr_Bais_factByMonth(int yearly, int monthly, string baisComCode)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select t1.htAmount,t1.ysAmount,t1.lrAmount,t1.yield,t1.yearly,t1.monthly,t2.easCode from Hr_Bais_fact t1
                            left join Hr_Middle_sys t2 on t2.baisCode=t1.baisComCode and t2.mType='公司' ";

                sql += " where t2.baisCode=" + baisComCode + " and t1.yearly=" + yearly + " and t1.monthly=" + monthly;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
         * 
         *  获取当前公司 月度Bhr实际
         *  baisComCode 当前公司 bais公司编码
         */
        public DataTable GetHr_Bhr_factByMonth(int yearly, int monthly, string baisComCode)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select t1.easComCode,t1.easComName,t1.easDeptCode,t1.easDeptName,t1.easPostCode,t1.easPostName,t1.postLevel,t1.postType,
                            t1.wage,t1.workDays,t1.yearly,t1.monthly from Hr_Bhr_fact t1 
                            left join Hr_Middle_sys t2 on t2.easCode=t1.easComCode and t2.mType='公司' ";

                sql += " where t2.baisCode=" + baisComCode + " and t1.yearly=" + yearly + " and t1.monthly=" + monthly;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        


        #region //规划、调整预算，项目进展，实际

        /*
         * 添加规划预算
         * 同一公司 同一年的同一月度 数据唯一
         */
        public int AddHr_Midghys(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                string sql = @"if not exists(select id from Hr_Midghys where easComCode=@easComCode and yearly=@yearly and monthly=@monthly)
                                    begin
                                        insert into Hr_Midghys(easComCode,cxNum,htAmount,ysAmount,lrAmount,yield,yieEffic,gjEffic,proTeams,workDays,yearly,monthly,mid,addPer,addTime) 
                                        values(@easComCode,@cxNum,@htAmount,@ysAmount,@lrAmount,@yield,@yieEffic,@gjEffic,@proTeams,@workDays,@yearly,@monthly,@mid,@addPer,@addTime)
                                    end
                                else
                                    begin
                                        update Hr_Midghys set cxNum=@cxNum,htAmount=@htAmount,ysAmount=@ysAmount,lrAmount=@lrAmount,yield=@yield,yieEffic=@yieEffic,gjEffic=@gjEffic,
                                        proTeams=@proTeams,workDays=@workDays,yearly=@yearly,monthly=@monthly,mid=@mid,addPer=@addPer,addTime=@addTime 
                                        where easComCode=@easComCode and yearly=@yearly and monthly=@monthly
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@easComCode", diction["easComCode"]));
                cmd.Parameters.Add(new SqlParameter("@cxNum", double.Parse(diction["cxNum"])));
                cmd.Parameters.Add(new SqlParameter("@htAmount", double.Parse(diction["htAmount"])));
                cmd.Parameters.Add(new SqlParameter("@ysAmount", double.Parse(diction["ysAmount"])));
                cmd.Parameters.Add(new SqlParameter("@lrAmount", double.Parse(diction["lrAmount"])));
                cmd.Parameters.Add(new SqlParameter("@yield", double.Parse(diction["yield"])));
                cmd.Parameters.Add(new SqlParameter("@yieEffic", double.Parse(diction["yieEffic"])));
                cmd.Parameters.Add(new SqlParameter("@gjEffic", double.Parse(diction["gjEffic"])));
                cmd.Parameters.Add(new SqlParameter("@proTeams", double.Parse(diction["proTeams"])));
                cmd.Parameters.Add(new SqlParameter("@workDays", double.Parse(diction["workDays"])));
                cmd.Parameters.Add(new SqlParameter("@yearly", int.Parse(diction["yearly"])));
                cmd.Parameters.Add(new SqlParameter("@monthly", int.Parse(diction["monthly"])));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                int result = cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /*
         * 获取规划预算
         * 开始日期月度到结束日期月度
         */
        public DataTable GetHr_Midghys(DateTime start, DateTime end, int mid)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                int yearly1 = start.Year;
                int monthly1 = start.Month;

                int yearly2 = end.Year;
                int monthly2 = end.Month;


                string sql = @"select t1.id,t1.cxNum,t1.htAmount,t1.ysAmount,t1.lrAmount,t1.yield,t1.yieEffic,t1.gjEffic,t1.proTeams,t1.workDays,t1.yearly,t1.monthly,t1.mid,t1.addTime,
                        t2.easName comName from Hr_Midghys t1 left join Hr_Middle_sys t2 on t2.easCode=t1.easComCode and t2.mType='公司' where t1.mid=" + mid;

                if (yearly1 < yearly2)
                {
                    sql += " and ((t1.yearly=" + yearly1 + " and t1.monthly>=" + monthly1 + ") or (t1.yearly=" + yearly2 + " and t1.monthly<=" + monthly2 + "))";
                }
                else
                {
                    sql += " and t1.yearly=" + yearly1 + " and t1.monthly>=" + monthly1 + " and t1.monthly<=" + monthly2;
                }
                
                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        //删除 规划预算数据
        public void DelHr_MidghysById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "delete Hr_Midghys where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();
                
            }
            catch (Exception e)
            {
                throw e;
            }
        }




        /*
         * 获取 公司 月度规划预算
         * 
         */
        public DataTable GetHr_MidghysByMonth(int yearly, int monthly, string baisComCode)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select t1.id,t1.easComCode,t1.cxNum,t1.htAmount,t1.ysAmount,t1.lrAmount,t1.yield,t1.yieEffic,t1.gjEffic,t1.proTeams,t1.workDays,t1.yearly,t1.monthly,
                        t2.easName comName from Hr_Midghys t1 left join Hr_Middle_sys t2 on t2.easCode=t1.easComCode and t2.mType='公司' ";

                sql += " where t2.baisCode=" + baisComCode + " and t1.yearly=" + yearly + " and t1.monthly=" + monthly;


                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        

        /*
         * 添加 调整预算
         * 同一公司 同一年的同一月度 数据唯一
         */
        public int AddHr_Midtzys(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                string sql = @"if not exists(select id from Hr_Midtzys where easComCode=@easComCode and yearly=@yearly and monthly=@monthly)
                                    begin
                                        insert into Hr_Midtzys(easComCode,cxNum,htAmount,ysAmount,lrAmount,yield,yieEffic,gjEffic,proTeams,workDays,yearly,monthly,mid,addPer,addTime) 
                                        values(@easComCode,@cxNum,@htAmount,@ysAmount,@lrAmount,@yield,@yieEffic,@gjEffic,@proTeams,@workDays,@yearly,@monthly,@mid,@addPer,@addTime)
                                    end
                                else
                                    begin
                                        update Hr_Midtzys set cxNum=@cxNum,htAmount=@htAmount,ysAmount=@ysAmount,lrAmount=@lrAmount,yield=@yield,yieEffic=@yieEffic,gjEffic=@gjEffic,
                                        proTeams=@proTeams,workDays=@workDays,yearly=@yearly,monthly=@monthly,mid=@mid,addPer=@addPer,addTime=@addTime 
                                        where easComCode=@easComCode and yearly=@yearly and monthly=@monthly 
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@easComCode", diction["easComCode"]));
                cmd.Parameters.Add(new SqlParameter("@cxNum", double.Parse(diction["cxNum"])));
                cmd.Parameters.Add(new SqlParameter("@htAmount", double.Parse(diction["htAmount"])));
                cmd.Parameters.Add(new SqlParameter("@ysAmount", double.Parse(diction["ysAmount"])));
                cmd.Parameters.Add(new SqlParameter("@lrAmount", double.Parse(diction["lrAmount"])));
                cmd.Parameters.Add(new SqlParameter("@yield", double.Parse(diction["yield"])));
                cmd.Parameters.Add(new SqlParameter("@yieEffic", double.Parse(diction["yieEffic"])));
                cmd.Parameters.Add(new SqlParameter("@gjEffic", double.Parse(diction["gjEffic"])));
                cmd.Parameters.Add(new SqlParameter("@proTeams", double.Parse(diction["proTeams"])));
                cmd.Parameters.Add(new SqlParameter("@workDays", double.Parse(diction["workDays"])));
                cmd.Parameters.Add(new SqlParameter("@yearly", int.Parse(diction["yearly"])));
                cmd.Parameters.Add(new SqlParameter("@monthly", int.Parse(diction["monthly"])));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                int result = cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /*
         * 获取规划预算
         * 开始日期月度到结束日期月度
         */
        public DataTable GetHr_Midtzys(DateTime start, DateTime end, int mid)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                int yearly1 = start.Year;
                int monthly1 = start.Month;

                int yearly2 = end.Year;
                int monthly2 = end.Month;


                string sql = @"select t1.id,t1.cxNum,t1.htAmount,t1.ysAmount,t1.lrAmount,t1.yield,t1.yieEffic,t1.gjEffic,t1.proTeams,t1.workDays,t1.yearly,t1.monthly,t1.mid,t1.addTime,
                        t2.easName comName from Hr_Midtzys t1 left join Hr_Middle_sys t2 on t2.easCode=t1.easComCode and t2.mType='公司' where t1.mid=" + mid;

                if (yearly1 < yearly2)
                {
                    sql += " and ((t1.yearly=" + yearly1 + " and t1.monthly>=" + monthly1 + ") or (t1.yearly=" + yearly2 + " and t1.monthly<=" + monthly2 + "))";
                }
                else
                {
                    sql += " and t1.yearly=" + yearly1 + " and t1.monthly>=" + monthly1 + " and t1.monthly<=" + monthly2;
                }

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        //删除 调整预算数据
        public void DelHr_MidtzysById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "delete Hr_Midtzys where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();

            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
         * 修改 调整预算
         */
        public int EditHr_MidtzysById(int id, Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"update Hr_Midtzys set cxNum=@cxNum,htAmount=@htAmount,ysAmount=@ysAmount,lrAmount=@lrAmount,yield=@yield,
                            yieEffic=@yieEffic,gjEffic=@gjEffic,proTeams=@proTeams,workDays=@workDays,mid=@mid,addPer=@addPer,addTime=@addTime where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                cmd.Parameters.Add(new SqlParameter("@cxNum", double.Parse(diction["cxNum"])));
                cmd.Parameters.Add(new SqlParameter("@htAmount", double.Parse(diction["htAmount"])));
                cmd.Parameters.Add(new SqlParameter("@ysAmount", double.Parse(diction["ysAmount"])));
                cmd.Parameters.Add(new SqlParameter("@lrAmount", double.Parse(diction["lrAmount"])));
                cmd.Parameters.Add(new SqlParameter("@yield", double.Parse(diction["yield"])));
                cmd.Parameters.Add(new SqlParameter("@yieEffic", double.Parse(diction["yieEffic"])));
                cmd.Parameters.Add(new SqlParameter("@gjEffic", double.Parse(diction["gjEffic"])));
                cmd.Parameters.Add(new SqlParameter("@proTeams", double.Parse(diction["proTeams"])));
                cmd.Parameters.Add(new SqlParameter("@workDays", double.Parse(diction["workDays"])));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                int result = cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }



        /*
         * 添加 项目进展
         * 同一公司 同一年的同一月度 数据唯一
         */
        public int AddHr_Midxmjz(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                string sql = @"if not exists(select id from Hr_Midxmjz where easComCode=@easComCode and yearly=@yearly and monthly=@monthly)
                                    begin
                                        insert into Hr_Midxmjz(easComCode,cxNum,htAmount,ysAmount,lrAmount,proBudget,progjBudget,yieEffic,gjEffic,proTeams,workDays,yearly,monthly,mid,addPer,addTime) 
                                        values(@easComCode,@cxNum,@htAmount,@ysAmount,@lrAmount,@proBudget,@progjBudget,@yieEffic,@gjEffic,@proTeams,@workDays,@yearly,@monthly,@mid,@addPer,@addTime)
                                    end
                                else
                                    begin
                                        update Hr_Midxmjz set cxNum=@cxNum,htAmount=@htAmount,ysAmount=@ysAmount,lrAmount=@lrAmount,proBudget=@proBudget,progjBudget=@progjBudget,yieEffic=@yieEffic,
                                        gjEffic=@gjEffic,proTeams=@proTeams,workDays=@workDays,yearly=@yearly,monthly=@monthly,mid=@mid,addPer=@addPer,addTime=@addTime 
                                        where easComCode=@easComCode and yearly=@yearly and monthly=@monthly
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@easComCode", diction["easComCode"]));
                cmd.Parameters.Add(new SqlParameter("@cxNum", double.Parse(diction["cxNum"])));
                cmd.Parameters.Add(new SqlParameter("@htAmount", double.Parse(diction["htAmount"])));
                cmd.Parameters.Add(new SqlParameter("@ysAmount", double.Parse(diction["ysAmount"])));
                cmd.Parameters.Add(new SqlParameter("@lrAmount", double.Parse(diction["lrAmount"])));
                cmd.Parameters.Add(new SqlParameter("@proBudget", double.Parse(diction["proBudget"])));
                cmd.Parameters.Add(new SqlParameter("@progjBudget", double.Parse(diction["progjBudget"])));
                cmd.Parameters.Add(new SqlParameter("@yieEffic", double.Parse(diction["yieEffic"])));
                cmd.Parameters.Add(new SqlParameter("@gjEffic", double.Parse(diction["gjEffic"])));
                cmd.Parameters.Add(new SqlParameter("@proTeams", double.Parse(diction["proTeams"])));
                cmd.Parameters.Add(new SqlParameter("@workDays", double.Parse(diction["workDays"])));
                cmd.Parameters.Add(new SqlParameter("@yearly", int.Parse(diction["yearly"])));
                cmd.Parameters.Add(new SqlParameter("@monthly", int.Parse(diction["monthly"])));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                int result = cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        
        /*
         * 获取项目进展
         * 开始日期月度到结束日期月度
         */
        public DataTable GetHr_Midxmjz(DateTime start, DateTime end, int mid)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                int yearly1 = start.Year;
                int monthly1 = start.Month;

                int yearly2 = end.Year;
                int monthly2 = end.Month;


                string sql = @"select t1.id,t1.cxNum,t1.htAmount,t1.ysAmount,t1.lrAmount,t1.proBudget,t1.progjBudget,t1.yieEffic,t1.gjEffic,t1.proTeams,t1.workDays,t1.yearly,t1.monthly,
                        t1.mid,t1.addTime,t2.easName comName from Hr_Midxmjz t1 left join Hr_Middle_sys t2 on t2.easCode=t1.easComCode and t2.mType='公司' where t1.mid=" + mid;

                if (yearly1 < yearly2)
                {
                    sql += " and ((t1.yearly=" + yearly1 + " and t1.monthly>=" + monthly1 + ") or (t1.yearly=" + yearly2 + " and t1.monthly<=" + monthly2 + "))";
                }
                else
                {
                    sql += " and t1.yearly=" + yearly1 + " and t1.monthly>=" + monthly1 + " and t1.monthly<=" + monthly2;
                }

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        //删除 项目进展数据
        public void DelHr_MidxmjzById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "delete Hr_Midxmjz where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();

            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /*
         * 添加 项目进展
         * 同一公司 同一年的同一月度 数据唯一
         */
        public int AddHr_Midfact(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                string sql = @"if not exists(select id from Hr_Midfact where easComCode=@easComCode and yearly=@yearly and monthly=@monthly)
                                    begin
                                        insert into Hr_Midfact(easComCode,cxNum,htAmount,ysAmount,lrAmount,yield,yieEffic,gjEffic,proTeams,workDays,yearly,monthly,mid,addPer,addTime) 
                                        values(@easComCode,@cxNum,@htAmount,@ysAmount,@lrAmount,@yield,@yieEffic,@gjEffic,@proTeams,@workDays,@yearly,@monthly,@mid,@addPer,@addTime)
                                    end
                                else
                                    begin
                                        update Hr_Midfact set cxNum=@cxNum,htAmount=@htAmount,ysAmount=@ysAmount,lrAmount=@lrAmount,yield=@yield,yieEffic=@yieEffic,
                                        gjEffic=@gjEffic,proTeams=@proTeams,workDays=@workDays,yearly=@yearly,monthly=@monthly,mid=@mid,addPer=@addPer,addTime=@addTime 
                                        where easComCode=@easComCode and yearly=@yearly and monthly=@monthly 
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@easComCode", diction["easComCode"]));
                cmd.Parameters.Add(new SqlParameter("@cxNum", double.Parse(diction["cxNum"])));
                cmd.Parameters.Add(new SqlParameter("@htAmount", double.Parse(diction["htAmount"])));
                cmd.Parameters.Add(new SqlParameter("@ysAmount", double.Parse(diction["ysAmount"])));
                cmd.Parameters.Add(new SqlParameter("@lrAmount", double.Parse(diction["lrAmount"])));
                cmd.Parameters.Add(new SqlParameter("@yield", double.Parse(diction["yield"])));
                cmd.Parameters.Add(new SqlParameter("@yieEffic", double.Parse(diction["yieEffic"])));
                cmd.Parameters.Add(new SqlParameter("@gjEffic", double.Parse(diction["gjEffic"])));
                cmd.Parameters.Add(new SqlParameter("@proTeams", double.Parse(diction["proTeams"])));
                cmd.Parameters.Add(new SqlParameter("@workDays", double.Parse(diction["workDays"])));
                cmd.Parameters.Add(new SqlParameter("@yearly", int.Parse(diction["yearly"])));
                cmd.Parameters.Add(new SqlParameter("@monthly", int.Parse(diction["monthly"])));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                int result = cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
         * 获取实际
         * 开始日期月度到结束日期月度
         */
        public DataTable GetHr_Midfact(DateTime start, DateTime end, int mid)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                int yearly1 = start.Year;
                int monthly1 = start.Month;

                int yearly2 = end.Year;
                int monthly2 = end.Month;


                string sql = @"select t1.id,t1.cxNum,t1.htAmount,t1.ysAmount,t1.lrAmount,t1.yield,t1.yieEffic,t1.gjEffic,t1.proTeams,t1.workDays,t1.yearly,t1.monthly,
                        t1.mid,t1.addTime,t2.easName comName from Hr_Midfact t1 left join Hr_Middle_sys t2 on t2.easCode=t1.easComCode and t2.mType='公司' where t1.mid=" + mid;

                if (yearly1 < yearly2)
                {
                    sql += " and ((t1.yearly=" + yearly1 + " and t1.monthly>=" + monthly1 + ") or (t1.yearly=" + yearly2 + " and t1.monthly<=" + monthly2 + "))";
                }
                else
                {
                    sql += " and t1.yearly=" + yearly1 + " and t1.monthly>=" + monthly1 + " and t1.monthly<=" + monthly2;
                }

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        //删除实际数据
        public void DelHr_MidfactById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "delete Hr_Midfact where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();

            }
            catch (Exception e)
            {
                throw e;
            }
        }


        #endregion


        /*
         * 
         *  获取 公司 月度调整预算
         */
        public DataTable GetHr_MidtzysByMonth(int yearly, int monthly, string baisComCode)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select t1.id,t1.easComCode,t1.cxNum,t1.htAmount,t1.ysAmount,t1.lrAmount,t1.yield,t1.yieEffic,t1.gjEffic,t1.proTeams,t1.workDays,t1.yearly,t1.monthly,
                            t2.easName comName from Hr_Midtzys t1 left join Hr_Middle_sys t2 on t2.easCode=t1.easComCode and t2.mType='公司' ";

                sql += " where t2.baisCode=" + baisComCode + " and t1.yearly=" + yearly + " and t1.monthly=" + monthly;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
         * 
         *  获取当前公司 月度Bhr实际
         *  baisComCode 当前公司 bais公司编码
         *  转化到 岗位配备规则编码
         *  部门，岗位 取岗位配备规则编码
         */
        public DataTable GetHr_Bhr_factByMonth2(int yearly, int monthly, string baisComCode)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select t1.easComCode,t1.easComName,t1.easDeptCode,t1.easDeptName,t1.easPostCode,t1.easPostName,t1.postLevel,t1.postType,
                                t1.wage,t1.workDays,t1.yearly,t1.monthly,t3.ruleCode ruleDeptCode,t4.ruleCode rulePostCode from Hr_Bhr_fact t1 
                                left join Hr_Middle_sys t2 on t2.easCode=t1.easComCode and t2.mType='公司'
                                left join Hr_Middle_rule t3 on t3.easCode=t1.easDeptCode and t3.mType='部门'
                                left join Hr_Middle_rule t4 on t3.easCode=t1.easPostCode and t4.mType='岗位' ";

                sql += " where t2.baisCode=" + baisComCode + " and t1.yearly=" + yearly;

                if (monthly != 0)
                {
                    sql += " and t1.monthly=" + monthly;
                }

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
         * 
         *  获取当前公司 月度Crm调整预算
         *  baisComCode 当前公司 bais公司编码
         *  转化到 岗位配备规则编码
         *  部门，岗位 取岗位配备规则编码
         */
        public DataTable GetHr_Crm_tzysByMonth(int yearly, int monthly, string baisComCode)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select t1.crmComCode,t1.crmDeptCode,t1.htAmount,t1.syAmount,t1.ghsyAmount,t1.tzsyAmount,t1.yearly,t1.monthly,
                            t2.easCode easComCode,t3.easCode easDeptCode,t4.ruleCode ruleDeptCode from Hr_Crm_tzys t1 
                            left join Hr_Middle_sys t2 on t2.crmCode=t1.crmComCode and t2.mType='公司'
                            left join Hr_Middle_sys t3 on t3.crmCode=t1.crmDeptCode and t3.mType='部门'
                            left join Hr_Middle_rule t4 on t4.easCode=t3.easCode and t4.mType='部门' ";
                sql += " where t2.baisCode=" + baisComCode + " and t1.yearly=" + yearly + " and t1.monthly=" + monthly;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        

        #region //（非）市场人数

        /*
         * 添加 非市场人数
         * 同一公司 同一年的同一月度 数据唯一
         */
        public int AddHr_Midysrs(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                string sql = @"if not exists(select id from Hr_Midysrs where easComCode=@easComCode and ruleDeptCode=@ruleDeptCode and rulePostCode=@rulePostCode and yearly=@yearly and monthly=@monthly)
                                    begin
                                        insert into Hr_Midysrs(easComCode,ruleDeptCode,rulePostCode,coreQuota,coreActual,boneQuota,boneActual,floatQuota,floatActual,floatFore,floatghys,floattzys,
                                        yearly,monthly,mid,addPer,addTime) 
                                        values(@easComCode,@ruleDeptCode,@rulePostCode,@coreQuota,@coreActual,@boneQuota,@boneActual,@floatQuota,@floatActual,@floatFore,@floatghys,@floattzys,
                                        @yearly,@monthly,@mid,@addPer,@addTime)
                                    end
                                else
                                    begin
                                        update Hr_Midysrs set coreQuota=@coreQuota,coreActual=@coreActual,boneQuota=@boneQuota,boneActual=@boneActual,floatQuota=@floatQuota,
                                        floatActual=@floatActual,floatFore=@floatFore,floatghys=@floatghys,floattzys=@floattzys,mid=@mid,addPer=@addPer,addTime=@addTime 
                                        where easComCode=@easComCode and ruleDeptCode=@ruleDeptCode and rulePostCode=@rulePostCode and yearly=@yearly and monthly=@monthly 
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@easComCode", diction["easComCode"]));
                cmd.Parameters.Add(new SqlParameter("@ruleDeptCode", diction["ruleDeptCode"]));
                cmd.Parameters.Add(new SqlParameter("@rulePostCode", diction["rulePostCode"]));
                cmd.Parameters.Add(new SqlParameter("@coreQuota", int.Parse(diction["coreQuota"])));
                cmd.Parameters.Add(new SqlParameter("@coreActual", int.Parse(diction["coreActual"])));
                cmd.Parameters.Add(new SqlParameter("@boneQuota", int.Parse(diction["boneQuota"])));
                cmd.Parameters.Add(new SqlParameter("@boneActual", int.Parse(diction["boneActual"])));
                cmd.Parameters.Add(new SqlParameter("@floatQuota", int.Parse(diction["floatQuota"])));
                cmd.Parameters.Add(new SqlParameter("@floatActual", int.Parse(diction["floatActual"])));
                cmd.Parameters.Add(new SqlParameter("@floatFore", int.Parse(diction["floatFore"])));
                cmd.Parameters.Add(new SqlParameter("@floatghys", int.Parse(diction["floatghys"])));
                cmd.Parameters.Add(new SqlParameter("@floattzys", int.Parse(diction["floattzys"])));
                cmd.Parameters.Add(new SqlParameter("@yearly", int.Parse(diction["yearly"])));
                cmd.Parameters.Add(new SqlParameter("@monthly", int.Parse(diction["monthly"])));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                int result = cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        
        /*
         * 获取 非市场人数
         * 开始日期月度到结束日期月度
         */
        public DataTable GetHr_Midysrs(DateTime start, DateTime end, int mid)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                int yearly1 = start.Year;
                int monthly1 = start.Month;

                int yearly2 = end.Year;
                int monthly2 = end.Month;


                string sql = @"select t1.id,t1.easComCode,t1.ruleDeptCode,t1.rulePostCode,t1.coreQuota,t1.coreActual,t1.boneQuota,t1.boneActual,t1.floatQuota,t1.floatActual,
                            t1.floatFore,t1.floatghys,t1.floattzys,t1.yearly,t1.monthly,t1.mid,t2.easName comName,t3.ruleName deptName,t4.ruleName postName,t5.postLevel
                            from Hr_Midysrs t1 left join Hr_Middle_sys t2 on t2.easCode=t1.easComCode and t2.mType='公司'
                            left join Hr_Middle_rule t3 on t3.ruleCode=t1.ruleDeptCode and t3.mType='部门'
                            left join Hr_Middle_rule t4 on t4.ruleCode=t1.rulePostCode and t4.mType='岗位' 
                            left join Hr_Rule_ysgw t5 on t5.postCode=t1.rulePostCode  where t1.mid=" + mid;

                if (yearly1 < yearly2)
                {
                    sql += " and ((t1.yearly=" + yearly1 + " and t1.monthly>=" + monthly1 + ") or (t1.yearly=" + yearly2 + " and t1.monthly<=" + monthly2 + "))";
                }
                else
                {
                    sql += " and t1.yearly=" + yearly1 + " and t1.monthly>=" + monthly1 + " and t1.monthly<=" + monthly2;
                }

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //删除 非市场人数 数据
        public void DelHr_MidysrsById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "delete Hr_Midysrs where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();

            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /*
         * 添加 市场人数
         * 同一公司 同一年的同一月度 数据唯一
         */
        public int AddHr_Midhtrs(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                string sql = @"if not exists(select id from Hr_Midhtrs where easComCode=@easComCode and ruleDeptCode=@ruleDeptCode and rulePostCode=@rulePostCode and yearly=@yearly and monthly=@monthly)
                                    begin
                                        insert into Hr_Midhtrs(easComCode,ruleDeptCode,rulePostCode,core_ghys,core_tzys,bone_ghys,bone_tzys,yearly,monthly,mid,addPer,addTime) 
                                        values(@easComCode,@ruleDeptCode,@rulePostCode,@core_ghys,@core_tzys,@bone_ghys,@bone_tzys,@yearly,@monthly,@mid,@addPer,@addTime)
                                    end
                                else
                                    begin
                                        update Hr_Midhtrs set core_ghys=@core_ghys,core_tzys=@core_tzys,bone_ghys=@bone_ghys,bone_tzys=@bone_tzys,mid=@mid,addPer=@addPer,addTime=@addTime 
                                        where easComCode=@easComCode and ruleDeptCode=@ruleDeptCode and rulePostCode=@rulePostCode and yearly=@yearly and monthly=@monthly 
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@easComCode", diction["easComCode"]));
                cmd.Parameters.Add(new SqlParameter("@ruleDeptCode", diction["ruleDeptCode"]));
                cmd.Parameters.Add(new SqlParameter("@rulePostCode", diction["rulePostCode"]));
                cmd.Parameters.Add(new SqlParameter("@core_ghys", int.Parse(diction["core_ghys"])));
                cmd.Parameters.Add(new SqlParameter("@core_tzys", int.Parse(diction["core_tzys"])));
                cmd.Parameters.Add(new SqlParameter("@bone_ghys", int.Parse(diction["bone_ghys"])));
                cmd.Parameters.Add(new SqlParameter("@bone_tzys", int.Parse(diction["bone_tzys"])));
                cmd.Parameters.Add(new SqlParameter("@yearly", int.Parse(diction["yearly"])));
                cmd.Parameters.Add(new SqlParameter("@monthly", int.Parse(diction["monthly"])));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                int result = cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
         * 获取 市场人数
         * 开始日期月度到结束日期月度
         */
        public DataTable GetHr_Midhtrs(DateTime start, DateTime end, int mid)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                int yearly1 = start.Year;
                int monthly1 = start.Month;

                int yearly2 = end.Year;
                int monthly2 = end.Month;


                string sql = @"select t1.id,t1.easComCode,t1.ruleDeptCode,t1.rulePostCode,t1.core_ghys,t1.core_tzys,t1.bone_ghys,t1.bone_tzys,
                            t1.yearly,t1.monthly,t1.mid,t2.easName comName,t3.ruleName deptName,t4.ruleName postName from Hr_Midhtrs t1 
                            left join Hr_Middle_sys t2 on t2.easCode=t1.easComCode and t2.mType='公司'
                            left join Hr_Middle_rule t3 on t3.ruleCode=t1.ruleDeptCode and t3.mType='部门'
                            left join Hr_Middle_rule t4 on t4.ruleCode=t1.rulePostCode and t4.mType='岗位' where t1.mid=" + mid;

                if (yearly1 < yearly2)
                {
                    sql += " and ((t1.yearly=" + yearly1 + " and t1.monthly>=" + monthly1 + ") or (t1.yearly=" + yearly2 + " and t1.monthly<=" + monthly2 + "))";
                }
                else
                {
                    sql += " and t1.yearly=" + yearly1 + " and t1.monthly>=" + monthly1 + " and t1.monthly<=" + monthly2;
                }

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //删除 市场人数 数据
        public void DelHr_MidhtrsById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "delete Hr_Midhtrs where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();

            }
            catch (Exception e)
            {
                throw e;
            }
        }



        #endregion


        /*
         *  非市场人数
         *  baisComCode 当前公司 bais公司编码
         */
        public DataTable GetHr_MidysrsByMonth(int yearly, int monthly, string baisComCode)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select t1.easComCode,t1.ruleDeptCode,t1.rulePostCode,t1.coreQuota,t1.boneQuota,t1.floatQuota,
                            t1.coreActual,t1.boneActual,t1.floatActual,t1.floatghys,t1.floattzys from Hr_Midysrs t1
                            left join Hr_Middle_sys t2 on t2.easCode=t1.easComCode and t2.mType='公司' ";
                sql += " where t2.baisCode=" + baisComCode + " and t1.yearly=" + yearly + " and t1.monthly=" + monthly;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
         *  市场人数
         *  baisComCode 当前公司 bais公司编码
         */
        public DataTable GetHr_MidhtrsByMonth(int yearly, int monthly, string baisComCode)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select t1.easComCode,t1.ruleDeptCode,t1.rulePostCode,t1.core_ghys,t1.bone_ghys,t1.core_tzys,t1.bone_tzys from Hr_Midhtrs t1
                            left join Hr_Middle_sys t2 on t2.easCode=t1.easComCode and t2.mType='公司' ";
                sql += " where t2.baisCode=" + baisComCode + " and t1.yearly=" + yearly + " and t1.monthly=" + monthly;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        #region //人工费

        /*
         * 
         *  获取当前公司 月度 bais 人工收入
         *  baisComCode 当前公司 bais公司编码
         */
        public DataTable GetHr_Bais_rgsrByMonth(int yearly, int monthly, string baisComCode)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select t1.baisComCode,t1.costType,t1.planBudget,t1.adjustBudget,t1.proBudget,t1.quotaLabor,t1.proportion,t1.yearly,t1.monthly,t2.easCode from Hr_Bais_rgsr t1 
                            left join Hr_Middle_sys t2 on t2.baisCode=t1.baisComCode and t2.mType='公司' ";
                sql += " where t2.baisCode=" + baisComCode + " and t1.yearly=" + yearly + " and t1.monthly=" + monthly; 

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /*
         * 
         *  获取当前公司 月度 bais 人工收入
         *  baisComCode 当前公司 bais公司编码
         */
        public DataTable GetHr_Bais_rgzcByMonth(int yearly, int monthly, string baisComCode)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = @"select t1.baisComCode,t1.costType,t1.quotaLabor,t1.yearly,t1.monthly,t2.easCode from Hr_Bais_rgzc t1 
                            left join Hr_Middle_sys t2 on t2.baisCode=t1.baisComCode and t2.mType='公司' ";
                sql += " where t2.baisCode=" + baisComCode + " and t1.yearly=" + yearly + " and t1.monthly=" + monthly;

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        


        /*
         * 添加 人工费-收入
         * 同一公司 同一年的同一月度 数据唯一
         */
        public int AddHr_Midrgsr(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                string sql = @"if not exists(select id from Hr_Midrgsr where easComCode=@easComCode and yearly=@yearly and monthly=@monthly)
                                    begin
                                        insert into Hr_Midrgsr(easComCode,costType,planBudget,adjustBudget,proBudget,quotaLabor,proportion,yearly,monthly,mid,addPer,addTime) 
                                        values(@easComCode,@costType,@planBudget,@adjustBudget,@proBudget,@quotaLabor,@proportion,@yearly,@monthly,@mid,@addPer,@addTime)
                                    end
                                else
                                    begin
                                        update Hr_Midrgsr set costType=@costType,planBudget=@planBudget,adjustBudget=@adjustBudget,proBudget=@proBudget,quotaLabor=@quotaLabor,
                                        proportion=@proportion,yearly=@yearly,monthly=@monthly,mid=@mid,addPer=@addPer,addTime=@addTime 
                                        where easComCode=@easComCode and yearly=@yearly and monthly=@monthly
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@easComCode", diction["easComCode"]));
                cmd.Parameters.Add(new SqlParameter("@costType", diction["costType"]));
                cmd.Parameters.Add(new SqlParameter("@planBudget", double.Parse(diction["planBudget"])));
                cmd.Parameters.Add(new SqlParameter("@adjustBudget", double.Parse(diction["adjustBudget"])));
                cmd.Parameters.Add(new SqlParameter("@proBudget", double.Parse(diction["proBudget"])));
                cmd.Parameters.Add(new SqlParameter("@quotaLabor", double.Parse(diction["quotaLabor"])));
                cmd.Parameters.Add(new SqlParameter("@proportion", double.Parse(diction["proportion"])));
                cmd.Parameters.Add(new SqlParameter("@yearly", int.Parse(diction["yearly"])));
                cmd.Parameters.Add(new SqlParameter("@monthly", int.Parse(diction["monthly"])));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                int result = cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        
        /*
        * 获取人工费-收入
        * 开始日期月度到结束日期月度
        */
        public DataTable GetHr_Midrgsr(DateTime start, DateTime end, int mid)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                int yearly1 = start.Year;
                int monthly1 = start.Month;

                int yearly2 = end.Year;
                int monthly2 = end.Month;


                string sql = @"select t1.id,t1.costType,t1.planBudget,t1.adjustBudget,t1.proBudget,t1.quotaLabor,t1.proportion,t1.yearly,t1.monthly,
                        t1.mid,t1.addTime,t2.easName comName from Hr_Midrgsr t1 left join Hr_Middle_sys t2 on t2.easCode=t1.easComCode and t2.mType='公司' where t1.mid=" + mid;

                if (yearly1 < yearly2)
                {
                    sql += " and ((t1.yearly=" + yearly1 + " and t1.monthly>=" + monthly1 + ") or (t1.yearly=" + yearly2 + " and t1.monthly<=" + monthly2 + "))";
                }
                else
                {
                    sql += " and t1.yearly=" + yearly1 + " and t1.monthly>=" + monthly1 + " and t1.monthly<=" + monthly2;
                }

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        
        //删除人工费-收入数据
        public void DelHr_MidrgsrById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "delete Hr_Midrgsr where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();

            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /*
         * 添加 人工费-支出
         * 同一公司 同一年的同一月度 数据唯一
         */
        public int AddHr_Midrgzc(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                string sql = @"if not exists(select id from Hr_Midrgzc where easComCode=@easComCode and ruleDeptCode=@ruleDeptCode and rulePostCode=@rulePostCode and yearly=@yearly and monthly=@monthly)
                                    begin
                                        insert into Hr_Midrgzc(easComCode,ruleDeptCode,rulePostCode,costType,planBudget,adjustBudget,proBudget,yearly,monthly,mid,addPer,addTime) 
                                        values(@easComCode,@ruleDeptCode,@rulePostCode,@costType,@planBudget,@adjustBudget,@proBudget,@yearly,@monthly,@mid,@addPer,@addTime)
                                    end
                                else
                                    begin
                                        update Hr_Midrgzc set costType=@costType,planBudget=@planBudget,adjustBudget=@adjustBudget,proBudget=@proBudget,
                                        yearly=@yearly,monthly=@monthly,mid=@mid,addPer=@addPer,addTime=@addTime 
                                        where easComCode=@easComCode and ruleDeptCode=@ruleDeptCode and rulePostCode=@rulePostCode and yearly=@yearly and monthly=@monthly
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@easComCode", diction["easComCode"]));
                cmd.Parameters.Add(new SqlParameter("@ruleDeptCode", diction["ruleDeptCode"]));
                cmd.Parameters.Add(new SqlParameter("@rulePostCode", diction["rulePostCode"]));
                cmd.Parameters.Add(new SqlParameter("@costType", diction["costType"]));
                cmd.Parameters.Add(new SqlParameter("@planBudget", double.Parse(diction["planBudget"])));
                cmd.Parameters.Add(new SqlParameter("@adjustBudget", double.Parse(diction["adjustBudget"])));
                cmd.Parameters.Add(new SqlParameter("@proBudget", double.Parse(diction["proBudget"])));
                cmd.Parameters.Add(new SqlParameter("@yearly", int.Parse(diction["yearly"])));
                cmd.Parameters.Add(new SqlParameter("@monthly", int.Parse(diction["monthly"])));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                int result = cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /*
        * 获取人工费-支出
        * 开始日期月度到结束日期月度
        */
        public DataTable GetHr_Midrgzc(DateTime start, DateTime end, int mid)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                int yearly1 = start.Year;
                int monthly1 = start.Month;

                int yearly2 = end.Year;
                int monthly2 = end.Month;


                string sql = @"select t1.id,t1.costType,t1.planBudget,t1.adjustBudget,t1.proBudget,t1.yearly,t1.monthly,t1.mid,t1.addTime,
                            t2.easName comName,t3.ruleName deptName,t4.ruleName postName from Hr_Midrgzc t1 
                            left join Hr_Middle_sys t2 on t2.easCode=t1.easComCode and t2.mType='公司' 
                            left join Hr_Middle_rule t3 on t3.ruleCode=t1.ruleDeptCode and t3.mType='部门'
                            left join Hr_Middle_rule t4 on t4.ruleCode=t1.rulePostCode and t4.mType='岗位'
                            where t1.mid=" + mid;

                if (yearly1 < yearly2)
                {
                    sql += " and ((t1.yearly=" + yearly1 + " and t1.monthly>=" + monthly1 + ") or (t1.yearly=" + yearly2 + " and t1.monthly<=" + monthly2 + "))";
                }
                else
                {
                    sql += " and t1.yearly=" + yearly1 + " and t1.monthly>=" + monthly1 + " and t1.monthly<=" + monthly2;
                }

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        //删除人工费-支出数据
        public void DelHr_MidrgzcById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "delete Hr_Midrgzc where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();

            }
            catch (Exception e)
            {
                throw e;
            }
        }

        

        /*
         * 添加 人工费-支出实际
         * 同一公司 同一年的同一月度 数据唯一
         */
        public int AddHr_Midrgsj(Dictionary<string, string> diction)
        {
            try
            {
                SqlConnection sqlCon = SqlCon2();
                sqlCon.Open();

                string sql = @"if not exists(select id from Hr_Midrgsj where easComCode=@easComCode and yearly=@yearly and monthly=@monthly)
                                    begin
                                        insert into Hr_Midrgsj(easComCode,costType,quotaLabor,yearly,monthly,mid,addPer,addTime) 
                                        values(@easComCode,@costType,@quotaLabor,@yearly,@monthly,@mid,@addPer,@addTime)
                                    end
                                else
                                    begin
                                        update Hr_Midrgsj set costType=@costType,quotaLabor=@quotaLabor,yearly=@yearly,monthly=@monthly,
                                        mid=@mid,addPer=@addPer,addTime=@addTime 
                                        where easComCode=@easComCode and yearly=@yearly and monthly=@monthly
                                    end";

                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@easComCode", diction["easComCode"]));
                cmd.Parameters.Add(new SqlParameter("@costType", diction["costType"]));
                cmd.Parameters.Add(new SqlParameter("@quotaLabor", double.Parse(diction["quotaLabor"])));
                cmd.Parameters.Add(new SqlParameter("@yearly", int.Parse(diction["yearly"])));
                cmd.Parameters.Add(new SqlParameter("@monthly", int.Parse(diction["monthly"])));
                cmd.Parameters.Add(new SqlParameter("@mid", int.Parse(diction["mid"])));
                cmd.Parameters.Add(new SqlParameter("@addPer", diction["addPer"]));
                cmd.Parameters.Add(new SqlParameter("@addTime", DateTime.Now));

                int result = cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlCon.Close();

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        
        /*
        * 获取人工费-支出实际
        * 开始日期月度到结束日期月度
        */
        public DataTable GetHr_Midrgsj(DateTime start, DateTime end, int mid)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                int yearly1 = start.Year;
                int monthly1 = start.Month;

                int yearly2 = end.Year;
                int monthly2 = end.Month;


                string sql = @"select t1.id,t1.costType,t1.quotaLabor,t1.yearly,t1.monthly,t1.mid,t1.addTime,t2.easName comName from Hr_Midrgsj t1 
                            left join Hr_Middle_sys t2 on t2.easCode=t1.easComCode and t2.mType='公司' where t1.mid=" + mid;

                if (yearly1 < yearly2)
                {
                    sql += " and ((t1.yearly=" + yearly1 + " and t1.monthly>=" + monthly1 + ") or (t1.yearly=" + yearly2 + " and t1.monthly<=" + monthly2 + "))";
                }
                else
                {
                    sql += " and t1.yearly=" + yearly1 + " and t1.monthly>=" + monthly1 + " and t1.monthly<=" + monthly2;
                }

                SqlDataAdapter myda = new SqlDataAdapter(sql, con); // 实例化适配器
                DataTable dt = new DataTable(); // 实例化数据表
                myda.Fill(dt); // 保存数据 

                myda.Dispose();
                con.Close();

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        
        //删除人工费-支出实际数据
        public void DelHr_MidrgsjById(int id)
        {
            try
            {
                SqlConnection con = SqlCon2();
                con.Open();

                string sql = "delete Hr_Midrgsj where id=@id";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.Add(new SqlParameter("@id", id));
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();

            }
            catch (Exception e)
            {
                throw e;
            }
        }

        #endregion

    }
}
