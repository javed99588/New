using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using MVC_CRUD.Models;
using ClosedXML.Excel;
using System.Web.UI.WebControls;
using System.IO;
using System.Web.UI;
using OfficeOpenXml;

namespace MVC_CRUD.Controllers
{
    public class EmployeeController : Controller
    {
        string connStr = ConfigurationManager.ConnectionStrings["myConnectionString"].ConnectionString;

        public IEnumerable<Employee> GetAllEmployees()
        {
            List<Employee> lstemployee = new List<Employee>();
            using (SqlConnection con = new SqlConnection(connStr))
            {
                SqlCommand cmd = new SqlCommand("select * from Employee", con);
                con.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    Employee employee = new Employee();

                    employee.EmpCode = Convert.ToInt32(rdr["EmpCode"]);
                    employee.Name = rdr["Name"].ToString();
                    employee.Address= rdr["Address"].ToString();
                    employee.Designation = rdr["Designation"].ToString();
                    employee.Gender= rdr["Gender"].ToString();
                    lstemployee.Add(employee);
                }
                con.Close();
            }
            return lstemployee;
        }

      

        [HttpGet]
        public ActionResult Index()
        {

                return View(GetAllEmployees());
        }

       

        // GET: Employee/Create
        public ActionResult Create()
        {
            var list = new List<String>() { "Manager", "Assistant", "Tester", "Software Engineer" };
            ViewBag.list = list;
            return View(new Employee());
        }

        // POST: Employee/Create
        [HttpPost]
        public ActionResult Create(Employee emp)
        {
            using (SqlConnection sqlCon = new SqlConnection(connStr))
            {
                sqlCon.Open();
                string query = "insert into Employee values(@Name,@Address,@Designation,@Gender)";
                SqlCommand cmd = new SqlCommand(query, sqlCon);
                cmd.Parameters.AddWithValue("@Name", emp.Name);
                cmd.Parameters.AddWithValue("@Address", emp.Address);
                cmd.Parameters.AddWithValue("@Designation", emp.Designation);
                cmd.Parameters.AddWithValue("@Gender", emp.Gender);
                cmd.ExecuteNonQuery();

            }
            return RedirectToAction("Index");
        }

        // GET: Employee/Edit/5
        public ActionResult Edit(int id)
        {
            var list = new List<String>() { "Manager", "Assistant", "Tester", "Software Engineer" };
            ViewBag.list = list;

            Employee emp = new Employee();

            DataTable dt = new DataTable();
            using (SqlConnection sqlCon = new SqlConnection(connStr))
            {
                sqlCon.Open();
                string query = "select * from Employee where EmpCode=@EmpCode";
                SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
                sqlDa.SelectCommand.Parameters.AddWithValue("@EmpCode", id);
                sqlDa.Fill(dt);
            }
            if(dt.Rows.Count==1)
            {
                emp.EmpCode = Convert.ToInt32(dt.Rows[0][0].ToString());
                emp.Name = (dt.Rows[0][1].ToString());
                emp.Address = (dt.Rows[0][2].ToString());
                emp.Designation = (dt.Rows[0][3].ToString());
                emp.Gender = (dt.Rows[0][4].ToString());
                return View(emp);

            }
            else
            {
                return RedirectToAction("Index");
            }
        }

        // POST: Employee/Edit/5
        [HttpPost]
        public ActionResult Edit(Employee emp)
        {
            using (SqlConnection sqlCon = new SqlConnection(connStr))
            {
                sqlCon.Open();
                string query = "update Employee set Name=@Name, Address=@Address, Designation=@Designation, Gender=@Gender where EmpCode=@EmpCode";
                SqlCommand cmd = new SqlCommand(query, sqlCon);
                cmd.Parameters.AddWithValue("@EmpCode", emp.EmpCode);
                cmd.Parameters.AddWithValue("@Name", emp.Name);
                cmd.Parameters.AddWithValue("@Address", emp.Address);
                cmd.Parameters.AddWithValue("@Designation", emp.Designation);
                cmd.Parameters.AddWithValue("@Gender", emp.Gender);
                cmd.ExecuteNonQuery();

            }
            return RedirectToAction("Index");
        }

        // GET: Employee/Delete/5
        public ActionResult Delete(int id)
        {
            using (SqlConnection sqlCon = new SqlConnection(connStr))
            {
                sqlCon.Open();
                string query = "DELETE from Employee where EmpCode=@EmpCode";
                SqlCommand cmd = new SqlCommand(query, sqlCon);
                cmd.Parameters.AddWithValue("@EmpCode", id);
                
                cmd.ExecuteNonQuery();

            }
            return RedirectToAction("Index");
        }

        /// <summary>
        /// option 1 to export the data in to excel
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportToExcel()
        {
            var gv = new GridView();
            gv.DataSource = GetAllEmployees();
            gv.DataBind();
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=DemoExcel.xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";
            StringWriter objStringWriter = new StringWriter();
            HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
            gv.RenderControl(objHtmlTextWriter);
            Response.Output.Write(objStringWriter.ToString());
            Response.Flush();
            Response.End();
            return View();
        }
        /// <summary>
        /// optionn2 to export rthe data into excel
        /// </summary>
        public void ExportListUsingEPPlus()
        {

            var data = GetAllEmployees();
            ExcelPackage excel = new ExcelPackage();
            var workSheet = excel.Workbook.Worksheets.Add("Sheet1");
            workSheet.Cells[1, 1].LoadFromCollection(data, true);
            using (var memoryStream = new MemoryStream())
            {
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                
                Response.AddHeader("content-disposition", "attachment;  filename=Employee.xlsx");
                excel.SaveAs(memoryStream);
                memoryStream.WriteTo(Response.OutputStream);
                Response.Flush();
                Response.End();
            }
        }

    }
}
