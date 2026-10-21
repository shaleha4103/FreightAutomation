using System;
using System.Web.Mvc;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;

namespace FreightAutomation.Controllers
{
    public class AdminLoginController : Controller
    {
      
        [HttpGet]
        public ActionResult Login()
        {
            return View();
        }


        // ==========================================
        // LOGIN VALIDATION
        // ==========================================

        [HttpPost]
        public ActionResult LoginInput(string username, string password)
        {
            string connectionString =
                ConfigurationManager
                .ConnectionStrings["CreateCon"]
                .ConnectionString;

            using (SqlConnection con =
                new SqlConnection(connectionString))
            {
                using (SqlCommand cmd =
                    new SqlCommand("UserLogin_Checking", con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.Add(
                        "@username",
                        SqlDbType.NVarChar, 20).Value =
                        username.Trim();

                    cmd.Parameters.Add(
                        "@password",
                        SqlDbType.NVarChar, 20).Value =
                        password;

                    con.Open();

                    int count =
                        Convert.ToInt32(cmd.ExecuteScalar());

                    if (count == 1)
                    {
                        Session["Username"] = username.Trim();

                        return RedirectToAction(
                            "Upload",
                            "FreightAuto");
                    }

                    ViewBag.Error =
                        "Invalid username or password.";

                    return View("Login");
                }
            }
        }


        // ==========================================
        // LOGOUT
        // ==========================================

        public ActionResult Logout()
        {
            Session.Clear();
            Session.Abandon();

            return RedirectToAction(
                "Login",
                "AdminLogin");
        }
    }
}