using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using HanyAddressBookGet.Models;
using System.IO;
using System.Web.UI;

namespace HanyAddressBookGet.Controllers
{
    public class AddressbookController : Controller
    {
        private AddressbookDBEntities db = new AddressbookDBEntities();

        // GET: /Addressbook/
        public ActionResult Index()
        {
            var employee = db.Employee.Include(e => e.Department).Include(e => e.Title);
            return View(employee.ToList());
        }

        // GET: /Addressbook/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Employee employee = db.Employee.Find(id);
            if (employee == null)
            {
                return HttpNotFound();
            }
            return View(employee);
        }

        // GET: /Addressbook/Create
        public ActionResult Create()
        {
            ViewBag.DeptID = new SelectList(db.Department, "DeptID", "DeptName");
            ViewBag.TitleID = new SelectList(db.Title, "TitleID", "TitleName");
            return View();
        }

        // POST: /Addressbook/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "EmpID,EmpName,TitleID,DeptID,EmpMobile,EmpHomeTel,EmpDateofBirth,EmpAddress,EmpEmail")] Employee employee, HttpPostedFileBase upload)
        {
            if (ModelState.IsValid)
            {
                if (upload != null && upload.ContentLength > 0)

                    using (var reader = new System.IO.BinaryReader(upload.InputStream))
                    {
                        employee.EmpPhoto = reader.ReadBytes(upload.ContentLength);
                    }
                db.Employee.Add(employee);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.DeptID = new SelectList(db.Department, "DeptID", "DeptName", employee.DeptID);
            ViewBag.TitleID = new SelectList(db.Title, "TitleID", "TitleName", employee.TitleID);
            return View(employee);
        }

        // GET: /Addressbook/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Employee employee = db.Employee.Find(id);
            if (employee == null)
            {
                return HttpNotFound();
            }
            ViewBag.DeptID = new SelectList(db.Department, "DeptID", "DeptName", employee.DeptID);
            ViewBag.TitleID = new SelectList(db.Title, "TitleID", "TitleName", employee.TitleID);
            return View(employee);
        }

        // POST: /Addressbook/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "EmpID,EmpName,TitleID,DeptID,EmpMobile,EmpHomeTel,EmpDateofBirth,EmpAddress,EmpEmail")] Employee employee, HttpPostedFileBase upload)
        {
            if (ModelState.IsValid)
            {
                if (upload != null && upload.ContentLength > 0)

                    using (var reader = new System.IO.BinaryReader(upload.InputStream))
                    {
                        employee.EmpPhoto = reader.ReadBytes(upload.ContentLength);
                    }
                db.Entry(employee).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.DeptID = new SelectList(db.Department, "DeptID", "DeptName", employee.DeptID);
            ViewBag.TitleID = new SelectList(db.Title, "TitleID", "TitleName", employee.TitleID);
            return View(employee);
        }

        // GET: /Addressbook/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Employee employee = db.Employee.Find(id);
            if (employee == null)
            {
                return HttpNotFound();
            }
            return View(employee);
        }

        // POST: /Addressbook/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Employee employee = db.Employee.Find(id);
            db.Employee.Remove(employee);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        //Export To Excel Method
        public void ExportClientsListToExcel()
        {
            var grid = new System.Web.UI.WebControls.GridView();

            var query = from d in db.Employee
                        select d;
            grid.DataSource = query.ToList();


            grid.DataBind();

            Response.ClearContent();
            Response.AddHeader("content-disposition", "attachment; filename=Exported_Addressbook.xls");
            Response.ContentType = "application/excel";
            StringWriter sw = new StringWriter();
            HtmlTextWriter htw = new HtmlTextWriter(sw);

            grid.RenderControl(htw);

            Response.Write(sw.ToString());

            Response.End();

        }

        //Show Image
        public ActionResult showImg(int id)
        {
            var image = (from m in db.Employee
                         where m.EmpID == id
                         select m.EmpPhoto).FirstOrDefault();

            var stream = new MemoryStream(image.ToArray());

            return new FileStreamResult(stream, "image/jpeg");
        }
    }
}
