using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using HanyAddressBookGet.Models;

namespace HanyAddressBookGet.Controllers
{
    public class TitlesController : Controller
    {
        private AddressbookDBEntities db = new AddressbookDBEntities();

        // GET: /Titles/
        public ActionResult Index()
        {
            return View(db.Title.ToList());
        }

        // GET: /Titles/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Title title = db.Title.Find(id);
            if (title == null)
            {
                return HttpNotFound();
            }
            return View(title);
        }

        // GET: /Titles/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: /Titles/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include="TitleID,TitleName")] Title title)
        {
            if (ModelState.IsValid)
            {
                db.Title.Add(title);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(title);
        }

        // GET: /Titles/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Title title = db.Title.Find(id);
            if (title == null)
            {
                return HttpNotFound();
            }
            return View(title);
        }

        // POST: /Titles/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include="TitleID,TitleName")] Title title)
        {
            if (ModelState.IsValid)
            {
                db.Entry(title).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(title);
        }

        // GET: /Titles/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Title title = db.Title.Find(id);
            if (title == null)
            {
                return HttpNotFound();
            }
            return View(title);
        }

        // POST: /Titles/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Title title = db.Title.Find(id);
            db.Title.Remove(title);
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
    }
}
