using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using GRC.Models;
using GRC.Models.Security;

namespace GRC.Controllers
{
    public class TypeReclamationsController : BaseController
    {
        private GRCModelContainer db = new GRCModelContainer();

        // GET: TypeReclamations
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult Index()
        {
            return View(db.TypeReclamation.ToList());
        }

        // GET: TypeReclamations/Details/5
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            TypeReclamation typeReclamation = db.TypeReclamation.Find(id);
            if (typeReclamation == null)
            {
                return HttpNotFound();
            }
            return View(typeReclamation);
        }

        // GET: TypeReclamations/Create
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult Create()
        {
            return View();
        }

        // POST: TypeReclamations/Create
        // Afin de déjouer les attaques par sur-validation, activez les propriétés spécifiques que vous voulez lier. Pour 
        // plus de détails, voir  http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "TypeReclamationId,Libelle,Description")] TypeReclamation typeReclamation)
        {
            if (ModelState.IsValid)
            {
                db.TypeReclamation.Add(typeReclamation);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(typeReclamation);
        }

        // GET: TypeReclamations/Edit/5
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            TypeReclamation typeReclamation = db.TypeReclamation.Find(id);
            if (typeReclamation == null)
            {
                return HttpNotFound();
            }
            return View(typeReclamation);
        }

        // POST: TypeReclamations/Edit/5
        // Afin de déjouer les attaques par sur-validation, activez les propriétés spécifiques que vous voulez lier. Pour 
        // plus de détails, voir  http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult Edit([Bind(Include = "TypeReclamationId,Libelle,Description")] TypeReclamation typeReclamation)
        {
            if (ModelState.IsValid)
            {
                db.Entry(typeReclamation).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(typeReclamation);
        }

        // GET: TypeReclamations/Delete/5
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            TypeReclamation typeReclamation = db.TypeReclamation.Find(id);
            if (typeReclamation == null)
            {
                return HttpNotFound();
            }
            return View(typeReclamation);
        }

        // POST: TypeReclamations/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult DeleteConfirmed(int id)
        {
            TypeReclamation typeReclamation = db.TypeReclamation.Find(id);
            db.TypeReclamation.Remove(typeReclamation);
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
