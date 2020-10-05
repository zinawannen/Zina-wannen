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
    public class CheckPointsController : Controller
    {
        private GRCModelContainer db = new GRCModelContainer();

        // GET: CheckPoints+
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult Index()
        {
            return View(db.AuditClotureCheckPoints.ToList());
        }

        // GET: CheckPoints/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            AuditClotureCheckPoints auditClotureCheckPoints = db.AuditClotureCheckPoints.Find(id);
            if (auditClotureCheckPoints == null)
            {
                return HttpNotFound();
            }
            return View(auditClotureCheckPoints);
        }

        // GET: CheckPoints/Create
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult Create()
        {
            return View();
        }

        // POST: CheckPoints/Create
        // Afin de déjouer les attaques par sur-validation, activez les propriétés spécifiques que vous voulez lier. Pour 
        // plus de détails, voir  http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "PointId,LibellePointFr,LibellePointEn,Actif,Ordre,Obligatoire")] AuditClotureCheckPoints auditClotureCheckPoints)
        {
            if (ModelState.IsValid)
            {
                var maxOrder = db.AuditClotureCheckPoints.Max(x => x.Ordre);
                auditClotureCheckPoints.Ordre = maxOrder.HasValue ? maxOrder.Value + 1 : 1;
                db.AuditClotureCheckPoints.Add(auditClotureCheckPoints);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(auditClotureCheckPoints);
        }

        // GET: CheckPoints/Edit/5
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            AuditClotureCheckPoints auditClotureCheckPoints = db.AuditClotureCheckPoints.Find(id);
            if (auditClotureCheckPoints == null)
            {
                return HttpNotFound();
            }
            return View(auditClotureCheckPoints);
        }

        // POST: CheckPoints/Edit/5
        // Afin de déjouer les attaques par sur-validation, activez les propriétés spécifiques que vous voulez lier. Pour 
        // plus de détails, voir  http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult Edit([Bind(Include = "PointId,LibellePointFr,LibellePointEn,Actif,Ordre,Obligatoire")] AuditClotureCheckPoints auditClotureCheckPoints)
        {
            if (ModelState.IsValid)
            {
                db.Entry(auditClotureCheckPoints).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(auditClotureCheckPoints);
        }

        // GET: CheckPoints/Delete/5
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            AuditClotureCheckPoints auditClotureCheckPoints = db.AuditClotureCheckPoints.Find(id);
            if (auditClotureCheckPoints == null)
            {
                return HttpNotFound();
            }
            return View(auditClotureCheckPoints);
        }

        // POST: CheckPoints/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult DeleteConfirmed(int id)
        {
            AuditClotureCheckPoints auditClotureCheckPoints = db.AuditClotureCheckPoints.Find(id);
            db.AuditClotureCheckPoints.Remove(auditClotureCheckPoints);
            db.SaveChanges();
            return RedirectToAction("Index");
        }
        
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult Activation(int id)
        {
            AuditClotureCheckPoints auditClotureCheckPoints = db.AuditClotureCheckPoints.Find(id);
            auditClotureCheckPoints.Actif = !auditClotureCheckPoints.Actif;
            db.Entry(auditClotureCheckPoints).State = EntityState.Modified;
            db.SaveChanges();
            return RedirectToAction("Index");
        }
        
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult Order()
        {
            var auditClotureCheckPoints = db.AuditClotureCheckPoints.OrderBy(x => x.Ordre);
            return View(auditClotureCheckPoints);
        }
        [HttpPost]
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult Order(string data, FormCollection formCollection)
        {
            var s = data.Replace("Ordered[]=", "").Split('&');
            //try
            //{
            int order = 1;
            foreach (var id in s)
            {
                var c = db.AuditClotureCheckPoints.Find(int.Parse(id));
                c.Ordre = order;
                db.Entry(c).State = EntityState.Modified;
                order++;
            }
            db.SaveChanges();
            //}
            //catch (System.Data.Entity.Validation.DbEntityValidationException e) { }
            //catch (EntityException e) { }
            //catch (Exception e) { }   
            return RedirectToAction("Order");
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
