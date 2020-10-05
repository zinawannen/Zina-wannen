using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using GRC.Models;

namespace GRC.Controllers
{
    public class RapportCRUDsController : BaseController
    {
        private GRCModelContainer db = new GRCModelContainer();

        // GET: /RapportCRUDs/
        public ActionResult Index()
        {
            return View(db.RapportSet.ToList());
        }

        // GET: /RapportCRUDs/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Rapport rapport = db.RapportSet.Find(id);
            if (rapport == null)
            {
                return HttpNotFound();
            }
            return View(rapport);
        }

        // GET: /RapportCRUDs/Create
        public ActionResult Create()
        {
            ViewBag.PersonnelId = new SelectList(db.PersonnelSet.Select(x => new { x.PersonnelId, NomComplet = x.Nom + " " + x.Prenom }), "PersonnelId", "NomComplet");
            return View();
        }

        // POST: /RapportCRUDs/Create
        // Afin de déjouer les attaques par sur-validation, activez les propriétés spécifiques que vous voulez lier. Pour 
        // plus de détails, voir  http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(Rapport rapport)
        {
            if (ModelState.IsValid)
            {
                db.RapportSet.Add(rapport);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(rapport);
        }

        // GET: /RapportCRUDs/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Rapport rapport = db.RapportSet.Find(id);
            if (rapport == null)
            {
                return HttpNotFound();
            }
            return View(rapport);
        }

        // POST: /RapportCRUDs/Edit/5
        // Afin de déjouer les attaques par sur-validation, activez les propriétés spécifiques que vous voulez lier. Pour 
        // plus de détails, voir  http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "RapportId,Produit,NumSerie,DateOuverture,DateCloture,PiloteId,D2NumAncienRapport,D2Quoi,D2Qui,D2Ou,D2Comment,D2Combien,D2Pourquoi,D2ProblemeRecurent,D2CommentaireDepassementDelai,D3CommentaireDepassementDelai,D5CommentaireDepassementDelai,D7CommentaireDepassementDelai,D8CommentaireDepassementDelai")] Rapport rapport)
        {
            if (ModelState.IsValid)
            {
                db.Entry(rapport).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(rapport);
        }

        // GET: /RapportCRUDs/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Rapport rapport = db.RapportSet.Find(id);
            if (rapport == null)
            {
                return HttpNotFound();
            }
            return View(rapport);
        }

        // POST: /RapportCRUDs/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Rapport rapport = db.RapportSet.Find(id);
            db.RapportSet.Remove(rapport);
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
