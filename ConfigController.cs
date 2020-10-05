using GRC.Models;
using GRC.Models.Security;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;

namespace GRC.Controllers
{
    public class ConfigController : Controller
    {
        private GRCModelContainer db = new GRCModelContainer();

        // GET: Config
        public ActionResult Index()
        {
            return View();
        }

        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult Objectifs()
        {
            var objectifs = db.Configs.Where(x => x.Key.StartsWith("DelaiObjectif.") && x.IsActive == true);
            return View(objectifs);
        }

        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult ListeDiffusionObligatoire(string id)
        {
            id = id.ToUpper();
            ViewBag.Etape = id;
            var ListeDiffusionObligatoire = db.ListeDiffusionObligatoire.Where(x => x.EtapeDiffusion == id).FirstOrDefault();
            return View(ListeDiffusionObligatoire);
        }
        [HttpPost]
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult AjoutMembreListeDiffusionObligatoire(string id, FormCollection form)
        {
            if (form.AllKeys.Contains("membre") && form.AllKeys.Contains("etape"))
            {
                var membre = form["membre"];
                var etape = form["etape"];
                int membreId;
                if (int.TryParse(membre, out membreId) && !string.IsNullOrEmpty(etape))
                {
                    var listeDiffObligatoire = db.ListeDiffusionObligatoire.Where(x => x.EtapeDiffusion == etape).FirstOrDefault();
                    if (listeDiffObligatoire != null)
                    {
                        var p = db.PersonnelSet.Find(membreId);
                        if (p != null)
                        {
                            if (!listeDiffObligatoire.PersonnelSet.Contains(p))
                            {
                                listeDiffObligatoire.PersonnelSet.Add(p);
                            }                            
                        }
                    }
                    else
                    {
                        listeDiffObligatoire = new ListeDiffusionObligatoire();
                        listeDiffObligatoire.EtapeDiffusion = etape ; 

                        var p = db.PersonnelSet.Find(membreId);
                        if (p != null)
                        {
                            listeDiffObligatoire.PersonnelSet.Add(p);
                        }
                        db.ListeDiffusionObligatoire.Add(listeDiffObligatoire);
                    }
                    db.SaveChanges();
                }
            }            

            return RedirectToAction("ListeDiffusionObligatoire", new { id = id });
        }
        [Authorize, GRCAuthorize(Roles = "Admin")]
        public ActionResult SupprimerMembreListeDiffusionObligatoire(string id, int membreId)
        {
            if (membreId > 0)
            {
                var p = db.PersonnelSet.Find(membreId);
                if (p != null)
                {
                    var listeDiffObligatoire = db.ListeDiffusionObligatoire.Where(x => x.EtapeDiffusion == id).FirstOrDefault();
                    if (listeDiffObligatoire != null)
                    {
                        listeDiffObligatoire.PersonnelSet.Remove(p);
                        db.SaveChanges();
                    }
                }
            }
            return RedirectToAction("ListeDiffusionObligatoire", new { id = id });
        }
        [HttpPost]
        public JsonResult EditConfigField(string pk, string name, string value)
        {
            var r = new
            {
                state = false,
                message = "Pas d'exécution."
            };
            //if (pk > 0)
            if (!string.IsNullOrEmpty(pk))
            {
                var config = db.Configs.Find(pk);
                if (config == null)
                {
                    r = new
                    {
                        state = false,
                        message = "Clé pour Config erroné."
                    };
                }
                else
                {
                    try
                    {

                        config.Val = value;
                        db.Entry(config).State = EntityState.Modified;

                        db.SaveChanges();

                        r = new
                        {
                            state = true,
                            message = "Valeur Config changé."
                        };
                        
                    }
                    catch (Exception e)
                    {
                        r = new
                        {
                            state = false,
                            message = "Exception : " + e.Message
                        };
                    }
                }
            }
            else
            {
                r = new
                {
                    state = false,
                    message = "Clé pour Config invalide."
                };
            }

            return Json(r, JsonRequestBehavior.AllowGet);
        }


        #region TypesReclamations

        // GET: TypeReclamations
        public ActionResult TypesReclamations()
        {
            return View(db.TypeReclamation);
        }

        // GET: Config/TypesReclamationsDetails/5
        public ActionResult TypesReclamationsDetails(int? id)
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
        public ActionResult TypesReclamationsCreer()
        {
            return View();
        }

        // POST: Config/TypesReclamationsCreer
        // Afin de déjouer les attaques par sur-validation, activez les propriétés spécifiques que vous voulez lier. Pour 
        // plus de détails, voir  http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult TypesReclamationsCreer([Bind(Include = "TypeReclamationId,Libelle,Description")] TypeReclamation typeReclamation)
        {
            if (ModelState.IsValid)
            {
                db.TypeReclamation.Add(typeReclamation);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(typeReclamation);
        }


        public ActionResult TypesReclamationsModifier(int? id)
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

        // POST: Config/TypesReclamationsModifier/5
        // Afin de déjouer les attaques par sur-validation, activez les propriétés spécifiques que vous voulez lier. Pour 
        // plus de détails, voir  http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult TypesReclamationsModifier([Bind(Include = "TypeReclamationId,Libelle,Description")] TypeReclamation typeReclamation)
        {
            if (ModelState.IsValid)
            {
                db.Entry(typeReclamation).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(typeReclamation);
        }

        // GET: Config/Delete/5
        public ActionResult TypesReclamationsSupprimer(int? id)
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

        // POST: Config/TypesReclamationsSupprimer/5
        [HttpPost, ActionName("TypesReclamationsSupprimer")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            TypeReclamation typeReclamation = db.TypeReclamation.Find(id);
            db.TypeReclamation.Remove(typeReclamation);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        #endregion

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