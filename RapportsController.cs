using Aspose.Cells;
using GRC.Models;
using GRC.Models.Security;
using PagedList;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Data.Entity.Validation;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Threading.Tasks;
using System.Transactions;
using System.Web;
using System.Web.Mvc;



namespace GRC.Controllers
{
    public class RapportsController : BaseController
    {
        private GRCModelContainer db = new GRCModelContainer();
        //
        // GET: /Rapport/
        //[Authorize]

        // *******************************************************************************************************************
        // Méthode modifiée par Imen OUERTANI
        // Le 09/11/2015 
        // Paramètres de la méthode => Liste des Rapports ("null" si aucucne recherche n'était effectuée)
        public ActionResult Index(List<Rapport> ListRapp) //Index()
        {
            //****************************************************************************************************************

            var UAP = TempData["UAP"];
            var Pilote = TempData["PersonnelId"];
            var Annee = TempData["Annee"];
            var type = TempData["TypeRapport"];

            if (Annee == null)
            {
                Annee = DateTime.Now.Year;
                ViewBag.Default = Annee;
            }

            //****************************************************************************************************
            //var AnneeDeMiseEnService = 2015;
            var list_Annee = new List<string>();
            for (int i = 2015; i <= DateTime.Now.Year; i++)
            {
                list_Annee.Add(i.ToString());
            }
            //****************************************************************************************************

            var list_UAP = db.Configs.Find("UAP").Val.Split(',');

            ViewBag.Annee = new SelectList(list_Annee, DateTime.Now.Year.ToString());

            var personnelListSelect = db.Roles.Where(x => x.RoleName == "Pilote").FirstOrDefault().Users.Select(u => u.Personne).ToList();

            //db.Users.Select(x => x.Roles.Where(u => u.RoleName == "Pilote")).ToList();           


            //********************
            var typeRapport = new List<string>();
            typeRapport.Add("Alerte");
            typeRapport.Add("Reclamation");

            ViewBag.TypeRapport = new SelectList(typeRapport, type);
            ///********************************************************************************************

            if (ListRapp != null)
            {
                ViewBag.UAP = new SelectList(list_UAP, UAP);
                ViewBag.PersonnelId = new SelectList(personnelListSelect, "PersonnelId", "NomComplet", Pilote);
                ViewBag.Annee = new SelectList(list_Annee, Annee);
                //var pageNumber = page ?? 1;
                //var onePageOfRapports = ListRapp.OrderBy(x => x.DateOuverture).ToPagedList(pageNumber, 5);
                //ViewBag.OnePageOfRapports = onePageOfRapports;
                //return View(onePageOfRapports);
                return View(ListRapp.OrderBy(x => x.DateOuverture));
            }
            else
            {
                ViewBag.UAP = new SelectList(list_UAP, "");
                ViewBag.PersonnelId = new SelectList(personnelListSelect, "PersonnelId", "NomComplet", string.Empty);

                //IEnumerable<Rapport> rapports = await db.RapportSet.Where(x => x.IsActif == true).ToListAsync();
                var rapports = db.RapportSet.Where(x => x.IsActif == true && x.DateOuverture.Year == DateTime.Now.Year).ToList();

                if (User != null)
                {
                    if (User.roles.Contains("Admin"))
                    {
                        rapports = db.RapportSet.Where(x => x.DateOuverture.Year == DateTime.Now.Year).ToList();
                    }
                    else
                    {
                        rapports = db.RapportSet.Where(x => x.IsActif == true || x.PersonnelId == User.PersonnelId && x.DateOuverture.Year == DateTime.Now.Year).ToList();
                    }
                }
                var objectifs = db.Configs.Where(x => x.Key.StartsWith("DelaiObjectif.") && x.IsActive == true).ToList();

                //var pageNumber = page ?? 1; // if no page was specified in the querystring, default to the first page (1)
                //var onePageOfRapports = rapports.OrderBy(x => x.DateOuverture).ToPagedList(pageNumber, 5); // will only contain 25 products max because of the pageSize

                ViewBag.Objectifs = objectifs;
                //ViewBag.OnePageOfRapports = onePageOfRapports;
                //return View(onePageOfRapports);
                return View(rapports.OrderByDescending(x => x.DateOuverture));
            }
        }

        //*********************************************************************************
        // Méthode ajoutée par Imen OUERTANI
        // Le 09/11/2015 
        // Filtrage des rapports
        [HttpPost, ActionName("Index")]
        public ActionResult Recherche(FormCollection form) //Index(int? page)
        {

            string sUAP = form["UAP"];
            string Pilote = form["PersonnelId"];
            string sAnnee = form["Annee"];
            string type = form["TypeRapport"];
            bool alerte = false;
            if (type=="Alerte")
            {
                alerte = true;
            }
            int sPilote = string.IsNullOrEmpty(Pilote) ? 0 : Int32.Parse(Pilote);
       //   var Annee = Int32.Parse(sAnnee);
            //int sUAP = string.IsNullOrEmpty(UAP) ? 0 : Int32.Parse(UAP);

            TempData["UAP"] = sUAP;
            TempData["Pilote"] = Pilote;
            TempData["Annee"] = sAnnee;
            TempData["TypeRapport"] = type;
            //********************************************************************************************************
            var LAnnee = new List<string>();

            for (int i = 2015; i <= DateTime.Now.Year; i++)
            {
                LAnnee.Add(i.ToString());
            }
            if (sAnnee == "")
            {
                sAnnee = DateTime.Now.Year.ToString();
                ViewBag.Default = sAnnee;
            }
            int year = Int32.Parse(sAnnee);
            ViewBag.Annee = new SelectList(LAnnee, sAnnee);
         
            var list_UAP = db.Configs.Find("UAP").Val.Split(',');
            ViewBag.UAP = new SelectList(list_UAP, sUAP);

            var personnelListSelect = db.Roles.Where(x => x.RoleName == "Pilote").FirstOrDefault().Users.Select(u => u.Personne).ToList();
            ViewBag.PersonnelId = new SelectList(personnelListSelect, "PersonnelId", "NomComplet", Pilote);
           

            ViewData["Annee"] = new SelectList(LAnnee);

            var typeRapport = new List<string>();
            typeRapport.Add("Alerte");
            typeRapport.Add("Reclamation");

            ViewBag.TypeRapport = new SelectList(typeRapport, type);

            //****************************************************************************************************************

            List<Rapport> rapports = db.RapportSet.Where(x => x.IsActif == true && x.DateOuverture.Year == year).ToList();

            if (User != null)
            {
                if (User.roles.Contains("Admin"))
                {
                    rapports = db.RapportSet.Where(x => x.DateOuverture.Year == year).ToList();
                }
                else
                {
                    rapports = db.RapportSet.Where(x => x.IsActif == true || x.PersonnelId == User.PersonnelId && x.DateOuverture.Year == year).ToList();
                }
            }

            //**************** Recherche des Rapports par UAP et Pilote passés en paramètre **********************************

            if (sUAP != "" && Pilote != "" && type != "")
            {
                if (alerte == true)
                {
                    rapports = db.RapportSet.Where(x => x.IsActif == true && x.Pilote.PersonnelId == sPilote && x.UAP == sUAP && x.DateOuverture.Year == year && x.Alerte == true).ToList();
                }
                
                else 
                {
                    rapports = db.RapportSet.Where(x => x.IsActif == true && x.Pilote.PersonnelId == sPilote && x.UAP == sUAP && x.DateOuverture.Year == year && x.Alerte == false).ToList();
                }
            }
            else if (sUAP == "" && Pilote != "" && type != "")
            {
                if (alerte == true)
                {

                    rapports = db.RapportSet.Where(x => x.IsActif == true && x.Pilote.PersonnelId == sPilote && x.DateOuverture.Year == year && x.Alerte == true).ToList();
                }
                else
                {
                    rapports = db.RapportSet.Where(x => x.IsActif == true && x.Pilote.PersonnelId == sPilote && x.DateOuverture.Year == year && x.Alerte == false).ToList();

                }

            }
            else if (sUAP == "" && Pilote == "" && type != "")
            {
                if (alerte == true)
                {
                    rapports = db.RapportSet.Where(x => x.IsActif == true && x.DateOuverture.Year == year && x.Alerte == true).ToList();
                }
                else
                {
                    rapports = db.RapportSet.Where(x => x.IsActif == true && x.DateOuverture.Year == year && x.Alerte == false).ToList();

                }
            }
            else if (sUAP != "" && Pilote == "" && type == "")
            {
                rapports = db.RapportSet.Where(x => x.IsActif == true && x.UAP == sUAP && x.DateOuverture.Year == year).ToList();
            }
            else if (sUAP != "" && Pilote == "" && type != "")
            {
                if (alerte == true)
                {
                    rapports = db.RapportSet.Where(x => x.IsActif == true && x.DateOuverture.Year == year && x.UAP == sUAP && x.Alerte == true).ToList();
                }
                else
                {
                    rapports = db.RapportSet.Where(x => x.IsActif == true && x.DateOuverture.Year == year && x.UAP == sUAP && x.Alerte == false).ToList();

                }
            }
            else if (sUAP == "" && Pilote == "" && type == "")
            {
                rapports = db.RapportSet.Where(x => x.IsActif == true && x.DateOuverture.Year == year).ToList();
            }
            else if (sUAP != "" && Pilote != "" && type == "")
            {
                rapports = db.RapportSet.Where(x => x.IsActif == true && x.DateOuverture.Year == year && x.UAP==sUAP && x.Pilote.PersonnelId==sPilote).ToList();
            }

            var objectifs = db.Configs.Where(x => x.Key.StartsWith("DelaiObjectif.") && x.IsActive == true).ToList();
            ViewBag.Objectifs = objectifs;

           
            return View(rapports.OrderByDescending(x => x.DateOuverture));
           
        }

        //*********************************************************************************

        [Authorize, GRCAuthorize(Roles = "Pilote")]
        public ActionResult Creer()
        {
            CreerRapportInit();

            return View();
        }
        [Authorize, GRCAuthorize(Roles = "Pilote")]
        [HttpPost, ActionName("Creer")]
        public ActionResult CreerNouvelRapport(Rapport rapport)
        {
            if (ModelState.IsValid)
            {
                try {

                    rapport.DateMaj = DateTime.Now;
                    rapport.DateOuverture = DateTime.Now;
                    rapport.IsActif = true;

                    rapport.D6_Defauts_Client = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0";
                    rapport.D6_Defauts_Fournisseur = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0";
                    rapport.D6_Defauts_Interne = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0";

                    db.RapportSet.Add(rapport);
                    db.SaveChanges();

                    return RedirectToAction("Rapport", new { id = rapport.RapportId });
                }
                catch (System.Data.Entity.Validation.DbEntityValidationException dbEx)
                {
                    var viewErrors = "";
                    foreach (System.Data.Entity.Validation.DbEntityValidationResult entityErr in dbEx.EntityValidationErrors)
                    {
                        foreach (System.Data.Entity.Validation.DbValidationError error in entityErr.ValidationErrors)
                        {
                            // Console.WriteLine("Error Property Name {0} : Error Message: {1}",
                            // error.PropertyName, error.ErrorMessage);
                            viewErrors += "<li>Validation Exception for <b>" + error.PropertyName + "</b> : Message: <b>" + error.ErrorMessage + "</b></li>";
                        }
                    }
                    ViewBag.viewErrors = viewErrors;
                }
                catch (Exception error)
                {
                    var viewErrors = "";
                    viewErrors += "<li>Error Adding new item :" + error.Message + " : Inner Exception: " + error.InnerException + "<br /><b>Please notify your system administrator or developer.</b></li>";
                    ViewBag.viewErrors = viewErrors;
                }
            }
            CreerRapportInit();
            return View(rapport);
        }

        private void CreerRapportInit()
        {
            var list_UAP = db.Configs.Find("UAP").Val.Split(',');
            ViewBag.UAP = new SelectList(list_UAP);//new SelectList(new List<string> { "ALSTOM", "Schneider", "PE", "USINE" });
            ViewBag.TypeIncident = new SelectList(new List<string> { "Client Direct "," Client Final", "Interne", "Fournisseur" });

            var personnelListSelect = db.PersonnelSet.Select(x => new
            {
                x.PersonnelId,
                NomComplet = ((x.Nom ?? "") + " " + (x.Prenom ?? "")).Trim()
            }).ToList();
            //personnelListSelect.Insert(0, new { PersonnelId = 0, NomComplet = "..." });

            ViewBag.PersonnelId = new SelectList(
                personnelListSelect,
                "PersonnelId",
                "NomComplet", string.Empty);
            //ViewBag.UAP = new SelectList(new List<string> { "ALSTOM", "Schneider", "PE", "USINE" });

            ViewBag.TypeReclamationId = new SelectList(
                db.TypeReclamation.Select(x => new
                {
                    x.TypeReclamationId,
                    x.Libelle
                }),
                "TypeReclamationId",
                "Libelle", string.Empty);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        /// http://tech.pro/tutorial/1252/asynchronous-controllers-in-asp-net-mvc
        /// 
        public async Task<ActionResult> Rapport(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Rapport rapport = await db.RapportSet.FindAsync(id);
            if (rapport == null)
            {
                return HttpNotFound();
            }
            var persons = db.PersonnelSet.Select(x => new
            {
                PersonnelId = x.PersonnelId,
                x.Nom, x.Prenom,
                NomComplet = string.Concat(x.Nom, " ", x.Prenom ?? "")
            });

            ViewBag.PersonnelSelectList = new SelectList(
                await persons.ToListAsync(),
                "PersonnelId",
                "NomComplet", string.Empty);

            return View(rapport);
        }

        public ActionResult CloturerEtapeConfirmation(int? id, string Etape)
        {
            var listeDiffusion = db.RapportSet.Find(id)
                .ListesDiffusion.Where(x => x.EtapeDiffusion.Equals(Etape)).FirstOrDefault();


            return PartialView(listeDiffusion);
        }

        public ActionResult CloturerEtape(int? id, string Etape)
        {
            var rapport = db.RapportSet.Find(id);
            //bool stat = false;

            //if (rapport.D3_1_Investigation_1Tracabilite_Reaction == "Non") stat = true;


            //if (stat)
            //{
            //    //clot
            //}
            //else
            //{
            //    // redirect to erreur clot 
            //}
            switch (Etape)
            {
                case "1D":
                    rapport.D1_DateCloture = DateTime.Now;
                    break;
                case "2D":
                    rapport.D2_DateCloture = DateTime.Now;
                    break;
                case "3D":
                    rapport.D3_DateCloture = DateTime.Now;
                    break;
                case "4D":
                    rapport.D4_DateCloture = DateTime.Now;
                    break;
                case "5D":
                    rapport.D5_DateCloture = DateTime.Now;
                    break;
                case "6D":
                    rapport.D6_DateCloture = DateTime.Now;
                    break;
                case "7D":
                    rapport.D7_DateCloture = DateTime.Now;
                  
                    break;
                default:
                    break;
            }
            
            // TODO : Procédure d'envoi de mails à la liste de diffusion

            if (db.Entry(rapport).State == EntityState.Modified)
            {
                db.SaveChanges();

                // TODO envoi de mails à la liste de diffusion de l'étape
                this.SendMailCloturerEtape(rapport, Etape);
            }
            return RedirectToAction("Rapport", new { id = id });
        }

        private void SendMailCloturerEtape(Rapport rapport, string Etape)
        {
            MailMessage mail = new MailMessage();
            SmtpClient smtpClient = new SmtpClient();


            // les paramètres de SMTP
            string mail_smtp_host = db.Configs.Find("mail_smtp_host").Val;

            int mail_smtp_port = 25; int.TryParse(db.Configs.Find("mail_smtp_port").Val, out mail_smtp_port);

            string mail_smtp_credentials_username = db.Configs.Find("mail_smtp_credentials_username").Val;
            string mail_smtp_credentials_password = db.Configs.Find("mail_smtp_credentials_password").Val;

            bool mail_smtp_enablessl = false; bool.TryParse(db.Configs.Find("mail_smtp_enablessl").Val, out mail_smtp_enablessl);


            // les paramètres du mail destinateur
            string mail_from_mail = db.Configs.Find("mail_from_mail").Val;
            string mail_from_name = db.Configs.Find("mail_from_name").Val;
            var mail_cc_mail = db.Configs.Find("mail_cc_mail").Val.Replace(" ", "").Split(';');
            var mail_cc_name = db.Configs.Find("mail_cc_name").Val.Split(';'); ;


            if (!string.IsNullOrEmpty(mail_smtp_host))
            {
                try
                {
                    mail.From = new MailAddress(mail_from_mail, mail_from_name);

                    //if (!string.IsNullOrEmpty(mail_cc_mail))
                    //    mail.CC.Add(new MailAddress(mail_cc_mail, mail_cc_name));
                    if (mail_cc_mail != null)
                    {
                        if (mail_cc_mail.Length > 0)
                        {
                            for (int i = 0; i < mail_cc_mail.Length; i++)
                            {
                                mail.CC.Add(new MailAddress(
                                    mail_cc_mail[i],
                                    mail_cc_name != null ? (mail_cc_name.Length >= i ? mail_cc_name[i] : "") : ""
                                    ));
                            }
                        }
                    }

                    // Affectation des destinataires du mail
                    var desti = rapport.MembreEquipeSet;
                    foreach (var item in desti)
                    {
                        var P = item.PersonnelSet;
                        if (!string.IsNullOrEmpty(P.Mail))
                            mail.To.Add(new MailAddress(P.Mail, P.NomComplet));
                    }
                    var listeDiffusion = rapport.ListesDiffusion.Where(x => x.EtapeDiffusion.Equals(Etape)).FirstOrDefault();
                    if (listeDiffusion != null)
                    {
                       var destinataires = listeDiffusion.PersonnelSet;
                       
                        foreach (var dest in destinataires)
                        {
                            if (!string.IsNullOrEmpty(dest.Mail))
                                mail.To.Add(new MailAddress(dest.Mail, dest.NomComplet));
                        }
                    }

                    var listeDiffusionObligatoire = db.ListeDiffusionObligatoire.Where(x => x.EtapeDiffusion.Equals(Etape)).FirstOrDefault();
                    if (listeDiffusionObligatoire != null)
                    {
                        var destinataires = listeDiffusionObligatoire.PersonnelSet;

                        foreach (var dest in destinataires)
                        {
                            if (!string.IsNullOrEmpty(dest.Mail))
                                mail.To.Add(new MailAddress(dest.Mail, dest.NomComplet));
                        }
                    }

                    mail.Subject = "GRC - Clôture d'étape du rapport : " + Etape;
                    mail.IsBodyHtml = true;
                    mail.Body = "Notifiaction pour clôture d'étape";

                    // Préparation du corps du mail
                    var viewData = new Dictionary<string, string>();
                    viewData.Add("Etape", Etape);
                    string htmlContent = this.RenderViewToString(
                        "~/Views/MailTemplates/CloturerEtape.cshtml",
                        rapport,
                        new ViewDataDictionary() { { "Etape", Etape } }
                );

                    AlternateView av = AlternateView.CreateAlternateViewFromString(htmlContent, null, System.Net.Mime.MediaTypeNames.Text.Html);

                    string Logo8D = Server.MapPath(Url.Content("~/Images/Logo/Logo8D.png")); //string.Format("http://{0}{1}{2}", Request.Url.Host, Request.Url.Port > 0 ? ":" + Request.Url.Port : "", Url.Content("/Images/Caper.png"));
                    //            
                    if (Logo8D != string.Empty)
                    {
                        LinkedResource lr = new LinkedResource(Logo8D, System.Net.Mime.MediaTypeNames.Image.Jpeg);
                        lr.ContentId = "image2";
                        av.LinkedResources.Add(lr);
                    }

                    // Ajout du corps du mail
                    mail.AlternateViews.Add(av);

                    // configuration de l'SMTP
                    smtpClient.Host = mail_smtp_host;// "smtp.gmail.com";
                    if (mail_smtp_port > 0) smtpClient.Port = mail_smtp_port;// 587;
                    if (!string.IsNullOrEmpty(mail_smtp_credentials_username))
                    {
                        smtpClient.Credentials = new System.Net.NetworkCredential(
                            mail_smtp_credentials_username,
                            mail_smtp_credentials_password);//"from@gmail.com","Password");
                    }
                    smtpClient.EnableSsl = mail_smtp_enablessl;// true;

                    // et ça s'envole
                    smtpClient.Send(mail);
                }
                catch (Exception ex)
                {
                    ModelState.AddModelError("",
                           "Une erreur es survenue lors de l'envoi de mail. Exception:" + ex.Message);

                }
            }
        }

        [HttpPost]
        public ActionResult Desactiver(int id)
        {
            var rapport = db.RapportSet.Find(id);
            if (rapport != null)
            {
                rapport.IsActif = false;
                rapport.DateDesactivation = DateTime.Now;

                db.SaveChanges();
                TempData["success"] = "Le rapport " + rapport.ReferenceRapport + " a été désactivé et mis à la corbeille.";
            }
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult Restaurer(int id)
        {
            var rapport = db.RapportSet.Find(id);
            if (rapport != null)
            {
                rapport.IsActif = true;
                rapport.DateReactivation = DateTime.Now;

                db.SaveChanges();
                TempData["success"] = "Le rapport " + rapport.ReferenceRapport + " a été restauré.";
            }
            return RedirectToAction("Index");
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

            using (TransactionScope scope = new TransactionScope())
            {
                var FileStoragePath =
                    db.Configs.Find("FileStoragePath").Val;
                //WebConfigurationManager.AppSettings["FileStoragePath"];

                try
                {
                    foreach (var item in db.ListesDiffusion.Where(x => x.RapportId == rapport.RapportId))
                    {
                        item.PersonnelSet.Clear();
                        db.ListesDiffusion.Remove(item);
                    }
                    foreach (var item in db.D4Causes.Where(x => x.RapportId == rapport.RapportId))
                        db.D4Causes.Remove(item);

                    foreach (var item in db.D5ActionsCorrectives.Where(x => x.RapportId == rapport.RapportId))
                        db.D5ActionsCorrectives.Remove(item);

                    foreach (var item in db.D7ActionsGeneralisation.Where(x => x.RapportId == rapport.RapportId))
                        db.D7ActionsGeneralisation.Remove(item);

                    foreach (var item in db.MembreEquipeSet.Where(x => x.RapportId == rapport.RapportId))
                        db.MembreEquipeSet.Remove(item);

                    foreach (var item in db.RapportAuditClotureCheckList.Where(x => x.RapportId == rapport.RapportId))
                        db.RapportAuditClotureCheckList.Remove(item);

                    foreach (var item in db.PhotoSet.Where(x => x.Rapport.RapportId == rapport.RapportId))
                    {
                        // supprimer les images enregistrées

                        //var imgPath = Path.Combine(
                        //    FileStoragePath,
                        //        "Rapports",             // Dossier des Rapports
                        //        "R_" + item.Rapport.RapportId, // Dossier du rapport en cours
                        //        "D2_Photos",            // Dossier des Photos
                        //        item.Type.ToUpper(),          // Dossier des Photos OK ou NOK
                        //        item.Key + item.ExtensionFichier
                        //        );

                        //System.IO.File.Delete(imgPath);

                        db.PhotoSet.Remove(item);
                        //if (db.Entry(item).State != EntityState.Deleted)
                        //    db.Entry(item).State = EntityState.Deleted; 
                    }

                    foreach (var item in db.Documents.Where(x => x.Rapport.RapportId == rapport.RapportId))
                    {
                        // supprimer les documents joint au rapport.
                        //var doc = db.Documents.Find(id);


                        //var docPath = Path.Combine(
                        //    FileStoragePath,
                        //        "Rapports",             // Dossier des Rapports
                        //        "R_" + item.Rapport.RapportId, // Dossier du rapport en cours
                        //        "Documents",            // Dossier des Documents 
                        //        item.Key + item.ExtensionFichier
                        //        );

                        //System.IO.File.Delete(docPath);

                        db.Documents.Remove(item);
                        //if (db.Entry(item).State != EntityState.Deleted)
                        //    db.Entry(item).State = EntityState.Deleted;
                        //db.SaveChanges();
                    }

                    // supprimer le dossier du rapport
                    var rapportPath = Path.Combine(
                            FileStoragePath,
                                "Rapports",             // Dossier des Rapports
                                "R_" + rapport.RapportId
                                );
                    if (System.IO.Directory.Exists(rapportPath))
                        System.IO.Directory.Delete(rapportPath, true);

                    db.RapportSet.Remove(rapport);

                    db.SaveChanges();

                    scope.Complete();
                }
                catch (Exception e)
                {
                    TempData["error"] = "Une erreur est survenue lors de la suppression du Rapport:<br />"
                        + e.Message + " (Source : " + e.Source + ")"
                        + (e.InnerException != null ? "<br>Inner Exception: <br>" + e.InnerException.Message + " (Source : " + e.InnerException.Source + ")" : "")
                        + "<br>Merci de notifier l'administrateur de l'application ou dans le cas échéant un membre SI (Dev).";
                    scope.Dispose();
                    return RedirectToAction("Delete", new { id = id });
                    //response = "{\"status\":\"danger\",\"message\":\"Error :" + e.Message + " \"}";
                }
                scope.Dispose();
            }

            TempData["success"] = "La suppression du rapport est effecuée avec succès.";
            TempData["info"] = "Pour info, le rapport supprimé ne pourra plus être récupéré.";

            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult CloturerRapport(int id, string CommentVisaManagement, FormCollection form)
        {
            var rapport = db.RapportSet.Find(id);

            //var d = form["CommentVisaManagement"];

            rapport.DateCloture = DateTime.Now;
            rapport.CommentVisaCloture = CommentVisaManagement;

            db.Entry(rapport).State = EntityState.Modified;
            db.SaveChanges();

            TempData["success"] = "Le rapport a été clôturé avec succès. ";


            // TODO envoi de mails à la liste de diffusion 8D
            this.SendMailCloturerRapport(rapport);

            return RedirectToAction("Rapport", new { id = id });
        }

        private void SendMailCloturerRapport(Rapport rapport)
        {
            MailMessage mail = new MailMessage();
            SmtpClient smtpClient = new SmtpClient();


            // les paramètres de SMTP
            string mail_smtp_host = db.Configs.Find("mail_smtp_host").Val;

            int mail_smtp_port = 25; int.TryParse(db.Configs.Find("mail_smtp_port").Val, out mail_smtp_port);

            string mail_smtp_credentials_username = db.Configs.Find("mail_smtp_credentials_username").Val;
            string mail_smtp_credentials_password = db.Configs.Find("mail_smtp_credentials_password").Val;

            bool mail_smtp_enablessl = false; bool.TryParse(db.Configs.Find("mail_smtp_enablessl").Val, out mail_smtp_enablessl);


            // les paramètres du mail destinateur
            string mail_from_mail = db.Configs.Find("mail_from_mail").Val;
            string mail_from_name = db.Configs.Find("mail_from_name").Val;
            var mail_cc_mail = db.Configs.Find("mail_cc_mail").Val.Replace(" ", "").Split(';');
            var mail_cc_name = db.Configs.Find("mail_cc_name").Val.Split(';'); ;


            if (!string.IsNullOrEmpty(mail_smtp_host))
            {
                try
                {
                    mail.From = new MailAddress(mail_from_mail, mail_from_name);

                    //if (!string.IsNullOrEmpty(mail_cc_mail))
                    //    mail.CC.Add(new MailAddress(mail_cc_mail, mail_cc_name));
                    if (mail_cc_mail != null)
                    {
                        if (mail_cc_mail.Length > 0)
                        {
                            for (int i = 0; i < mail_cc_mail.Length; i++)
                            {
                                mail.CC.Add(new MailAddress(
                                    mail_cc_mail[i],
                                    mail_cc_name != null ? (mail_cc_name.Length >= i ? mail_cc_name[i] : "") : ""
                                    ));
                            }
                        }
                    }

                    // Affectation des destinataires du mail

                    var listeDiffusion = rapport.ListesDiffusion.Where(x => x.EtapeDiffusion.Equals("8D")).FirstOrDefault();
                    if (listeDiffusion != null)
                    {
                        var destinataires = listeDiffusion.PersonnelSet;
                        var desti = rapport.MembreEquipeSet;

                        foreach (var item in desti)
                        {
                            var P = item.PersonnelSet;
                            if (!string.IsNullOrEmpty(P.Mail))
                                mail.To.Add(new MailAddress(P.Mail, P.NomComplet));
                        }
                        foreach (var dest in destinataires)
                        {
                            if (!string.IsNullOrEmpty(dest.Mail))
                                mail.To.Add(new MailAddress(dest.Mail, dest.NomComplet));
                        }
                    }
                    // Mettre en copie l'Auditeur
                    if (!string.IsNullOrEmpty(rapport.Auditeur.Mail))
                        mail.CC.Add(new MailAddress(rapport.Auditeur.Mail, rapport.Auditeur.NomComplet));
                    // Mettre le responsable désignateur de l'auditeur en copie
                    if (!string.IsNullOrEmpty(rapport.ResponsableDesignantAuditeur.Mail))
                        mail.CC.Add(new MailAddress(rapport.ResponsableDesignantAuditeur.Mail, rapport.ResponsableDesignantAuditeur.NomComplet));


                    mail.Subject = "GRC - Clôture du rapport d'audit";
                    mail.IsBodyHtml = true;
                    mail.Body = "Notifiaction pour clôture d'étape";

                    // Préparation du corps du mail
                    //var viewData = new Dictionary<string, string>();
                    //viewData.Add("Etape", Etape);
                    string htmlContent = this.RenderViewToString(
                        "~/Views/MailTemplates/CloturerRapport.cshtml",
                        rapport
                        //,new ViewDataDictionary(viewData)
                        );

                    AlternateView av = AlternateView.CreateAlternateViewFromString(htmlContent, null, System.Net.Mime.MediaTypeNames.Text.Html);

                    string Logo8D = Server.MapPath(Url.Content("~/Images/Logo/Logo8D.png")); //string.Format("http://{0}{1}{2}", Request.Url.Host, Request.Url.Port > 0 ? ":" + Request.Url.Port : "", Url.Content("/Images/Caper.png"));
                    //            
                    if (Logo8D != string.Empty)
                    {
                        LinkedResource lr = new LinkedResource(Logo8D, System.Net.Mime.MediaTypeNames.Image.Jpeg);
                        lr.ContentId = "image2";
                        av.LinkedResources.Add(lr);
                    }

                    // Ajout du corps du mail
                    mail.AlternateViews.Add(av);

                    // configuration de l'SMTP
                    smtpClient.Host = mail_smtp_host;// "smtp.gmail.com";
                    if (mail_smtp_port > 0) smtpClient.Port = mail_smtp_port;// 587;
                    if (!string.IsNullOrEmpty(mail_smtp_credentials_username))
                    {
                        smtpClient.Credentials = new System.Net.NetworkCredential(
                            mail_smtp_credentials_username,
                            mail_smtp_credentials_password);//"from@gmail.com","Password");
                    }
                    smtpClient.EnableSsl = mail_smtp_enablessl;// true;

                    // et ça s'envole
                    smtpClient.Send(mail);
                }
                catch (Exception ex)
                {
                    ModelState.AddModelError("",
                           "Une erreur est survenue lors de l'envoi de mail. Exception:" + ex.Message);

                }
            }
        }

        public ActionResult DemandeAudit(int id)
        {
            var rapport = db.RapportSet.Find(id);

            rapport.DemandeAuditClotureDate = DateTime.Now;
            rapport.DemandeAuditCloture = true;

            db.Entry(rapport).State = EntityState.Modified;
            db.SaveChanges();

            // envoyer un mail au responsables
            this.SendMailDemandeAudit(rapport);

            return RedirectToAction("Rapport", new { id = id });
        }

        private void SendMailDemandeAudit(Rapport rapport)
        {
            MailMessage mail = new MailMessage();
            SmtpClient smtpClient = new SmtpClient();


            // les paramètres de SMTP
            string mail_smtp_host = db.Configs.Find("mail_smtp_host").Val;

            int mail_smtp_port = 25; int.TryParse(db.Configs.Find("mail_smtp_port").Val, out mail_smtp_port);

            string mail_smtp_credentials_username = db.Configs.Find("mail_smtp_credentials_username").Val;
            string mail_smtp_credentials_password = db.Configs.Find("mail_smtp_credentials_password").Val;

            bool mail_smtp_enablessl = false; bool.TryParse(db.Configs.Find("mail_smtp_enablessl").Val, out mail_smtp_enablessl);


            // les paramètres du mail destinateur
            string mail_from_mail = db.Configs.Find("mail_from_mail").Val;
            string mail_from_name = db.Configs.Find("mail_from_name").Val;
            var mail_cc_mail = db.Configs.Find("mail_cc_mail").Val.Replace(" ", "").Split(';');
            var mail_cc_name = db.Configs.Find("mail_cc_name").Val.Split(';'); ;


            if (!string.IsNullOrEmpty(mail_smtp_host))
            {
                try
                {
                    mail.From = new MailAddress(mail_from_mail, mail_from_name);

                    //if (!string.IsNullOrEmpty(mail_cc_mail))
                    //    mail.CC.Add(new MailAddress(mail_cc_mail, mail_cc_name));
                    if (mail_cc_mail != null)
                    {
                        if (mail_cc_mail.Length > 0)
                        {
                            for (int i = 0; i < mail_cc_mail.Length; i++)
                            {
                                mail.CC.Add(new MailAddress(
                                    mail_cc_mail[i],
                                    mail_cc_name != null ? (mail_cc_name.Length >= i ? mail_cc_name[i] : "") : ""
                                    ));
                            }
                        }
                    }

                    // Affectation des destinataires du mail
                    var RolesList = new List<string> { "Auditeur", "Admin", "DQSSE", "RSQ" };
                    var destinataires =
                        (from users in db.Users
                         from roles in users.Roles
                         where RolesList.Contains(roles.RoleName)
                         select users.Personne).Distinct().ToList();
                    //db.Users.Where(x => x.Roles.Select(r=>r.RoleName).Intersect(new List<string>{ "" }).Count()>0);

                    foreach (var dest in destinataires)
                    {
                        if (!string.IsNullOrEmpty(dest.Mail))
                            mail.To.Add(new MailAddress(dest.Mail, dest.NomComplet));
                    }

                    mail.Subject = "GRC - Nouvelle demande d'audit de clôture";
                    mail.IsBodyHtml = true;
                    mail.Body = "Notifiaction pour une nouvelle demande d'audit de clôture";

                    string htmlContent = this.RenderViewToString("~/Views/MailTemplates/DemandeAudit.cshtml", rapport);

                    // Préparation du corps du mail
                    AlternateView av = AlternateView.CreateAlternateViewFromString(htmlContent, null, System.Net.Mime.MediaTypeNames.Text.Html);

                    string Logo8D = Server.MapPath(Url.Content("~/Images/Logo/Logo8D.png")); //string.Format("http://{0}{1}{2}", Request.Url.Host, Request.Url.Port > 0 ? ":" + Request.Url.Port : "", Url.Content("/Images/Caper.png"));
                    //            
                    if (Logo8D != string.Empty)
                    {
                        LinkedResource lr = new LinkedResource(Logo8D, System.Net.Mime.MediaTypeNames.Image.Jpeg);
                        lr.ContentId = "image2";
                        av.LinkedResources.Add(lr);
                    }

                    // Ajout du corps du mail
                    mail.AlternateViews.Add(av);

                    // configuration de l'SMTP
                    smtpClient.Host = mail_smtp_host;// "smtp.gmail.com";
                    if (mail_smtp_port > 0) smtpClient.Port = mail_smtp_port;// 587;
                    if (!string.IsNullOrEmpty(mail_smtp_credentials_username))
                    {
                        smtpClient.Credentials = new System.Net.NetworkCredential(
                            mail_smtp_credentials_username,
                            mail_smtp_credentials_password);//"from@gmail.com","Password");
                    }
                    smtpClient.EnableSsl = mail_smtp_enablessl;// true;

                    // et ça s'envole
                    smtpClient.Send(mail);
                }
                catch (Exception ex)
                {
                    ModelState.AddModelError("",
                           "Une erreur es survenue lors de l'envoi de mail. Exception:" + ex.Message);

                }
            }



        }

        //private void sendMails(Workflow wf)
        //{
        //    MailMessage mail = new MailMessage();
        //    SmtpClient smtpClient = new SmtpClient();

        //    // les paramètres de SMTP
        //    string mail_smtp_host = context.Parametres
        //                .Where(x => x.Param == "mail_smtp_host").FirstOrDefault().Valeur;

        //    int mail_smtp_port = 25; int.TryParse(context.Parametres
        //                .Where(x => x.Param == "mail_smtp_port").FirstOrDefault().Valeur, out mail_smtp_port);

        //    string mail_smtp_credentials_username = context.Parametres
        //                .Where(x => x.Param == "mail_smtp_credentials_username").FirstOrDefault().Valeur;
        //    string mail_smtp_credentials_password = context.Parametres
        //                .Where(x => x.Param == "mail_smtp_credentials_password").FirstOrDefault().Valeur;

        //    bool mail_smtp_enablessl = false; bool.TryParse(context.Parametres
        //                .Where(x => x.Param == "mail_smtp_enablessl").FirstOrDefault().Valeur, out mail_smtp_enablessl);

        //    // les paramètres du mail destinateur
        //    string mail_from_mail = context.Parametres
        //                .Where(x => x.Param == "mail_from_mail").FirstOrDefault().Valeur;
        //    string mail_from_name = context.Parametres
        //                .Where(x => x.Param == "mail_from_name").FirstOrDefault().Valeur;
        //    string mail_cc_mail = context.Parametres
        //                .Where(x => x.Param == "mail_cc_mail").FirstOrDefault().Valeur;
        //    string mail_cc_name = context.Parametres
        //                .Where(x => x.Param == "mail_cc_name").FirstOrDefault().Valeur;

        //    string logoPath = context.Parametres
        //                .Where(x => x.Param == "site_chemin_logo").FirstOrDefault().Valeur;

        //    if (!string.IsNullOrEmpty(mail_smtp_host))
        //    {
        //        try
        //        {
        //            mail.From = new MailAddress(mail_from_mail, mail_from_name);

        //            if (!string.IsNullOrEmpty(mail_cc_mail))
        //                mail.CC.Add(new MailAddress(mail_cc_mail, mail_cc_name));

        //            // Affectation des destinataires du mail
        //            foreach (var signature in wf.Signatures)
        //            {
        //                if (!string.IsNullOrEmpty(signature.Signataire.Mail))
        //                    mail.To.Add(new MailAddress(signature.Signataire.Mail, signature.Signataire.NomComplet));
        //            }
        //            mail.Subject = "CAPER : Notifiaction pour validation de signature de la fiche [" + wf.FicheVersion.Fiche.Code + "-" + wf.FicheVersion.Indice + "]";
        //            mail.IsBodyHtml = true;
        //            mail.Body = "Notifiaction pour validation de signature de fiche en attente";


        //            // attacher le ficher à valider au mail
        //            //FileStream fs = new FileStream("E:\\TestFolder\\test.pdf",FileMode.Open, FileAccess.Read);
        //            //Attachment a = new Attachment(fs, "test.pdf", MediaTypeNames.Application.Octet);
        //            //mail.Attachments.Add(a);

        //            // préparation du corps du mail
        //            string htmlContent = "<html><body>" +
        //                "<h1 style='color:darkred'>Notifiaction pour validation de Workflow</h1>" +
        //                "<hr>" +
        //                "<p>Une fiche vient d'être émise, vous êtres invité à la valider.</p>" +
        //                "<table style='border-collapse:collapse;border:1px solid black'>" +
        //                "<tr style='border-collapse:collapse;border:1px solid black'>" +
        //                "    <th style='font-weight:bold;text-align:right'>Fiche : </th><td>" + wf.FicheVersion.Fiche.Code + "-" + wf.FicheVersion.Indice + "</td></tr>" +
        //                "<tr style='border-collapse:collapse;border:1px solid black'>" +
        //                "    <th style='font-weight:bold;text-align:right'>Emetteur : </th><td>" + wf.FicheVersion.Emetteur.NomComplet + "</td></tr>" +
        //                "<tr style='border-collapse:collapse;border:1px solid black'>" +
        //                "    <th style='font-weight:bold;text-align:right'>Date Emission : </th><td>" + wf.FicheVersion.DateEmission + "</td></tr>" +
        //                "<tr style='border-collapse:collapse;border:1px solid black'>" +
        //                "    <th style='font-weight:bold;text-align:right'>Description de la Fiche : </th><td>" + wf.FicheVersion.Fiche.Designation + "</td></tr>" +
        //                "<tr style='border-collapse:collapse;border:1px solid black'>" +
        //                "    <th style='font-weight:bold;text-align:right'>Description de la Version : </th><td>" + wf.FicheVersion.Description + "</td></tr>" +
        //                "</table>" +
        //                "<hr>" +
        //                "<p>Procéder à <a href=\"" + string.Format("http://{0}{1}{2}", Request.Url.Host, Request.Url.Port > 0 ? ":" + Request.Url.Port : "", Url.Action("Signer", "Workflows")) + "\">la validation de signatures</a>.</p>" +
        //                "<hr>" +
        //                "<img src=\"cid:image2\">" +
        //                "<img src=\"cid:image1\">" +
        //                "</body></html>";

        //            AlternateView av = AlternateView.CreateAlternateViewFromString(htmlContent, null, MediaTypeNames.Text.Html);

        //            // attacher le logo au corps du mail
        //            if (System.IO.File.Exists(logoPath))
        //            {
        //                LinkedResource lr = new LinkedResource(logoPath, MediaTypeNames.Image.Jpeg);
        //                lr.ContentId = "image1";
        //                av.LinkedResources.Add(lr);
        //            }
        //            string caperLogo = Server.MapPath(Url.Content("~/Images/Caper.png")); //string.Format("http://{0}{1}{2}", Request.Url.Host, Request.Url.Port > 0 ? ":" + Request.Url.Port : "", Url.Content("/Images/Caper.png"));
        //            if (caperLogo != string.Empty)
        //            {
        //                LinkedResource lr = new LinkedResource(caperLogo, MediaTypeNames.Image.Jpeg);
        //                lr.ContentId = "image2";
        //                av.LinkedResources.Add(lr);
        //            }

        //            // ajout du corps du mail
        //            mail.AlternateViews.Add(av);

        //            // configuration de l'SMTP
        //            smtpClient.Host = mail_smtp_host;// "smtp.gmail.com";
        //            if (mail_smtp_port > 0) smtpClient.Port = mail_smtp_port;// 587;
        //            if (!string.IsNullOrEmpty(mail_smtp_credentials_username))
        //            {
        //                smtpClient.Credentials = new System.Net.NetworkCredential(
        //                    mail_smtp_credentials_username,
        //                    mail_smtp_credentials_password);//"from@gmail.com","Password");
        //            }
        //            smtpClient.EnableSsl = mail_smtp_enablessl;// true;

        //            // et ça s'envole
        //            smtpClient.Send(mail);

        //        }
        //        catch (Exception ex)
        //        {
        //            ModelState.AddModelError("",
        //                   "Une erreur es survenue lors de l'envoi de mail. Exception:" + ex.Message);

        //        }
        //    }
        //    else
        //    {
        //        ModelState.AddModelError("",
        //                   "Un problème de paramètrage dans la Base de données, merci de contacter votre administrateur système.");
        //    }
        //}

        [HttpPost]
        public JsonResult EditField(int pk, string name, string value)
        {
            var r = new
            {
                state = false,
                message = "Pas d'exécution."
            };
            if (pk > 0)
            {
                Rapport rapport = db.RapportSet.Find(pk);
                if (rapport == null)
                {
                    r = new
                    {
                        state = false,
                        message = "Clé pour Rapport erroné."
                    };
                }
                else
                {
                    try
                    {
                        Type t = rapport.GetType().GetProperty(name).GetType();
                        System.Reflection.PropertyInfo propertyInfo = rapport.GetType().GetProperty(name);//, t).SetValue(rapport, value);
                        //propertyInfo.SetValue(rapport, Convert.ChangeType(value, propertyInfo.PropertyType), null);

                        Type type = Nullable.GetUnderlyingType(propertyInfo.PropertyType)
                            ?? propertyInfo.PropertyType;

                        object safeValue = (value == null) ? null
                                                           : Convert.ChangeType(value, type);

                        propertyInfo.SetValue(rapport, safeValue, null);

                        //Type type = your_class.GetType();
                        //PropertyInfo propinfo = type.GetProperty("hasBeenPaid");
                        //value = propinfo.GetValue(your_class, null);

                        rapport.DateMaj = DateTime.Now;
                        db.Entry(rapport).State = EntityState.Modified;

                        db.SaveChanges();

                        r = new
                        {
                            state = true,
                            message = "Valeur changé."
                        };
                    }
                    catch (Exception e)
                    {
                        r = new
                        {
                            state = false,
                            message = "Exception : " + e.Message + "\n" + e.StackTrace + "\n" + e.InnerException
                        };
                    }
                }
            }
            else
            {
                r = new
                {
                    state = false,
                    message = "Clé pour Rapport invalide."
                };
            }

            return Json(r, JsonRequestBehavior.AllowGet);
        }
        //http://www.albahari.com/nutshell/predicatebuilder.aspx

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pk">L'identifiant du Rapport, RapportId</param>
        /// <param name="name">Le nom du champ à modifier composé de Cause_[Type]_[Cause5M]_[Attribut à modifier]_[Ordre]</param>
        /// <param name="value"></param>
        /// <returns></returns>
        /// 

     


        [HttpPost]
        public JsonResult Edit4DCauses(int pk, string name, string value)
        {

            var nameSplit = name.Split('_');
            var r = new
            {
                state = false,
                message = "Pas d'exécution."
            };
            #region tab3
            if (nameSplit[1] == "RevueAMDEC")
            {
                string type1 = nameSplit.Length > 1 ? nameSplit[1] : "";
                string attribut1 = nameSplit.Length > 2? nameSplit[2] : "";
               // int ordre1; string s_ordre1 = nameSplit.Length > 3 ? nameSplit[3] : ""; int.TryParse(s_ordre1, out ordre1);


                Response.StatusCode = (int)HttpStatusCode.Accepted;
                if (pk > 0)
                {
                    Rapport rapport = db.RapportSet.Find(pk);
                    if (rapport == null)
                    {
                        r = new
                        {
                            state = false,
                            message = "Clé pour Rapport erroné."
                        };
                        Response.StatusCode = (int)HttpStatusCode.BadRequest;
                    }
                    else
                    {
                        //if (ordre1 > 0)
                        {

                        }
                        var cause = rapport.D4Causes.Where(x =>
                            x.Type == type1 
                           
                        //    x.Ordre == ordre1
                            ).FirstOrDefault();
                        try
                        {
                            if (cause == null)
                            {
                                // nouvelle cause à ajouter
                                cause = new D4Causes
                                {
                                    Type = type1,
                                    
                                   // Ordre = ordre1,
                                    //RapportId = pk

                                };
                                Type t = cause.GetType().GetProperty(attribut1).GetType();
                                System.Reflection.PropertyInfo propertyInfo = cause.GetType().GetProperty(attribut1);//, t).SetValue(rapport, value);
                                                                                                                    //propertyInfo.SetValue(rapport, Convert.ChangeType(value, propertyInfo.PropertyType), null);

                                Type propertyType = Nullable.GetUnderlyingType(propertyInfo.PropertyType)
                                    ?? propertyInfo.PropertyType;

                                object safeValue = (value == null) ? null
                                                                   : Convert.ChangeType(value, propertyType);

                                propertyInfo.SetValue(cause, safeValue, null);

                                //Type type = your_class.GetType();
                                //PropertyInfo propinfo = type.GetProperty("hasBeenPaid");
                                //value = propinfo.GetValue(your_class, null);

                                rapport.D4Causes.Add(cause);

                            }
                            else
                            {

                                Type t = cause.GetType().GetProperty(attribut1).GetType();
                                System.Reflection.PropertyInfo propertyInfo = cause.GetType().GetProperty(attribut1);

                                Type propertyType = Nullable.GetUnderlyingType(propertyInfo.PropertyType)
                                    ?? propertyInfo.PropertyType;

                                object safeValue = (value == null) ? null
                                                                   : Convert.ChangeType(value, propertyType);

                                propertyInfo.SetValue(cause, safeValue, null);
                                db.Entry(cause).State = System.Data.Entity.EntityState.Modified;
                            }

                            rapport.DateMaj = DateTime.Now;
                            db.Entry(rapport).State = EntityState.Modified;


                            try
                            {
                                db.SaveChanges();
                            }
                            catch (DbEntityValidationException ex)
                            {
                                foreach (var entityValidationErrors in ex.EntityValidationErrors)
                                {
                                    foreach (var validationError in entityValidationErrors.ValidationErrors)
                                    {
                                        Response.Write("Property: " + validationError.PropertyName + " Error: " + validationError.ErrorMessage);
                                    }
                                }
                            }

                            r = new
                            {
                                state = true,
                                message = "Valeur changé."
                            };
                            Response.StatusCode = (int)HttpStatusCode.OK;
                        }
                        catch (Exception e)
                        {
                            r = new
                            {
                                state = false,
                                message = "Exception : " + e.Message
                            };
                            Response.StatusCode = (int)HttpStatusCode.InternalServerError;
                        }
                    }
                }
                else
                {
                    r = new
                    {
                        state = false,
                        message = "Clé pour Rapport invalide."
                    };
                    Response.StatusCode = (int)HttpStatusCode.BadRequest;
                }

            }
            #endregion
            #region tb1&2
            else
            {
                string type = nameSplit.Length > 1 ? nameSplit[1] : "";
                string cause5M = nameSplit.Length > 2 ? nameSplit[2] : "";
                string attribut = nameSplit.Length > 3 ? nameSplit[3] : "";
                int ordre; string s_ordre = nameSplit.Length > 4 ? nameSplit[4] : ""; int.TryParse(s_ordre, out ordre);

                Response.StatusCode = (int)HttpStatusCode.Accepted;
                if (pk > 0)
                {
                    Rapport rapport = db.RapportSet.Find(pk);
                    if (rapport == null)
                    {
                        r = new
                        {
                            state = false,
                            message = "Clé pour Rapport erroné."
                        };
                        Response.StatusCode = (int)HttpStatusCode.BadRequest;
                    }
                    else
                    {
                        if (ordre > 0)
                        {

                        }
                        var cause = rapport.D4Causes.Where(x =>
                            x.Type == type &&
                            x.Cause5M == cause5M &&
                            x.Ordre == ordre
                            ).FirstOrDefault();
                        try
                        {
                            if (cause == null)
                            {
                                // nouvelle cause à ajouter
                                cause = new D4Causes
                                {
                                    Type = type,
                                    Cause5M = cause5M,
                                    Ordre = ordre,
                                    //RapportId = pk

                                };
                                Type t = cause.GetType().GetProperty(attribut).GetType();
                                System.Reflection.PropertyInfo propertyInfo = cause.GetType().GetProperty(attribut);//, t).SetValue(rapport, value);
                                                                                                                    //propertyInfo.SetValue(rapport, Convert.ChangeType(value, propertyInfo.PropertyType), null);

                                Type propertyType = Nullable.GetUnderlyingType(propertyInfo.PropertyType)
                                    ?? propertyInfo.PropertyType;

                                object safeValue = (value == null) ? null
                                                                   : Convert.ChangeType(value, propertyType);

                                propertyInfo.SetValue(cause, safeValue, null);

                                //Type type = your_class.GetType();
                                //PropertyInfo propinfo = type.GetProperty("hasBeenPaid");
                                //value = propinfo.GetValue(your_class, null);

                                rapport.D4Causes.Add(cause);

                            }
                            else
                            {

                                Type t = cause.GetType().GetProperty(attribut).GetType();
                                System.Reflection.PropertyInfo propertyInfo = cause.GetType().GetProperty(attribut);

                                Type propertyType = Nullable.GetUnderlyingType(propertyInfo.PropertyType)
                                    ?? propertyInfo.PropertyType;

                                object safeValue = (value == null) ? null
                                                                   : Convert.ChangeType(value, propertyType);

                                propertyInfo.SetValue(cause, safeValue, null);
                                db.Entry(cause).State = System.Data.Entity.EntityState.Modified;
                            }

                            rapport.DateMaj = DateTime.Now;
                            db.Entry(rapport).State = EntityState.Modified;

                            try
                            {
                                db.SaveChanges();
                            }
                            catch (DbEntityValidationException ex)
                            {
                                foreach (var entityValidationErrors in ex.EntityValidationErrors)
                                {
                                    foreach (var validationError in entityValidationErrors.ValidationErrors)
                                    {
                                        Response.Write("Property: " + validationError.PropertyName + " Error: " + validationError.ErrorMessage);
                                    }
                                }
                            }

                            r = new
                            {
                                state = true,
                                message = "Valeur changé."
                            };
                            Response.StatusCode = (int)HttpStatusCode.OK;
                        }
                        catch (Exception e)
                        {
                            r = new
                            {
                                state = false,
                                message = "Exception : " + e.Message
                            };
                            Response.StatusCode = (int)HttpStatusCode.InternalServerError;
                        }
                    }
                }
                else
                {
                    r = new
                    {
                        state = false,
                        message = "Clé pour Rapport invalide."
                    };
                    Response.StatusCode = (int)HttpStatusCode.BadRequest;
                }

            }
            #endregion
            return Json(r, JsonRequestBehavior.AllowGet);
        }
    

        [HttpPost]
        public JsonResult EditD5Actions(int pk, string name, string value)
        {
            var nameSplit = name.Split('_');
            string attribut = nameSplit.Length > 1 ? nameSplit[1] : ""; 
            int actionId; string s_actionId = nameSplit.Length > 2 ? nameSplit[2] : ""; int.TryParse(s_actionId, out actionId);

            var r = new
            {
                state = false,
                message = "Pas d'exécution."
            };
            Response.StatusCode = (int)HttpStatusCode.Accepted;
            if (actionId > 0)
            {
                var action = db.D5ActionsCorrectives.Find(actionId);
                if (action == null)
                {
                    r = new
                    {
                        state = false,
                        message = "Clé pour D5ActionsCorrectives erroné."
                    };
                    Response.StatusCode = (int)HttpStatusCode.BadRequest;
                }
                else
                {
                    try
                    {
                        System.Reflection.PropertyInfo propertyInfo = action.GetType().GetProperty(attribut);

                        Type propertyType = Nullable.GetUnderlyingType(propertyInfo.PropertyType)
                            ?? propertyInfo.PropertyType;

                        object safeValue = (string.IsNullOrEmpty(value)) ? null
                                                           : Convert.ChangeType(value, propertyType);

                        propertyInfo.SetValue(action, safeValue, null);

                        //this.MarshalEditObject(attribut, value, ref (action as object));

                        db.Entry(action).State = System.Data.Entity.EntityState.Modified;

                        var rapport = db.RapportSet.Find(action.RapportId);
                        rapport.DateMaj = DateTime.Now;
                        db.Entry(rapport).State = EntityState.Modified;

                        db.SaveChanges();

                        r = new
                        {
                            state = true,
                            message = "Valeur changé."
                        };
                        Response.StatusCode = (int)HttpStatusCode.OK;
                    }
                    catch (Exception e)
                    {
                        r = new
                        {
                            state = false,
                            message = "Exception : " + e.Message
                        };
                        Response.StatusCode = (int)HttpStatusCode.InternalServerError;
                    }
                }
            }
            else
            {
                r = new
                {
                    state = false,
                    message = "Clé pour D5ActionsCorrectives invalide."
                };
                Response.StatusCode = (int)HttpStatusCode.BadRequest;
            }

            return Json(r, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public JsonResult EditD7Actions(int pk, string name, string value)
        {
            var nameSplit = name.Split('_');
            string attribut = nameSplit.Length > 1 ? nameSplit[1] : "";
            int actionId; string s_actionId = nameSplit.Length > 2 ? nameSplit[2] : ""; int.TryParse(s_actionId, out actionId);

            var r = new
            {
                state = false,
                message = "Pas d'exécution."
            };
            Response.StatusCode = (int)HttpStatusCode.Accepted;
            if (actionId > 0)
            {
                var action = db.D7ActionsGeneralisation.Find(actionId);
                if (action == null)
                {
                    r = new
                    {
                        state = false,
                        message = "Clé pour D7ActionsGeneralisation erroné."
                    };
                    Response.StatusCode = (int)HttpStatusCode.BadRequest;
                }
                else
                {
                    try
                    {
                        System.Reflection.PropertyInfo propertyInfo = action.GetType().GetProperty(attribut);

                        Type propertyType = Nullable.GetUnderlyingType(propertyInfo.PropertyType)
                            ?? propertyInfo.PropertyType;

                        object safeValue = (value == null) ? null
                                                           : Convert.ChangeType(value, propertyType);

                        propertyInfo.SetValue(action, safeValue, null);

                        //this.MarshalEditObject(attribut, value, ref (action as object));

                        db.Entry(action).State = System.Data.Entity.EntityState.Modified;


                        var rapport = db.RapportSet.Find(action.RapportId);
                        rapport.DateMaj = DateTime.Now;
                        db.Entry(rapport).State = EntityState.Modified;

                        db.SaveChanges();

                        r = new
                        {
                            state = true,
                            message = "Valeur changé."
                        };
                        Response.StatusCode = (int)HttpStatusCode.OK;
                    }
                    catch (Exception e)
                    {
                        r = new
                        {
                            state = false,
                            message = "Exception : " + e.Message
                        };
                        Response.StatusCode = (int)HttpStatusCode.InternalServerError;
                    }
                }
            }
            else
            {
                r = new
                {
                    state = false,
                    message = "Clé pour D7ActionsGeneralisation invalide."
                };
                Response.StatusCode = (int)HttpStatusCode.BadRequest;
            }

            return Json(r, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public JsonResult GetInfoPersonneJson(int id)
        {
            var personne = db.PersonnelSet.Find(id);
            return Json(new
            {
                personne.Matricule,
                personne.Nom,
                personne.Prenom,
                personne.NomComplet,
                personne.CentreCout,
                personne.ChefHierarchique,
                personne.Fonction,
                personne.Mail,
                personne.Service,
                personne.Tel,
                personne.Windows
            }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult PersonnelListJson(string q, int page, int page_limit)
        {
            var persons = db.PersonnelSet
                .Where(x => x.Nom.Contains(q)||x.Prenom.Contains(q))
                .Select(x => new
                {
                    id = x.PersonnelId,
                    text = string.Concat(x.Nom, " ", x.Prenom ?? "") /*,
                    PersonnelId = x.PersonnelId,
                    x.Nom,
                    x.Prenom,
                    NomComplet = string.Concat(x.Nom, " ", x.Prenom ?? ""))*/
                }).ToList();
                

            return Json(persons, JsonRequestBehavior.AllowGet);
        }

        public JsonResult FonctionsListJson(string q, int page, int page_limit)
        {
            var persons = db.FonctionSet
                .Where(x => x.Titre.Contains(q) && x.Actif==true)
                .Select(x => new
                {
                    id = x.FonctionId,
                    text = x.Titre
                }).ToList();


            return Json(persons, JsonRequestBehavior.AllowGet);
        }

        public ActionResult NouveauMembreEquipe(int rapportId, FormCollection form)
        {
            int membre;
            int fonction;            
            int.TryParse(form["membre"], out membre);
            int.TryParse(form["fonction"], out fonction);
            if (membre > 0 /*&& fonction > 0*/)
            {
                var m = new MembreEquipe
                {
                    RapportId = rapportId,
                    FonctionId = fonction > 0 ? fonction as Nullable<int> : null,
                    PersonnelId = membre
                };
                db.MembreEquipeSet.Add(m);
                /// test list diffusion ///
              
                var rapport = db.RapportSet.Find(rapportId);
                rapport.DateMaj = DateTime.Now;
                db.Entry(rapport).State = EntityState.Modified;
              //  rapport.ListesDiffusion.Add(m);
                db.SaveChanges();
            }

            return RedirectToAction("Rapport", new { id = rapportId });
        }

        [HttpPost]
        public JsonResult SupprimerMembreEquipe(int membreEquipeId)
        {
            var r = new
            {
                state = false,
                message = "Pas d'exécution."
            };
            if (membreEquipeId > 0)
            {
                var membre = db.MembreEquipeSet.Find(membreEquipeId);

                if (membre != null)
                {
                    db.MembreEquipeSet.Remove(membre);
                    db.Entry(membre).State = System.Data.Entity.EntityState.Deleted;

                    var rapport = db.RapportSet.Find(membre.RapportId);
                    rapport.DateMaj = DateTime.Now;
                    db.Entry(rapport).State = EntityState.Modified;

                    db.SaveChanges();
                    r = new
                    {
                        state = true,
                        message = "Membre Supprimé."
                    };
                }
                else
                {
                    r = new
                    {
                        state = false,
                        message = "MembreEquipeSet retourne NULL. membreEquipeId = " + membreEquipeId
                    };
                }
            }
            else
            {
                r = new
                {
                    state = false,
                    message = "membreEquipeId = " + membreEquipeId
                };
            }
            return Json(r, JsonRequestBehavior.AllowGet);
        }

        public JsonResult UpdateEquipe(int id, FormCollection form)
        {
            var r = new
            {
                state = false,
                message = "Pas d'exécution."
            };
            if (id > 0)
            {
                Rapport rapport = db.RapportSet.Find(id);

                r = new
                {
                    state = true,
                    message = "...."
                };
            }

            return Json(r, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult AjouterD5ActionCorrective(D5ActionsCorrectives newAction, FormCollection form)
        {
            //var action = new D5ActionsCorrectives();
            try
            {
                if (newAction.RapportId > 0)
                {
                    db.D5ActionsCorrectives.Add(newAction);
                    db.SaveChanges();
                }
                else
                {
                    ViewBag.error = "newAction.RapportId <= 0 ";
                    Response.StatusCode = (int)HttpStatusCode.BadRequest;
                }
                //var formCollection = form.AllKeys.ToDictionary(k => k, v => form[v]);

                //foreach (var pair in formCollection)
                //{
                //    System.Reflection.PropertyInfo propertyInfo = action.GetType().GetProperty(pair.Key);
                //    Type propertyType = Nullable.GetUnderlyingType(propertyInfo.PropertyType)
                //       ?? propertyInfo.PropertyType;
                //    object safeValue = (pair.Value == null) ? null : Convert.ChangeType(pair.Value, propertyType);
                //    action.GetType().GetProperty(pair.Key).SetValue(action, safeValue, null);
                //}

            }
            catch (Exception e)
            {
                ViewBag.error = e.Message;
                Response.StatusCode = (int)HttpStatusCode.BadRequest;
            }
            return PartialView(newAction);
        }
        [HttpPost]
        public ActionResult AjouterD7GeneralisationActions(D7ActionsGeneralisation newAction, FormCollection form)
        {
            //var action = new D5ActionsCorrectives();
            try
            {
                if (newAction.RapportId > 0)
                {
                    db.D7ActionsGeneralisation.Add(newAction);
                    db.SaveChanges();
                }
                else
                {
                    ViewBag.error = "newAction.RapportId <= 0 ";
                    Response.StatusCode = (int)HttpStatusCode.BadRequest;
                }
                //var formCollection = form.AllKeys.ToDictionary(k => k, v => form[v]);

                //foreach (var pair in formCollection)
                //{
                //    System.Reflection.PropertyInfo propertyInfo = action.GetType().GetProperty(pair.Key);
                //    Type propertyType = Nullable.GetUnderlyingType(propertyInfo.PropertyType)
                //       ?? propertyInfo.PropertyType;
                //    object safeValue = (pair.Value == null) ? null : Convert.ChangeType(pair.Value, propertyType);
                //    action.GetType().GetProperty(pair.Key).SetValue(action, safeValue, null);
                //}

            }
            catch (Exception e)
            {
                ViewBag.error = e.Message;
                Response.StatusCode = (int)HttpStatusCode.BadRequest;
            }
            return PartialView(newAction);
        }
        
        [HttpPost]
        public JsonResult SupprimerD5ActionCorrective(int actionId)
        {
            var r = new
            {
                state = false,
                message = "Pas d'exécution."
            };
            if (actionId > 0)
            {
                var action = db.D5ActionsCorrectives.Find(actionId);

                if (action != null)
                {
                    db.D5ActionsCorrectives.Remove(action);
                    db.Entry(action).State = System.Data.Entity.EntityState.Deleted;
                    db.SaveChanges();
                    r = new
                    {
                        state = true,
                        message = "Action Supprimé."
                    };
                }
                else
                {
                    r = new
                    {
                        state = false,
                        message = "D5ActionsCorrectives retourne NULL. membreEquipeId = " + actionId
                    };
                }
            }
            else
            {
                r = new
                {
                    state = false,
                    message = "actionId = " + actionId
                };
            }
            return Json(r, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public JsonResult SupprimerD7GeneralisationActions(int actionId)
        {
            var r = new
            {
                state = false,
                message = "Pas d'exécution."
            };
            if (actionId > 0)
            {
                var action = db.D7ActionsGeneralisation.Find(actionId);

                if (action != null)
                {
                    db.D7ActionsGeneralisation.Remove(action);
                    if (db.Entry(action).State != System.Data.Entity.EntityState.Deleted)
                        db.Entry(action).State = System.Data.Entity.EntityState.Deleted;

                    db.SaveChanges();
                    r = new
                    {
                        state = true,
                        message = "D7ActionsGeneralisation Supprimé."
                    };
                }
                else
                {
                    r = new
                    {
                        state = false,
                        message = "Action non existante (D7ActionsGeneralisation retourne NULL). membreEquipeId = " + actionId
                    };
                }
            }
            else
            {
                r = new
                {
                    state = false,
                    message = "actionId = " + actionId
                };
            }
            return Json(r, JsonRequestBehavior.AllowGet);
        }


        public ActionResult ListeDiffusion(int id, string etape = "")
        {
            var rapport = db.RapportSet.Find(id);
            ListesDiffusion listeDiffusion = null;
            if (!string.IsNullOrEmpty(etape))
            {
                // TODO : Rendre la liste de diffusion non modifiable à l'affichage : cacher les boutons ADD et Remove item

                listeDiffusion = rapport.ListesDiffusion.Where(x => x.EtapeDiffusion.Equals(etape)).FirstOrDefault();
                if(listeDiffusion==null)
                {
                    listeDiffusion = new ListesDiffusion
                    {
                        EtapeDiffusion = etape,
                        RapportId = id
                    };

                    db.ListesDiffusion.Add(listeDiffusion);
                    db.SaveChanges();
                }
            }
            var CondLD = db.ListeDiffusionObligatoire.Where(x => x.EtapeDiffusion == etape).FirstOrDefault();
            
                return PartialView(new ListesDiffusionViewModel
                {
                    ListesDiffusion = listeDiffusion,
                    ListeDiffusionObligatoire = CondLD ?? null
                });
           
        }
        public class ListesDiffusionViewModel
        {
            public ListeDiffusionObligatoire ListeDiffusionObligatoire { get; set; }
            public ListesDiffusion ListesDiffusion { get; set; }
        }

        public JsonResult ListeDiffusionSupprimerMembre(int listeDiffusionId ,int membreId)
        {
            var r = new
            {
                state = false,
                message = "Pas d'exécution."
            };
            if (listeDiffusionId > 0)
            {
                var liste = db.ListesDiffusion.Find(listeDiffusionId);

                if (liste != null)
                {
                    if (membreId > 0)
                    {
                        var membre = liste.PersonnelSet.Where(x=>x.PersonnelId==membreId).FirstOrDefault() ;
                        if (membre != null)
                        {
                            liste.PersonnelSet.Remove(membre);                            
                            
                            db.SaveChanges();
                            r = new
                            {
                                state = true,
                                message = "Membre Supprimé."
                            };
                        }
                        else
                        {
                            r = new
                            {
                                state = false,
                                message = "membreId = " + membreId + " n'appartiens pas à la liste listeDiffusionId= " + listeDiffusionId
                            };
                        }
                    }
                    else
                    {
                        r = new
                        {
                            state = false,
                            message = "membreId = " + membreId
                        };
                    }
                }
                else
                {
                    r = new
                    {
                        state = false,
                        message = "Liste de diffusion non existante (D7ActionsGeneralisation retourne NULL). listeDiffusionId = " + listeDiffusionId
                    };
                }
            }
            else
            {
                r = new
                {
                    state = false,
                    message = "listeDiffusionId = " + listeDiffusionId
                };
            }
            return Json(r, JsonRequestBehavior.AllowGet);
        }

        public JsonResult ListeDiffusionAjouterMembre(int listeDiffusionId ,int membreId)
        {
            var r = new
            {
                state = false,
                message = "Pas d'exécution."
            };
            if (listeDiffusionId > 0)
            {
                var liste = db.ListesDiffusion.Find(listeDiffusionId);

                if (liste != null)
                {
                    if (membreId > 0)
                    {
                        var p = db.PersonnelSet.Find(membreId);
                        if (p != null)
                        {
                            liste.PersonnelSet.Add(p);
                            db.SaveChanges();

                            r = new
                            {
                                state = true,
                                message = "Membre Ajouté."
                            };
                        }
                    }
                    else
                    {
                        r = new
                        {
                            state = false,
                            message = "membreId = " + membreId
                        };
                    }
                }
                else
                {
                    r = new
                    {
                        state = false,
                        message = "Liste de diffusion non existante (D7ActionsGeneralisation retourne NULL). listeDiffusionId = " + listeDiffusionId
                    };
                }
            }
            else
            {
                r = new
                {
                    state = false,
                    message = "listeDiffusionId = " + listeDiffusionId
                };
            }
            return Json(r, JsonRequestBehavior.AllowGet);
        }

        //http://www.codeproject.com/Articles/690136/All-About-TransactionScope
        //http://transactionalfilemgr.codeplex.com/
        //http://stackoverflow.com/questions/7939339/how-to-write-a-transaction-to-cover-moving-a-file-and-inserting-record-in-databa
        //https://github.com/blueimp/jQuery-File-Upload/wiki/Options
        //http://garfbradaz.github.io/MvcImage/
        //http://garfbradazweb.wordpress.com/2011/08/03/mvc-3-image-previewjquery/
        [HttpPost]
        public ContentResult UploadFiles(int id, string type)
        {
            if (id > 0)
            {
                if (!string.IsNullOrEmpty(type.Trim()))
                {
                    type = type.Trim().ToUpper();

                    var rapport = db.RapportSet.Find(id);
                    var r = new List<UploadFilesResult>();
                    var rids = new List<int>();

                    //String path = ConfigurationManager.AppSettings["configFile"];

                    var FileStoragePath = // chemin physique d'entrepos de fichiers pour GRC
                    db.Configs.Find("FileStoragePath").Val;
                    //WebConfigurationManager.AppSettings["FileStoragePath"];

                    string ImgKey = // Clé d'identification d'une image
                        "P_" + type +"_" + DateTime.Now.Ticks.ToString();

                    string destinationPath = 
                        Path.Combine(
                            FileStoragePath,
                            "Rapports",             // Dossier des Rapports
                            "R_" + rapport.RapportId, // Dossier du rapport en cours
                            "D2_Photos",            // Dossier des Photos
                            type          // Dossier des Photos OK ou NOK
                        );
                    //IFileManager fileManager = new TxFileManager();
                    using (TransactionScope scope = new TransactionScope())
                    {
                        foreach (string file in Request.Files)
                        {
                            // Traitement physique du fichier
                            HttpPostedFileBase hpf = Request.Files[file] as HttpPostedFileBase;
                            if (hpf.ContentLength == 0)
                                continue;

                            string fileName = Path.GetFileName(hpf.FileName);
                            string fileExt = Path.GetExtension(hpf.FileName);

                            string savedFileName = // chemin complet d'enregistrement du fichier
                                Path.Combine(
                                    destinationPath, 
                                    ImgKey + fileExt
                                );
                            if (!Directory.Exists(destinationPath))
                                Directory.CreateDirectory(destinationPath);

                            hpf.SaveAs(savedFileName);

                            r.Add(new UploadFilesResult()
                            {
                                Name = hpf.FileName,
                                Length = hpf.ContentLength,
                                Type = hpf.ContentType
                            });

                            System.Drawing.Image img = System.Drawing.Image.FromStream(hpf.InputStream);//.FromFile(@"c:\ggs\ggs Access\images\members\1.jpg");

                            // TODO : Traitement du fichier coté Base, on ajoute une nouvelle ligne dans la table des photos
                            var image = new Photo
                            {
                                Type = type,
                                Rapport = rapport,
                                MimeType = hpf.ContentType,
                                NomFichier = hpf.FileName,
                                ExtensionFichier = fileExt,
                                Key = ImgKey,
                                Width = img.Width,
                                Height = img.Height
                            };
                            db.PhotoSet.Add(image);
                            db.SaveChanges();
                            rids.Add(image.PhotoId);
                        }
                        scope.Complete();
                    }
                    return Content("{\"name\":\"" + r[0].Name + "\",\"type\":\"" + r[0].Type + "\",\"size\":\"" + string.Format("{0} bytes", r[0].Length) + "\",\"PhotoId\":\"" + rids[0].ToString() + "\"}", "application/json");
                }
                else
                {
                    return Content("{\"name\":\"[Nothing To Add]\",\"type\":\"[Pas de Type d'image donné (OK/NOK)]\",\"size\":\"[-0]\"}", "application/json");
                }
            }
            else
            {
                return Content("{\"name\":\"[Nothing To Add]\",\"type\":\"[Pas de Id Rapport donné]\",\"size\":\"[-0]\"}", "application/json");
            }
        }
        [HttpPost]
        public JsonResult DeleteImage(int id)
        {
            var r = new
            {
                status = "warning",
                message = "Nothing done."
            };
            // suppression de l'image avec l'id donné
            //string response = "{\"status\":\"warning\",\"message\":\"Nothing done.\"}";
            using (TransactionScope scope = new TransactionScope())
            {
                try
                {
                    var img = db.PhotoSet.Find(id);

                    var FileStoragePath =
                        db.Configs.Find("FileStoragePath").Val;
                    //WebConfigurationManager.AppSettings["FileStoragePath"];

                    var imgPath = Path.Combine(
                        FileStoragePath,
                            "Rapports",             // Dossier des Rapports
                            "R_" + img.Rapport.RapportId, // Dossier du rapport en cours
                            "D2_Photos",            // Dossier des Photos
                            img.Type.ToUpper(),          // Dossier des Photos OK ou NOK
                            img.Key + img.ExtensionFichier
                            );

                    System.IO.File.Delete(imgPath);

                    db.PhotoSet.Remove(img);
                    if (db.Entry(img).State != EntityState.Deleted)
                        db.Entry(img).State = EntityState.Deleted;
                    db.SaveChanges();

                    scope.Complete();
                    r = new
                    {
                        status = "success",
                        message = "File deleted."
                    };
                    //response = "{\"status\":\"success\",\"message\":\"File deleted\"}";
                }
                catch(Exception e)
                {
                    r = new
                    {
                        status = "danger",
                        message = "Error :" + e.Message + " "
                    };
                    //response = "{\"status\":\"danger\",\"message\":\"Error :" + e.Message + " \"}";
                }
            }
            return Json(r, JsonRequestBehavior.AllowGet);
            //return Content(response, "application/json");
        }

        public FileStreamResult GetImage(int id)
        {
            var dbImage = db.PhotoSet.Find(id);
            var FileStoragePath =
                    db.Configs.Find("FileStoragePath").Val;
                    //WebConfigurationManager.AppSettings["FileStoragePath"];
            //var s = System.IO.File.OpenRead("");
            //var stream = new MemoryStream(s.);
            string filePath = Path.Combine(
                FileStoragePath,
                "Rapports",                     // Dossier des Rapports
                "R_" + dbImage.Rapport.RapportId, // Dossier du rapport en cours
                "D2_Photos",                    // Dossier des Photos
                dbImage.Type.ToUpper(),         // Dossier des Photos OK ou NOK
                dbImage.Key + dbImage.ExtensionFichier
                );

            //FileStream stream = System.IO.File.Open(filePath, FileMode.Open);
            FileStream stream = System.IO.File.OpenRead(filePath);
            return File(stream, dbImage.MimeType, dbImage.NomFichier);

            //FileStream stream = System.IO.File.Open(@"C:\Users\G557454\Pictures\MSL\DSC_0001_BD.jpg", FileMode.Open);
            //return File(stream, "image/png", "DSC_0001_BD.jpg");
            //return File(stream, "text/plain", "your_file_name.txt");   
        }

        public ActionResult PhotosCroquis(int id, string type)
        {
            type = type.Trim().ToUpper();
            ViewBag.Type = type;
            var rapport = db.RapportSet.Find(id);
            var photos = rapport.Photo.Where(x => x.Type == type).ToList();
            return PartialView(photos);
        }

        #region Documents
         
        public ActionResult Documents(int id, string type)
        {
            var docs = db.RapportSet.Find(id).Documents.ToList();
            switch (type)
            {
                case "json":
                    return Json(docs.Select(x => new
                    {
                        x.DateAjout,
                        x.Description,
                        x.DocumentId,
                        x.ExtensionFichier,
                        x.Key,
                        x.Rapport_RapportId
                    }).ToList(), JsonRequestBehavior.AllowGet);  
                case "html":
                default:
                    return PartialView(docs); 
            }
        }

        public FileStreamResult Document(int id)
        {
            var doc = db.Documents.Find(id);
            var FileStoragePath =
                    db.Configs.Find("FileStoragePath").Val;
                    //WebConfigurationManager.AppSettings["FileStoragePath"];
            //var s = System.IO.File.OpenRead("");
            //var stream = new MemoryStream(s.);
            string filePath = Path.Combine(
                FileStoragePath,
                "Rapports",                     // Dossier des Rapports
                "R_" + doc.Rapport.RapportId, // Dossier du rapport en cours
                "Documents",                    // Dossier des Photos 
                doc.Key + doc.ExtensionFichier
                );

            FileStream stream = System.IO.File.Open(filePath, FileMode.Open);
            return File(stream, doc.MimeType, doc.NomFichier);
        }

        public ContentResult DocumentUpload(int id)
        {
            if (id > 0)
            {
                var rapport = db.RapportSet.Find(id);
                var r = new List<UploadFilesResult>();
                var rids = new List<int>();

                var FileStoragePath = // chemin physique d'entrepos de fichiers pour GRC
                    db.Configs.Find("FileStoragePath").Val;
                    //WebConfigurationManager.AppSettings["FileStoragePath"];

                string DocKey = // Clé d'identification d'un document
                    "D_" + DateTime.Now.Ticks.ToString() + "_" + rapport.RapportId;

                string destinationPath =
                    Path.Combine(
                        FileStoragePath,
                        "Rapports",             // Dossier des Rapports
                        "R_" + rapport.RapportId, // Dossier du rapport en cours
                        "Documents"            // Dossier des Photos
                    );
                //IFileManager fileManager = new TxFileManager();
                using (TransactionScope scope = new TransactionScope())
                {
                    foreach (string file in Request.Files)
                    {
                        // Traitement physique du fichier
                        HttpPostedFileBase hpf = Request.Files[file] as HttpPostedFileBase;
                        if (hpf.ContentLength == 0)
                            continue;

                        string fileName = Path.GetFileName(hpf.FileName);
                        string fileExt = Path.GetExtension(hpf.FileName);

                        string savedFileName = // chemin complet d'enregistrement du fichier
                            Path.Combine(
                                destinationPath,
                                DocKey + fileExt
                            );
                        if (!Directory.Exists(destinationPath))
                            Directory.CreateDirectory(destinationPath);

                        hpf.SaveAs(savedFileName);

                        r.Add(new UploadFilesResult()
                        {
                            Name = hpf.FileName,
                            Length = hpf.ContentLength,
                            Type = hpf.ContentType
                        });


                        // TODO : Traitement du fichier coté Base, on ajoute une nouvelle ligne dans la table des photos
                        var doc = new Documents
                        { 
                            Rapport = rapport,
                            MimeType = hpf.ContentType,
                            NomFichier = hpf.FileName,
                            ExtensionFichier = fileExt,
                            Key = DocKey,
                            DateAjout = DateTime.Now
                        };
                        db.Documents.Add(doc);
                        db.SaveChanges();
                        rids.Add(doc.DocumentId);
                    }
                    scope.Complete();
                }
                return Content("{\"name\":\"" + r[0].Name + "\",\"type\":\"" + r[0].Type + "\",\"size\":\"" + string.Format("{0} bytes", r[0].Length) + "\",\"PhotoId\":\"" + rids[0].ToString() + "\"}", "application/json");

            }
            else
            {
                return Content("{\"name\":\"[Nothing To Add]\",\"type\":\"[Pas de Id Rapport donné]\",\"size\":\"[-0]\"}", "application/json");
            }
        }
        public JsonResult DocumentDelete(int id)
        {
            var r = new
            {
                status = "warning",
                message = "Nothing done."
            };
            // suppression de l'image avec l'id donné
            //string response = "{\"status\":\"warning\",\"message\":\"Nothing done.\"}";
            using (TransactionScope scope = new TransactionScope())
            {
                try
                {
                    var doc = db.Documents.Find(id);

                    var FileStoragePath =
                        db.Configs.Find("FileStoragePath").Val;
                        //WebConfigurationManager.AppSettings["FileStoragePath"];

                    var docPath = Path.Combine(
                        FileStoragePath,
                            "Rapports",             // Dossier des Rapports
                            "R_" + doc.Rapport.RapportId, // Dossier du rapport en cours
                            "Documents",            // Dossier des Documents 
                            doc.Key + doc.ExtensionFichier
                            );

                    System.IO.File.Delete(docPath);

                    db.Documents.Remove(doc);
                    if (db.Entry(doc).State != EntityState.Deleted)
                        db.Entry(doc).State = EntityState.Deleted;
                    db.SaveChanges();

                    scope.Complete();
                    r = new
                    {
                        status = "success",
                        message = "Document deleted."
                    };
                    //response = "{\"status\":\"success\",\"message\":\"File deleted\"}";
                }
                catch (Exception e)
                {
                    r = new
                    {
                        status = "danger",
                        message = "Error :" + e.Message + " "
                    };
                    //response = "{\"status\":\"danger\",\"message\":\"Error :" + e.Message + " \"}";
                }
            }
            return Json(r, JsonRequestBehavior.AllowGet);
            //return Content(response, "application/json");
        }
        #endregion


        #region ############ Excel ################

        public ActionResult Excel(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var rapport = db.RapportSet.Find(id);
   
            var culture = CultureHelper.GetCurrentCulture();

            string fileName = "24_035_114-L_Rapport_8D_FR.xls";

            if (culture.ToLower().Contains("en"))
            {
                fileName = "24_035_114-L_Rapport_8D_EN.xls";
            }

            Workbook workbook = new Workbook(Path.Combine(baseDir, "TemplatesDocs", fileName));
            Worksheet sheet_1_2D = workbook.Worksheets[0];
            Worksheet sheet_3D = workbook.Worksheets[1];
            Worksheet sheet_4D = workbook.Worksheets[2];
            Worksheet sheet_5_8D = workbook.Worksheets[3];

            // *********************************************************
            //          Entêtes du document (pour les différentes pages)
            // *********************************************************
                     
            string InfoPiloteRapport = "Pilote :" + (rapport.Pilote != null ? rapport.Pilote.NomComplet : "") + Environment.NewLine
                        + "Tél :" + rapport.Pilote.Tel + Environment.NewLine
                    + "Réf 8D :" + rapport.Reference8D + Environment.NewLine
                    + "Réf Réclamation Client :" + rapport.ReferenceReclamationClient + Environment.NewLine
                    + "Alerte:" + rapport.Alerte + Environment.NewLine
                    ;

            if (culture.ToLower().Contains("en"))
            {
                InfoPiloteRapport = "Pilot :" + (rapport.Pilote != null ? rapport.Pilote.NomComplet : "") + Environment.NewLine
                   + "Tel :" + rapport.Pilote.Tel + Environment.NewLine
                   + "Ref 8D :" + rapport.Reference8D + Environment.NewLine
                   + "Ref Réclamation Client :" + rapport.ReferenceReclamationClient + Environment.NewLine
                   + "Type:" + rapport.Alerte + Environment.NewLine
                   ;
            }

            sheet_1_2D.Cells["Y1"].PutValue(InfoPiloteRapport);
            sheet_3D.Cells["Y1"].PutValue(InfoPiloteRapport);
            sheet_4D.Cells["Y1"].PutValue(InfoPiloteRapport);
            sheet_5_8D.Cells["Y1"].PutValue(InfoPiloteRapport);


            // *********************************************************
            //          Feuille 1-2 D
            // *********************************************************

            //Add a checkbox to the first worksheet in the workbook.
            int index_CB_Client = sheet_1_2D.CheckBoxes.Add(5, 16, 16, 200);
            Aspose.Cells.Drawing.CheckBox checkbox_Client = sheet_1_2D.CheckBoxes[index_CB_Client];
            checkbox_Client.Text = Ressources.Rapport.Client;// "Client";
            checkbox_Client.TextVerticalAlignment = TextAlignmentType.Top;
            checkbox_Client.Value = rapport.TypeIncident == "Client" ? true : false;

            int index_CB_Interne = sheet_1_2D.CheckBoxes.Add(6, 16, 16, 120);
            Aspose.Cells.Drawing.CheckBox checkbox_Interne = sheet_1_2D.CheckBoxes[index_CB_Interne];
            checkbox_Interne.Text = Ressources.Rapport.Interne;//"Interne";
            checkbox_Interne.TextVerticalAlignment = TextAlignmentType.Top;
            checkbox_Interne.Value = rapport.TypeIncident == "Interne" ? true : false;

            int index_CB_Frs = sheet_1_2D.CheckBoxes.Add(7, 16, 16, 120);
            Aspose.Cells.Drawing.CheckBox checkbox_Frs = sheet_1_2D.CheckBoxes[index_CB_Frs];
            checkbox_Frs.Text = Ressources.Rapport.Fournisseur;//"Fournisseur";
            checkbox_Frs.TextVerticalAlignment = TextAlignmentType.Top;
            checkbox_Frs.Value = rapport.TypeIncident == "Fournisseur" ? true : false;
            var dateFab = rapport.DateFabrication.ToString();
            var dateDabri = dateFab.Split(' ');
            //Date ouverture
            sheet_1_2D.Cells["I4"].PutValue("'" + rapport.DateOuverture.ToString());
            //Identification produit :							
            sheet_1_2D.Cells["I5"].PutValue(rapport.Produit);
            //Référence : 
            sheet_1_2D.Cells["I6"].PutValue(rapport.Reference);
            //N° série :
            sheet_1_2D.Cells["I7"].PutValue(rapport.NumSerie);
            //Date de fabrication :	
            sheet_1_2D.Cells["I8"].PutValue(dateDabri[0]);

            // ****** 2 – Description du problème :  	

            //Quoi?			
            sheet_1_2D.Cells["E15"].PutValue(rapport.D2Quoi);
            //Qui?			
            sheet_1_2D.Cells["E16"].PutValue(rapport.D2Qui);
            //Où?			
            sheet_1_2D.Cells["E17"].PutValue(rapport.D2Ou);
            //Quand?			
            sheet_1_2D.Cells["E18"].PutValue(rapport.D2Quand);
            //Comment?			
            sheet_1_2D.Cells["E19"].PutValue(rapport.D2Comment);
            //Combien?			
            sheet_1_2D.Cells["E20"].PutValue(rapport.D2Combien);
            //Pourquoi?			
            sheet_1_2D.Cells["E21"].PutValue(rapport.D2Pourquoi);


            //Le problème est-il récurrent (problème similaire déjà détecté sur autre ligne/ autre produit/ autre Client) ?

            Aspose.Cells.Drawing.RadioButton Radio_Oui = sheet_1_2D.Shapes.AddRadioButton(21, 10, 10, 0, 20, 80);
            Radio_Oui.Text = Ressources.Labels.Oui;
            Radio_Oui.TextVerticalAlignment = TextAlignmentType.Top;
            Radio_Oui.IsChecked = rapport.D2ProblemeRecurent;

            Aspose.Cells.Drawing.RadioButton Radio_Non = sheet_1_2D.Shapes.AddRadioButton(21, 30, 10, 0, 20, 80);
            Radio_Non.Text = Ressources.Labels.Non;
            Radio_Non.TextVerticalAlignment = TextAlignmentType.Top;
            Radio_Non.IsChecked = !rapport.D2ProblemeRecurent;

            // N° de l’ancien rapport 8D
            sheet_1_2D.Cells["AA22"].PutValue(rapport.D2NumAncienRapport);



            // ****** 1 – Equipe 

            if (rapport.MembreEquipeSet.Count > 0)
            {
                int rowStart = 12;
                sheet_1_2D.Cells.InsertRows(rowStart, rapport.MembreEquipeSet.Count - 1);

                foreach (var e in rapport.MembreEquipeSet)
                {
                    sheet_1_2D.Cells.Merge(rowStart - 1, 0, 1, 8);
                    sheet_1_2D.Cells.Merge(rowStart - 1, 8, 1, 7);

                    sheet_1_2D.Cells.Merge(rowStart - 1, 15, 1, 10);
                    sheet_1_2D.Cells.Merge(rowStart - 1, 25, 1, 7);

                    sheet_1_2D.Cells["A" + rowStart].PutValue(e.PersonnelSet != null ? e.PersonnelSet.NomComplet : "");
                    sheet_1_2D.Cells["I" + rowStart].PutValue(e.PersonnelSet != null ? e.PersonnelSet.Fonction : "");

                    rowStart++;
                }
            }


            // ***** Photos Croquis

            int maxWidth = 305;
            int maxHeight = 400;

            int imgsRowStart = 25 + rapport.MembreEquipeSet.Count;

            int paddingImgs = 0;
            foreach (var img in rapport.Photo.Where(x => x.Type == "OK"))
            {
                int imgIndex = sheet_1_2D.Pictures.Add(
                    paddingImgs > 0
                          ? paddingImgs +1
                        : imgsRowStart - 1, 0
                    ,this.GetImage(img.PhotoId).FileStream);
                var picture = sheet_1_2D.Pictures[imgIndex];


                var ratioX = (double)maxWidth / picture.OriginalWidth;
                var ratioY = (double)maxHeight / picture.OriginalHeight;
                var ratio = Math.Min(ratioX, ratioY);

                var newWidth = (int)(picture.OriginalWidth * ratio);
                var newHeight = (int)(picture.OriginalHeight * ratio);



                picture.Width = newWidth;
                picture.Height = newHeight;

                //var rescalePoucent = (int)(((double)maxWidth / picture.OriginalWidth) * 100);
                //picture.WidthScale = rescalePoucent > 100 ? 100 : rescalePoucent;
                //picture.HeightScale = rescalePoucent > 100 ? 100 : rescalePoucent;


                paddingImgs = picture.ActualLowerRightRow;
            }

            paddingImgs = 0;
            foreach (var img in rapport.Photo.Where(x => x.Type == "NOK"))
            {
                int imgIndex = sheet_1_2D.Pictures.Add(
                    paddingImgs > 0
                        ? paddingImgs + 1
                        : imgsRowStart - 1, 15
                    , this.GetImage(img.PhotoId).FileStream);
                var picture = sheet_1_2D.Pictures[imgIndex];


                var ratioX = (double)maxWidth / picture.OriginalWidth;
                var ratioY = (double)maxHeight / picture.OriginalHeight;
                var ratio = Math.Min(ratioX, ratioY);

                var newWidth = (int)(picture.OriginalWidth * ratio);
                var newHeight = (int)(picture.OriginalHeight * ratio);



                picture.Width = newWidth;
                picture.Height = newHeight;

                paddingImgs = picture.ActualLowerRightRow;
            }

            // http://www.microtuts.com/c-resize-an-image-proportionally-specify-max-widthheight/


            //int pictureIndex = sheet_1_2D.Pictures.Add(1, 5, GetImage(12).FileStream);

            //Aspose.Cells.Drawing.Picture picture = sheet_1_2D.Pictures[pictureIndex];

            //int pictureHeight = int.Parse((picture.OriginalHeight * (305 / picture.OriginalWidth)).ToString().Split('.')[0]);

            //picture.Left = 2;
            //picture.Top = 60;
            //picture.Width = picture.OriginalWidth>305 ? 305 : picture.OriginalWidth;
            //picture.Height = pictureHeight;

            //int imgsRowStart = 25 + rapport.MembreEquipeSet.Count;

            //int pictureIndex = sheet_1_2D.Pictures.Add(imgsRowStart, 0, GetImage(14).FileStream);
            //Aspose.Cells.Drawing.Picture picture2 = sheet_1_2D.Pictures[pictureIndex];

            //int maxWidth = 305;
            //int maxHeight = 400;

            //var ratioX = (double)maxWidth / picture2.OriginalWidth;
            //var ratioY = (double)maxHeight / picture2.OriginalHeight;
            //var ratio = Math.Min(ratioX, ratioY);

            //var newWidth = (int)(picture2.OriginalWidth * ratio);
            //var newHeight = (int)(picture2.OriginalHeight * ratio);

            //var rescalePoucent = (int)(((double)maxWidth / picture2.OriginalWidth) * 100);

            //picture2.Left = 2;
            //picture2.Top = 60 + picture.OriginalHeight;
            //picture2.Width = picture2.OriginalWidth > maxWidth ? maxWidth : picture2.OriginalWidth;
            //picture2.Height = 3;
            //picture2.IsLockAspectRatio = true;
            //picture2.UpperLeftColumn = 0;
            //picture2.UpperLeftRow = 0;
            //picture2.Top = 10;
            //picture2.Left = 16;

            //picture2.WidthScale = rescalePoucent > 100 ? 100 : rescalePoucent;
            //picture2.HeightScale = rescalePoucent > 100 ? 100 : rescalePoucent;




            // *********************************************************
            //          Feuille 3D
            // *********************************************************

            // ****** 3. Définir et mettre en place les actions de sécurisation																															
            // ****** 3.1 Actions d’investigation 		
            workbook.ChangePalette(System.Drawing.Color.Black, 55);
            Style blackBackgroundStyle = workbook.Styles[workbook.Styles.Add()];//new Style{BackgroundColor = System.Drawing.Color.Black};
            blackBackgroundStyle.BackgroundColor = System.Drawing.Color.Black;


            switch (rapport.D3_1_Investigation_1Tracabilite_Statut)
            {
                case "oui": sheet_3D.Cells["Q6"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["Q6"].PutValue("X"); break;
                case "non": sheet_3D.Cells["S6"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["Q6"].PutValue("X"); break;
                case "NA": sheet_3D.Cells["U6"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["Q6"].PutValue("X"); break;
            }
            switch (rapport.D3_1_Investigation_2AffectSimilaires_Statut)
            {
                case "oui": sheet_3D.Cells["Q7"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["Q7"].PutValue("X"); break;
                case "non": sheet_3D.Cells["S7"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["S7"].PutValue("X"); break;
                case "NA": sheet_3D.Cells["U7"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["U7"].PutValue("X"); break;
            }
            switch (rapport.D3_1_Investigation_3ProduitsExped_Statut)
            {
                case "oui": sheet_3D.Cells["Q8"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["Q8"].PutValue("X"); break;
                case "non": sheet_3D.Cells["S8"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["S8"].PutValue("X"); break;
                case "NA": sheet_3D.Cells["U8"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["U8"].PutValue("X"); break;
            }
            switch (rapport.D3_1_Investigation_4ProcessDefaut_Statut)
            {
                case "oui": sheet_3D.Cells["Q9"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["Q9"].PutValue("X"); break;
                case "non": sheet_3D.Cells["S9"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["S9"].PutValue("X"); break;
                case "NA": sheet_3D.Cells["U9"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["U9"].PutValue("X"); break;
            }
            switch (rapport.D3_1_Investigation_5FrsNoAlert_Statut)
            {
                case "oui": sheet_3D.Cells["Q10"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["Q10"].PutValue("X"); break;
                case "non": sheet_3D.Cells["S10"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["S10"].PutValue("X"); break;
                case "NA": sheet_3D.Cells["U10"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["U10"].PutValue("X"); break;
            }
            switch (rapport.D3_1_Investigation_6PbPresentCheezClient_Statut)
            {
                case "oui": sheet_3D.Cells["Q11"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["Q11"].PutValue("X"); break;
                case "non": sheet_3D.Cells["S11"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["S11"].PutValue("X"); break;
                case "NA": sheet_3D.Cells["U11"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["U11"].PutValue("X"); break;
            }
            switch (rapport.D3_1_Investigation_7_Statut)
            {
                case "oui": sheet_3D.Cells["Q12"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["Q12"].PutValue("X"); break;
                case "non": sheet_3D.Cells["S12"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["S12"].PutValue("X"); break;
                case "NA": sheet_3D.Cells["U12"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["U12"].PutValue("X"); break;
            }
            switch (rapport.D3_1_Investigation_8_Statut)
            {
                case "oui": sheet_3D.Cells["Q13"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["Q13"].PutValue("X"); break;
                case "non": sheet_3D.Cells["S13"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["S13"].PutValue("X"); break;
                case "NA": sheet_3D.Cells["U13"].SetStyle(blackBackgroundStyle); sheet_3D.Cells["U13"].PutValue("X"); break;
            }

            //sheet_1_2D.Cells["W6"].PutValue(rapport.D3_1_Investigation_1Tracabilite_Statut);
            //sheet_1_2D.Cells["W6"].PutValue(rapport.D3_1_Investigation_2AffectSimilaires_Statut);
            //sheet_1_2D.Cells["W6"].PutValue(rapport.D3_1_Investigation_3ProduitsExped_Statut);
            //sheet_1_2D.Cells["W6"].SetStyle(blackBackgroundStyle);// PutValue(rapport.D3_1_Investigation_4ProcessDefaut_Statut);
            //sheet_1_2D.Cells["W6"].PutValue(rapport.D3_1_Investigation_5FrsNoAlert_Statut);
            //sheet_1_2D.Cells["W6"].PutValue(rapport.D3_1_Investigation_6PbPresentCheezClient_Statut);

            sheet_3D.Cells["W6"].PutValue(rapport.D3_1_Investigation_1Tracabilite_Reaction);
            sheet_3D.Cells["W7"].PutValue(rapport.D3_1_Investigation_2AffectSimilaires_Reaction);
            sheet_3D.Cells["W8"].PutValue(rapport.D3_1_Investigation_3ProduitsExped_Reaction);
            sheet_3D.Cells["W9"].PutValue(rapport.D3_1_Investigation_4ProcessDefaut_Reaction);
            sheet_3D.Cells["W10"].PutValue(rapport.D3_1_Investigation_5FrsNoAlert_Reaction);
            sheet_3D.Cells["W11"].PutValue(rapport.D3_1_Investigation_6PbPresentCheezClient_Reaction);

            // ****** 3.2 Containment Check list :
            // start from 15Y 15AA 15AC
            // end to 38Y 38AA 38AC
            int rowNumStart = 17;
            sheet_3D.Cells["Y" + rowNumStart].PutValue(rapport.D3_2_Communiquer_1_Applicable ? "X" : "");
            sheet_3D.Cells["AA" + rowNumStart].PutValue(rapport.PiloteD3_2_Communiquer_1 != null ? rapport.PiloteD3_2_Communiquer_1.NomComplet : "");
            sheet_3D.Cells["AC" + rowNumStart].PutValue(rapport.D3_2_Communiquer_1_Comment);

            ++rowNumStart;
            sheet_3D.Cells["Y" + rowNumStart].PutValue(rapport.D3_2_Stopper_1_1_Applicable ? "X" : "");
            sheet_3D.Cells["AA" + rowNumStart].PutValue(rapport.PiloteD3_2_Stopper_1_1 != null ? rapport.PiloteD3_2_Stopper_1_1.NomComplet : "");
            sheet_3D.Cells["AC" + rowNumStart].PutValue(rapport.D3_2_Stopper_1_1_Comment);

            ++rowNumStart;
            sheet_3D.Cells["Y" + rowNumStart].PutValue(rapport.D3_2_Stopper_1_2_Applicable ? "X" : "");
            sheet_3D.Cells["AA" + rowNumStart].PutValue(rapport.PiloteD3_2_Stopper_1_2 != null ? rapport.PiloteD3_2_Stopper_1_2.NomComplet : "");
            sheet_3D.Cells["AC" + rowNumStart].PutValue(rapport.D3_2_Stopper_1_2_Comment);

            ++rowNumStart;
            sheet_3D.Cells["Y" + rowNumStart].PutValue(rapport.D3_2_Stopper_1_3_Applicable ? "X" : "");
            sheet_3D.Cells["AA" + rowNumStart].PutValue(rapport.PiloteD3_2_Stopper_1_3 != null ? rapport.PiloteD3_2_Stopper_1_3.NomComplet : "");
            sheet_3D.Cells["AC" + rowNumStart].PutValue(rapport.D3_2_Stopper_1_3_Comment);

            ++rowNumStart;
            sheet_3D.Cells["Y" + rowNumStart].PutValue(rapport.D3_2_Stopper_1_4_Applicable ? "X" : "");
            sheet_3D.Cells["AA" + rowNumStart].PutValue(rapport.PiloteD3_2_Stopper_1_4 != null ? rapport.PiloteD3_2_Stopper_1_4.NomComplet : "");
            sheet_3D.Cells["AC" + rowNumStart].PutValue(rapport.D3_2_Stopper_1_4_Comment);

            ++rowNumStart;
            sheet_3D.Cells["Y" + rowNumStart].PutValue(rapport.D3_2_Stopper_2_1_Applicable ? "X" : "");
            sheet_3D.Cells["AA" + rowNumStart].PutValue(rapport.PiloteD3_2_Stopper_2_1 != null ? rapport.PiloteD3_2_Stopper_2_1.NomComplet : "");
            sheet_3D.Cells["AC" + rowNumStart].PutValue(rapport.D3_2_Stopper_2_1_Comment);

            ++rowNumStart;
            sheet_3D.Cells["Y" + rowNumStart].PutValue(rapport.D3_2_Stopper_2_2_Applicable ? "X" : "");
            sheet_3D.Cells["AA" + rowNumStart].PutValue(rapport.PiloteD3_2_Stopper_2_2 != null ? rapport.PiloteD3_2_Stopper_2_2.NomComplet : "");
            sheet_3D.Cells["AC" + rowNumStart].PutValue(rapport.D3_2_Stopper_2_2_Comment);

            ++rowNumStart;
            sheet_3D.Cells["Y" + rowNumStart].PutValue(rapport.D3_2_Stopper_2_3_Applicable ? "X" : "");
            sheet_3D.Cells["AA" + rowNumStart].PutValue(rapport.PiloteD3_2_Stopper_2_3 != null ? rapport.PiloteD3_2_Stopper_2_3.NomComplet : "");
            sheet_3D.Cells["AC" + rowNumStart].PutValue(rapport.D3_2_Stopper_2_3_Comment);

            ++rowNumStart;
            sheet_3D.Cells["Y" + rowNumStart].PutValue(rapport.D3_2_Stopper_3_Applicable ? "X" : "");
            sheet_3D.Cells["AA" + rowNumStart].PutValue(rapport.PiloteD3_2_Stopper_3 != null ? rapport.PiloteD3_2_Stopper_3.NomComplet : "");
            sheet_3D.Cells["AC" + rowNumStart].PutValue(rapport.D3_2_Stopper_3_Comment);
            rowNumStart += 3;// fusion de 4 lignes


            ++rowNumStart;
            sheet_3D.Cells["Y" + rowNumStart].PutValue(rapport.D3_2_Trier_1_Applicable ? "X" : "");
            sheet_3D.Cells["AA" + rowNumStart].PutValue(rapport.PiloteD3_2_Trier_1 != null ? rapport.PiloteD3_2_Trier_1.NomComplet : "");
            sheet_3D.Cells["AC" + rowNumStart].PutValue(rapport.D3_2_Trier_1_Comment);

            ++rowNumStart;
            sheet_3D.Cells["Y" + rowNumStart].PutValue(rapport.D3_2_Trier_2_Applicable ? "X" : "");
            sheet_3D.Cells["AA" + rowNumStart].PutValue(rapport.PiloteD3_2_Trier_2 != null ? rapport.PiloteD3_2_Trier_2.NomComplet : "");
            sheet_3D.Cells["AC" + rowNumStart].PutValue(rapport.D3_2_Trier_2_Comment);

            ++rowNumStart;
            sheet_3D.Cells["Y" + rowNumStart].PutValue(rapport.D3_2_Trier_3_Applicable ? "X" : "");
            sheet_3D.Cells["AA" + rowNumStart].PutValue(rapport.PiloteD3_2_Trier_3 != null ? rapport.PiloteD3_2_Trier_3.NomComplet : "");
            sheet_3D.Cells["AC" + rowNumStart].PutValue(rapport.D3_2_Trier_3_Comment);

            ++rowNumStart;
            sheet_3D.Cells["Y" + rowNumStart].PutValue(rapport.D3_2_Trier_4_1_Applicable ? "X" : "");
            sheet_3D.Cells["AA" + rowNumStart].PutValue(rapport.PiloteD3_2_Trier_4_1 != null ? rapport.PiloteD3_2_Trier_4_1.NomComplet : "");
            sheet_3D.Cells["AC" + rowNumStart].PutValue(rapport.D3_2_Trier_4_1_Comment);

            ++rowNumStart;
            sheet_3D.Cells["Y" + rowNumStart].PutValue(rapport.D3_2_Trier_4_2_Applicable ? "X" : "");
            sheet_3D.Cells["AA" + rowNumStart].PutValue(rapport.PiloteD3_2_Trier_4_2 != null ? rapport.PiloteD3_2_Trier_4_2.NomComplet : "");
            sheet_3D.Cells["AC" + rowNumStart].PutValue(rapport.D3_2_Trier_4_2_Comment);

            ++rowNumStart;
            sheet_3D.Cells["Y" + rowNumStart].PutValue(rapport.D3_2_Trier_4_3_Applicable ? "X" : "");
            sheet_3D.Cells["AA" + rowNumStart].PutValue(rapport.PiloteD3_2_Trier_4_3 != null ? rapport.PiloteD3_2_Trier_4_3.NomComplet : "");
            sheet_3D.Cells["AC" + rowNumStart].PutValue(rapport.D3_2_Trier_4_3_Comment);

            ++rowNumStart;
            sheet_3D.Cells["Y" + rowNumStart].PutValue(rapport.D3_2_Trier_4_4_Applicable ? "X" : "");
            sheet_3D.Cells["AA" + rowNumStart].PutValue(rapport.PiloteD3_2_Trier_4_4 != null ? rapport.PiloteD3_2_Trier_4_4.NomComplet : "");
            sheet_3D.Cells["AC" + rowNumStart].PutValue(rapport.D3_2_Trier_4_4_Comment);

            ++rowNumStart;
            sheet_3D.Cells["Y" + rowNumStart].PutValue(rapport.D3_2_Trier_4_5_Applicable ? "X" : "");
            sheet_3D.Cells["AA" + rowNumStart].PutValue(rapport.PiloteD3_2_Trier_4_5 != null ? rapport.PiloteD3_2_Trier_4_5.NomComplet : "");
            sheet_3D.Cells["AC" + rowNumStart].PutValue(rapport.D3_2_Trier_4_5_Comment);



            ++rowNumStart;
            sheet_3D.Cells["Y" + rowNumStart].PutValue(rapport.D3_2_Reparer_1_Applicable ? "X" : "");
            sheet_3D.Cells["AA" + rowNumStart].PutValue(rapport.PiloteD3_2_Reparer_1 != null ? rapport.PiloteD3_2_Reparer_1.NomComplet : "");
            sheet_3D.Cells["AC" + rowNumStart].PutValue(rapport.D3_2_Reparer_1_Comment);

            ++rowNumStart;
            sheet_3D.Cells["Y" + rowNumStart].PutValue(rapport.D3_2_Reparer_2_Applicable ? "X" : "");
            sheet_3D.Cells["AA" + rowNumStart].PutValue(rapport.PiloteD3_2_Reparer_2 != null ? rapport.PiloteD3_2_Reparer_2.NomComplet : "");
            sheet_3D.Cells["AC" + rowNumStart].PutValue(rapport.D3_2_Reparer_2_Comment);

            ++rowNumStart;
            sheet_3D.Cells["Y" + rowNumStart].PutValue(rapport.D3_2_Reparer_3_Applicable ? "X" : "");
            sheet_3D.Cells["AA" + rowNumStart].PutValue(rapport.PiloteD3_2_Reparer_3 != null ? rapport.PiloteD3_2_Reparer_3.NomComplet : "");
            sheet_3D.Cells["AC" + rowNumStart].PutValue(rapport.D3_2_Reparer_3_Comment);

            ++rowNumStart;
            sheet_3D.Cells["Y" + rowNumStart].PutValue(rapport.D3_2_Reparer_4_Applicable ? "X" : "");
            sheet_3D.Cells["AA" + rowNumStart].PutValue(rapport.PiloteD3_2_Reparer_4 != null ? rapport.PiloteD3_2_Reparer_4.NomComplet : "");
            sheet_3D.Cells["AC" + rowNumStart].PutValue(rapport.D3_2_Reparer_4_Comment);

            var ChezClientTauxDefaut = 0.0 ;
            var InterneTauxDefaut = 0.0 ;
            var ChezFournisseurTauxDefaut = 0.0 ;

            if (rapport.D3_3_ChezClien_QteNonConforme != null && rapport.D3_3_ChezClien_QteTriee != null)
            {
                if (rapport.D3_3_ChezClien_QteNonConforme == "Empty" || rapport.D3_3_ChezClien_QteTriee == "Empty" || rapport.D3_3_ChezClien_QteNonConforme == "" || rapport.D3_3_ChezClien_QteTriee == "") { }
                else if (float.Parse(rapport.D3_3_ChezClien_QteNonConforme) != 0 && float.Parse(rapport.D3_3_ChezClien_QteTriee) != 0)
                {
                    ChezClientTauxDefaut = Math.Round(float.Parse(rapport.D3_3_ChezClien_QteNonConforme) / float.Parse(rapport.D3_3_ChezClien_QteTriee) * 100, 2);
                }
            }
            if (rapport.D3_3_EnInterne_QteNonConforme != null && rapport.D3_3_EnInterne_QteTriee != null)
            {
                if (rapport.D3_3_EnInterne_QteNonConforme == "Empty" || rapport.D3_3_EnInterne_QteTriee == "Empty" || rapport.D3_3_EnInterne_QteNonConforme == "" || rapport.D3_3_EnInterne_QteTriee == "") { }
                else if (float.Parse(rapport.D3_3_EnInterne_QteNonConforme) != 0 && float.Parse(rapport.D3_3_EnInterne_QteTriee) != 0)
                {
                    InterneTauxDefaut = Math.Round(float.Parse(rapport.D3_3_EnInterne_QteNonConforme) / float.Parse(rapport.D3_3_EnInterne_QteTriee) * 100, 2);
                }
            }
            if (rapport.D3_3_ChezFournisseur_QteNonConforme != null && rapport.D3_3_ChezFournisseur_QteTriee != null)
            {
                if (rapport.D3_3_ChezFournisseur_QteNonConforme == "Empty" || rapport.D3_3_ChezFournisseur_QteTriee == "Empty" || rapport.D3_3_ChezFournisseur_QteNonConforme == "" || rapport.D3_3_ChezFournisseur_QteTriee == "") { }
                else if (float.Parse(rapport.D3_3_ChezFournisseur_QteNonConforme) != 0 && float.Parse(rapport.D3_3_ChezFournisseur_QteTriee) != 0)
                {
                    ChezFournisseurTauxDefaut = Math.Round(float.Parse(rapport.D3_3_ChezFournisseur_QteNonConforme) / float.Parse(rapport.D3_3_ChezFournisseur_QteTriee) * 100, 2);
                }
            }




            // ****** 3.3 Résultats de tri : 

            //Chez le client :										
            //Quantité triée :										
            sheet_3D.Cells["P42"].PutValue(rapport.D3_3_ChezClien_QteTriee);
            //Quantité non conforme :					
            sheet_3D.Cells["Y42"].PutValue(rapport.D3_3_ChezClien_QteNonConforme);
            //TAux de défaut
            sheet_3D.Cells["AF42"].PutValue(ChezClientTauxDefaut + "%");


            //En interne (toute la chaîne logistique à prendre en compte) :				
            //Quantité triée :										
            sheet_3D.Cells["P43"].PutValue(rapport.D3_3_EnInterne_QteTriee);
            //Quantité non conforme :					
            sheet_3D.Cells["Y43"].PutValue(rapport.D3_3_EnInterne_QteNonConforme);
            //TAux de défaut
            sheet_3D.Cells["AF43"].PutValue(InterneTauxDefaut + "%");

            //Chez le fournisseur										
            //Quantité triée :										
            sheet_3D.Cells["P44"].PutValue(rapport.D3_3_ChezFournisseur_QteTriee);
            //Quantité non conforme :					
            sheet_3D.Cells["Y44"].PutValue(rapport.D3_3_ChezFournisseur_QteNonConforme);
            //TAux de défaut
            sheet_3D.Cells["AF44"].PutValue(ChezFournisseurTauxDefaut + "%");

            // ****** 3.4 Identification du produit sécurisé expédié au client :
            sheet_3D.Cells["T46"].PutValue(rapport.D3_4_Identification);

            // *********************************************************
            //          Feuille 4D
            // *********************************************************


            // ****** OCCURRENCE DU DEFAUT/  PROBLEME

            var cause = rapport.D4Causes.Where(x =>
                        x.Type == "Occurence" &&
                        x.Cause5M == "Matiere" &&
                        x.Ordre == 1
                        )
                        .FirstOrDefault();

            int row_4D_4 = 7;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }
            // *********
            cause = rapport.D4Causes.Where(x =>
                        x.Type == "Occurence" &&
                        x.Cause5M == "Matiere" &&
                        x.Ordre == 2
                        )
                        .FirstOrDefault();

            row_4D_4 = 8;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }

            cause = rapport.D4Causes.Where(x =>
                        x.Type == "Occurence" &&
                        x.Cause5M == "Matiere" &&
                        x.Ordre == 3
                        )
                        .FirstOrDefault();

            row_4D_4 = 9;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }

            cause = rapport.D4Causes.Where(x =>
                        x.Type == "Occurence" &&
                        x.Cause5M == "Materiel" &&
                        x.Ordre == 1
                        )
                        .FirstOrDefault();

            row_4D_4 = 10;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }

            cause = rapport.D4Causes.Where(x =>
                        x.Type == "Occurence" &&
                        x.Cause5M == "Materiel" &&
                        x.Ordre == 2
                        )
                        .FirstOrDefault();

            row_4D_4 = 11;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }

            cause = rapport.D4Causes.Where(x =>
                        x.Type == "Occurence" &&
                        x.Cause5M == "Materiel" &&
                        x.Ordre == 3
                        )
                        .FirstOrDefault();

            row_4D_4 = 12;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }
            cause = rapport.D4Causes.Where(x =>
                        x.Type == "Occurence" &&
                        x.Cause5M == "Milieu" &&
                        x.Ordre == 1
                        )
                        .FirstOrDefault();

            row_4D_4 = 13;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }

            cause = rapport.D4Causes.Where(x =>
                        x.Type == "Occurence" &&
                        x.Cause5M == "Milieu" &&
                        x.Ordre == 2
                        )
                        .FirstOrDefault();

            row_4D_4 = 14;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }
            cause = rapport.D4Causes.Where(x =>
                        x.Type == "Occurence" &&
                        x.Cause5M == "Milieu" &&
                        x.Ordre == 3
                        )
                        .FirstOrDefault();

            row_4D_4 = 15;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }

            cause = rapport.D4Causes.Where(x =>
                        x.Type == "Occurence" &&
                        x.Cause5M == "Methodes" &&
                        x.Ordre == 1
                        )
                        .FirstOrDefault();

            row_4D_4 = 16;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }
            cause = rapport.D4Causes.Where(x =>
                        x.Type == "Occurence" &&
                        x.Cause5M == "Methodes" &&
                        x.Ordre == 2
                        )
                        .FirstOrDefault();

            row_4D_4 = 17;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }

            cause = rapport.D4Causes.Where(x =>
                        x.Type == "Occurence" &&
                        x.Cause5M == "Methodes" &&
                        x.Ordre == 3
                        )
                        .FirstOrDefault();

            row_4D_4 = 18;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }

            cause = rapport.D4Causes.Where(x =>
                        x.Type == "Occurence" &&
                        x.Cause5M == "MainDoeuvres" &&
                        x.Ordre == 1
                        )
                        .FirstOrDefault();

            row_4D_4 = 19;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }

            cause = rapport.D4Causes.Where(x =>
                        x.Type == "Occurence" &&
                        x.Cause5M == "MainDoeuvres" &&
                        x.Ordre == 2
                        )
                        .FirstOrDefault();

            row_4D_4 = 20;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }

            cause = rapport.D4Causes.Where(x =>
                        x.Type == "Occurence" &&
                        x.Cause5M == "MainDoeuvres" &&
                        x.Ordre == 3
                        )
                        .FirstOrDefault();

            row_4D_4 = 21;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }



            // ****** NON DETECTION DU DEFAUT/ PROBLEME

            cause = rapport.D4Causes.Where(x =>
                        x.Type == "NonDetection" &&
                        x.Cause5M == "Matiere" &&
                        x.Ordre == 1
                        )
                        .FirstOrDefault();

            row_4D_4 = 25;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }
            // *********
            cause = rapport.D4Causes.Where(x =>
                        x.Type == "NonDetection" &&
                        x.Cause5M == "Matiere" &&
                        x.Ordre == 2
                        )
                        .FirstOrDefault();

            row_4D_4++;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }

            cause = rapport.D4Causes.Where(x =>
                        x.Type == "NonDetection" &&
                        x.Cause5M == "Matiere" &&
                        x.Ordre == 3
                        )
                        .FirstOrDefault();

            row_4D_4++;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }

            cause = rapport.D4Causes.Where(x =>
                        x.Type == "NonDetection" &&
                        x.Cause5M == "Materiel" &&
                        x.Ordre == 1
                        )
                        .FirstOrDefault();

            row_4D_4++;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }

            cause = rapport.D4Causes.Where(x =>
                        x.Type == "NonDetection" &&
                        x.Cause5M == "Materiel" &&
                        x.Ordre == 2
                        )
                        .FirstOrDefault();

            row_4D_4++;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }

            cause = rapport.D4Causes.Where(x =>
                        x.Type == "NonDetection" &&
                        x.Cause5M == "Materiel" &&
                        x.Ordre == 3
                        )
                        .FirstOrDefault();

            row_4D_4++;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }
            cause = rapport.D4Causes.Where(x =>
                        x.Type == "NonDetection" &&
                        x.Cause5M == "Milieu" &&
                        x.Ordre == 1
                        )
                        .FirstOrDefault();

            row_4D_4++;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }

            cause = rapport.D4Causes.Where(x =>
                        x.Type == "NonDetection" &&
                        x.Cause5M == "Milieu" &&
                        x.Ordre == 2
                        )
                        .FirstOrDefault();

            row_4D_4++;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }
            cause = rapport.D4Causes.Where(x =>
                        x.Type == "NonDetection" &&
                        x.Cause5M == "Milieu" &&
                        x.Ordre == 3
                        )
                        .FirstOrDefault();

            row_4D_4++;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }

            cause = rapport.D4Causes.Where(x =>
                        x.Type == "NonDetection" &&
                        x.Cause5M == "Methodes" &&
                        x.Ordre == 1
                        )
                        .FirstOrDefault();

            row_4D_4++;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }
            cause = rapport.D4Causes.Where(x =>
                        x.Type == "NonDetection" &&
                        x.Cause5M == "Methodes" &&
                        x.Ordre == 2
                        )
                        .FirstOrDefault();

            row_4D_4++;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }

            cause = rapport.D4Causes.Where(x =>
                        x.Type == "NonDetection" &&
                        x.Cause5M == "Methodes" &&
                        x.Ordre == 3
                        )
                        .FirstOrDefault();

            row_4D_4++;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }

            cause = rapport.D4Causes.Where(x =>
                        x.Type == "NonDetection" &&
                        x.Cause5M == "MainDoeuvres" &&
                        x.Ordre == 1
                        )
                        .FirstOrDefault();

            row_4D_4++;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }

            cause = rapport.D4Causes.Where(x =>
                        x.Type == "NonDetection" &&
                        x.Cause5M == "MainDoeuvres" &&
                        x.Ordre == 2
                        )
                        .FirstOrDefault();

            row_4D_4++;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }

            cause = rapport.D4Causes.Where(x =>
                        x.Type == "NonDetection" &&
                        x.Cause5M == "MainDoeuvres" &&
                        x.Ordre == 3
                        )
                        .FirstOrDefault();

            row_4D_4++;
            if (cause != null)
            {
                sheet_4D.Cells["E" + row_4D_4].PutValue(cause.Cause);
                sheet_4D.Cells["O" + row_4D_4].PutValue(cause.Pourquoi1);
                sheet_4D.Cells["Q" + row_4D_4].PutValue(cause.Pourquoi2);
                sheet_4D.Cells["S" + row_4D_4].PutValue(cause.Pourquoi3);
                sheet_4D.Cells["U" + row_4D_4].PutValue(cause.Pourquoi4);
                sheet_4D.Cells["W" + row_4D_4].PutValue(cause.Pourquoi5);
                sheet_4D.Cells["Y" + row_4D_4].PutValue(cause.Retenue.HasValue ? (cause.Retenue.Value ? Ressources.Labels.Retenu : Ressources.Labels.NonRetenu) : "");
                sheet_4D.Cells["AA" + row_4D_4].PutValue(cause.Justification);
                sheet_4D.Cells["AE" + row_4D_4].PutValue(cause.NumAction);
            }
            if (rapport.ReclamationSpecifiqueClient== false )
            {
                cause = rapport.D4Causes.Where(x =>
                 x.Type == "RevueAMDEC"
                                                )
                                .FirstOrDefault();

                row_4D_4 = 43;
                // row_4D_4++;
                if (cause != null)
                {
                    sheet_4D.Cells["B" + row_4D_4].PutValue(cause.RefAMDEC);
                    sheet_4D.Cells["E" + row_4D_4].PutValue(cause.NouvelleVersion);
                    sheet_4D.Cells["O" + row_4D_4].PutValue(cause.VersionActuelle);

                }
            }
            
                // *********************************************************
                //          Feuille 5-8D
                // *********************************************************


                // ****** 8 – Clôture
                // Date de l’audit de clôture : I78
                sheet_5_8D.Cells["I63"].PutValue(rapport.DateCloture);
                // Commentaire et visa du management : I79 
                sheet_5_8D.Cells["I64"].PutValue(rapport.CommentVisaCloture);
                // Audit effectué par : Y78         
                sheet_5_8D.Cells["Y63"].PutValue(rapport.Auditeur != null ? rapport.Auditeur.NomComplet : "");
                // Date Y79
                sheet_5_8D.Cells["Y64"].PutValue(rapport.DateAudit);
                // Date de diffusion : I81
                sheet_5_8D.Cells["I66"].PutValue(rapport.DateDifusion);

                if (rapport.ReclamationSpecifiqueClient)
                {
                     cause = rapport.D4Causes.Where(x =>
                 x.Type == "RevueAMDEC"
                                                )
                                .FirstOrDefault();

             rowNumStart = 70;
          // row_4D_4++;
            if (cause != null)
            {
                sheet_5_8D.Cells["B" + rowNumStart].PutValue(cause.RefAMDEC);
                sheet_5_8D.Cells["E" + rowNumStart].PutValue(cause.VersionActuelle);
                sheet_5_8D.Cells["O" + rowNumStart].PutValue(cause.NouvelleVersion);

            }
                }

                // ****** 7.2 Standardisation

                rowNumStart = 54;

                sheet_5_8D.Cells["U" + rowNumStart].PutValue(rapport.ResponsableStandardisationAmdec != null ? rapport.ResponsableStandardisationAmdec.NomComplet : "");
                sheet_5_8D.Cells["Y" + rowNumStart].PutValue(rapport.Standatdisation_Amdec_Delai);
                sheet_5_8D.Cells["AC" + rowNumStart].PutValue(rapport.Standatdisation_Amdec_DateRealisation);

                rowNumStart++;
                sheet_5_8D.Cells["U" + rowNumStart].PutValue(rapport.ResponsableStandardisationPlanSurveillance != null ? rapport.ResponsableStandardisationPlanSurveillance.NomComplet : "");
                sheet_5_8D.Cells["Y" + rowNumStart].PutValue(rapport.Standatdisation_PlanSurveillance_Delai);
                sheet_5_8D.Cells["AC" + rowNumStart].PutValue(rapport.Standatdisation_PlanSurveillance_DateRealisation);

                rowNumStart++;
                sheet_5_8D.Cells["U" + rowNumStart].PutValue(rapport.ResponsableStandardisationFicheInstruction != null ? rapport.ResponsableStandardisationFicheInstruction.NomComplet : "");
                sheet_5_8D.Cells["Y" + rowNumStart].PutValue(rapport.Standatdisation_FicheInstruction_Delai);
                sheet_5_8D.Cells["AC" + rowNumStart].PutValue(rapport.Standatdisation_FicheInstruction_DateRealisation);

                rowNumStart++;
                sheet_5_8D.Cells["U" + rowNumStart].PutValue(rapport.ResponsableStandardisationDefautheque != null ? rapport.ResponsableStandardisationDefautheque.NomComplet : "");
                sheet_5_8D.Cells["Y" + rowNumStart].PutValue(rapport.Standatdisation_Defautheque_Delai);
                sheet_5_8D.Cells["AC" + rowNumStart].PutValue(rapport.Standatdisation_Defautheque_DateRealisation);

                rowNumStart++;
                sheet_5_8D.Cells["A" + rowNumStart].PutValue("Autre :" + rapport.Standatdisation_Autre1_Description);
                sheet_5_8D.Cells["U" + rowNumStart].PutValue(rapport.ResponsableStandardisationAutre1 != null ? rapport.ResponsableStandardisationAutre1.NomComplet : "");
                sheet_5_8D.Cells["Y" + rowNumStart].PutValue(rapport.Standatdisation_Autre1_Delai);
                sheet_5_8D.Cells["AC" + rowNumStart].PutValue(rapport.Standatdisation_Autre1_DateRealisation);

                rowNumStart++;
                sheet_5_8D.Cells["A" + rowNumStart].PutValue("Autre :" + rapport.Standatdisation_Autre2_Description);
                sheet_5_8D.Cells["U" + rowNumStart].PutValue(rapport.ResponsableStandardisationAutre2 != null ? rapport.ResponsableStandardisationAutre2.NomComplet : "");
                sheet_5_8D.Cells["Y" + rowNumStart].PutValue(rapport.Standatdisation_Autre2_Delai);
                sheet_5_8D.Cells["AC" + rowNumStart].PutValue(rapport.Standatdisation_Autre2_DateRealisation);

                rowNumStart++;
                sheet_5_8D.Cells["A" + rowNumStart].PutValue("Autre :" + rapport.Standatdisation_Autre3_Description);
                sheet_5_8D.Cells["U" + rowNumStart].PutValue(rapport.ResponsableStandardisationAutre3 != null ? rapport.ResponsableStandardisationAutre3.NomComplet : "");
                sheet_5_8D.Cells["Y" + rowNumStart].PutValue(rapport.Standatdisation_Autre3_Delai);
                sheet_5_8D.Cells["AC" + rowNumStart].PutValue(rapport.Standatdisation_Autre3_DateRealisation);


                // ****** 7.1 Généralisation des actions correctives

                if (rapport.D7ActionsGeneralisation.Count> 0)
                {
                    int rowStart = 50;
                    sheet_5_8D.Cells.InsertRows(rowStart, rapport.D7ActionsGeneralisation.Count - 1);

                    foreach (var action in rapport.D7ActionsGeneralisation)
                    {
                        sheet_5_8D.Cells.Merge(rowStart - 1, 0, 1, 7);
                        sheet_5_8D.Cells.Merge(rowStart - 1, 7, 1, 13);
                        sheet_5_8D.Cells.Merge(rowStart - 1, 20, 1, 4);
                        sheet_5_8D.Cells.Merge(rowStart - 1, 24, 1, 4);
                        sheet_5_8D.Cells.Merge(rowStart - 1, 28, 1, 4);


                        sheet_5_8D.Cells["A" + rowStart].PutValue(action.Perimetre);
                        sheet_5_8D.Cells["H" + rowStart].PutValue(action.Action);
                        sheet_5_8D.Cells["U" + rowStart].PutValue(action.PersonnelSet != null ? action.PersonnelSet.NomComplet : "");
                        sheet_5_8D.Cells["Y" + rowStart].PutValue(action.Delai);
                        sheet_5_8D.Cells["AC" + rowStart].PutValue(action.DateRealisation);

                        rowStart++;
                    }
                }


                // ****** 6 – Efficacité du plan d'action global

                string[] array_Fournisseur = rapport.D6_Defauts_Fournisseur != null ? rapport.D6_Defauts_Fournisseur.Split(',') : null;
                string[] array_Interne = rapport.D6_Defauts_Interne != null ? rapport.D6_Defauts_Interne.Split(','): null;
                string[] array_Client = rapport.D6_Defauts_Client != null ? rapport.D6_Defauts_Client.Split(',') : null;

                for (int i = 0; i < 31; i++)
                {
                    if (array_Fournisseur != null)
                        sheet_5_8D.Cells[12, i + 1].PutValue(int.Parse(array_Fournisseur.Length >= i ? array_Fournisseur[i] : ""));//B21
                    if (array_Interne != null)
                        sheet_5_8D.Cells[13, i + 1].PutValue(int.Parse(array_Interne.Length >= i ? array_Interne[i] : ""));//B22
                    if (array_Client != null)
                        sheet_5_8D.Cells[14, i + 1].PutValue(int.Parse(array_Client.Length >= i ? array_Client[i] : ""));//B23
                }


                // ****** 5 – Actions correctives 
                // Date de fabrication Premiers produits certifiés conformes (estimation) : 
                sheet_5_8D.Cells["X7"].PutValue(rapport.D5_Date_FabPremierProduitConforme);


                if (rapport.D5ActionsCorrectives.Count > 0)
                {
                    int rowStart = 6;
                    sheet_5_8D.Cells.InsertRows(rowStart, rapport.D5ActionsCorrectives.Count - 1);

                    foreach (var action in rapport.D5ActionsCorrectives)
                    {

                        sheet_5_8D.Cells.Merge(rowStart - 1, 1, 1, 11);
                        sheet_5_8D.Cells.Merge(rowStart - 1, 12, 1, 4);
                        sheet_5_8D.Cells.Merge(rowStart - 1, 16, 1, 3);
                        sheet_5_8D.Cells.Merge(rowStart - 1, 19, 1, 3);
                        sheet_5_8D.Cells.Merge(rowStart - 1, 22, 1, 4);
                        sheet_5_8D.Cells.Merge(rowStart - 1, 26, 1, 6);

                        sheet_5_8D.Cells["A" + rowStart.ToString()].PutValue(action.NumAction);
                        sheet_5_8D.Cells["B" + rowStart.ToString()].PutValue(action.Action);
                        sheet_5_8D.Cells["M" + rowStart.ToString()].PutValue(action.Responsable != null ? action.Responsable.NomComplet : "");
                        sheet_5_8D.Cells["Q" + rowStart.ToString()].PutValue(action.Delai);
                        sheet_5_8D.Cells["T" + rowStart.ToString()].PutValue(action.DateRealisation);
                        sheet_5_8D.Cells["W" + rowStart.ToString()].PutValue(action.VerifEfficacite);
                        sheet_5_8D.Cells["AA" + rowStart.ToString()].PutValue(action.MoyenVerification);

                        rowStart++;
                    }
                }

                // *********************************************************
                //          Documents Annexes
                // *********************************************************


                // TODO : pour chaque document on va créer un onglet

                // TODO : dans chaque ongle on met le nom et description du document

                var FileStoragePath =
                        db.Configs.Find("FileStoragePath").Val;
                //WebConfigurationManager.AppSettings["FileStoragePath"];

                // Image de présentation du document
                string ImageUrl = Path.Combine(baseDir, "Images", "doc.jpg"); //Define a string variable to store the image path.

                FileStream fs = System.IO.File.OpenRead(ImageUrl); //Get the picture into the streams.

                byte[] imageData = new Byte[fs.Length];//Define a byte array.
                fs.Read(imageData, 0, imageData.Length);//Obtain the picture into the array of bytes from streams.


                foreach (var doc in rapport.Documents)
                {
                    int indiceOnglet = workbook.Worksheets.Add();

                    Worksheet onglet_Doc = workbook.Worksheets[indiceOnglet];
                    onglet_Doc.Name = doc.NomFichier.Length < 31 ? doc.NomFichier : doc.NomFichier.Substring(0, 30);

                    string filePath = Path.Combine(
                    FileStoragePath,
                    "Rapports",                     // Dossier des Rapports
                    "R_" + doc.Rapport.RapportId, // Dossier du rapport en cours
                    "Documents",                    // Dossier des Photos 
                    doc.Key + doc.ExtensionFichier
                    );

                    sheet_5_8D.Cells[0, 1].PutValue(doc.NomFichier);
                    sheet_5_8D.Cells[1, 1].PutValue(doc.Description);

                    if (doc.MimeType.Contains("image"))
                    {

                        int pictureIndex = onglet_Doc.Pictures.Add(3, 1, this.Document(doc.DocumentId).FileStream);
                        Aspose.Cells.Drawing.Picture pic = onglet_Doc.Pictures[pictureIndex];
                    }
                    else
                    {
                        FileStream fstream = System.IO.File.OpenRead(filePath);
                        byte[] objectBytes = new Byte[fstream.Length];
                        fstream.Read(objectBytes, 0, objectBytes.Length);


                        onglet_Doc.OleObjects.Add(3, 1, 100, 105, imageData);
                        onglet_Doc.OleObjects[0].ObjectData = objectBytes;
                        //onglet_Doc.OleObjects[0].FileType = Aspose.Cells.Drawing.OleFileType.MapiMessage;
                        onglet_Doc.OleObjects[0].SourceFullName = doc.NomFichier;
                    }
                }

                //int indiceWorksheet = workbook.Worksheets.Add();//Adding a new worksheet to the Workbook object            
                //Worksheet sheet_Docs = workbook.Worksheets[indiceWorksheet];//Obtaining the reference of the newly added worksheet by passing its sheet index
                //sheet_Docs.Name = "Documents";//Setting the name of the newly added worksheet




                //string path = Path.Combine(baseDir, "Scripts", "jquery-1.10.2.min.map");//"dataDir" + "book1.xls"; //Get an excel file path in a variable.

                //fs = System.IO.File.OpenRead(path);//Get the file into the streams.

                //byte[] objectData = new Byte[fs.Length];//Define an array of bytes.            
                //fs.Read(objectData, 0, objectData.Length);//Store the file from streams.

                //sheet_Docs.OleObjects.Add(14, 3, 200, 220, imageData);
                //sheet_Docs.OleObjects[0].ObjectData = objectData;
                //sheet_Docs.OleObjects[0].Name = "nom nnnddaadada";
                //sheet_Docs.OleObjects[0].Text = "text nnnddaadada";
                //sheet_Docs.OleObjects[0].Title = "titre nnnddaadada";
                //sheet_Docs.OleObjects[0].AlternativeText = "AlternativeText nnnddaadada";

                // *********************************************************
                //          Finalisation
                // *********************************************************
                workbook.FileFormat = FileFormatType.Xlsx;
                Byte[] buffer = workbook.SaveToStream().ToArray();

                return File(buffer,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "Rapport_8D " + rapport.ReferenceRapport + ".xls");
                    }

             #endregion

        private void MarshalEditObject(string attribut, string value, ref object o)
        {
            System.Reflection.PropertyInfo propertyInfo = o.GetType().GetProperty(attribut);

            Type propertyType = Nullable.GetUnderlyingType(propertyInfo.PropertyType)
                ?? propertyInfo.PropertyType;

            object safeValue = (value == null) ? null
                                               : Convert.ChangeType(value, propertyType);

            propertyInfo.SetValue(o, safeValue, null);
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