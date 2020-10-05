using GRC.Models;
using GRC.Models.Security;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Net.Mail;
using System.Web;
using System.Web.Mvc;

namespace GRC.Controllers
{
    public class AuditController : BaseController
    {
        private GRCModelContainer db = new GRCModelContainer();
        // GET: Audit
        public ActionResult Index()
        {
            return View();
        }

        //[Authorize, GRCAuthorize(Roles = "Admin,DQSSE,RSQ,Auditeur")]
        public ActionResult CheckList(int id)

        {
            //Récupérer la checklist du rapport
            var rapport = db.RapportSet.Find(id);
            var checklist = rapport.RapportAuditClotureCheckList;//.ToList();

            var chechPontsList = db.AuditClotureCheckPoints.OrderBy(x=>x.Ordre).ToList();
            // si le rapport ne possède pas de checklist encore on lui crée la checklist
            if (checklist.Count > 0) 
            { 
                // rien pour l'instant
            }
            else
            {
                foreach (var point in chechPontsList)
                {
                    rapport.RapportAuditClotureCheckList.Add(new RapportAuditClotureCheckList
                    {
                        CheckPointId = point.PointId,
                        AuditClotureCheckPoints = point,
                        DateMaj = DateTime.Now
                    });
                }
                db.Entry(rapport).State = System.Data.Entity.EntityState.Modified;
                db.SaveChanges();

                return RedirectToAction("CheckList", new { id = id });
            }

            // il est préférable que dans la vue on boucle sur la checklist du rapport au lieu des checkpoints déja enregistrés
            // ....fait.

            //ViewBag.ChechPontsList = chechPontsList;

            return View(rapport);
        }

        //[HttpPost]
        //public JsonResult EditField(int pk, string name, string value)
        //{
        //    var r = new
        //    {
        //        state = false,
        //        message = "Pas d'exécution."
        //    };
        //    if (pk > 0)
        //    {
        //        RapportAuditClotureCheckList checklist = db.RapportAuditClotureCheckList.Find(pk);
        //        if (checklist == null)
        //        {
        //            r = new
        //            {
        //                state = false,
        //                message = "Clé pour checklist erroné."
        //            };
        //        }
        //        else
        //        {
        //            try
        //            {
        //                Type t = checklist.GetType().GetProperty(name).GetType();
        //                System.Reflection.PropertyInfo propertyInfo = checklist.GetType().GetProperty(name);//, t).SetValue(rapport, value);
        //                //propertyInfo.SetValue(rapport, Convert.ChangeType(value, propertyInfo.PropertyType), null);

        //                Type type = Nullable.GetUnderlyingType(propertyInfo.PropertyType)
        //                    ?? propertyInfo.PropertyType;

        //                object safeValue = (value == null) ? null
        //                                                   : Convert.ChangeType(value, type);

        //                propertyInfo.SetValue(checklist, safeValue, null);

        //                //Type type = your_class.GetType();
        //                //PropertyInfo propinfo = type.GetProperty("hasBeenPaid");
        //                //value = propinfo.GetValue(your_class, null);

        //                checklist.DateMaj = DateTime.Now;
        //                db.Entry(checklist).State = EntityState.Modified;

        //                db.SaveChanges();

        //                r = new
        //                {
        //                    state = true,
        //                    message = "Valeur changé."
        //                };
        //            }
        //            catch (Exception e)
        //            {
        //                r = new
        //                {
        //                    state = false,
        //                    message = "Exception : " + e.Message
        //                };
        //            }
        //        }
        //    }
        //    else
        //    {
        //        r = new
        //        {
        //            state = false,
        //            message = "Clé pour checklist invalide."
        //        };
        //    }

        //    return Json(r, JsonRequestBehavior.AllowGet);
        //}

        [HttpPost]
        [Authorize, GRCAuthorize(Roles = "Admin,DQSSE,RSQ,Auditeur")]
        public ActionResult CloturerAudit(FormCollection form)
        {
            try
            {
                var RapportId = int.Parse(form["RapportId"]);
                var rapport = db.RapportSet.Find(RapportId);
                var statut = form["Decision"];
                rapport.DateAudit = DateTime.Now;
                rapport.AuditDecision = true;
                //if ((statut == "NA") || (statut == "OUI"))
                if (statut == "OUI")
                {


                    db.Entry(rapport).State = EntityState.Modified;
                    db.SaveChanges();

                    TempData["success"] = "L'audit du rapport " + rapport.ReferenceRapport + " a été clôturé avec succès";


                    this.SendMailCloturerAudit(rapport);
                }

                else
                {
                    rapport.IsActif = false;
                    rapport.DateDesactivation = DateTime.Now;

                    db.SaveChanges();
                    TempData["Faillure"] = "L'audit du rapport " + rapport.ReferenceRapport + " a été bloqué ";
                }
            }
            catch (Exception e)
            {
                TempData["error"] = "Exception système : " + e.Message;
            }

            return RedirectToAction("EnAttente");
        }

        private void SendMailCloturerAudit(Rapport rapport)
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

                    // Mettre le responsable désignateur de l'auditeur en copie
                    if (!string.IsNullOrEmpty(rapport.ResponsableDesignantAuditeur.Mail))
                        mail.CC.Add(new MailAddress(rapport.ResponsableDesignantAuditeur.Mail, rapport.ResponsableDesignantAuditeur.NomComplet));
                    // Mettre en copie l'Auditeur
                    if (!string.IsNullOrEmpty(rapport.Auditeur.Mail))
                        mail.CC.Add(new MailAddress(rapport.Auditeur.Mail, rapport.Auditeur.NomComplet));


                    var RolesList = new List<string> { "Auditeur", "Admin", "DQSSE", "RSQ" };
                    

                    // Affectation des destinataires du mail
                    if (!string.IsNullOrEmpty(rapport.Pilote.Mail))
                        mail.To.Add(new MailAddress(rapport.Pilote.Mail, rapport.Pilote.NomComplet));
                    


                    mail.Subject = "GRC - Décision de Clôure de Rapport d'audit";
                    mail.IsBodyHtml = true;
                    mail.Body = "Notifiaction pour la Décision de Clôure de Rapport d'audit";

                    string htmlContent = this.RenderViewToString("~/Views/MailTemplates/CloturerAudit.cshtml", rapport);

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
                var CheckListItem = db.RapportAuditClotureCheckList.Find(pk);
                if (CheckListItem == null)
                {
                    r = new
                    {
                        state = false,
                        message = "Clé pour RapportAuditClotureCheckList erroné."
                    };
                }
                else
                {
                    try
                    {
                        Type t = CheckListItem.GetType().GetProperty(name).GetType();
                        System.Reflection.PropertyInfo propertyInfo = CheckListItem.GetType().GetProperty(name);//, t).SetValue(rapport, value);
                        //propertyInfo.SetValue(rapport, Convert.ChangeType(value, propertyInfo.PropertyType), null);

                        Type type = Nullable.GetUnderlyingType(propertyInfo.PropertyType)
                            ?? propertyInfo.PropertyType;

                        object safeValue = (value == null) ? null
                                                           : Convert.ChangeType(value, type);

                        propertyInfo.SetValue(CheckListItem, safeValue, null);

                        //Type type = your_class.GetType();
                        //PropertyInfo propinfo = type.GetProperty("hasBeenPaid");
                        //value = propinfo.GetValue(your_class, null);

                        CheckListItem.DateMaj = DateTime.Now;
                        db.Entry(CheckListItem).State = EntityState.Modified;

                        db.SaveChanges();

                        r = new
                        {
                            state = true,
                            message = "CheckListItem changé."
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
                    message = "Clé pour RapportAuditClotureCheckList invalide."
                };
            }

            return Json(r, JsonRequestBehavior.AllowGet);
        }

        [Authorize, GRCAuthorize(Roles = "Admin,DQSSE,RSQ")]
        public ActionResult EnCours()
        {
            var rapports/*EnAttenteAffectaion*/ = db.RapportSet/*.Where(x =>
                x.DemandeAuditCloture == true &&
                !(x.DesignationAuditeurDate.HasValue || x.DesignationAuditeurResponsableId.HasValue)
                )*/
                ;

            return View(rapports.OrderByDescending(x => x.DateOuverture));
          //  return View(rapports.OrderByDescending/*EnAttenteAffectaion*/);
        }

        [HttpPost]
        [Authorize, GRCAuthorize(Roles = "Admin,DQSSE,RSQ")]
        public ActionResult AffecterAuditeurRapport(int RapportId, int AuditeurId)
        {
            if (RapportId > 0 && AuditeurId > 0)
            {
                var rapport = db.RapportSet.Find(RapportId);
                var auditeur = db.PersonnelSet.Find(AuditeurId);


                rapport.Auditeur = auditeur;
                rapport.AuditeurId = auditeur.PersonnelId;
                rapport.DesignationAuditeurDate = DateTime.Now;

                var responsableDesignateur = db.PersonnelSet.Find(User.PersonnelId);
                rapport.DesignationAuditeurResponsableId = User.PersonnelId;
                rapport.ResponsableDesignantAuditeur = responsableDesignateur;

                db.Entry(rapport).State = EntityState.Modified;
                db.SaveChanges();

                // TODO : on envoi un mail à la personne désigné comme Auditeur

                this.SendMailAffectationAuditeur(rapport, auditeur, responsableDesignateur);

                TempData["success"] = "<b>" + auditeur.NomComplet + "</b> a été choisi pour auditer le Rapport <b>" + rapport.ReferenceRapport + "</b>.";
            }
            else
            {
                TempData["error"] = "Erreur lors de l'affectation de l'auditeur au rapport sélectionné.";
            }
            return RedirectToAction("EnCours");
        }

        private void SendMailAffectationAuditeur(Rapport rapport, Personnel auditeur, Personnel responsableDesignateur)
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

                    // Mettre le responsable désignateur de l'auditeur en copie
                    if (!string.IsNullOrEmpty(rapport.ResponsableDesignantAuditeur.Mail))
                        mail.CC.Add(new MailAddress(rapport.ResponsableDesignantAuditeur.Mail, rapport.ResponsableDesignantAuditeur.NomComplet));
                    // Mettre en copie le Pilote du rapport
                    if (!string.IsNullOrEmpty(rapport.Pilote.Mail))
                        mail.CC.Add(new MailAddress(rapport.Pilote.Mail, rapport.Pilote.NomComplet));


                    // Affectation des destinataires du mail
                    var RolesList = new List<string> { "Auditeur", "Admin", "DQSSE", "RSQ" };
                    //var destinataires =
                    //    (from users in db.Users
                    //     from roles in users.Roles
                    //     where RolesList.Contains(roles.RoleName)
                    //     select users.Personne).Distinct().ToList();
                    ////db.Users.Where(x => x.Roles.Select(r=>r.RoleName).Intersect(new List<string>{ "" }).Count()>0);

                    //foreach (var dest in destinataires)
                    //{
                    if (!string.IsNullOrEmpty(rapport.Auditeur.Mail))
                        mail.To.Add(new MailAddress(rapport.Auditeur.Mail, rapport.Auditeur.NomComplet));
                    //}

                    mail.Subject = "GRC - Affectation Audit de Clôture de Rapport";
                    mail.IsBodyHtml = true;
                    mail.Body = "Notifiaction pour Affectation Audit de Clôture de Rapport";

                    string htmlContent = this.RenderViewToString("~/Views/MailTemplates/AffectationAuditeur.cshtml", rapport);

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

        [Authorize, GRCAuthorize(Roles = "Auditeur,Admin,DQSSE,RSQ")]
        public ActionResult EnAttente()
        {
            
            var rapports = db.RapportSet.Where(x =>
                x.DemandeAuditCloture == true 
                //&& x.DesignationAuditeurResponsableId == User.PersonnelId
                )
                ;
            return View(rapports.OrderByDescending(x => x.DateOuverture));
        }

        [Authorize, GRCAuthorize(Roles = "Admin,DQSSE,RSQ")]
        public ActionResult Auditeurs()
        {
            return View();
        }

        [Authorize, GRCAuthorize(Roles = "Admin,DQSSE,RSQ")]
        public ActionResult DesignationAuditeurs()
        {
            return View();
        }


    }
}