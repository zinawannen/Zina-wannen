using GRC.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GRC.Controllers
{
    public class StatistiquesController : Controller
    {
        private GRCModelContainer db = new GRCModelContainer();
        //
        // GET: /Statistiques/
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// Une représentation graphique du nombre de réclamation par zone (UAPs, usine) incluant aussi le type d’incident.
        /// </summary>
        /// <returns></returns>
        public ActionResult ReclamationsParZone()
        {
            var series = new Dictionary<string, List<int>>();
            var seriesL = new List<object>();
            var seriesAlertes = new Dictionary<string, List<int>>();
            var seriesLAlertes = new List<object>();

            var list_UAP = db.Configs.Find("UAP").Val.Split(',');
            var list_TypesIncidents = db.Configs.Find("TypesIncidents").Val.Split(',');

            var lst = new List<object>();
            #region Reclamations par Zone 
            foreach (var typeInc in list_TypesIncidents)
            {
                var ttypeInc = typeInc.Trim();
                var listCountUAPs = new List<int>();


                foreach (var uap in list_UAP)
                {
                    var q = db.RapportSet.Where(x =>
                        x.UAP == uap
                        && x.TypeIncident == ttypeInc
                        && x.IsActif == true
                        && x.Alerte == false).Count();
                    listCountUAPs.Add(q);
                }
                seriesL.Add(new
                {
                    name = ttypeInc,
                    data = listCountUAPs
                });
            }
            ViewBag.series = new System.Web.Script.Serialization.JavaScriptSerializer()
                .Serialize(Json(seriesL).Data);
            ViewBag.xAxis = new System.Web.Script.Serialization.JavaScriptSerializer()
                .Serialize(Json(list_UAP).Data);

            #endregion

            #region Alertes par Zone 

            foreach (var typeInc in list_TypesIncidents)
            {
                var ttypeInc = typeInc.Trim();
                var listCountUAPs = new List<int>();


                foreach (var uap in list_UAP)
                {
                    var q = db.RapportSet.Where(x =>
                        x.UAP == uap
                        && x.TypeIncident == ttypeInc
                        && x.IsActif == true
                        && x.Alerte == true).Count();
                    listCountUAPs.Add(q);
                }
                seriesLAlertes.Add(new
                {
                    name = ttypeInc,
                    data = listCountUAPs
                });
            }
            ViewBag.seriesAlertes = new System.Web.Script.Serialization.JavaScriptSerializer()
                .Serialize(Json(seriesLAlertes).Data);
            ViewBag.xAxisAlertes = new System.Web.Script.Serialization.JavaScriptSerializer()
                .Serialize(Json(list_UAP).Data);
            #endregion

            return View();
        }

        /// <summary>
        /// Une représentation graphique du nombre de réclamation par type de réclamation (CMS, vague, …)
        /// </summary>
        /// <returns></returns>
        public ActionResult ReclamationsParTypeReclamation(FormCollection form)
        {

            string UAP = form["UAP"];

            string Annee = form["Annee"];
            var list_Annee = new List<string>();
            for (int i = 2017; i <= DateTime.Now.Year; i++)
            {
                list_Annee.Add(i.ToString());
            }
            ViewBag.Annee = new SelectList(list_Annee, Annee);
            if (Annee == null)
            {
                Annee = DateTime.Now.Year.ToString();
                ViewBag.Default = Annee;
            }

            if (Annee == "")
            {
                Annee = DateTime.Now.Year.ToString();
                ViewBag.Default = Annee;
            }

            int year = Int32.Parse(Annee);
            DateTime dateDeb1 = new DateTime(year, 1, 1);

            // var UAP = TempData["UAP"];

            var list_UAP = db.Configs.Find("UAP").Val.Split(',');
            ViewBag.UAP = new SelectList(list_UAP, UAP);



            if (UAP != null)
            {

                dynamic data1 =
                    (from rapport in db.RapportSet
                     where rapport.IsActif == true
                     && rapport.Alerte == false
                     && rapport.DateOuverture >= dateDeb1 && rapport.UAP == UAP
                     group rapport by rapport.TypeReclamation.Libelle into reclamations
                     select new
                     {
                         Reclamation = reclamations.Key,
                         Nbre = reclamations.Count()
                     })
                     .ToList().Select(x => x.ToExpando());//.AsEnumerable().Select(x => new { x.Reclamation, x.Nbre });
                                                          ////var data = q.ToList();
                                                          //return View(data);            }

                dynamic data =
              (from rapport in db.RapportSet
               where rapport.IsActif == true && rapport.Alerte == true && rapport.DateOuverture >= dateDeb1 && rapport.UAP == UAP
               group rapport by rapport.TypeReclamation.Libelle into reclamations
               select new
               {
                   Reclamation = reclamations.Key,
                   Nbre = reclamations.Count()
               })
               .ToList().Select(x => x.ToExpando());//.AsEnumerable().Select(x => new { x.Reclamation, x.Nbre });

               
                ViewBag.data = data;
                ViewBag.data1 = data1;
            }
            else
            {

                dynamic data1 =
                    (from rapport in db.RapportSet
                     where rapport.IsActif == true
                     && rapport.Alerte == false
                     && rapport.DateOuverture >= dateDeb1
                     group rapport by rapport.TypeReclamation.Libelle into reclamations
                     select new
                     {
                         Reclamation = reclamations.Key,
                         Nbre = reclamations.Count()
                     })
                     .ToList().Select(x => x.ToExpando());


                dynamic data =
              (from rapport in db.RapportSet
               where rapport.IsActif == true && rapport.Alerte == true && rapport.DateOuverture >= dateDeb1
               group rapport by rapport.TypeReclamation.Libelle into reclamations
               select new
               {
                   Reclamation = reclamations.Key,
                   Nbre = reclamations.Count()
               })
               .ToList().Select(x => x.ToExpando());//.AsEnumerable().Select(x => new { x.Reclamation, x.Nbre });

                ViewBag.data1 = data1;
                ViewBag.data = data;                                       ////var data = q.ToList();
                                                                           //return View(data);  
            }


            return View();
        }

        /// <summary>
        /// Evolution globale par type d’incident
        /// </summary>
        /// <returns></returns>
        public ActionResult EvolutionParTypeIncident()

        {
            var dateDeb = DateTime.Now.AddYears(-1);

            //DateTime dateDeb1 = new DateTime(year, 1, 1);
            var cYear = dateDeb.Year;
            var cMonth = dateDeb.Month;
            var dateIntervalList = new List<string>();

            while (true) //!(cYear == DateTime.Now.Year && cMonth == DateTime.Now.Year))
            {
                while (true) //cMonth != 13 || !(cYear == DateTime.Now.Year && cMonth - 1 == DateTime.Now.Year))
                {
                    dateIntervalList.Add(cMonth + "/" + cYear);

                    if (cMonth == 12) break;
                    if (cYear == DateTime.Now.Year && cMonth == DateTime.Now.Month) break;

                    ++cMonth;
                }
                if (cYear == DateTime.Now.Year && cMonth == DateTime.Now.Month) break;
                ++cYear;
                cMonth = 1;

            }




           // http://stackoverflow.com/questions/223713/can-i-pass-an-anonymous-type-to-my-asp-net-mvc-view

            #region Reclamation Par incident
            var recs =
                (from rapport in db.RapportSet
                 where rapport.DateOuverture > dateDeb
                 && rapport.IsActif == true
                 //&& rapport.Alerte == false
                 group rapport by new
                 {
                     rapport.TypeIncident,
                     rapport.DateOuverture.Year,
                     rapport.DateOuverture.Month
                 } into incidents
                 orderby new
                 {
                     incidents.Key.Year,
                     incidents.Key.Month,
                     incidents.Key.TypeIncident
                 }
                 select new
                 {
                     Incident = incidents.Key.TypeIncident,
                     Annee = incidents.Key.Year,
                     Mois = incidents.Key.Month,
                     Nbre = incidents.Count()
                 })
                 .ToList(); //.Select(x => x.ToExpando());
            var res = new List<object>();
            var list_TypesIncidents = db.Configs.Find("TypesIncidents").Val.Split(',');

            foreach (var item in dateIntervalList)
            {
                var rec = recs.Where(x => x.Mois + "/" + x.Annee == item);

                if (rec == null)
                {
                    res.Add(new
                    {
                        Date = item,
                        Client = 0,
                        ClientDirect = 0,
                        ClientFinal = 0,
                        Interne = 0,
                        Fournisseur = 0
                    });
                }
                else
                {
                    var Client = rec.Where(x => x.Incident == "Client").FirstOrDefault();
                    var ClientDirect = rec.Where(x => x.Incident == "Client Direct").FirstOrDefault();
                    var ClientFinal = rec.Where(x => x.Incident == "Client Final").FirstOrDefault();
                    var Interne = rec.Where(x => x.Incident == "Interne").FirstOrDefault();
                    var Fournisseur = rec.Where(x => x.Incident == "Fournisseur").FirstOrDefault();

                    res.Add(new
                    {
                        Date = item,
                        Client = Client == null ? 0 : Client.Nbre,
                        ClientDirect = ClientDirect == null ? 0 : ClientDirect.Nbre,
                        ClientFinal = ClientFinal == null ? 0 : ClientFinal.Nbre,
                        Interne = Interne == null ? 0 : Interne.Nbre,
                        Fournisseur = Fournisseur == null ? 0 : Fournisseur.Nbre
                    });
                }
            }
            ViewBag.ReclamationsParIncident = res.Select(x => x.ToExpando());
            #endregion
            #region Alertes Par incident
            var recsAlertes =
                (from rapport in db.RapportSet
                 where rapport.DateOuverture > dateDeb
                 && rapport.IsActif == true
                 && rapport.Alerte == true
                 group rapport by new
                 {
                     rapport.TypeIncident,
                     rapport.DateOuverture.Year,
                     rapport.DateOuverture.Month
                 } into incidents
                 orderby new
                 {
                     incidents.Key.Year,
                     incidents.Key.Month,
                     incidents.Key.TypeIncident
                 }
                 select new
                 {
                     Incident = incidents.Key.TypeIncident,
                     Annee = incidents.Key.Year,
                     Mois = incidents.Key.Month,
                     Nbre = incidents.Count()
                 })
                 .ToList(); //.Select(x => x.ToExpando());
            var resAlertes = new List<object>();
            // var list_TypesIncidents = db.Configs.Find("TypesIncidents").Val.Split(',');

            foreach (var item in dateIntervalList)
            {
                var rec1 = recsAlertes.Where(x => x.Mois + "/" + x.Annee == item);

                if (rec1 == null)
                {
                    resAlertes.Add(new
                    {
                        Date = item,
                        Client = 0,
                        ClientDirect = 0,
                        ClientFinal = 0,
                        Interne = 0,
                        Fournisseur = 0
                    });
                }
                else
                {
                    var Client = rec1.Where(x => x.Incident == "Client").FirstOrDefault();
                    var ClientDirect = rec1.Where(x => x.Incident == "Client Direct").FirstOrDefault();
                    var ClientFinal = rec1.Where(x => x.Incident == "Client Final").FirstOrDefault();
                    var Interne = rec1.Where(x => x.Incident == "Interne").FirstOrDefault();
                    var Fournisseur = rec1.Where(x => x.Incident == "Fournisseur").FirstOrDefault();

                    resAlertes.Add(new
                    {
                        Date = item,
                        Client = Client == null ? 0 : Client.Nbre,
                        ClientDirect = ClientDirect == null ? 0 : ClientDirect.Nbre,
                        ClientFinal = ClientFinal == null ? 0 : ClientFinal.Nbre,
                        Interne = Interne == null ? 0 : Interne.Nbre,
                        Fournisseur = Fournisseur == null ? 0 : Fournisseur.Nbre
                    });
                }
            }
            ViewBag.AlertesParIncident = resAlertes.Select(x => x.ToExpando());
            #endregion
            //  return View(res.Select(x => x.ToExpando()));
            return View();
        }

        public ActionResult EvolutionParUAP(string Annee)
        {
            //////*////
            var list_Annee = new List<string>();
            for (int i = 2017; i <= DateTime.Now.Year; i++)
            {
                list_Annee.Add(i.ToString());
            }
            ViewBag.Annee = new SelectList(list_Annee, Annee);

            //  int annee = DateTime.Now.Year;


            if (Annee == null)
            {
                Annee = DateTime.Now.Year.ToString();
                ViewBag.Default = Annee;
            }


            int year = Int32.Parse(Annee);
            DateTime dateDeb1 = new DateTime(year, 1, 1);
            DateTime firstDaylastYear = new DateTime(dateDeb1.Year, 1, 1);
            var clastYear = dateDeb1.Year;
            var clastYearMonth = firstDaylastYear.Month;

            //  var cYear = dateDeb.Year;
            // var cMonth = dateDeb.Month;
            var dateIntervalList = new List<string>();

            //while (true) //!(cYear == DateTime.Now.Year && cMonth == DateTime.Now.Year))
            //{
            while (true)//cMonth != 13 || !(cYear == DateTime.Now.Year && cMonth - 1 == DateTime.Now.Year))
            {
                dateIntervalList.Add(clastYearMonth + "/" + clastYear);

                if (clastYearMonth == 12) break;
                if (clastYear == DateTime.Now.Year && clastYearMonth == DateTime.Now.Month) break;

                ++clastYearMonth;
            }



            #region Evolution des reclamations Client
            var recs =
                (from rapport in db.RapportSet
                 where rapport.DateOuverture >= dateDeb1
                 && rapport.IsActif == true
                 && rapport.Alerte == false
                 group rapport by new
                 {
                     rapport.UAP,
                     rapport.DateOuverture.Year,
                     rapport.DateOuverture.Month
                 } into incidents
                 orderby new
                 {
                     incidents.Key.Year,
                     incidents.Key.Month,
                     incidents.Key.UAP
                 }
                 select new
                 {
                     UAP = incidents.Key.UAP,
                     Annee = incidents.Key.Year,
                     Mois = incidents.Key.Month,
                     Nbre = incidents.Count()
                 })
                 .ToList(); //.Select(x => x.ToExpando());
            var res = new List<object>();
            foreach (var item in dateIntervalList)
            {
                var rec = recs.Where(x => x.Mois + "/" + x.Annee == item);

                if (rec == null)
                {
                    res.Add(new
                    {
                        VALEO = 0,
                        IEE = 0,
                        FRESINUS = 0,
                        GE = 0,
                        PES = 0,
                        NumItalie = 0,
                        Cahor = 0,
                        Zender = 0,
                        // TPO = 0,
                        // Denso = 0,
                        MontBlanc = 0,
                        SANDEN = 0,
                        BSH = 0,
                        Schneider = 0,
                        Tot = 0,
                    });
                }
                else
                {

                    var Schneider = rec.Where(x => (x.UAP.Contains("Schneider") | x.UAP.Contains("MontBlanc"))).FirstOrDefault();

                    var VALEO = rec.Where(x => x.UAP == "VALEO").FirstOrDefault();
                    var IEE = rec.Where(x => x.UAP == "IEE").FirstOrDefault();
                    var FRESINUS = rec.Where(x => x.UAP == "FRESINUS").FirstOrDefault();
                    var GE = rec.Where(x => x.UAP == "GE" | x.UAP == "ALSTOM").FirstOrDefault();
                    var BSH = rec.Where(x => x.UAP == "BSH").FirstOrDefault();
                    var SANDEN = rec.Where(x => x.UAP == "SANDEN").FirstOrDefault();

                    var Zender = rec.Where(x => x.UAP == "Zender").FirstOrDefault();
                    var Cahor = rec.Where(x => x.UAP == "Cahor").FirstOrDefault();
                    var NumItalie = rec.Where(x => x.UAP == "GE").FirstOrDefault();
                    var PES = rec.Where(x => x.UAP == "PES").FirstOrDefault();

                    var tot = 0;
                    var totrec = new List<int>(); ;
                    ///*rec.Count();*/
                    foreach (var i in rec)
                    {
                        tot += i.Nbre;
                        //totrec.Add(tot);
                    }

                    res.Add(new
                    {
                        Date = item,
                        VALEO = VALEO == null ? 0 : VALEO.Nbre,
                        IEE = IEE == null ? 0 : IEE.Nbre,
                        FRESINUS = FRESINUS == null ? 0 : FRESINUS.Nbre,
                        GE = GE == null ? 0 : GE.Nbre,
                        BSH = BSH == null ? 0 : BSH.Nbre,
                        SANDEN = SANDEN == null ? 0 : SANDEN.Nbre,

                        Schneider = Schneider == null ? 0 : Schneider.Nbre,

                        Zender = Zender == null ? 0 : Zender.Nbre,
                        Cahor = Cahor == null ? 0 : Cahor.Nbre,
                        NumItalie = NumItalie == null ? 0 : NumItalie.Nbre,
                        PES = PES == null ? 0 : PES.Nbre,
                        tot = tot == 0 ? 0 : tot
                    });
                }
            }
            // ViewBag.EvolutionDesReclamationsClient = res.Select(x => x.ToExpando());
            ViewBag.EvolutionParUAP = res.Select(x => x.ToExpando());
            #endregion
            #region Paretos Alertes Clients
            var recsAlertes =
            (from rapport in db.RapportSet
             where rapport.DateOuverture >= dateDeb1
             && rapport.IsActif == true
              && rapport.Alerte == true
             group rapport by new
             {
                 rapport.UAP,
                 rapport.DateOuverture.Year,
                 rapport.DateOuverture.Month
             } into incidents
             orderby new
             {
                 incidents.Key.Year,
                 incidents.Key.Month,
                 incidents.Key.UAP
             }
             select new
             {
                 UAP = incidents.Key.UAP,
                 Annee = incidents.Key.Year,
                 Mois = incidents.Key.Month,
                 Nbre = incidents.Count()
             })
             .ToList();//.Select(x => x.ToExpando());
            var resAlertes = new List<object>();
            foreach (var i in dateIntervalList)
            {
                var recAlertes = recsAlertes.Where(x => x.Mois + "/" + x.Annee == i);

                if (recAlertes == null)
                {
                    resAlertes.Add(new
                    {
                        VALEO = 0,
                        IEE = 0,
                        FRESINUS = 0,
                        GE = 0,
                        PES = 0,
                        NumItalie = 0,
                        Cahor = 0,
                        Zender = 0,
                        // TPO = 0,
                        // Denso = 0,
                        MontBlanc = 0,
                        SANDEN = 0,
                        BSH = 0,
                        Schneider = 0,
                        Tot = 0,

                    });
                }
                else
                {
                    var Schneider = recAlertes.Where(x => (x.UAP.Contains("Schneider") | x.UAP.Contains("MontBlanc"))).FirstOrDefault();

                    var VALEO = recAlertes.Where(x => x.UAP == "VALEO").FirstOrDefault();
                    var IEE = recAlertes.Where(x => x.UAP == "IEE").FirstOrDefault();
                    var FRESINUS = recAlertes.Where(x => x.UAP == "FRESINUS").FirstOrDefault();
                    var GE = recAlertes.Where(x => x.UAP == "GE" | x.UAP == "ALSTOM").FirstOrDefault();
                    var BSH = recAlertes.Where(x => x.UAP == "BSH").FirstOrDefault();
                    var SANDEN = recAlertes.Where(x => x.UAP == "SANDEN").FirstOrDefault();

                    var Zender = recAlertes.Where(x => x.UAP == "Zender").FirstOrDefault();
                    var Cahor = recAlertes.Where(x => x.UAP == "Cahor").FirstOrDefault();
                    var NumItalie = recAlertes.Where(x => x.UAP == "GE").FirstOrDefault();
                    var PES = recAlertes.Where(x => x.UAP == "PES").FirstOrDefault();

                    var tot = 0;
                    var totrec = new List<int>(); ;
                    ///*rec.Count();*/
                    foreach (var item in recAlertes)
                    {
                        tot += item.Nbre;
                        //totrec.Add(tot);
                    }

                    resAlertes.Add(new
                    {
                        Date = i,
                        VALEO = VALEO == null ? 0 : VALEO.Nbre,
                        IEE = IEE == null ? 0 : IEE.Nbre,
                        FRESINUS = FRESINUS == null ? 0 : FRESINUS.Nbre,
                        GE = GE == null ? 0 : GE.Nbre,
                        BSH = BSH == null ? 0 : BSH.Nbre,
                        SANDEN = SANDEN == null ? 0 : SANDEN.Nbre,

                        Schneider = Schneider == null ? 0 : Schneider.Nbre,

                        Zender = Zender == null ? 0 : Zender.Nbre,
                        Cahor = Cahor == null ? 0 : Cahor.Nbre,
                        NumItalie = NumItalie == null ? 0 : NumItalie.Nbre,
                        PES = PES == null ? 0 : PES.Nbre,
                        tot = tot == 0 ? 0 : tot
                    });
                }
            }
            ViewBag.EvolutionParUAPAlertes = resAlertes.Select(x => x.ToExpando());
            #endregion

            return View();
        }

        public ActionResult TauxReactivite(FormCollection form)
        {
            //****************************************************************************************************
            //var AnneeDeMiseEnService = 2015;
            string UAP = form["UAP"];

            string Annee = form["Annee"];
            var list_Annee = new List<string>();
            for (int i = 2017; i <= DateTime.Now.Year; i++)
            {
                list_Annee.Add(i.ToString());
            }
            ViewBag.Annee = new SelectList(list_Annee, Annee);
            if (Annee == null)
            {
                Annee = DateTime.Now.Year.ToString();
                ViewBag.Default = Annee;
            }

            if (Annee == "")
            {
                Annee = DateTime.Now.Year.ToString();
                ViewBag.Default = Annee;
            }

            int year = Int32.Parse(Annee);
            DateTime dateDeb1 = new DateTime(year, 1, 1);

            // var UAP = TempData["UAP"];

            var list_UAP = db.Configs.Find("UAP").Val.Split(',');
            ViewBag.UAP = new SelectList(list_UAP, UAP);

            List<string> Etape = new List<string>();
            Dictionary<string, string> Eta = new Dictionary<string, string>();

            var series = new Dictionary<string, List<int>>();
            var seriesL = new List<object>();
            int poucentage = 0;
            var objectifs = db.Configs.Where(x => x.Key.StartsWith("DelaiObjectif.") && x.IsActive == true).ToList();
            List<Rapport> rapports = new List<Rapport>();
            if (UAP == null | UAP=="")
            {
               rapports = db.RapportSet.Where(x => x.IsActif == true && x.DateOuverture.Year == year).ToList();
            }
            else
            {
                rapports = db.RapportSet.Where(x => x.IsActif == true && x.DateOuverture.Year == year && x.UAP == UAP).ToList();
            }
            var RapportsClotures = new List<int>();
            for (int i = 0; i < objectifs.Count();)
            {

                foreach (var item in objectifs)
                {
                    if (item.Key.Contains("DelaiObjectif"))
                    {
                        Etape.Add(item.Key.Substring(item.Key.Length - 2));


                        Eta.Add(Etape[i], item.Val);
                        i++;
                    }
                }


            }
            var listCountrapports = new List<int>();
            var DelaiCloture = new Reactivity();

            foreach (var etap in Eta)
            {

                int result = Int32.Parse(etap.Value);
                foreach (var r in rapports)
                {
                    /// timespan for each step **//////
                    //modification 17/06/2019//
                    var delai_2D = (int)(r.D2_DateCloture - r.DateOuverture).GetValueOrDefault().TotalDays;
                    var delai_3D = (int)(r.D3_DateCloture - r.DateOuverture).GetValueOrDefault().TotalDays;
                    var delai_4D = (int)(r.D4_DateCloture - r.DateOuverture).GetValueOrDefault().TotalDays;
                    var delai_5D = (int)(r.D5_DateCloture - r.DateOuverture).GetValueOrDefault().TotalDays;
                    var delai_7D = (int)(r.D7_DateCloture - r.DateOuverture).GetValueOrDefault().TotalDays;
                    var delai_8D = (int)(r.D8_DateCloture - r.DateOuverture).GetValueOrDefault().TotalDays; 

                    //if ((r.D2_DateCloture == null) || (r.D3_DateCloture == null) || r.D4_DateCloture == null || r.D5_DateCloture == null || r.D7_DateCloture == null || r.D8_DateCloture)
                    //{

                    //}
                    //if (!delai_2D.HasValue || (!delai_3D.HasValue) || (!delai_4D.HasValue) || (!delai_5D.HasValue) || (!delai_7D.HasValue) || (!delai_8D.HasValue))
                    //{
                    //}

                    Dictionary<string, int?> delai = new Dictionary<string, int?>();
                    delai.Add("delai_2D",delai_2D);
                    delai.Add("delai_3D",delai_3D);
                    delai.Add("delai_4D",delai_4D);
                    delai.Add("delai_5D",delai_5D);
                    delai.Add("delai_7D",delai_7D);
                    delai.Add("delai_8D",delai_8D);

                    foreach (var del in delai)
                    {
                       
                        if( del.Value.HasValue & del.Key.Contains(etap.Key))
                        {

                      
                    Dictionary<string ,int?> timespan = new Dictionary<string, int?>();

                            //timespan.Add("delai_2D", delai_2D.HasValue ? delai_2D.Value : 0);
                            //timespan.Add("delai_3D", delai_3D.HasValue ? delai_3D.Value : 0);
                            //timespan.Add("delai_4D", delai_4D.HasValue ? delai_4D.Value : 0);
                            //timespan.Add("delai_5D", delai_5D.HasValue ? delai_5D.Value : 0);
                            //timespan.Add("delai_7D", delai_7D.HasValue ? delai_7D.Value : 0);
                            //timespan.Add("delai_8D", delai_8D.HasValue ? delai_8D.Value : 0);
                            timespan.Add("delai_2D", delai_2D);
                            timespan.Add("delai_3D", delai_3D);
                            timespan.Add("delai_4D", delai_4D);
                            timespan.Add("delai_5D", delai_5D );
                            timespan.Add("delai_7D", delai_7D);
                            timespan.Add("delai_8D", delai_8D);



                            foreach (var item in timespan)
                    {
                        if (item.Key.Contains(etap.Key) & item.Value  < result)
                        {
                            listCountrapports.Add(r.RapportId);
                        }
                    }

                        }

                    }

                }
                poucentage = (100 * listCountrapports.Count) / rapports.Count;
                listCountrapports.Clear(); ;
                seriesL.Add(new
                {
                   // Data = Etape,
                    Data = etap.Key,
                    Name = poucentage,
                });


            }

            ViewBag.series = seriesL.Select(x => x.ToExpando());
            //new System.Web.Script.Serialization.JavaScriptSerializer()
            //    .Serialize(Json(seriesL).Data); 
            ViewBag.xAxis = new System.Web.Script.Serialization.JavaScriptSerializer()
                  .Serialize(Json(Etape).Data);
            ViewBag.Data = Etape;
           
             return View();
        }
    }
}
           


            //****************************************************************************************************
            // List<Rapport> rapportsDeAnneeEncours = db.RapportSet.Where(x => x.IsActif == true && x.DateOuverture.Year == annee && x.DateCloture != null).ToList();

            //List<Rapport> rapportsDeAnneeEncours = db.RapportSet.Where(x => x.IsActif == true && x.DateCloture != null).ToList();
            //int Objectif = Int32.Parse(db.Configs.Find("DelaiObjectif.8D").Val);

            //ViewBag.Objectif = Objectif.ToString();
            //var res = new List<object>();
            //foreach (var item in rapportsDeAnneeEncours)
            //{

            //    int DureeTrait = 0;
            //    DureeTrait = TimeSpan.Parse((item.DateCloture - item.DateOuverture).ToString()).Days;

            //    res.Add(new
            //    {
            //        NumRapport = item.RapportId,
            //        UAP = item.UAP,
            //        Pilote = item.Pilote.NomComplet,
            //        Duree = DureeTrait
            //    });

            //}
            //List<Rapport> rapportsNonCloturés = db.RapportSet.Where(x => x.IsActif == true && x.DateCloture == null).ToList();
            //var resul1 = new List<object>();
            //foreach (var item in rapportsNonCloturés)
            //{

            //    int DureeTraitement = 0;
            //    DureeTraitement = TimeSpan.Parse((DateTime.Now - item.DateOuverture).ToString()).Days;

            //    resul1.Add(new
            //    {
            //        NumRapport = item.RapportId,
            //        UAP = item.UAP,
            //        Pilote = item.Pilote.NomComplet,
            //        Duree = DureeTraitement
            //    });

            //}
            //ViewBag.RapportCloture = resul1.Select(x => x.ToExpando());




       

        //[HttpPost, ActionName("TauxReactivite")]
        //public ActionResult OK(FormCollection form) 
        //{
        //    string sAnnee = form["Annee"];
        //    return TauxReactivite(sAnnee);
        //}
   

