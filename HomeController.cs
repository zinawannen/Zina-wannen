using GRC.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace GRC.Controllers
{
    public class HomeController : BaseController
    {
        private GRCModelContainer db = new GRCModelContainer();

        public /*async Task<ActionResult>*/ ActionResult Index(string Annee)

        {



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
            //    if (clastYear == DateTime.Now.Year /*&& clastYearMonth == DateTime.Now.Month*/) break;
            //    //++clastYear;
            //    clastYearMonth = 1;
            //}

            #region Paretos Reclamations

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
                 An = incidents.Key.Year,
                 Mois = incidents.Key.Month,
                 Nbre = incidents.Count()
             })
             .ToList();//.Select(x => x.ToExpando());
            var res = new List<object>();
            foreach (var i in dateIntervalList)
            {
                var rec = recs.Where(x => x.Mois + "/" + x.An == i);

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
                    foreach (var item in rec)
                    {
                        tot += item.Nbre;
                        //totrec.Add(tot);
                    }

                    // = tot;
                    res.Add(new
                    {


                        Date = i,
                        VALEO = VALEO == null ? 0 : VALEO.Nbre,
                        IEE = IEE == null ? 0 : IEE.Nbre,
                        FRESINUS = FRESINUS == null ? 0 : FRESINUS.Nbre,
                        GE = GE == null ? 0 : GE.Nbre,
                        BSH = BSH == null ? 0 : BSH.Nbre,
                        SANDEN = SANDEN == null ? 0 : SANDEN.Nbre,
                       
                        Schneider = Schneider == null ?0 : Schneider.Nbre,
                      
                        Zender = Zender == null ? 0 : Zender.Nbre,
                        Cahor = Cahor == null ? 0 : Cahor.Nbre,
                        NumItalie = NumItalie == null ? 0 : NumItalie.Nbre,
                        PES = PES == null ? 0 : PES.Nbre,
                        tot = tot == 0 ? 0 : tot
                              });
                    }
            }

            ViewBag.EvolutionParUAP = res.Select(x => x.ToExpando());
           // ViewBag.TotalReclamations = recs.ToString();

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

            //******************************************************
            List<Rapport> rapportsDeAnneeEncours = db.RapportSet.Where(x => x.IsActif == true && x.DateOuverture.Year == year && x.DateCloture != null && x.Alerte == false).ToList();
            //List<Rapport> rapportsDeAnneeEncours = db.RapportSet.Where(x => x.IsActif == true && x.DateOuverture.Year < DateTime.Now.Year && x.DateCloture != null).ToList();

            List<Rapport> rapports = db.RapportSet.Where(x => x.IsActif == true && x.DateCloture != null).ToList();
            List<Rapport> rapportsNonCloturés = db.RapportSet.Where(x => x.IsActif == true && x.DateCloture == null).ToList();

            var resul = new List<object>();
            foreach (var item in rapports)
            {

                int DureeTrait = 0;
                DureeTrait = TimeSpan.Parse((item.DateCloture - item.DateOuverture).ToString()).Days;

                resul.Add(new
                {
                    NumRapport = item.RapportId,
                    UAP = item.UAP,
                    Pilote = item.Pilote.NomComplet,
                    Duree = DureeTrait
                });

            }

            ViewBag.TauxReactivite = resul.Select(x => x.ToExpando());


            var resul1 = new List<object>();
            foreach (var item in rapportsNonCloturés)
            {

                int DureeTraitement = 0;
                DureeTraitement = TimeSpan.Parse((DateTime.Now - item.DateOuverture).ToString()).Days;

                resul1.Add(new
                {
                    NumRapport = item.RapportId,
                    UAP = item.UAP,
                    Pilote = item.Pilote.NomComplet,
                    Duree = DureeTraitement
                });

            }
            ViewBag.RapportCloture = resul1.Select(x => x.ToExpando());

            //*******************************************************
            //***** ReclamationsParZone
            #region Reclamations par type incident
            var series = new Dictionary<string, List<int>>();
            var seriesL = new List<object>();
            var list_UAP = db.Configs.Find("UAP").Val.Split(',');
            List<string> UAPS = list_UAP.ToList();
            var list_TypesIncidents = db.Configs.Find("TypesIncidents").Val.Split(',');
            List<string> ls = list_TypesIncidents.ToList();
            // ls.Add("Client Direct / Client Final");

            var lst = new List<object>();
            foreach (var typeInc in ls)
            {
                var ttypeInc = typeInc.Trim();
                var listCountUAPs = new List<int>();
                foreach (var uap in UAPS)
                {
                    var q = db.RapportSet.Where(x => x.UAP == uap && x.TypeIncident == ttypeInc && x.Alerte == false && x.IsActif == true && x.DateOuverture.Year == DateTime.Now.Year).Count();
                    listCountUAPs.Add(q);
                }
                seriesL.Add(new
                {
                    name = ttypeInc,
                    data = listCountUAPs
                });

            }

            ViewBag.ReclamationsParZone_series = new System.Web.Script.Serialization.JavaScriptSerializer()
                .Serialize(Json(seriesL).Data);
            ViewBag.ReclamationsParZone_xAxis = new System.Web.Script.Serialization.JavaScriptSerializer()
                .Serialize(Json(UAPS).Data);
            #endregion
            #region Alertes par type Incident
            var seriesAlertes = new Dictionary<string, List<int>>();
            var seriesLAlertes = new List<object>();
            //  var list_UAP = db.Configs.Find("UAP").Val.Split(',');
            //  var list_TypesIncidents = db.Configs.Find("TypesIncidents").Val.Split(',');
          //  var Alerte = true;

            var lstAlertes = new List<object>();
            foreach (var typeInc in ls)
            {
                var ttypeInc = typeInc.Trim();
                var listCountUAPs = new List<int>();
                foreach (var uap in UAPS)
                {
                    var q = db.RapportSet.Where(x => x.UAP == uap && x.TypeIncident == ttypeInc && x.IsActif == true && x.Alerte == true && x.DateOuverture.Year == DateTime.Now.Year).Count();
                    listCountUAPs.Add(q);
                }
                seriesLAlertes.Add(new
                {
                    name = ttypeInc,
                    data = listCountUAPs
                });
            }

            ViewBag.ReclamationsParZone_seriesAlertes = new System.Web.Script.Serialization.JavaScriptSerializer()
                .Serialize(Json(seriesLAlertes).Data);


            ViewBag.ReclamationsParZone_xAxisAlertes = new System.Web.Script.Serialization.JavaScriptSerializer()
                .Serialize(Json(UAPS).Data);

            #endregion

            //******************************************************
            List<Rapport> rapportsAlertesDeAnneeEncours = db.RapportSet.Where(x => x.IsActif == true && x.DateOuverture.Year == year && x.DateCloture != null && x.Alerte == true).ToList();
            //List<Rapport> rapportsDeAnneeEncours = db.RapportSet.Where(x => x.IsActif == true && x.DateOuverture.Year == DateTime.Now.Year && x.DateCloture != null).ToList();


            List<Rapport> rapportsAlertes = db.RapportSet.Where(x => x.IsActif == true && x.DateOuverture.Year < DateTime.Now.Year && x.DateCloture != null && x.Alerte == true).ToList();
            var resulAl = new List<object>();
            foreach (var item in rapportsAlertes)
            {

                int DureeTrait = 0;
                DureeTrait = TimeSpan.Parse((item.DateCloture - item.DateOuverture).ToString()).Days;

                resulAl.Add(new
                {
                    NumRapport = item.RapportId,
                    UAP = item.UAP,
                    Pilote = item.Pilote.NomComplet,
                    Duree = DureeTrait
                });

            }
            ViewBag.TauxReactiviteAlertes = resulAl.Select(x => x.ToExpando());

            //*******************************************************


            return View();
        }

        public ActionResult OK(FormCollection form) //
        {
            string sAnnee = form["Annee"];
            return Index(sAnnee);
        }
        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}